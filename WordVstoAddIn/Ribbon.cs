using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WordClient.Core.Models;
using WordClient.Core.Services;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordVstoAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI? _ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordVstoAddIn.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public async void OnAuthenticateClick(Office.IRibbonControl control)
        {
            try
            {
                using var api = new ApiClient();
                bool isValid = await api.ValidateAuthAsync();
                if (isValid)
                {
                    MessageBox.Show(
                        "Authentication Successful!\r\nAPI Key is valid and server connection is active.",
                        "Authentication",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        "Authentication Failed!\r\nPlease check the API Key configured in App.config.",
                        "Authentication Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Unable to connect to the server: {ex.Message}",
                    "Connection Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void OnLoadDocumentClick(Office.IRibbonControl control)
        {
            try
            {
                using var api = new ApiClient();
                var docDto = await api.GetDocumentAsync("doc-001");

                using var wordService = new WordDocumentService(Globals.ThisAddIn?.Application);
                wordService.GenerateDocumentFromDto(docDto);

                MessageBox.Show(
                    $"Document '{docDto.Title}' (ID: {docDto.Id}) was loaded successfully from the server!",
                    "Document Loaded",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error loading document from server: {ex.Message}",
                    "Load Document Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void OnSendSelectedTextClick(Office.IRibbonControl control)
        {
            try
            {
                using var wordService = new WordDocumentService(Globals.ThisAddIn?.Application);
                string selectedText = wordService.GetSelectedText();

                if (string.IsNullOrWhiteSpace(selectedText))
                {
                    MessageBox.Show(
                        "Please select some text in the Word document first.",
                        "No Text Selected",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                using var api = new ApiClient();
                string result = await api.ProcessSelectedTextAsync(selectedText, "uppercase");

                // Insert result into selected range
                wordService.InsertResult(result, InsertLocation.SelectedRange);

                MessageBox.Show(
                    "Selected text processed and inserted back into the document!",
                    "Text Processed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error processing selected text: {ex.Message}",
                    "Process Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void OnSendDocumentClick(Office.IRibbonControl control)
        {
            try
            {
                using var wordService = new WordDocumentService(Globals.ThisAddIn?.Application);
                string docText = wordService.GetDocumentText();

                if (string.IsNullOrWhiteSpace(docText))
                {
                    MessageBox.Show(
                        "The current document is empty.",
                        "Empty Document",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                using var api = new ApiClient();
                var response = await api.ProcessDocumentAsync("doc-active", docText, "uppercase");

                // Insert response at end of document
                wordService.InsertResult(
                    $"\r\n--- SERVER PROCESSED RESULT ---\r\n{response.Result}\r\n---------------------------------",
                    InsertLocation.EndOfDocument);

                MessageBox.Show(
                    $"Full document processed successfully!\r\nWords: {response.WordsCount}, Paragraphs: {response.ParagraphsCount}.\r\nResult inserted at document end.",
                    "Document Processed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error sending full document: {ex.Message}",
                    "Send Document Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void OnCheckDocumentClick(Office.IRibbonControl control)
        {
            try
            {
                using var wordService = new WordDocumentService(Globals.ThisAddIn?.Application);
                var localReport = wordService.InspectDocument();

                using var api = new ApiClient();
                var serverCheck = await api.CheckDocumentAsync(new DocumentCheckRequest
                {
                    DocumentId = "active-doc",
                    Text = wordService.GetDocumentText(),
                    TableCount = localReport.TableCount
                });

                string message = $"Document Inspection Report:\r\n" +
                                 $"• Words: {localReport.WordCount}\r\n" +
                                 $"• Paragraphs: {localReport.ParagraphCount}\r\n" +
                                 $"• Empty Paragraphs: {localReport.EmptyParagraphCount}\r\n" +
                                 $"• Tables: {localReport.TableCount}\r\n" +
                                 $"• Spelling Errors: {localReport.SpellingErrorsCount}\r\n" +
                                 $"• Server Summary: {serverCheck.Summary}";

                MessageBox.Show(message, "Document Inspection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error checking document: {ex.Message}",
                    "Check Document Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void OnSaveTemplateClick(Office.IRibbonControl control)
        {
            try
            {
                using var wordService = new WordDocumentService(Globals.ThisAddIn?.Application);
                var doc = wordService.GetActiveDocument();

                var templateDto = new TemplateDto
                {
                    Id = $"tpl-{DateTime.UtcNow:yyyyMMdd-HHmmss}",
                    Name = $"Template from {doc.Name}",
                    Description = "Saved via Word VSTO Ribbon Add-In",
                    Content = new DocumentContentDto
                    {
                        Paragraphs = new List<ParagraphDto>
                        {
                            new ParagraphDto { Style = "Heading1", Text = "Saved Template Title" },
                            new ParagraphDto { Style = "Normal", Text = wordService.GetDocumentText() }
                        }
                    },
                    Settings = new DocumentSettingsDto { Font = "Times New Roman", FontSize = 12 }
                };

                using var api = new ApiClient();
                var saved = await api.SaveTemplateAsync(templateDto);

                MessageBox.Show(
                    $"Template '{saved.Name}' (ID: {saved.Id}) saved successfully to the server database!",
                    "Template Saved",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error saving template: {ex.Message}",
                    "Save Template Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using var stream = asm.GetManifestResourceStream(resourceNames[i]);
                    if (stream != null)
                    {
                        using var textReader = new StreamReader(stream);
                        return textReader.ReadToEnd();
                    }
                }
            }

            // Fallback to reading file directly from disk if not embedded
            var localPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ribbon.xml");
            if (File.Exists(localPath))
            {
                return File.ReadAllText(localPath);
            }

            return string.Empty;
        }

        #endregion
    }
}
