using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
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

        private static string FormatInsertLocation(InsertLocation loc) => loc switch
        {
            InsertLocation.SelectedRange => "выделенный диапазон",
            InsertLocation.EndOfDocument => "конец документа",
            InsertLocation.BookmarkOrContentControl => "закладку / Content Control",
            _ => loc.ToString()
        };

        public async void OnAuthenticateClick(Office.IRibbonControl control)
        {
            try
            {
                using var api = new ApiClient();
                var auth = await api.ValidateAuthDetailedAsync();
                if (auth != null && auth.Valid)
                {
                    MessageBox.Show(
                        $"Авторизация прошла успешно!\r\n" +
                        $"Сервер: {api.Config.ServerUrl}\r\n" +
                        $"Пользователь: {auth.FullName ?? auth.User} ({auth.Role ?? "User"})\r\n" +
                        $"Лог клиента: {ClientLogger.LogFilePath}",
                        "Авторизация",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        $"Ошибка авторизации!\r\nСервер: {api.Config.ServerUrl}\r\n{auth?.Message ?? "Проверьте API-ключ в файле App.config."}",
                        "Ошибка авторизации",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Ribbon Authenticate error", ex);
                MessageBox.Show(
                    $"Не удалось подключиться к серверу: {ex.Message}",
                    "Ошибка подключения",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void OnLoadDocumentClick(Office.IRibbonControl control)
        {
            try
            {
                using var api = new ApiClient();
                List<DocumentSummaryDto> docs;
                List<TemplateDto> templates;

                try
                {
                    docs = await api.GetDocumentsAsync();
                }
                catch
                {
                    docs = new List<DocumentSummaryDto>();
                }

                try
                {
                    templates = await api.GetTemplatesAsync();
                }
                catch
                {
                    templates = new List<TemplateDto>();
                }

                var selection = RibbonDialogs.PromptLoadDocument(docs, templates, "doc-001");
                if (selection == null) return;

                using var wordService = new WordDocumentService(Globals.ThisAddIn?.Application);

                if (selection.Mode == LoadSourceMode.LocalDocx)
                {
                    using var openDlg = new OpenFileDialog
                    {
                        Title = "Открыть существующий документ Word (.docx)",
                        Filter = "Документы Word (*.docx)|*.docx|Все файлы (*.*)|*.*",
                        CheckFileExists = true
                    };

                    if (openDlg.ShowDialog() != DialogResult.OK) return;

                    Word.Document? openedDoc = null;
                    try
                    {
                        openedDoc = wordService.OpenDocument(openDlg.FileName);
                    }
                    finally
                    {
                        WordDocumentService.ReleaseComObject(openedDoc);
                    }

                    MessageBox.Show(
                        $"Документ «{Path.GetFileName(openDlg.FileName)}» успешно открыт!",
                        "Документ открыт",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                string? savePath = null;
                if (selection.SaveAsDocx)
                {
                    using var saveDlg = new SaveFileDialog
                    {
                        Title = "Сохранить сформированный документ Word",
                        Filter = "Документ Word (*.docx)|*.docx",
                        DefaultExt = "docx",
                        FileName = selection.Mode == LoadSourceMode.ServerTemplate
                            ? $"{selection.Template?.Id ?? "template"}.docx"
                            : $"{selection.DocumentId ?? "doc-001"}.docx"
                    };

                    if (saveDlg.ShowDialog() != DialogResult.OK) return;
                    savePath = saveDlg.FileName;
                }

                if (selection.Mode == LoadSourceMode.ServerTemplate && selection.Template != null)
                {
                    Word.Document? tplDoc = null;
                    try
                    {
                        tplDoc = wordService.GenerateDocumentFromTemplate(selection.Template, savePath);
                    }
                    finally
                    {
                        WordDocumentService.ReleaseComObject(tplDoc);
                    }

                    MessageBox.Show(
                        $"Шаблон «{selection.Template.Name}» успешно сформирован в Word!" +
                        (savePath != null ? $"\r\nСохранён в: {savePath}" : string.Empty),
                        "Шаблон загружен",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                var docDto = await api.GetDocumentAsync(selection.DocumentId!);
                Word.Document? createdDoc = null;
                try
                {
                    createdDoc = wordService.GenerateDocumentFromDto(docDto, savePath);
                }
                finally
                {
                    WordDocumentService.ReleaseComObject(createdDoc);
                }

                MessageBox.Show(
                    $"Документ «{docDto.Title}» (ID: {docDto.Id}, v{docDto.Metadata?.Version ?? 1}) успешно загружен с сервера!" +
                    (savePath != null ? $"\r\nСохранён в: {savePath}" : string.Empty),
                    "Документ загружен",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Ribbon LoadDocument error", ex);
                MessageBox.Show(
                    $"Ошибка загрузки документа: {ex.Message}",
                    "Ошибка загрузки",
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
                        "Сначала выделите фрагмент текста в документе Word.",
                        "Текст не выделен",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                using var api = new ApiClient();
                var operations = await api.GetSupportedOperationsAsync();
                var options = RibbonDialogs.PromptProcessingOptions(operations, InsertLocation.SelectedRange);
                if (options == null) return;

                string result = await api.ProcessSelectedTextAsync(selectedText, options.Value.Operation);
                wordService.InsertResult(result, options.Value.Location);

                MessageBox.Show(
                    $"Выделенный текст обработан (операция: {options.Value.Operation}) и вставлен в {FormatInsertLocation(options.Value.Location)}!",
                    "Текст обработан",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Ribbon SendSelectedText error", ex);
                MessageBox.Show(
                    $"Ошибка обработки выделенного текста: {ex.Message}",
                    "Ошибка обработки",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void OnSendDocumentClick(Office.IRibbonControl control)
        {
            try
            {
                using var wordService = new WordDocumentService(Globals.ThisAddIn?.Application);
                using var api = new ApiClient();

                var fileInfo = wordService.GetActiveDocumentFileInfo();
                if (fileInfo.FileSizeBytes.HasValue && fileInfo.FileSizeBytes.Value > api.Config.MaxFileSizeBytes)
                {
                    MessageBox.Show(
                        $"Текущий документ «{fileInfo.Name}» ({fileInfo.FileSizeBytes.Value / 1024.0 / 1024.0:F2} МБ) превышает ограничение 100 МБ.",
                        "Превышен размер файла",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                string docText = wordService.GetDocumentText();
                if (string.IsNullOrWhiteSpace(docText))
                {
                    MessageBox.Show(
                        "Текущий документ пуст.",
                        "Пустой документ",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                var operations = await api.GetSupportedOperationsAsync();
                var options = RibbonDialogs.PromptSendDocumentOptions(operations, InsertLocation.EndOfDocument);
                if (options == null) return;

                var response = await api.ProcessDocumentAsync(fileInfo.Name, docText, options.Operation);
                wordService.InsertResult(
                    $"\r\n--- РЕЗУЛЬТАТ ОБРАБОТКИ СЕРВЕРОМ ({options.Operation}) ---\r\n{response.Result}\r\n---------------------------------",
                    options.Location);

                string uploadNote = string.Empty;
                if (options.UploadDocxFile)
                {
                    string? uploadPath = fileInfo.FullPath;
                    string? tempPath = null;
                    try
                    {
                        if (string.IsNullOrWhiteSpace(uploadPath) || !uploadPath!.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                        {
                            tempPath = Path.Combine(Path.GetTempPath(), $"word_upload_{DateTime.UtcNow:yyyyMMddHHmmss}.docx");
                            wordService.SaveActiveDocumentAs(tempPath);
                            uploadPath = tempPath;
                        }

                        var uploaded = await api.UploadDocumentFileAsync(uploadPath);
                        uploadNote = $"\r\nФайл .docx загружен на сервер: {uploaded.RelativePath} ({uploaded.SizeBytes} байт).";
                    }
                    finally
                    {
                        if (tempPath != null && File.Exists(tempPath))
                        {
                            try { File.Delete(tempPath); } catch { }
                        }
                    }
                }

                MessageBox.Show(
                    $"Весь документ обработан (операция: {options.Operation})!\r\nСлов: {response.WordsCount}, абзацев: {response.ParagraphsCount}.\r\nРезультат вставлен в {FormatInsertLocation(options.Location)}.{uploadNote}",
                    "Документ обработан",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Ribbon SendDocument error", ex);
                MessageBox.Show(
                    $"Ошибка отправки документа: {ex.Message}",
                    "Ошибка отправки",
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

                var sb = new StringBuilder();
                sb.AppendLine("=== Метаданные документа ===");
                foreach (var kvp in localReport.Metadata)
                {
                    string label = kvp.Key switch
                    {
                        "Title" => "Заголовок",
                        "Author" => "Автор",
                        "CreationDate" => "Дата создания",
                        _ => kvp.Key
                    };
                    sb.AppendLine($"• {label}: {kvp.Value}");
                }
                sb.AppendLine();
                sb.AppendLine("=== Статистика Word Interop ===");
                sb.AppendLine($"• Слов: {localReport.WordCount}");
                sb.AppendLine($"• Символов: {localReport.CharacterCount}");
                sb.AppendLine($"• Абзацев: {localReport.ParagraphCount}");
                sb.AppendLine($"• Пустых абзацев: {localReport.EmptyParagraphCount}" +
                              (localReport.EmptyParagraphIndices.Count > 0
                                  ? $" ({string.Join(", ", localReport.EmptyParagraphIndices)})"
                                  : string.Empty));
                sb.AppendLine($"• Таблиц: {localReport.TableCount}");
                sb.AppendLine($"• Ошибок орфографии: {localReport.SpellingErrorsCount}");
                sb.AppendLine();
                sb.AppendLine("=== Проверка на сервере ===");
                sb.AppendLine(serverCheck.Summary);
                if (serverCheck.Issues != null && serverCheck.Issues.Count > 0)
                {
                    sb.AppendLine("Найденные замечания:");
                    foreach (var issue in serverCheck.Issues)
                    {
                        sb.AppendLine($"  - {issue}");
                    }
                }
                else
                {
                    sb.AppendLine("Замечаний не обнаружено.");
                }

                RibbonDialogs.ShowInspectionReportDialog(sb.ToString(), () =>
                {
                    var findReplace = RibbonDialogs.PromptFindReplace();
                    if (findReplace == null) return;

                    using var frWord = new WordDocumentService(Globals.ThisAddIn?.Application);
                    int count = frWord.FindAndReplace(findReplace.Value.FindText, findReplace.Value.ReplaceText);
                    MessageBox.Show(
                        count > 0
                            ? $"Текст «{findReplace.Value.FindText}» заменён на «{findReplace.Value.ReplaceText}» в активном документе."
                            : $"Текст «{findReplace.Value.FindText}» не найден в активном документе.",
                        "Поиск и замена",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                });
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Ribbon CheckDocument error", ex);
                MessageBox.Show(
                    $"Ошибка проверки документа: {ex.Message}",
                    "Ошибка проверки",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void OnFindReplaceClick(Office.IRibbonControl control)
        {
            try
            {
                var findReplace = RibbonDialogs.PromptFindReplace();
                if (findReplace == null) return;

                using var wordService = new WordDocumentService(Globals.ThisAddIn?.Application);
                int count = wordService.FindAndReplace(findReplace.Value.FindText, findReplace.Value.ReplaceText);
                MessageBox.Show(
                    count > 0
                        ? $"Текст «{findReplace.Value.FindText}» заменён на «{findReplace.Value.ReplaceText}» в активном документе."
                        : $"Текст «{findReplace.Value.FindText}» не найден в активном документе.",
                    "Поиск и замена",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Ribbon FindReplace error", ex);
                MessageBox.Show(
                    $"Ошибка поиска и замены: {ex.Message}",
                    "Ошибка поиска и замены",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void OnSaveTemplateClick(Office.IRibbonControl control)
        {
            try
            {
                using var wordService = new WordDocumentService(Globals.ThisAddIn?.Application);
                var extractedDoc = wordService.ExtractDocumentStructure();

                using var api = new ApiClient();
                List<TemplateCategoryDto> categories;
                try
                {
                    categories = await api.GetCategoriesAsync();
                }
                catch
                {
                    categories = new List<TemplateCategoryDto>();
                }

                var dialogResult = RibbonDialogs.PromptSaveTemplate($"Шаблон: {extractedDoc.Title}", categories);
                if (dialogResult == null) return;

                var templateDto = new TemplateDto
                {
                    Id = $"tpl-{DateTime.UtcNow:yyyyMMdd-HHmmss}",
                    Name = dialogResult.Value.Name,
                    Description = dialogResult.Value.Description,
                    CategoryId = dialogResult.Value.CategoryId,
                    Content = extractedDoc.Content,
                    Settings = extractedDoc.Settings,
                    Metadata = extractedDoc.Metadata
                };

                var saved = await api.SaveTemplateAsync(templateDto);

                MessageBox.Show(
                    $"Шаблон «{saved.Name}» (ID: {saved.Id}) успешно сохранён!\r\nАбзацев: {saved.Content.Paragraphs.Count}, таблиц: {saved.Content.Tables.Count}.",
                    "Шаблон сохранён",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Ribbon SaveTemplate error", ex);
                MessageBox.Show(
                    $"Ошибка сохранения шаблона: {ex.Message}",
                    "Ошибка сохранения",
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
