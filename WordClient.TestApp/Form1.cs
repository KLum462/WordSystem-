using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordClient.Core.Configuration;
using WordClient.Core.Models;
using WordClient.Core.Services;

namespace WordClient.TestApp
{
    public partial class Form1 : Form
    {
        private TextBox _txtServerUrl = null!;
        private TextBox _txtApiKey = null!;
        private TextBox _txtDocId = null!;
        private ComboBox _cmbOperation = null!;
        private ComboBox _cmbInsertLocation = null!;
        private TextBox _txtLog = null!;
        private WordDocumentService? _wordService;

        public Form1()
        {
            InitializeComponent();
            ClientLogger.OnLog += AppendLog;
        }

        private void InitializeComponent()
        {
            this.Text = "Microsoft Word Interop & ASP.NET Core Integration Test Client";
            this.Size = new Size(880, 680);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10)
            };
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 130F));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 140F));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // Config Group Box
            var grpConfig = new GroupBox
            {
                Text = "Connection & Settings",
                Dock = DockStyle.Fill
            };
            var configLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 2,
                Padding = new Padding(6)
            };
            configLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F));
            configLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            configLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F));
            configLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            _txtServerUrl = new TextBox { Text = "http://localhost:5000", Dock = DockStyle.Fill };
            _txtApiKey = new TextBox { Text = "secret-internship-api-key-2026", Dock = DockStyle.Fill };
            _txtDocId = new TextBox { Text = "doc-001", Dock = DockStyle.Fill };

            _cmbOperation = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbOperation.Items.AddRange(new object[] { "uppercase", "lowercase", "reverse", "trim", "word_count", "title_case" });
            _cmbOperation.SelectedIndex = 0;

            configLayout.Controls.Add(new Label { Text = "Server URL:", Anchor = AnchorStyles.Left }, 0, 0);
            configLayout.Controls.Add(_txtServerUrl, 1, 0);
            configLayout.Controls.Add(new Label { Text = "API Key:", Anchor = AnchorStyles.Left }, 2, 0);
            configLayout.Controls.Add(_txtApiKey, 3, 0);

            configLayout.Controls.Add(new Label { Text = "Document ID:", Anchor = AnchorStyles.Left }, 0, 1);
            configLayout.Controls.Add(_txtDocId, 1, 1);
            configLayout.Controls.Add(new Label { Text = "Operation:", Anchor = AnchorStyles.Left }, 2, 1);
            configLayout.Controls.Add(_cmbOperation, 3, 1);

            grpConfig.Controls.Add(configLayout);

            // Ribbon Buttons Group Box
            var grpActions = new GroupBox
            {
                Text = "Required Ribbon Actions (Word 2016 Interop & API)",
                Dock = DockStyle.Fill
            };
            var flowActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(6),
                AutoScroll = true
            };

            var btnAuth = CreateButton("1. Authenticate", OnAuthenticateClick, Color.LightSteelBlue);
            var btnLoadDoc = CreateButton("2. Load Document", OnLoadDocumentClick, Color.LightGreen);
            var btnSendSelected = CreateButton("3. Send Selected Text", OnSendSelectedTextClick, Color.LightGoldenrodYellow);
            var btnSendDoc = CreateButton("4. Send Document", OnSendDocumentClick, Color.Khaki);
            var btnCheckDoc = CreateButton("5. Check Document", OnCheckDocumentClick, Color.LightSkyBlue);
            var btnSaveTemplate = CreateButton("6. Save Template", OnSaveTemplateClick, Color.LightSalmon);

            var btnOpenWord = CreateButton("Open/Create Word", OnOpenWordClick, Color.WhiteSmoke);
            var btnFindReplace = CreateButton("Find & Replace", OnFindReplaceClick, Color.WhiteSmoke);

            _cmbInsertLocation = new ComboBox { Width = 170, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbInsertLocation.Items.AddRange(new object[] { "Selected Range", "End of Document", "Bookmark/Content Control" });
            _cmbInsertLocation.SelectedIndex = 0;

            flowActions.Controls.Add(btnAuth);
            flowActions.Controls.Add(btnLoadDoc);
            flowActions.Controls.Add(btnSendSelected);
            flowActions.Controls.Add(btnSendDoc);
            flowActions.Controls.Add(btnCheckDoc);
            flowActions.Controls.Add(btnSaveTemplate);
            flowActions.Controls.Add(btnOpenWord);
            flowActions.Controls.Add(btnFindReplace);
            flowActions.Controls.Add(new Label { Text = "Insert Target:", AutoSize = true, Margin = new Padding(10, 8, 4, 0) });
            flowActions.Controls.Add(_cmbInsertLocation);

            grpActions.Controls.Add(flowActions);

            // Logs Group Box
            var grpLogs = new GroupBox
            {
                Text = "Execution Logs & Output",
                Dock = DockStyle.Fill
            };
            _txtLog = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point),
                BackColor = Color.FromArgb(245, 245, 245)
            };
            grpLogs.Controls.Add(_txtLog);

            mainPanel.Controls.Add(grpConfig, 0, 0);
            mainPanel.Controls.Add(grpActions, 0, 1);
            mainPanel.Controls.Add(grpLogs, 0, 2);

            this.Controls.Add(mainPanel);
        }

        private static Button CreateButton(string text, EventHandler onClick, Color backColor)
        {
            var btn = new Button
            {
                Text = text,
                Width = 150,
                Height = 36,
                Margin = new Padding(4),
                BackColor = backColor,
                FlatStyle = FlatStyle.System
            };
            btn.Click += onClick;
            return btn;
        }

        private ClientConfig GetCurrentConfig()
        {
            return new ClientConfig
            {
                ServerUrl = _txtServerUrl.Text.Trim(),
                ApiKey = _txtApiKey.Text.Trim(),
                TimeoutSeconds = 30,
                MaxRetries = 3
            };
        }

        private WordDocumentService GetWordService()
        {
            if (_wordService == null)
            {
                _wordService = new WordDocumentService();
            }
            return _wordService;
        }

        private void AppendLog(string message)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;

            this.BeginInvoke((Action)(() =>
            {
                _txtLog.AppendText(message + Environment.NewLine);
                _txtLog.SelectionStart = _txtLog.Text.Length;
                _txtLog.ScrollToCaret();
            }));
        }

        private async void OnAuthenticateClick(object? sender, EventArgs e)
        {
            try
            {
                using var api = new ApiClient(GetCurrentConfig());
                bool ok = await api.ValidateAuthAsync();
                MessageBox.Show(
                    ok ? "Authentication Succeeded! Server accepted API key." : "Authentication Failed! Invalid API key.",
                    "Authentication",
                    MessageBoxButtons.OK,
                    ok ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Auth error", ex);
                MessageBox.Show($"Authentication error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void OnLoadDocumentClick(object? sender, EventArgs e)
        {
            try
            {
                string docId = _txtDocId.Text.Trim();
                using var api = new ApiClient(GetCurrentConfig());
                var doc = await api.GetDocumentAsync(docId);

                var word = GetWordService();
                word.GenerateDocumentFromDto(doc);

                MessageBox.Show($"Document '{doc.Title}' loaded and generated in Word successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Load document error", ex);
                MessageBox.Show($"Failed to load document: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void OnSendSelectedTextClick(object? sender, EventArgs e)
        {
            try
            {
                var word = GetWordService();
                string selectedText = word.GetSelectedText();

                if (string.IsNullOrWhiteSpace(selectedText))
                {
                    MessageBox.Show("Please select some text in the active Word document first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string op = _cmbOperation.SelectedItem?.ToString() ?? "uppercase";
                using var api = new ApiClient(GetCurrentConfig());
                string result = await api.ProcessSelectedTextAsync(selectedText, op);

                var loc = (InsertLocation)_cmbInsertLocation.SelectedIndex;
                word.InsertResult(result, loc);

                MessageBox.Show($"Processed text ({op}) and inserted back into document at {loc}!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Send selected text error", ex);
                MessageBox.Show($"Error processing text: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void OnSendDocumentClick(object? sender, EventArgs e)
        {
            try
            {
                var word = GetWordService();
                string docText = word.GetDocumentText();

                if (string.IsNullOrWhiteSpace(docText))
                {
                    MessageBox.Show("The active Word document is empty.", "Empty Document", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string op = _cmbOperation.SelectedItem?.ToString() ?? "uppercase";
                using var api = new ApiClient(GetCurrentConfig());
                var resp = await api.ProcessDocumentAsync(_txtDocId.Text.Trim(), docText, op);

                var loc = (InsertLocation)_cmbInsertLocation.SelectedIndex;
                word.InsertResult($"\r\n--- SERVER PROCESSED RESULT ---\r\n{resp.Result}\r\n---------------------------------", loc);

                MessageBox.Show($"Entire document processed ({resp.WordsCount} words, {resp.ParagraphsCount} paragraphs) and inserted at {loc}!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Send document error", ex);
                MessageBox.Show($"Error sending document: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void OnCheckDocumentClick(object? sender, EventArgs e)
        {
            try
            {
                var word = GetWordService();
                var localReport = word.InspectDocument();

                using var api = new ApiClient(GetCurrentConfig());
                var serverCheck = await api.CheckDocumentAsync(new DocumentCheckRequest
                {
                    DocumentId = _txtDocId.Text.Trim(),
                    Text = word.GetDocumentText(),
                    TableCount = localReport.TableCount
                });

                string msg = $"Document Inspection Report:\r\n" +
                             $"• Word Count: {localReport.WordCount}\r\n" +
                             $"• Paragraph Count: {localReport.ParagraphCount}\r\n" +
                             $"• Empty Paragraphs: {localReport.EmptyParagraphCount}\r\n" +
                             $"• Tables Count: {localReport.TableCount}\r\n" +
                             $"• Spelling Errors: {localReport.SpellingErrorsCount}\r\n" +
                             $"• Server Check: {serverCheck.Summary}";

                MessageBox.Show(msg, "Inspection Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Check document error", ex);
                MessageBox.Show($"Error checking document: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void OnSaveTemplateClick(object? sender, EventArgs e)
        {
            try
            {
                var word = GetWordService();
                var doc = word.GetActiveDocument();

                var tpl = new TemplateDto
                {
                    Id = $"tpl-{DateTime.UtcNow:yyyyMMdd-HHmmss}",
                    Name = $"Template from {doc.Name}",
                    Description = "Captured via Desktop Client",
                    Content = new DocumentContentDto
                    {
                        Paragraphs = new List<ParagraphDto>
                        {
                            new ParagraphDto { Style = "Heading1", Text = "Template Heading" },
                            new ParagraphDto { Style = "Normal", Text = word.GetDocumentText() }
                        }
                    },
                    Settings = new DocumentSettingsDto { Font = "Times New Roman", FontSize = 12 }
                };

                using var api = new ApiClient(GetCurrentConfig());
                var saved = await api.SaveTemplateAsync(tpl);

                MessageBox.Show($"Template '{saved.Name}' (ID: {saved.Id}) saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Save template error", ex);
                MessageBox.Show($"Error saving template: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnOpenWordClick(object? sender, EventArgs e)
        {
            try
            {
                var word = GetWordService();
                word.CreateNewDocument();
                MessageBox.Show("New Microsoft Word document created and ready!", "Word Ready", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Open word error", ex);
                MessageBox.Show($"Error opening Word: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnFindReplaceClick(object? sender, EventArgs e)
        {
            try
            {
                var word = GetWordService();
                int count = word.FindAndReplace("First paragraph", "PROCESSED PARAGRAPH");
                MessageBox.Show($"Find and replace executed (replacements made: {count}).", "Find & Replace", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Find replace error", ex);
                MessageBox.Show($"Find & Replace error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            ClientLogger.OnLog -= AppendLog;
            _wordService?.Dispose();
        }
    }
}
