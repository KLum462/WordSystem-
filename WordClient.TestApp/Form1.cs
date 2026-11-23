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
        private TextBox _txtUsername = null!;
        private TextBox _txtPassword = null!;
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
            this.Text = "Клиент интеграции Microsoft Word Interop и ASP.NET Core";
            this.Size = new Size(920, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10)
            };
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 165F));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 165F));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // Config Group Box
            var grpConfig = new GroupBox
            {
                Text = "Подключение и настройки",
                Dock = DockStyle.Fill
            };
            var configLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 3,
                Padding = new Padding(6)
            };
            configLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115F));
            configLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            configLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115F));
            configLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            var loadedCfg = ClientConfig.LoadFromConfiguration();
            _txtServerUrl = new TextBox { Text = loadedCfg.ServerUrl, Dock = DockStyle.Fill };
            _txtApiKey = new TextBox { Text = loadedCfg.ApiKey, Dock = DockStyle.Fill };
            _txtUsername = new TextBox { Text = "admin", Dock = DockStyle.Fill };
            _txtPassword = new TextBox { Text = "admin123", UseSystemPasswordChar = true, Dock = DockStyle.Fill };
            _txtDocId = new TextBox { Text = "doc-001", Dock = DockStyle.Fill };

            _cmbOperation = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbOperation.Items.AddRange(new object[]
            {
                "uppercase", "lowercase", "reverse", "trim", "word_count", "title_case", "capitalize",
                "normalize", "anonymize", "extract_entities", "summarize", "stats"
            });
            _cmbOperation.SelectedIndex = 0;

            configLayout.Controls.Add(new Label { Text = "Адрес сервера:", Anchor = AnchorStyles.Left }, 0, 0);
            configLayout.Controls.Add(_txtServerUrl, 1, 0);
            configLayout.Controls.Add(new Label { Text = "API-ключ:", Anchor = AnchorStyles.Left }, 2, 0);
            configLayout.Controls.Add(_txtApiKey, 3, 0);

            configLayout.Controls.Add(new Label { Text = "Логин:", Anchor = AnchorStyles.Left }, 0, 1);
            configLayout.Controls.Add(_txtUsername, 1, 1);
            configLayout.Controls.Add(new Label { Text = "Пароль:", Anchor = AnchorStyles.Left }, 2, 1);
            configLayout.Controls.Add(_txtPassword, 3, 1);

            configLayout.Controls.Add(new Label { Text = "ID документа:", Anchor = AnchorStyles.Left }, 0, 2);
            configLayout.Controls.Add(_txtDocId, 1, 2);
            configLayout.Controls.Add(new Label { Text = "Операция:", Anchor = AnchorStyles.Left }, 2, 2);
            configLayout.Controls.Add(_cmbOperation, 3, 2);

            grpConfig.Controls.Add(configLayout);

            // Ribbon Buttons Group Box
            var grpActions = new GroupBox
            {
                Text = "Действия (Word 2016 Interop и API)",
                Dock = DockStyle.Fill
            };
            var flowActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(6),
                AutoScroll = true
            };

            var btnAuth = CreateButton("1. Авторизация", OnAuthenticateClick, Color.LightSteelBlue);
            var btnChangePass = CreateButton("Сменить пароль", OnChangePasswordClick, Color.Lavender);
            var btnLoadDoc = CreateButton("2. Загрузить документ", OnLoadDocumentClick, Color.LightGreen);
            var btnSendSelected = CreateButton("3. Отправить выделенное", OnSendSelectedTextClick, Color.LightGoldenrodYellow);
            var btnSendDoc = CreateButton("4. Отправить документ", OnSendDocumentClick, Color.Khaki);
            var btnCheckDoc = CreateButton("5. Проверить документ", OnCheckDocumentClick, Color.LightSkyBlue);
            var btnSaveTemplate = CreateButton("6. Сохранить шаблон", OnSaveTemplateClick, Color.LightSalmon);

            var btnOpenWord = CreateButton("Создать документ Word", OnOpenWordClick, Color.WhiteSmoke);
            var btnOpenDocx = CreateButton("Открыть .docx", OnOpenDocxClick, Color.WhiteSmoke);
            var btnSaveDocx = CreateButton("Сохранить .docx", OnSaveDocxClick, Color.WhiteSmoke);
            var btnUploadDocx = CreateButton("Загрузить .docx", OnUploadDocxClick, Color.Moccasin);
            var btnLoadTemplate = CreateButton("Загрузить шаблон", OnLoadTemplateClick, Color.LightCyan);
            var btnFindReplace = CreateButton("Найти и заменить", OnFindReplaceClick, Color.WhiteSmoke);

            _cmbInsertLocation = new ComboBox { Width = 190, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbInsertLocation.Items.AddRange(new object[] { "Выделенный диапазон", "Конец документа", "Закладка / Content Control" });
            _cmbInsertLocation.SelectedIndex = 0;

            flowActions.Controls.Add(btnAuth);
            flowActions.Controls.Add(btnChangePass);
            flowActions.Controls.Add(btnLoadDoc);
            flowActions.Controls.Add(btnSendSelected);
            flowActions.Controls.Add(btnSendDoc);
            flowActions.Controls.Add(btnCheckDoc);
            flowActions.Controls.Add(btnSaveTemplate);
            flowActions.Controls.Add(btnOpenWord);
            flowActions.Controls.Add(btnOpenDocx);
            flowActions.Controls.Add(btnSaveDocx);
            flowActions.Controls.Add(btnUploadDocx);
            flowActions.Controls.Add(btnLoadTemplate);
            flowActions.Controls.Add(btnFindReplace);
            flowActions.Controls.Add(new Label { Text = "Куда вставить:", AutoSize = true, Margin = new Padding(10, 8, 4, 0) });
            flowActions.Controls.Add(_cmbInsertLocation);

            grpActions.Controls.Add(flowActions);

            // Logs Group Box
            var grpLogs = new GroupBox
            {
                Text = "Журнал выполнения и логи",
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
                Width = 165,
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

        private static string FormatLocation(InsertLocation loc) => loc switch
        {
            InsertLocation.SelectedRange => "выделенный диапазон",
            InsertLocation.EndOfDocument => "конец документа",
            InsertLocation.BookmarkOrContentControl => "закладку / Content Control",
            _ => loc.ToString()
        };

        private async void OnAuthenticateClick(object? sender, EventArgs e)
        {
            try
            {
                using var api = new ApiClient(GetCurrentConfig());
                AuthValidateResponse? auth;

                if (!string.IsNullOrWhiteSpace(_txtUsername.Text) && !string.IsNullOrEmpty(_txtPassword.Text))
                {
                    auth = await api.LoginAsync(_txtUsername.Text.Trim(), _txtPassword.Text);
                    if (auth.Valid && !string.IsNullOrWhiteSpace(auth.ApiKey))
                    {
                        _txtApiKey.Text = auth.ApiKey;
                    }
                }
                else
                {
                    auth = await api.ValidateAuthDetailedAsync();
                }

                bool ok = auth?.Valid ?? false;
                MessageBox.Show(
                    ok
                        ? $"Авторизация прошла успешно!\r\nПользователь: {auth!.FullName ?? auth.User} ({auth.Role ?? "User"})"
                        : $"Ошибка авторизации! {auth?.Message ?? "Неверные учётные данные."}",
                    "Авторизация",
                    MessageBoxButtons.OK,
                    ok ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Auth error", ex);
                MessageBox.Show($"Ошибка авторизации: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void OnChangePasswordClick(object? sender, EventArgs e)
        {
            try
            {
                var input = PromptChangePasswordDialog(_txtUsername.Text.Trim());
                if (input == null) return;

                using var api = new ApiClient(GetCurrentConfig());
                var msg = await api.ChangePasswordAsync(input.Value.Username, input.Value.CurrentPassword, input.Value.NewPassword);
                _txtPassword.Text = input.Value.NewPassword;
                MessageBox.Show(msg, "Смена пароля", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Change password error", ex);
                MessageBox.Show($"Ошибка смены пароля: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static (string Username, string CurrentPassword, string NewPassword)? PromptChangePasswordDialog(string defaultUsername)
        {
            using var form = new Form
            {
                Text = "Смена пароля",
                Width = 420,
                Height = 230,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lblUser = new Label { Left = 14, Top = 18, Width = 125, Text = "Логин:" };
            var txtUser = new TextBox { Left = 145, Top = 15, Width = 245, Text = string.IsNullOrWhiteSpace(defaultUsername) ? "admin" : defaultUsername };

            var lblCurr = new Label { Left = 14, Top = 52, Width = 125, Text = "Текущий пароль:" };
            var txtCurr = new TextBox { Left = 145, Top = 49, Width = 245, UseSystemPasswordChar = true };

            var lblNew = new Label { Left = 14, Top = 86, Width = 125, Text = "Новый пароль:" };
            var txtNew = new TextBox { Left = 145, Top = 83, Width = 245, UseSystemPasswordChar = true };

            var lblConf = new Label { Left = 14, Top = 120, Width = 125, Text = "Повторите пароль:" };
            var txtConf = new TextBox { Left = 145, Top = 117, Width = 245, UseSystemPasswordChar = true };

            var btnOk = new Button { Text = "Сохранить", Left = 198, Top = 154, Width = 92, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Отмена", Left = 298, Top = 154, Width = 92, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { lblUser, txtUser, lblCurr, txtCurr, lblNew, txtNew, lblConf, txtConf, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            if (form.ShowDialog() != DialogResult.OK) return null;
            if (txtNew.Text != txtConf.Text)
            {
                MessageBox.Show("Пароли не совпадают.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
            if (string.IsNullOrWhiteSpace(txtUser.Text) || string.IsNullOrEmpty(txtCurr.Text) || string.IsNullOrWhiteSpace(txtNew.Text))
                return null;

            return (txtUser.Text.Trim(), txtCurr.Text, txtNew.Text);
        }

        private async void OnLoadDocumentClick(object? sender, EventArgs e)
        {
            try
            {
                string docId = _txtDocId.Text.Trim();
                using var api = new ApiClient(GetCurrentConfig());
                var doc = await api.GetDocumentAsync(docId);

                var word = GetWordService();
                var createdDoc = word.GenerateDocumentFromDto(doc);
                WordDocumentService.ReleaseComObject(createdDoc);

                MessageBox.Show($"Документ «{doc.Title}» успешно загружен и сформирован в Word!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Load document error", ex);
                MessageBox.Show($"Не удалось загрузить документ: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    MessageBox.Show("Сначала выделите фрагмент текста в активном документе Word.", "Текст не выделен", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string op = _cmbOperation.SelectedItem?.ToString() ?? "uppercase";
                using var api = new ApiClient(GetCurrentConfig());
                string result = await api.ProcessSelectedTextAsync(selectedText, op, _txtDocId.Text.Trim());

                var loc = (InsertLocation)_cmbInsertLocation.SelectedIndex;
                word.InsertResult(result, loc);

                MessageBox.Show($"Текст обработан ({op}) и вставлен в {FormatLocation(loc)}!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Send selected text error", ex);
                MessageBox.Show($"Ошибка обработки текста: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    MessageBox.Show("Активный документ Word пуст.", "Пустой документ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string op = _cmbOperation.SelectedItem?.ToString() ?? "uppercase";
                using var api = new ApiClient(GetCurrentConfig());
                var resp = await api.ProcessDocumentAsync(_txtDocId.Text.Trim(), docText, op);

                var loc = (InsertLocation)_cmbInsertLocation.SelectedIndex;
                word.InsertResult($"\r\n--- РЕЗУЛЬТАТ ОБРАБОТКИ СЕРВЕРОМ ---\r\n{resp.Result}\r\n---------------------------------", loc);

                MessageBox.Show($"Весь документ обработан (слов: {resp.WordsCount}, абзацев: {resp.ParagraphsCount}) и вставлен в {FormatLocation(loc)}!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Send document error", ex);
                MessageBox.Show($"Ошибка отправки документа: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                string metaLines = string.Empty;
                foreach (var kvp in localReport.Metadata)
                {
                    string label = kvp.Key switch
                    {
                        "Title" => "Заголовок",
                        "Author" => "Автор",
                        "CreationDate" => "Дата создания",
                        _ => kvp.Key
                    };
                    metaLines += $"• {label}: {kvp.Value}\r\n";
                }

                string issuesText = serverCheck.Issues != null && serverCheck.Issues.Count > 0
                    ? "\r\n• Замечания:\r\n  - " + string.Join("\r\n  - ", serverCheck.Issues)
                    : "\r\n• Замечаний не обнаружено.";

                string msg = $"Отчёт проверки документа:\r\n" +
                             $"--- Метаданные ---\r\n{metaLines}" +
                             $"--- Статистика ---\r\n" +
                             $"• Слов: {localReport.WordCount}\r\n" +
                             $"• Абзацев: {localReport.ParagraphCount}\r\n" +
                             $"• Пустых абзацев: {localReport.EmptyParagraphCount}\r\n" +
                             $"• Таблиц: {localReport.TableCount}\r\n" +
                             $"• Ошибок орфографии: {localReport.SpellingErrorsCount}\r\n" +
                             $"• Проверка на сервере: {serverCheck.Summary}{issuesText}";

                MessageBox.Show(msg, "Проверка завершена", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Check document error", ex);
                MessageBox.Show($"Ошибка проверки документа: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void OnSaveTemplateClick(object? sender, EventArgs e)
        {
            try
            {
                var word = GetWordService();
                var extracted = word.ExtractDocumentStructure();

                using var api = new ApiClient(GetCurrentConfig());
                List<TemplateCategoryDto> categories;
                try
                {
                    categories = await api.GetCategoriesAsync();
                }
                catch
                {
                    categories = new List<TemplateCategoryDto>();
                }

                var tplInput = PromptSaveTemplateDialog($"Шаблон: {extracted.Title}", categories);
                if (tplInput == null) return;

                var tpl = new TemplateDto
                {
                    Id = $"tpl-{DateTime.UtcNow:yyyyMMdd-HHmmss}",
                    Name = tplInput.Value.Name,
                    Description = tplInput.Value.Description,
                    CategoryId = tplInput.Value.CategoryId,
                    Content = extracted.Content,
                    Settings = extracted.Settings,
                    Metadata = extracted.Metadata
                };

                var saved = await api.SaveTemplateAsync(tpl);

                MessageBox.Show(
                    $"Шаблон «{saved.Name}» (ID: {saved.Id}) успешно сохранён!\r\nАбзацев: {saved.Content.Paragraphs.Count}, таблиц: {saved.Content.Tables.Count}.",
                    "Готово",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Save template error", ex);
                MessageBox.Show($"Ошибка сохранения шаблона: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void OnLoadTemplateClick(object? sender, EventArgs e)
        {
            try
            {
                using var api = new ApiClient(GetCurrentConfig());
                var templates = await api.GetTemplatesAsync();
                if (templates.Count == 0)
                {
                    MessageBox.Show("На сервере пока нет сохранённых шаблонов.", "Загрузка шаблона", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var selected = PromptSelectTemplateDialog(templates);
                if (selected == null) return;

                var word = GetWordService();
                var tplDoc = word.GenerateDocumentFromTemplate(selected);
                WordDocumentService.ReleaseComObject(tplDoc);

                MessageBox.Show(
                    $"Шаблон «{selected.Name}» (абзацев: {selected.Content.Paragraphs.Count}, таблиц: {selected.Content.Tables.Count}) загружен в Microsoft Word!",
                    "Шаблон загружен",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Load template error", ex);
                MessageBox.Show($"Ошибка загрузки шаблона: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnOpenWordClick(object? sender, EventArgs e)
        {
            try
            {
                var word = GetWordService();
                var doc = word.CreateNewDocument();
                WordDocumentService.ReleaseComObject(doc);
                MessageBox.Show("Создан новый пустой документ Microsoft Word!", "Word готов", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Open word error", ex);
                MessageBox.Show($"Ошибка запуска Word: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnOpenDocxClick(object? sender, EventArgs e)
        {
            try
            {
                using var openDlg = new OpenFileDialog
                {
                    Title = "Открыть существующий документ .docx",
                    Filter = "Документы Word (*.docx)|*.docx|Все файлы (*.*)|*.*",
                    CheckFileExists = true
                };
                if (openDlg.ShowDialog() != DialogResult.OK) return;

                var word = GetWordService();
                var opened = word.OpenDocument(openDlg.FileName);
                WordDocumentService.ReleaseComObject(opened);
                MessageBox.Show($"Документ «{System.IO.Path.GetFileName(openDlg.FileName)}» открыт в Microsoft Word!", "Документ открыт", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Open .docx error", ex);
                MessageBox.Show($"Ошибка открытия .docx: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnSaveDocxClick(object? sender, EventArgs e)
        {
            try
            {
                using var saveDlg = new SaveFileDialog
                {
                    Title = "Сохранить активный документ Word как .docx",
                    Filter = "Документ Word (*.docx)|*.docx",
                    DefaultExt = "docx",
                    FileName = $"{_txtDocId.Text.Trim()}.docx"
                };
                if (saveDlg.ShowDialog() != DialogResult.OK) return;

                var word = GetWordService();
                word.SaveActiveDocumentAs(saveDlg.FileName);
                MessageBox.Show($"Документ сохранён:\r\n{saveDlg.FileName}", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Save .docx error", ex);
                MessageBox.Show($"Ошибка сохранения .docx: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void OnUploadDocxClick(object? sender, EventArgs e)
        {
            try
            {
                using var openDlg = new OpenFileDialog
                {
                    Title = "Выберите файл .docx для загрузки на сервер (App_Data/uploads)",
                    Filter = "Документы Word (*.docx)|*.docx",
                    CheckFileExists = true
                };
                if (openDlg.ShowDialog() != DialogResult.OK) return;

                using var api = new ApiClient(GetCurrentConfig());
                var uploaded = await api.UploadDocumentFileAsync(openDlg.FileName);
                MessageBox.Show(
                    $"Файл «{uploaded.FileName}» ({uploaded.SizeBytes} байт) успешно загружен!\r\nПуть на сервере: {uploaded.RelativePath}",
                    "Загрузка завершена",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Upload .docx error", ex);
                MessageBox.Show($"Ошибка загрузки .docx: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnFindReplaceClick(object? sender, EventArgs e)
        {
            try
            {
                var input = PromptFindReplaceDialog();
                if (input == null) return;

                var word = GetWordService();
                int count = word.FindAndReplace(input.Value.FindText, input.Value.ReplaceText);
                MessageBox.Show(
                    count > 0
                        ? $"Текст «{input.Value.FindText}» заменён на «{input.Value.ReplaceText}» в активном документе."
                        : $"Текст «{input.Value.FindText}» не найден в активном документе.",
                    "Поиск и замена",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("Find replace error", ex);
                MessageBox.Show($"Ошибка поиска и замены: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static (string FindText, string ReplaceText)? PromptFindReplaceDialog()
        {
            using var form = new Form
            {
                Text = "Поиск и замена в Word",
                Width = 430,
                Height = 205,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lblFind = new Label { Left = 14, Top = 14, Width = 388, Text = "Найти текст:" };
            var txtFind = new TextBox { Left = 14, Top = 36, Width = 388 };
            var lblRep = new Label { Left = 14, Top = 68, Width = 388, Text = "Заменить на:" };
            var txtRep = new TextBox { Left = 14, Top = 90, Width = 388 };

            var btnOk = new Button { Text = "Заменить всё", Left = 198, Top = 128, Width = 110, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Отмена", Left = 316, Top = 128, Width = 86, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { lblFind, txtFind, lblRep, txtRep, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            if (form.ShowDialog() != DialogResult.OK) return null;
            if (string.IsNullOrEmpty(txtFind.Text)) return null;
            return (txtFind.Text, txtRep.Text);
        }

        private static (string Name, string Description, int? CategoryId)? PromptSaveTemplateDialog(
            string defaultName,
            List<TemplateCategoryDto> categories)
        {
            using var form = new Form
            {
                Text = "Сохранить активный документ как шаблон",
                Width = 450,
                Height = 265,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lblName = new Label { Left = 14, Top = 14, Width = 406, Text = "Название шаблона:" };
            var txtName = new TextBox { Left = 14, Top = 36, Width = 406, Text = defaultName };

            var lblDesc = new Label { Left = 14, Top = 68, Width = 406, Text = "Описание:" };
            var txtDesc = new TextBox { Left = 14, Top = 90, Width = 406, Text = "Извлечено со стилями и таблицами из Microsoft Word" };

            var lblCat = new Label { Left = 14, Top = 122, Width = 406, Text = "Категория:" };
            var cmbCat = new ComboBox { Left = 14, Top = 144, Width = 406, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbCat.Items.Add("(Без категории)");
            foreach (var c in categories)
            {
                cmbCat.Items.Add($"{c.Id}: {c.Name}");
            }
            cmbCat.SelectedIndex = categories.Count > 0 ? 1 : 0;

            var btnOk = new Button { Text = "Сохранить", Left = 226, Top = 184, Width = 96, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Отмена", Left = 330, Top = 184, Width = 90, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { lblName, txtName, lblDesc, txtDesc, lblCat, cmbCat, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            if (form.ShowDialog() != DialogResult.OK) return null;
            if (string.IsNullOrWhiteSpace(txtName.Text)) return null;

            int? categoryId = cmbCat.SelectedIndex > 0 && cmbCat.SelectedIndex - 1 < categories.Count
                ? categories[cmbCat.SelectedIndex - 1].Id
                : null;

            return (txtName.Text.Trim(), txtDesc.Text.Trim(), categoryId);
        }

        private static TemplateDto? PromptSelectTemplateDialog(List<TemplateDto> templates)
        {
            using var form = new Form
            {
                Text = "Загрузка шаблона с сервера",
                Width = 470,
                Height = 185,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lbl = new Label { Left = 14, Top = 14, Width = 426, Text = "Выберите шаблон для генерации в Microsoft Word:" };
            var cmb = new ComboBox { Left = 14, Top = 40, Width = 426, DropDownStyle = ComboBoxStyle.DropDownList };
            foreach (var t in templates)
            {
                cmb.Items.Add($"{t.Id} — {t.Name}");
            }
            cmb.SelectedIndex = 0;

            var btnOk = new Button { Text = "Загрузить", Left = 248, Top = 94, Width = 96, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Отмена", Left = 352, Top = 94, Width = 88, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { lbl, cmb, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            if (form.ShowDialog() != DialogResult.OK || cmb.SelectedIndex < 0) return null;
            return templates[cmb.SelectedIndex];
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            ClientLogger.OnLog -= AppendLog;
            _wordService?.Dispose();
        }
    }
}
