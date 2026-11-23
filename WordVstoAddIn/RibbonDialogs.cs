using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using WordClient.Core.Models;
using WordClient.Core.Services;

namespace WordVstoAddIn
{
    internal enum LoadSourceMode
    {
        ServerDocument,
        ServerTemplate,
        LocalDocx
    }

    internal sealed class LoadDocumentSelection
    {
        public LoadSourceMode Mode { get; set; }
        public string? DocumentId { get; set; }
        public TemplateDto? Template { get; set; }
        public bool SaveAsDocx { get; set; }
    }

    internal sealed class SendDocumentOptions
    {
        public string Operation { get; set; } = "uppercase";
        public InsertLocation Location { get; set; } = InsertLocation.EndOfDocument;
        public bool UploadDocxFile { get; set; }
    }

    internal static class RibbonDialogs
    {
        public static LoadDocumentSelection? PromptLoadDocument(
            List<DocumentSummaryDto> documents,
            List<TemplateDto> templates,
            string defaultId = "doc-001")
        {
            using var form = new Form
            {
                Text = "Загрузка документа или шаблона",
                Width = 500,
                Height = 295,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var rbServerDoc = new RadioButton { Left = 16, Top = 14, Width = 450, Text = "Сформировать .docx по JSON документа с сервера", Checked = true };
            var cmbDocs = new ComboBox { Left = 34, Top = 38, Width = 432, DropDownStyle = ComboBoxStyle.DropDown };
            foreach (var doc in documents)
            {
                cmbDocs.Items.Add($"{doc.Id} — {doc.Title} (v{doc.Version})");
            }
            if (cmbDocs.Items.Count > 0) cmbDocs.SelectedIndex = 0;
            else cmbDocs.Text = defaultId;

            var rbTemplate = new RadioButton { Left = 16, Top = 74, Width = 450, Text = "Сформировать .docx из сохранённого шаблона на сервере" };
            var cmbTemplates = new ComboBox { Left = 34, Top = 98, Width = 432, DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false };
            foreach (var tpl in templates)
            {
                cmbTemplates.Items.Add($"{tpl.Id} — {tpl.Name}");
            }
            if (cmbTemplates.Items.Count > 0) cmbTemplates.SelectedIndex = 0;
            else rbTemplate.Enabled = false;

            var rbLocal = new RadioButton { Left = 16, Top = 136, Width = 450, Text = "Открыть существующий локальный файл .docx (проверка блокировки и 100 МБ)" };

            var chkSaveAs = new CheckBox
            {
                Left = 16,
                Top = 172,
                Width = 450,
                Text = "Сохранить сформированный документ в файл .docx (SaveAs2)..."
            };

            void SyncState()
            {
                cmbDocs.Enabled = rbServerDoc.Checked;
                cmbTemplates.Enabled = rbTemplate.Checked;
                chkSaveAs.Enabled = !rbLocal.Checked;
            }

            rbServerDoc.CheckedChanged += (_, _) => SyncState();
            rbTemplate.CheckedChanged += (_, _) => SyncState();
            rbLocal.CheckedChanged += (_, _) => SyncState();

            var btnOk = new Button { Text = "ОК", Left = 290, Top = 212, Width = 84, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Отмена", Left = 382, Top = 212, Width = 84, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { rbServerDoc, cmbDocs, rbTemplate, cmbTemplates, rbLocal, chkSaveAs, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            if (form.ShowDialog() != DialogResult.OK) return null;

            if (rbLocal.Checked)
            {
                return new LoadDocumentSelection { Mode = LoadSourceMode.LocalDocx };
            }

            if (rbTemplate.Checked && cmbTemplates.SelectedIndex >= 0 && cmbTemplates.SelectedIndex < templates.Count)
            {
                return new LoadDocumentSelection
                {
                    Mode = LoadSourceMode.ServerTemplate,
                    Template = templates[cmbTemplates.SelectedIndex],
                    SaveAsDocx = chkSaveAs.Checked
                };
            }

            var selected = cmbDocs.Text.Trim();
            if (string.IsNullOrWhiteSpace(selected)) return null;
            int sep = selected.IndexOf(" — ", StringComparison.Ordinal);
            var docId = sep > 0 ? selected.Substring(0, sep).Trim() : selected;

            return new LoadDocumentSelection
            {
                Mode = LoadSourceMode.ServerDocument,
                DocumentId = docId,
                SaveAsDocx = chkSaveAs.Checked
            };
        }

        public static (string Operation, InsertLocation Location)? PromptProcessingOptions(
            List<string> operations,
            InsertLocation defaultLocation)
        {
            using var form = new Form
            {
                Text = "Параметры обработки текста на сервере",
                Width = 440,
                Height = 225,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lblOp = new Label { Left = 16, Top = 16, Width = 390, Text = "Операция:" };
            var cmbOp = new ComboBox { Left = 16, Top = 38, Width = 390, DropDownStyle = ComboBoxStyle.DropDownList };
            foreach (var op in operations)
            {
                cmbOp.Items.Add(op);
            }
            if (cmbOp.Items.Count > 0) cmbOp.SelectedIndex = 0;

            var lblLoc = new Label { Left = 16, Top = 76, Width = 390, Text = "Куда вставить результат:" };
            var cmbLoc = new ComboBox { Left = 16, Top = 98, Width = 390, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbLoc.Items.AddRange(new object[] { "Выделенный диапазон", "Конец документа", "Закладка / Content Control" });
            cmbLoc.SelectedIndex = (int)defaultLocation;

            var btnOk = new Button { Text = "Обработать", Left = 216, Top = 140, Width = 96, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Отмена", Left = 320, Top = 140, Width = 86, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { lblOp, cmbOp, lblLoc, cmbLoc, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            if (form.ShowDialog() != DialogResult.OK) return null;

            string opResult = cmbOp.SelectedItem?.ToString() ?? "uppercase";
            var locResult = (InsertLocation)Math.Max(0, cmbLoc.SelectedIndex);
            return (opResult, locResult);
        }

        public static SendDocumentOptions? PromptSendDocumentOptions(
            List<string> operations,
            InsertLocation defaultLocation)
        {
            using var form = new Form
            {
                Text = "Отправка всего документа на сервер",
                Width = 460,
                Height = 265,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lblOp = new Label { Left = 16, Top = 16, Width = 410, Text = "Операция обработки текста:" };
            var cmbOp = new ComboBox { Left = 16, Top = 38, Width = 410, DropDownStyle = ComboBoxStyle.DropDownList };
            foreach (var op in operations)
            {
                cmbOp.Items.Add(op);
            }
            if (cmbOp.Items.Count > 0) cmbOp.SelectedIndex = 0;

            var lblLoc = new Label { Left = 16, Top = 76, Width = 410, Text = "Куда вставить ответ сервера:" };
            var cmbLoc = new ComboBox { Left = 16, Top = 98, Width = 410, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbLoc.Items.AddRange(new object[] { "Выделенный диапазон", "Конец документа", "Закладка / Content Control" });
            cmbLoc.SelectedIndex = (int)defaultLocation;

            var chkUpload = new CheckBox
            {
                Left = 16,
                Top = 136,
                Width = 410,
                Text = "Также загрузить файл .docx на сервер (App_Data/uploads, до 100 МБ)"
            };

            var btnOk = new Button { Text = "Отправить", Left = 236, Top = 180, Width = 96, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Отмена", Left = 340, Top = 180, Width = 86, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { lblOp, cmbOp, lblLoc, cmbLoc, chkUpload, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            if (form.ShowDialog() != DialogResult.OK) return null;

            return new SendDocumentOptions
            {
                Operation = cmbOp.SelectedItem?.ToString() ?? "uppercase",
                Location = (InsertLocation)Math.Max(0, cmbLoc.SelectedIndex),
                UploadDocxFile = chkUpload.Checked
            };
        }

        public static void ShowInspectionReportDialog(string reportText, Action onFindReplaceRequested)
        {
            using var form = new Form
            {
                Text = "Отчёт проверки документа",
                Width = 580,
                Height = 430,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var txtReport = new TextBox
            {
                Left = 14,
                Top = 14,
                Width = 536,
                Height = 318,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Segoe UI", 9.5F),
                Text = reportText
            };

            var btnFindReplace = new Button
            {
                Text = "Найти и заменить...",
                Left = 14,
                Top = 344,
                Width = 155
            };
            btnFindReplace.Click += (_, _) => onFindReplaceRequested();

            var btnClose = new Button
            {
                Text = "Закрыть",
                Left = 456,
                Top = 344,
                Width = 94,
                DialogResult = DialogResult.OK
            };

            form.Controls.AddRange(new Control[] { txtReport, btnFindReplace, btnClose });
            form.AcceptButton = btnClose;
            form.ShowDialog();
        }

        public static (string FindText, string ReplaceText)? PromptFindReplace()
        {
            using var form = new Form
            {
                Text = "Поиск и замена в активном документе",
                Width = 430,
                Height = 205,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
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

        public static (string Name, string Description, int? CategoryId)? PromptSaveTemplate(
            string defaultName,
            List<TemplateCategoryDto> categories)
        {
            using var form = new Form
            {
                Text = "Сохранить документ как шаблон",
                Width = 450,
                Height = 275,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lblName = new Label { Left = 16, Top = 14, Width = 400, Text = "Название шаблона:" };
            var txtName = new TextBox { Left = 16, Top = 36, Width = 400, Text = defaultName };

            var lblDesc = new Label { Left = 16, Top = 70, Width = 400, Text = "Описание:" };
            var txtDesc = new TextBox { Left = 16, Top = 92, Width = 400, Text = "Извлечено со стилями и таблицами из Microsoft Word" };

            var lblCat = new Label { Left = 16, Top = 126, Width = 400, Text = "Категория:" };
            var cmbCat = new ComboBox { Left = 16, Top = 148, Width = 400, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbCat.Items.Add("(Без категории)");
            foreach (var cat in categories)
            {
                cmbCat.Items.Add($"{cat.Id}: {cat.Name}");
            }
            cmbCat.SelectedIndex = categories.Count > 0 ? 1 : 0;

            var btnOk = new Button { Text = "Сохранить", Left = 226, Top = 190, Width = 96, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Отмена", Left = 330, Top = 190, Width = 86, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { lblName, txtName, lblDesc, txtDesc, lblCat, cmbCat, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            if (form.ShowDialog() != DialogResult.OK) return null;

            string name = txtName.Text.Trim();
            if (string.IsNullOrWhiteSpace(name)) return null;

            int? categoryId = null;
            if (cmbCat.SelectedIndex > 0 && cmbCat.SelectedIndex - 1 < categories.Count)
            {
                categoryId = categories[cmbCat.SelectedIndex - 1].Id;
            }

            return (name, txtDesc.Text.Trim(), categoryId);
        }
    }
}
