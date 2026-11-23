using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using WordClient.Core.Models;
using Word = Microsoft.Office.Interop.Word;

namespace WordClient.Core.Services
{
    public enum InsertLocation
    {
        SelectedRange,
        EndOfDocument,
        BookmarkOrContentControl
    }

    public class WordInspectionReport
    {
        public int WordCount { get; set; }
        public int ParagraphCount { get; set; }
        public int EmptyParagraphCount { get; set; }
        public int TableCount { get; set; }
        public int CharacterCount { get; set; }
        public List<string> EmptyParagraphIndices { get; set; } = new List<string>();
        public bool HasSpellingErrors { get; set; }
        public int SpellingErrorsCount { get; set; }
        public Dictionary<string, string> Metadata { get; set; } = new Dictionary<string, string>();
    }

    public class WordDocumentFileInfo
    {
        public string Name { get; set; } = string.Empty;
        public string? FullPath { get; set; }
        public long? FileSizeBytes { get; set; }
    }

    public class WordDocumentService : IDisposable
    {
        private Word.Application? _wordApp;
        private bool _ownsApplicationInstance;
        private bool _shouldReleaseApplication;

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid lpclsid);

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject([MarshalAs(UnmanagedType.LPStruct)] Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        public WordDocumentService(Word.Application? existingApp = null)
        {
            if (existingApp != null)
            {
                _wordApp = existingApp;
                _ownsApplicationInstance = false;
            }
        }

        public Word.Application GetOrCreateApplication(bool makeVisible = true)
        {
            if (_wordApp != null)
            {
                try
                {
                    _wordApp.Visible = makeVisible;
                    return _wordApp;
                }
                catch
                {
                    _wordApp = null;
                }
            }

            try
            {
                // Try attaching to already running Word instance
                _wordApp = GetRunningWordApplication();
                _ownsApplicationInstance = false;
                _shouldReleaseApplication = true;
                ClientLogger.LogInfo("Connected to active Word instance.");
            }
            catch
            {
                // Word is not running - create a new instance
                try
                {
                    _wordApp = new Word.Application();
                    _ownsApplicationInstance = true;
                    _shouldReleaseApplication = true;
                    ClientLogger.LogInfo("Started new Microsoft Word process.");
                }
                catch (COMException ex)
                {
                    throw new InvalidOperationException("Microsoft Word is not installed or COM initialization failed.", ex);
                }
            }

            _wordApp.Visible = makeVisible;
            return _wordApp;
        }

        private static Word.Application GetRunningWordApplication()
        {
            int hrClsid = CLSIDFromProgID("Word.Application", out Guid clsid);
            if (hrClsid != 0)
            {
                throw new COMException("Failed to find ProgID Word.Application", hrClsid);
            }

            int hrActive = GetActiveObject(clsid, IntPtr.Zero, out object obj);
            if (hrActive != 0 || obj == null)
            {
                throw new COMException("Word is not currently running.", hrActive);
            }

            return (Word.Application)obj;
        }

        public Word.Document GetActiveDocument()
        {
            var app = GetOrCreateApplication();
            try
            {
                var doc = app.ActiveDocument;
                if (doc == null)
                {
                    throw new InvalidOperationException("No active document found in Microsoft Word. Please open or create a document first.");
                }
                return doc;
            }
            catch (COMException)
            {
                throw new InvalidOperationException("No active document found in Microsoft Word. Please open or create a document first.");
            }
        }

        public Word.Document CreateNewDocument()
        {
            var app = GetOrCreateApplication(makeVisible: true);
            Word.Documents? docs = null;
            try
            {
                docs = app.Documents;
                var doc = docs.Add();
                ClientLogger.LogInfo("Created a new blank Word document.");
                return doc;
            }
            finally
            {
                ReleaseComObject(docs);
            }
        }

        public Word.Document OpenDocument(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                throw new FileNotFoundException($"Target file was not found: {filePath}");
            }

            // Check file size (100 MB limit)
            var fileInfo = new FileInfo(filePath);
            if (fileInfo.Length > 104857600)
            {
                throw new InvalidOperationException($"The document file size ({fileInfo.Length / 1024.0 / 1024.0:F2} MB) exceeds the 100 MB limit.");
            }

            // Check if file is locked or in use
            if (IsFileLocked(filePath))
            {
                throw new IOException($"The target file '{Path.GetFileName(filePath)}' is locked or in use by another process.");
            }

            var app = GetOrCreateApplication(makeVisible: true);
            Word.Documents? docs = null;
            try
            {
                docs = app.Documents;
                var doc = docs.Open(filePath);
                ClientLogger.LogInfo($"Opened Word document: {filePath}");
                return doc;
            }
            finally
            {
                ReleaseComObject(docs);
            }
        }

        public void SaveActiveDocumentAs(string savePath)
        {
            if (string.IsNullOrWhiteSpace(savePath))
                throw new ArgumentException("Save path cannot be empty.", nameof(savePath));

            var fullPath = Path.GetFullPath(savePath);
            if (File.Exists(fullPath) && IsFileLocked(fullPath))
            {
                throw new IOException($"Target file '{Path.GetFileName(fullPath)}' is locked or in use by another process.");
            }

            var dir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var doc = GetActiveDocument();
            try
            {
                doc.SaveAs2(FileName: fullPath, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument);
                ClientLogger.LogInfo($"Saved active document as .docx: {fullPath}");
            }
            finally
            {
                ReleaseComObject(doc);
            }
        }

        public WordDocumentFileInfo GetActiveDocumentFileInfo()
        {
            var doc = GetActiveDocument();
            try
            {
                string name = doc.Name ?? "Document.docx";
                string? fullName = null;
                try
                {
                    fullName = doc.FullName;
                }
                catch
                {
                    fullName = null;
                }

                long? fileSize = null;
                if (!string.IsNullOrWhiteSpace(fullName) && File.Exists(fullName))
                {
                    fileSize = new FileInfo(fullName).Length;
                }

                return new WordDocumentFileInfo
                {
                    Name = name,
                    FullPath = (!string.IsNullOrWhiteSpace(fullName) && File.Exists(fullName)) ? fullName : null,
                    FileSizeBytes = fileSize
                };
            }
            finally
            {
                ReleaseComObject(doc);
            }
        }

        public Word.Document GenerateDocumentFromDto(DocumentDto dto, string? savePath = null)
        {
            if (dto == null)
                throw new ArgumentNullException(nameof(dto));

            var app = GetOrCreateApplication(makeVisible: true);
            Word.Documents? docs = null;
            Word.Document? doc = null;
            Word.Paragraphs? paragraphs = null;
            Word.Styles? styles = null;
            Word.Tables? tables = null;
            Word.Bookmarks? bookmarks = null;

            try
            {
                docs = app.Documents;
                doc = docs.Add();
                styles = doc.Styles;
                paragraphs = doc.Paragraphs;
                tables = doc.Tables;
                bookmarks = doc.Bookmarks;

                // Configure Normal style font defaults first
                if (dto.Settings != null)
                {
                    Word.Style? normalStyle = null;
                    Word.Font? normalFont = null;
                    try
                    {
                        normalStyle = styles[Word.WdBuiltinStyle.wdStyleNormal];
                        normalFont = normalStyle.Font;
                        if (!string.IsNullOrWhiteSpace(dto.Settings.Font))
                        {
                            normalFont.Name = dto.Settings.Font;
                        }
                        if (dto.Settings.FontSize > 0)
                        {
                            normalFont.Size = dto.Settings.FontSize;
                        }
                    }
                    catch
                    {
                        // Fallback applied per range below
                    }
                    finally
                    {
                        ReleaseComObject(normalFont);
                        ReleaseComObject(normalStyle);
                    }
                }

                // Add Title as Heading 1
                if (!string.IsNullOrWhiteSpace(dto.Title))
                {
                    Word.Paragraph? titlePara = null;
                    Word.Range? titleRange = null;
                    Word.Font? titleFont = null;
                    try
                    {
                        titlePara = paragraphs.Add();
                        titleRange = titlePara.Range;
                        titleRange.Text = dto.Title + "\n";
                        try
                        {
                            titleRange.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
                        }
                        catch
                        {
                            titleRange.Bold = 1;
                        }

                        if (!string.IsNullOrWhiteSpace(dto.Settings?.Font))
                        {
                            titleFont = titleRange.Font;
                            titleFont.Name = dto.Settings!.Font;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(titleFont);
                        ReleaseComObject(titleRange);
                        ReleaseComObject(titlePara);
                    }
                }

                // Insert paragraphs with styles
                if (dto.Content?.Paragraphs != null)
                {
                    foreach (var paraDto in dto.Content.Paragraphs)
                    {
                        Word.Paragraph? para = null;
                        Word.Range? range = null;
                        Word.Font? rangeFont = null;
                        try
                        {
                            para = paragraphs.Add();
                            range = para.Range;
                            range.Text = paraDto.Text + "\n";

                            bool isHeading = false;
                            if (!string.IsNullOrWhiteSpace(paraDto.Style))
                            {
                                try
                                {
                                    if (paraDto.Style.Equals("Heading1", StringComparison.OrdinalIgnoreCase) ||
                                        paraDto.Style.Equals("Heading 1", StringComparison.OrdinalIgnoreCase))
                                    {
                                        range.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
                                        isHeading = true;
                                    }
                                    else if (paraDto.Style.Equals("Heading2", StringComparison.OrdinalIgnoreCase) ||
                                             paraDto.Style.Equals("Heading 2", StringComparison.OrdinalIgnoreCase))
                                    {
                                        range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
                                        isHeading = true;
                                    }
                                    else
                                    {
                                        range.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
                                    }
                                }
                                catch
                                {
                                    // Style could not be applied; keep default
                                }
                            }

                            // Re-apply configured font/size after style assignment so settings are preserved
                            if (dto.Settings != null)
                            {
                                rangeFont = range.Font;
                                if (!string.IsNullOrWhiteSpace(dto.Settings.Font))
                                {
                                    rangeFont.Name = dto.Settings.Font;
                                }
                                if (!isHeading && dto.Settings.FontSize > 0)
                                {
                                    rangeFont.Size = dto.Settings.FontSize;
                                }
                            }
                        }
                        finally
                        {
                            ReleaseComObject(rangeFont);
                            ReleaseComObject(range);
                            ReleaseComObject(para);
                        }
                    }
                }

                // Add Tables
                if (dto.Content?.Tables != null)
                {
                    foreach (var tableDto in dto.Content.Tables)
                    {
                        if (tableDto.Rows != null && tableDto.Rows.Count > 0)
                        {
                            int numRows = tableDto.Rows.Count;
                            int numCols = 0;
                            foreach (var r in tableDto.Rows)
                            {
                                if (r != null && r.Count > numCols) numCols = r.Count;
                            }

                            if (numCols > 0)
                            {
                                Word.Range? endRange = null;
                                Word.Table? wordTable = null;
                                Word.Borders? borders = null;
                                try
                                {
                                    endRange = doc.Content;
                                    endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                                    wordTable = tables.Add(endRange, numRows, numCols);
                                    borders = wordTable.Borders;
                                    borders.Enable = 1;

                                    for (int rowIdx = 0; rowIdx < numRows; rowIdx++)
                                    {
                                        var rowData = tableDto.Rows[rowIdx];
                                        if (rowData == null) continue;

                                        for (int colIdx = 0; colIdx < rowData.Count && colIdx < numCols; colIdx++)
                                        {
                                            Word.Cell? cell = null;
                                            Word.Range? cellRange = null;
                                            Word.Font? cellFont = null;
                                            Word.Shading? shading = null;
                                            try
                                            {
                                                cell = wordTable.Cell(rowIdx + 1, colIdx + 1);
                                                cellRange = cell.Range;
                                                cellRange.Text = rowData[colIdx] ?? string.Empty;

                                                if (dto.Settings != null)
                                                {
                                                    cellFont = cellRange.Font;
                                                    if (!string.IsNullOrWhiteSpace(dto.Settings.Font))
                                                        cellFont.Name = dto.Settings.Font;
                                                    if (dto.Settings.FontSize > 0)
                                                        cellFont.Size = dto.Settings.FontSize;
                                                }

                                                // Header row styling
                                                if (rowIdx == 0)
                                                {
                                                    cellRange.Bold = 1;
                                                    shading = cell.Shading;
                                                    shading.BackgroundPatternColor = Word.WdColor.wdColorGray15;
                                                }
                                            }
                                            finally
                                            {
                                                ReleaseComObject(shading);
                                                ReleaseComObject(cellFont);
                                                ReleaseComObject(cellRange);
                                                ReleaseComObject(cell);
                                            }
                                        }
                                    }
                                }
                                finally
                                {
                                    ReleaseComObject(borders);
                                    ReleaseComObject(wordTable);
                                    ReleaseComObject(endRange);
                                }
                            }
                        }
                    }
                }

                // Add default bookmark for result insertion
                Word.Range? bookmarkRange = null;
                Word.Bookmark? createdBookmark = null;
                try
                {
                    bookmarkRange = doc.Content;
                    bookmarkRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    createdBookmark = bookmarks.Add("ResultBookmark", bookmarkRange);
                }
                catch
                {
                    // Optional bookmark
                }
                finally
                {
                    ReleaseComObject(createdBookmark);
                    ReleaseComObject(bookmarkRange);
                }

                // Save document if path is specified
                if (!string.IsNullOrWhiteSpace(savePath))
                {
                    var fullPath = Path.GetFullPath(savePath);
                    if (File.Exists(fullPath) && IsFileLocked(fullPath))
                    {
                        throw new IOException($"Target file '{Path.GetFileName(fullPath)}' is locked or in use by another process.");
                    }

                    var dir = Path.GetDirectoryName(fullPath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    doc.SaveAs2(FileName: fullPath, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument);
                    ClientLogger.LogInfo($"Document saved to: {fullPath}");
                }

                return doc;
            }
            finally
            {
                ReleaseComObject(bookmarks);
                ReleaseComObject(tables);
                ReleaseComObject(styles);
                ReleaseComObject(paragraphs);
                ReleaseComObject(docs);
            }
        }

        public Word.Document GenerateDocumentFromTemplate(TemplateDto template, string? savePath = null)
        {
            if (template == null)
                throw new ArgumentNullException(nameof(template));

            var dto = new DocumentDto
            {
                Id = template.Id,
                Title = template.Name,
                Metadata = template.Metadata ?? new DocumentMetadataDto(),
                Content = template.Content ?? new DocumentContentDto(),
                Settings = template.Settings ?? new DocumentSettingsDto()
            };

            return GenerateDocumentFromDto(dto, savePath);
        }

        public DocumentDto ExtractDocumentStructure(string? overrideId = null, string? overrideTitle = null)
        {
            var doc = GetActiveDocument();
            Word.Paragraphs? paragraphs = null;
            Word.Tables? tables = null;
            Word.Range? contentRange = null;

            try
            {
                contentRange = doc.Content;
                paragraphs = doc.Paragraphs;
                tables = doc.Tables;

                var extractedParagraphs = new List<ParagraphDto>();
                for (int i = 1; i <= paragraphs.Count; i++)
                {
                    Word.Paragraph? para = null;
                    Word.Range? range = null;
                    try
                    {
                        para = paragraphs[i];
                        range = para.Range;

                        // Skip paragraphs that belong to tables (extracted separately below)
                        bool inTable = false;
                        try
                        {
                            inTable = (bool)range.Information[Word.WdInformation.wdWithInTable];
                        }
                        catch
                        {
                            inTable = false;
                        }

                        if (inTable) continue;

                        var text = (range.Text ?? string.Empty).TrimEnd('\r', '\n', '\a');
                        if (string.IsNullOrWhiteSpace(text)) continue;

                        string styleName = "Normal";
                        try
                        {
                            dynamic styleObj = range.get_Style();
                            string rawStyle = styleObj?.NameLocal?.ToString() ?? string.Empty;
                            if (rawStyle.IndexOf("Heading 1", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                rawStyle.IndexOf("Заголовок 1", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                styleName = "Heading1";
                            }
                            else if (rawStyle.IndexOf("Heading 2", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                     rawStyle.IndexOf("Заголовок 2", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                styleName = "Heading2";
                            }
                        }
                        catch
                        {
                            styleName = "Normal";
                        }

                        extractedParagraphs.Add(new ParagraphDto
                        {
                            Style = styleName,
                            Text = text
                        });
                    }
                    finally
                    {
                        ReleaseComObject(range);
                        ReleaseComObject(para);
                    }
                }

                var extractedTables = new List<TableDto>();
                for (int tIdx = 1; tIdx <= tables.Count; tIdx++)
                {
                    Word.Table? table = null;
                    try
                    {
                        table = tables[tIdx];
                        int rowCount = table.Rows.Count;
                        int colCount = table.Columns.Count;
                        var rows = new List<List<string>>();

                        for (int r = 1; r <= rowCount; r++)
                        {
                            var rowValues = new List<string>();
                            for (int c = 1; c <= colCount; c++)
                            {
                                Word.Cell? cell = null;
                                Word.Range? cellRange = null;
                                try
                                {
                                    cell = table.Cell(r, c);
                                    cellRange = cell.Range;
                                    var cellText = (cellRange.Text ?? string.Empty).TrimEnd('\r', '\n', '\a');
                                    rowValues.Add(cellText);
                                }
                                catch
                                {
                                    rowValues.Add(string.Empty);
                                }
                                finally
                                {
                                    ReleaseComObject(cellRange);
                                    ReleaseComObject(cell);
                                }
                            }
                            rows.Add(rowValues);
                        }

                        extractedTables.Add(new TableDto
                        {
                            Name = $"Table{tIdx}",
                            Rows = rows
                        });
                    }
                    finally
                    {
                        ReleaseComObject(table);
                    }
                }

                string fontName = "Times New Roman";
                int fontSize = 12;
                try
                {
                    if (!string.IsNullOrWhiteSpace(contentRange.Font.Name))
                    {
                        fontName = contentRange.Font.Name;
                    }
                    if (contentRange.Font.Size > 0 && contentRange.Font.Size < 200)
                    {
                        fontSize = (int)Math.Round(contentRange.Font.Size);
                    }
                }
                catch
                {
                    // Keep default font settings
                }

                string author = Environment.UserName;
                string title = overrideTitle ?? Path.GetFileNameWithoutExtension(doc.Name);
                try
                {
                    dynamic props = doc.BuiltInDocumentProperties;
                    string? propAuthor = props["Author"]?.Value?.ToString();
                    if (!string.IsNullOrWhiteSpace(propAuthor)) author = propAuthor!;
                    if (string.IsNullOrWhiteSpace(overrideTitle))
                    {
                        string? propTitle = props["Title"]?.Value?.ToString();
                        if (!string.IsNullOrWhiteSpace(propTitle)) title = propTitle!;
                    }
                }
                catch
                {
                    // Keep fallback metadata
                }

                return new DocumentDto
                {
                    Id = overrideId ?? $"doc-{DateTime.UtcNow:yyyyMMdd-HHmmss}",
                    Title = string.IsNullOrWhiteSpace(title) ? "Untitled Document" : title,
                    Metadata = new DocumentMetadataDto
                    {
                        Author = author,
                        CreatedAt = DateTime.UtcNow.ToString("O"),
                        Description = $"Extracted from Microsoft Word ({extractedParagraphs.Count} paragraphs, {extractedTables.Count} tables)",
                        Version = 1
                    },
                    Content = new DocumentContentDto
                    {
                        Paragraphs = extractedParagraphs,
                        Tables = extractedTables
                    },
                    Settings = new DocumentSettingsDto
                    {
                        Font = fontName,
                        FontSize = fontSize
                    }
                };
            }
            finally
            {
                ReleaseComObject(contentRange);
                ReleaseComObject(tables);
                ReleaseComObject(paragraphs);
                ReleaseComObject(doc);
            }
        }

        public string GetSelectedText()
        {
            var app = GetOrCreateApplication();
            var doc = GetActiveDocument();

            Word.Selection? selection = null;
            Word.Range? range = null;

            try
            {
                selection = app.Selection;
                if (selection == null) return string.Empty;

                range = selection.Range;
                if (range == null) return string.Empty;

                var text = range.Text ?? string.Empty;
                return text.TrimEnd('\r', '\n', '\a');
            }
            finally
            {
                ReleaseComObject(range);
                ReleaseComObject(selection);
                ReleaseComObject(doc);
            }
        }

        public string GetDocumentText()
        {
            var doc = GetActiveDocument();
            Word.Range? contentRange = null;

            try
            {
                contentRange = doc.Content;
                var text = contentRange?.Text ?? string.Empty;
                return text;
            }
            finally
            {
                ReleaseComObject(contentRange);
                ReleaseComObject(doc);
            }
        }

        public void InsertResult(string resultText, InsertLocation location, string? bookmarkName = null)
        {
            var app = GetOrCreateApplication();
            var doc = GetActiveDocument();

            Word.Selection? selection = null;
            Word.Range? range = null;

            try
            {
                switch (location)
                {
                    case InsertLocation.SelectedRange:
                        selection = app.Selection;
                        range = selection.Range;
                        range.Text = resultText;
                        break;

                    case InsertLocation.EndOfDocument:
                        range = doc.Content;
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        range.Text = "\n" + resultText;
                        break;

                    case InsertLocation.BookmarkOrContentControl:
                        string targetBookmark = !string.IsNullOrWhiteSpace(bookmarkName) ? bookmarkName! : "ResultBookmark";
                        Word.Bookmarks? bookmarks = null;
                        Word.ContentControls? contentControls = null;
                        try
                        {
                            bookmarks = doc.Bookmarks;
                            if (bookmarks.Exists(targetBookmark))
                            {
                                Word.Bookmark? bmk = null;
                                Word.Range? bmkRange = null;
                                Word.Bookmark? recreated = null;
                                try
                                {
                                    bmk = bookmarks[targetBookmark];
                                    bmkRange = bmk.Range;
                                    bmkRange.Text = resultText;
                                    recreated = bookmarks.Add(targetBookmark, bmkRange);
                                }
                                finally
                                {
                                    ReleaseComObject(recreated);
                                    ReleaseComObject(bmkRange);
                                    ReleaseComObject(bmk);
                                }
                            }
                            else
                            {
                                contentControls = doc.ContentControls;
                                if (contentControls.Count > 0)
                                {
                                    Word.ContentControl? cc = null;
                                    Word.Range? ccRange = null;
                                    try
                                    {
                                        cc = contentControls[1];
                                        ccRange = cc.Range;
                                        ccRange.Text = resultText;
                                    }
                                    finally
                                    {
                                        ReleaseComObject(ccRange);
                                        ReleaseComObject(cc);
                                    }
                                }
                                else
                                {
                                    // Fallback to end of document
                                    range = doc.Content;
                                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                    range.Text = "\n" + resultText;
                                }
                            }
                        }
                        finally
                        {
                            ReleaseComObject(contentControls);
                            ReleaseComObject(bookmarks);
                        }
                        break;
                }

                ClientLogger.LogInfo($"Inserted processed text ({resultText.Length} chars) at {location}.");
            }
            finally
            {
                ReleaseComObject(range);
                ReleaseComObject(selection);
                ReleaseComObject(doc);
            }
        }

        public WordInspectionReport InspectDocument()
        {
            var doc = GetActiveDocument();
            var report = new WordInspectionReport();

            Word.Paragraphs? paragraphs = null;
            Word.Tables? tables = null;
            Word.Range? contentRange = null;

            try
            {
                contentRange = doc.Content;
                paragraphs = doc.Paragraphs;
                tables = doc.Tables;

                report.WordCount = contentRange.ComputeStatistics(Word.WdStatistic.wdStatisticWords);
                report.CharacterCount = contentRange.ComputeStatistics(Word.WdStatistic.wdStatisticCharacters);
                report.ParagraphCount = paragraphs.Count;
                report.TableCount = tables.Count;

                int emptyCount = 0;
                for (int i = 1; i <= paragraphs.Count; i++)
                {
                    var p = paragraphs[i];
                    var r = p.Range;
                    var t = r.Text?.TrimEnd('\r', '\n', '\a') ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(t))
                    {
                        emptyCount++;
                        report.EmptyParagraphIndices.Add($"Paragraph #{i}");
                    }
                    ReleaseComObject(r);
                    ReleaseComObject(p);
                }
                report.EmptyParagraphCount = emptyCount;

                // Spell check inspection
                try
                {
                    var spellingErrors = contentRange.SpellingErrors;
                    report.SpellingErrorsCount = spellingErrors.Count;
                    report.HasSpellingErrors = spellingErrors.Count > 0;
                    ReleaseComObject(spellingErrors);
                }
                catch
                {
                    report.HasSpellingErrors = false;
                }

                // Document Metadata
                try
                {
                    dynamic props = doc.BuiltInDocumentProperties;
                    report.Metadata["Title"] = props["Title"]?.Value?.ToString() ?? "Untitled";
                    report.Metadata["Author"] = props["Author"]?.Value?.ToString() ?? "Unknown";
                    report.Metadata["CreationDate"] = props["Creation Date"]?.Value?.ToString() ?? DateTime.UtcNow.ToString("O");
                }
                catch
                {
                    report.Metadata["Title"] = "Word Document";
                }

                return report;
            }
            finally
            {
                ReleaseComObject(contentRange);
                ReleaseComObject(tables);
                ReleaseComObject(paragraphs);
                ReleaseComObject(doc);
            }
        }

        public int FindAndReplace(string findText, string replaceText)
        {
            if (string.IsNullOrEmpty(findText)) return 0;

            var doc = GetActiveDocument();
            Word.Range? contentRange = null;
            Word.Find? find = null;
            Word.Replacement? replacement = null;

            try
            {
                contentRange = doc.Content;
                find = contentRange.Find;
                replacement = find.Replacement;
                find.ClearFormatting();
                replacement.ClearFormatting();
                find.Text = findText;
                replacement.Text = replaceText;

                object replaceAll = Word.WdReplace.wdReplaceAll;
                bool success = find.Execute(
                    Replace: ref replaceAll,
                    MatchCase: false,
                    MatchWholeWord: false,
                    MatchWildcards: false);

                ClientLogger.LogInfo($"Find and replace '{findText}' -> '{replaceText}' executed (Success: {success}).");
                return success ? 1 : 0;
            }
            finally
            {
                ReleaseComObject(replacement);
                ReleaseComObject(find);
                ReleaseComObject(contentRange);
                ReleaseComObject(doc);
            }
        }

        public static void ReleaseComObject(object? comObj)
        {
#if NET5_0_OR_GREATER
            if (!OperatingSystem.IsWindows()) return;
#endif
            if (comObj != null && Marshal.IsComObject(comObj))
            {
                try
                {
                    Marshal.ReleaseComObject(comObj);
                }
                catch
                {
                    // Ignore release errors
                }
            }
        }

        public static bool IsFileLocked(string filePath)
        {
            try
            {
                using var stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }

        public void Dispose()
        {
            if (_wordApp != null && _shouldReleaseApplication)
            {
                try
                {
                    if (_ownsApplicationInstance)
                    {
                        object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                        _wordApp.Quit(ref saveChanges);
                    }
                }
                catch
                {
                    // Ignore quit errors
                }
                finally
                {
                    ReleaseComObject(_wordApp);
                    _wordApp = null;
                }
            }

            // Force Garbage Collection to ensure all COM wrappers are finalized
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
