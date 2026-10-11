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

    public class WordDocumentService : IDisposable
    {
        private Word.Application? _wordApp;
        private bool _ownsApplicationInstance;

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
                ClientLogger.LogInfo("Connected to active Word instance.");
            }
            catch
            {
                // Word is not running - create a new instance
                try
                {
                    _wordApp = new Word.Application();
                    _ownsApplicationInstance = true;
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
            var doc = app.Documents.Add();
            ClientLogger.LogInfo("Created a new blank Word document.");
            return doc;
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
                throw new InvalidOperationException($"The document file size ({fileInfo.Length / 1024 / 1024} MB) exceeds the 100 MB limit.");
            }

            // Check if file is locked or in use
            if (IsFileLocked(filePath))
            {
                throw new IOException($"The target file '{Path.GetFileName(filePath)}' is locked or in use by another process.");
            }

            var app = GetOrCreateApplication(makeVisible: true);
            var doc = app.Documents.Open(filePath);
            ClientLogger.LogInfo($"Opened Word document: {filePath}");
            return doc;
        }

        public Word.Document GenerateDocumentFromDto(DocumentDto dto, string? savePath = null)
        {
            if (dto == null)
                throw new ArgumentNullException(nameof(dto));

            var app = GetOrCreateApplication(makeVisible: true);
            Word.Document? doc = null;
            Word.Paragraphs? paragraphs = null;
            Word.Styles? styles = null;
            Word.Tables? tables = null;

            try
            {
                doc = app.Documents.Add();
                styles = doc.Styles;
                paragraphs = doc.Paragraphs;
                tables = doc.Tables;

                // Configure default font settings
                var contentRange = doc.Content;
                if (dto.Settings != null)
                {
                    if (!string.IsNullOrWhiteSpace(dto.Settings.Font))
                    {
                        contentRange.Font.Name = dto.Settings.Font;
                    }
                    if (dto.Settings.FontSize > 0)
                    {
                        contentRange.Font.Size = dto.Settings.FontSize;
                    }
                }
                ReleaseComObject(contentRange);

                // Add Title as Heading 1
                if (!string.IsNullOrWhiteSpace(dto.Title))
                {
                    var titlePara = doc.Paragraphs.Add();
                    var titleRange = titlePara.Range;
                    titleRange.Text = dto.Title + "\n";
                    try
                    {
                        titleRange.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
                    }
                    catch
                    {
                        titleRange.Bold = 1;
                    }
                    ReleaseComObject(titleRange);
                    ReleaseComObject(titlePara);
                }

                // Insert paragraphs with styles
                if (dto.Content?.Paragraphs != null)
                {
                    foreach (var paraDto in dto.Content.Paragraphs)
                    {
                        var para = doc.Paragraphs.Add();
                        var range = para.Range;
                        range.Text = paraDto.Text + "\n";

                        if (!string.IsNullOrWhiteSpace(paraDto.Style))
                        {
                            try
                            {
                                if (paraDto.Style.Equals("Heading1", StringComparison.OrdinalIgnoreCase) ||
                                    paraDto.Style.Equals("Heading 1", StringComparison.OrdinalIgnoreCase))
                                {
                                    range.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
                                }
                                else if (paraDto.Style.Equals("Heading2", StringComparison.OrdinalIgnoreCase) ||
                                         paraDto.Style.Equals("Heading 2", StringComparison.OrdinalIgnoreCase))
                                {
                                    range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
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

                        ReleaseComObject(range);
                        ReleaseComObject(para);
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
                                var endRange = doc.Content;
                                endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                                var wordTable = doc.Tables.Add(endRange, numRows, numCols);
                                wordTable.Borders.Enable = 1;

                                for (int rowIdx = 0; rowIdx < numRows; rowIdx++)
                                {
                                    var rowData = tableDto.Rows[rowIdx];
                                    if (rowData == null) continue;

                                    for (int colIdx = 0; colIdx < rowData.Count && colIdx < numCols; colIdx++)
                                    {
                                        var cell = wordTable.Cell(rowIdx + 1, colIdx + 1);
                                        var cellRange = cell.Range;
                                        cellRange.Text = rowData[colIdx] ?? string.Empty;

                                        // Header row styling
                                        if (rowIdx == 0)
                                        {
                                            cellRange.Bold = 1;
                                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray15;
                                        }

                                        ReleaseComObject(cellRange);
                                        ReleaseComObject(cell);
                                    }
                                }

                                ReleaseComObject(wordTable);
                                ReleaseComObject(endRange);
                            }
                        }
                    }
                }

                // Add default bookmark for result insertion
                try
                {
                    var bookmarkRange = doc.Content;
                    bookmarkRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    doc.Bookmarks.Add("ResultBookmark", bookmarkRange);
                    ReleaseComObject(bookmarkRange);
                }
                catch
                {
                    // Optional bookmark
                }

                // Save document if path is specified
                if (!string.IsNullOrWhiteSpace(savePath))
                {
                    var fullPath = Path.GetFullPath(savePath);
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
                ReleaseComObject(tables);
                ReleaseComObject(styles);
                ReleaseComObject(paragraphs);
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
                        string targetBookmark = !string.IsNullOrWhiteSpace(bookmarkName) ? bookmarkName : "ResultBookmark";
                        if (doc.Bookmarks.Exists(targetBookmark))
                        {
                            var bmk = doc.Bookmarks[targetBookmark];
                            var bmkRange = bmk.Range;
                            bmkRange.Text = resultText;
                            doc.Bookmarks.Add(targetBookmark, bmkRange);
                            ReleaseComObject(bmkRange);
                            ReleaseComObject(bmk);
                        }
                        else if (doc.ContentControls.Count > 0)
                        {
                            var cc = doc.ContentControls[1];
                            cc.Range.Text = resultText;
                            ReleaseComObject(cc);
                        }
                        else
                        {
                            // Fallback to end of document
                            range = doc.Content;
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            range.Text = "\n" + resultText;
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

            try
            {
                contentRange = doc.Content;
                find = contentRange.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Text = findText;
                find.Replacement.Text = replaceText;

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
                ReleaseComObject(find);
                ReleaseComObject(contentRange);
                ReleaseComObject(doc);
            }
        }

        public static void ReleaseComObject(object? comObj)
        {
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
            if (_ownsApplicationInstance && _wordApp != null)
            {
                try
                {
                    object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                    _wordApp.Quit(ref saveChanges);
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
