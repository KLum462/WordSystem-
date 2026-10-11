using System.Collections.Generic;

namespace WordClient.Core.Models
{
    public class DocumentCheckRequest
    {
        public string? DocumentId { get; set; }
        public string Text { get; set; } = string.Empty;
        public List<string>? Paragraphs { get; set; }
        public int? TableCount { get; set; }
    }

    public class DocumentCheckResult
    {
        public int WordCount { get; set; }
        public int ParagraphCount { get; set; }
        public int EmptyParagraphCount { get; set; }
        public int TableCount { get; set; }
        public int CharacterCount { get; set; }
        public List<string> Issues { get; set; } = new List<string>();
        public string Summary { get; set; } = string.Empty;
    }
}
