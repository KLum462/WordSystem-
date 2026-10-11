namespace WordClient.Core.Models
{
    public class DocumentProcessRequest
    {
        public string? Id { get; set; }
        public string Text { get; set; } = string.Empty;
        public string Operation { get; set; } = "uppercase";
    }

    public class DocumentProcessResponse
    {
        public string Id { get; set; } = string.Empty;
        public string Result { get; set; } = string.Empty;
        public int CharactersCount { get; set; }
        public int WordsCount { get; set; }
        public int ParagraphsCount { get; set; }
    }
}
