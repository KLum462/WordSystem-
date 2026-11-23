namespace WordClient.Core.Models
{
    public class TextProcessRequest
    {
        public string? DocumentId { get; set; }
        public string Text { get; set; } = string.Empty;
        public string Operation { get; set; } = "uppercase";
    }

    public class TextProcessResponse
    {
        public int Id { get; set; }
        public string Result { get; set; } = string.Empty;
    }

    public class ErrorResponse
    {
        public string Error { get; set; } = string.Empty;
    }
}
