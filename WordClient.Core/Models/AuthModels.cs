namespace WordClient.Core.Models
{
    public class AuthValidateRequest
    {
        public string? ApiKey { get; set; }
    }

    public class AuthValidateResponse
    {
        public bool Valid { get; set; }
        public string Message { get; set; } = string.Empty;
        public string? User { get; set; }
        public string? FullName { get; set; }
        public string? Role { get; set; }
    }
}
