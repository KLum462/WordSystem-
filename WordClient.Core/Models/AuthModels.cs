namespace WordClient.Core.Models
{
    public class AuthValidateRequest
    {
        public string? ApiKey { get; set; }
    }

    public class LoginRequest
    {
        public string Username { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
    }

    public class ChangePasswordRequest
    {
        public string Username { get; set; } = string.Empty;
        public string CurrentPassword { get; set; } = string.Empty;
        public string NewPassword { get; set; } = string.Empty;
    }

    public class AuthValidateResponse
    {
        public bool Valid { get; set; }
        public string Message { get; set; } = string.Empty;
        public string? User { get; set; }
        public string? FullName { get; set; }
        public string? Role { get; set; }
        public string? ApiKey { get; set; }
    }
}
