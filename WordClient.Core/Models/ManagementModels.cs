using System;

namespace WordClient.Core.Models
{
    public class DocumentSummaryDto
    {
        public string Id { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
        public int Version { get; set; } = 1;
        public DateTime CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
    }

    public class TemplateCategoryDto
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public DateTime CreatedAt { get; set; }
    }

    public class CreateCategoryRequest
    {
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
    }

    public class CreateUserRequest
    {
        public string Username { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Role { get; set; } = "User";
        public string? ApiKey { get; set; }
    }

    public class UpdateUserStatusRequest
    {
        public bool IsActive { get; set; }
    }

    public class CreateCommentRequest
    {
        public string DocumentId { get; set; } = string.Empty;
        public string? Author { get; set; }
        public string Comment { get; set; } = string.Empty;
    }

    public class ResolveCommentRequest
    {
        public bool Resolved { get; set; } = true;
    }

    public class UpdateSettingRequest
    {
        public string Value { get; set; } = string.Empty;
        public string? Description { get; set; }
    }

    public class ProcessingHistoryDto
    {
        public int Id { get; set; }
        public string DocumentId { get; set; } = string.Empty;
        public string Operation { get; set; } = string.Empty;
        public int InputLength { get; set; }
        public int ResultLength { get; set; }
        public string? ClientIp { get; set; }
        public DateTime Timestamp { get; set; }
    }
}
