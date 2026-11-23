using System;

namespace WordClient.Core.Models
{
    public class TemplateDto
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public int? CategoryId { get; set; }
        public DocumentContentDto Content { get; set; } = new DocumentContentDto();
        public DocumentSettingsDto Settings { get; set; } = new DocumentSettingsDto();
        public DocumentMetadataDto Metadata { get; set; } = new DocumentMetadataDto();
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? UpdatedAt { get; set; }
    }
}
