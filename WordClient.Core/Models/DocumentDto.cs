using System;
using System.Collections.Generic;

namespace WordClient.Core.Models
{
    public class DocumentDto
    {
        public string Id { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public DocumentMetadataDto Metadata { get; set; } = new DocumentMetadataDto();
        public DocumentContentDto Content { get; set; } = new DocumentContentDto();
        public DocumentSettingsDto Settings { get; set; } = new DocumentSettingsDto();
    }

    public class DocumentMetadataDto
    {
        public string Author { get; set; } = string.Empty;
        public string CreatedAt { get; set; } = string.Empty;
        public string? Description { get; set; }
        public int? Version { get; set; } = 1;
    }

    public class DocumentContentDto
    {
        public List<ParagraphDto> Paragraphs { get; set; } = new List<ParagraphDto>();
        public List<TableDto> Tables { get; set; } = new List<TableDto>();
    }

    public class ParagraphDto
    {
        public string Style { get; set; } = "Normal";
        public string Text { get; set; } = string.Empty;
    }

    public class TableDto
    {
        public string Name { get; set; } = string.Empty;
        public List<List<string>> Rows { get; set; } = new List<List<string>>();
    }

    public class DocumentSettingsDto
    {
        public string Font { get; set; } = "Times New Roman";
        public int FontSize { get; set; } = 12;
    }
}
