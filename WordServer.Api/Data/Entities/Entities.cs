using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace WordServer.Api.Data.Entities
{
    [Table("Documents")]
    public class DocumentEntity
    {
        [Key]
        [MaxLength(100)]
        public string Id { get; set; } = string.Empty;

        [Required]
        [MaxLength(255)]
        public string Title { get; set; } = string.Empty;

        public string MetadataJson { get; set; } = "{}";

        public string ContentJson { get; set; } = "{}";

        public string SettingsJson { get; set; } = "{}";

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public DateTime? UpdatedAt { get; set; }
    }

    [Table("Templates")]
    public class TemplateEntity
    {
        [Key]
        [MaxLength(100)]
        public string Id { get; set; } = string.Empty;

        [Required]
        [MaxLength(255)]
        public string Name { get; set; } = string.Empty;

        [MaxLength(1000)]
        public string? Description { get; set; }

        public string DocumentStructureJson { get; set; } = "{}";

        public string SettingsJson { get; set; } = "{}";

        public string MetadataJson { get; set; } = "{}";

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public DateTime? UpdatedAt { get; set; }
    }

    [Table("Metadata")]
    public class MetadataEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [MaxLength(100)]
        public string? DocumentId { get; set; }

        [MaxLength(255)]
        public string Author { get; set; } = string.Empty;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public int Version { get; set; } = 1;

        [MaxLength(1000)]
        public string? Description { get; set; }
    }

    [Table("Settings")]
    public class SettingEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required]
        [MaxLength(100)]
        public string Key { get; set; } = string.Empty;

        [Required]
        public string Value { get; set; } = string.Empty;

        [MaxLength(500)]
        public string? Description { get; set; }

        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }

    [Table("Logs")]
    public class LogEntryEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        public DateTime Timestamp { get; set; } = DateTime.UtcNow;

        [MaxLength(50)]
        public string Level { get; set; } = "Information";

        [Required]
        public string Message { get; set; } = string.Empty;

        [MaxLength(255)]
        public string? Source { get; set; }

        public string? Exception { get; set; }
    }
}
