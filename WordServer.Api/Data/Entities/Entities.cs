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

        public int? CategoryId { get; set; }

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

    [Table("Users")]
    public class UserEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required]
        [MaxLength(100)]
        public string Username { get; set; } = string.Empty;

        [MaxLength(255)]
        public string FullName { get; set; } = string.Empty;

        [MaxLength(255)]
        public string Email { get; set; } = string.Empty;

        [Required]
        [MaxLength(255)]
        public string ApiKey { get; set; } = string.Empty;

        [MaxLength(500)]
        public string PasswordHash { get; set; } = string.Empty;

        [MaxLength(50)]
        public string Role { get; set; } = "Intern";

        public bool IsActive { get; set; } = true;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    [Table("DocumentVersions")]
    public class DocumentVersionEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required]
        [MaxLength(100)]
        public string DocumentId { get; set; } = string.Empty;

        public int VersionNumber { get; set; } = 1;

        [MaxLength(255)]
        public string Title { get; set; } = string.Empty;

        public string ContentJson { get; set; } = "{}";

        [MaxLength(255)]
        public string Author { get; set; } = string.Empty;

        [MaxLength(1000)]
        public string? ChangeNotes { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    [Table("TemplateCategories")]
    public class TemplateCategoryEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required]
        [MaxLength(100)]
        public string Name { get; set; } = string.Empty;

        [MaxLength(500)]
        public string? Description { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    [Table("ProcessingHistory")]
    public class ProcessingHistoryEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [MaxLength(100)]
        public string DocumentId { get; set; } = string.Empty;

        [MaxLength(50)]
        public string Operation { get; set; } = "uppercase";

        public int InputLength { get; set; }

        public int ResultLength { get; set; }

        [MaxLength(50)]
        public string? ClientIp { get; set; }

        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
    }

    [Table("DocumentComments")]
    public class DocumentCommentEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required]
        [MaxLength(100)]
        public string DocumentId { get; set; } = string.Empty;

        [MaxLength(255)]
        public string Author { get; set; } = string.Empty;

        [Required]
        [MaxLength(2000)]
        public string Comment { get; set; } = string.Empty;

        public bool Resolved { get; set; } = false;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
