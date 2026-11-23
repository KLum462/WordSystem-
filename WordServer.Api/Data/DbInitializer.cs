using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using WordClient.Core.Models;
using WordServer.Api.Data.Entities;
using WordServer.Api.Services;

namespace WordServer.Api.Data
{
    public static class DbInitializer
    {
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };

        public static async Task InitializeAsync(AppDbContext context)
        {
            // Ensure schema is created (for both migrations or direct ensure created)
            if (context.Database.IsRelational())
            {
                await context.Database.EnsureCreatedAsync();
                await EnsurePasswordHashColumnExistsAsync(context);
            }

            // Seed sample document doc-001 if not present
            if (!await context.Documents.AnyAsync(d => d.Id == "doc-001"))
            {
                var doc001 = new DocumentDto
                {
                    Id = "doc-001",
                    Title = "Report",
                    Metadata = new DocumentMetadataDto
                    {
                        Author = "Ivan",
                        CreatedAt = "2026-10-07T10:00:00Z",
                        Description = "Official quarterly report",
                        Version = 1
                    },
                    Content = new DocumentContentDto
                    {
                        Paragraphs = new List<ParagraphDto>
                        {
                            new ParagraphDto { Style = "Heading1", Text = "Heading" },
                            new ParagraphDto { Style = "Normal", Text = "First paragraph" },
                            new ParagraphDto { Style = "Normal", Text = "Second paragraph" }
                        },
                        Tables = new List<TableDto>
                        {
                            new TableDto
                            {
                                Name = "Table1",
                                Rows = new List<List<string>>
                                {
                                    new List<string> { "Column1", "Column2" },
                                    new List<string> { "Value1", "Value2" }
                                }
                            }
                        }
                    },
                    Settings = new DocumentSettingsDto
                    {
                        Font = "Times New Roman",
                        FontSize = 12
                    }
                };

                var entity = new DocumentEntity
                {
                    Id = doc001.Id,
                    Title = doc001.Title,
                    MetadataJson = JsonSerializer.Serialize(doc001.Metadata, JsonOptions),
                    ContentJson = JsonSerializer.Serialize(doc001.Content, JsonOptions),
                    SettingsJson = JsonSerializer.Serialize(doc001.Settings, JsonOptions),
                    CreatedAt = DateTime.UtcNow
                };

                context.Documents.Add(entity);

                // Add to Metadata table
                context.Metadata.Add(new MetadataEntity
                {
                    DocumentId = doc001.Id,
                    Author = doc001.Metadata.Author,
                    CreatedAt = DateTime.UtcNow,
                    Version = doc001.Metadata.Version ?? 1,
                    Description = doc001.Metadata.Description
                });
            }

            // Seed Template Categories before Templates so we can assign default CategoryId
            if (!await context.TemplateCategories.AnyAsync())
            {
                context.TemplateCategories.AddRange(
                    new TemplateCategoryEntity { Name = "Отчеты (Reports)", Description = "Корпоративные и квартальные отчеты", CreatedAt = DateTime.UtcNow },
                    new TemplateCategoryEntity { Name = "Договоры (Contracts)", Description = "Договоры подряда и соглашения", CreatedAt = DateTime.UtcNow },
                    new TemplateCategoryEntity { Name = "Служебные записки (Memos)", Description = "Внутренний документооборот", CreatedAt = DateTime.UtcNow }
                );
                await context.SaveChangesAsync();
            }

            var reportsCategoryId = await context.TemplateCategories
                .Where(c => c.Name.Contains("Reports"))
                .Select(c => (int?)c.Id)
                .FirstOrDefaultAsync();

            // Seed sample template tpl-001 if not present
            if (!await context.Templates.AnyAsync(t => t.Id == "tpl-001"))
            {
                var templateContent = new DocumentContentDto
                {
                    Paragraphs = new List<ParagraphDto>
                    {
                        new ParagraphDto { Style = "Heading1", Text = "Project Status Template" },
                        new ParagraphDto { Style = "Normal", Text = "Fill in the project details below:" }
                    },
                    Tables = new List<TableDto>
                    {
                        new TableDto
                        {
                            Name = "DeliverablesTable",
                            Rows = new List<List<string>>
                            {
                                new List<string> { "Deliverable", "Owner", "Status" },
                                new List<string> { "API Integration", "Intern", "Completed" },
                                new List<string> { "Word Interop Client", "Intern", "In Progress" }
                            }
                        }
                    }
                };

                var templateSettings = new DocumentSettingsDto
                {
                    Font = "Times New Roman",
                    FontSize = 12
                };

                var templateMetadata = new DocumentMetadataDto
                {
                    Author = "System",
                    CreatedAt = "2026-10-07T10:00:00Z",
                    Description = "Default corporate status template",
                    Version = 1
                };

                context.Templates.Add(new TemplateEntity
                {
                    Id = "tpl-001",
                    Name = "Quarterly Status Report Template",
                    Description = "Standard corporate template with milestones and deliverables table",
                    CategoryId = reportsCategoryId,
                    DocumentStructureJson = JsonSerializer.Serialize(templateContent, JsonOptions),
                    SettingsJson = JsonSerializer.Serialize(templateSettings, JsonOptions),
                    MetadataJson = JsonSerializer.Serialize(templateMetadata, JsonOptions),
                    CreatedAt = DateTime.UtcNow
                });
            }

            // Seed default settings
            if (!await context.Settings.AnyAsync())
            {
                context.Settings.AddRange(
                    new SettingEntity { Key = "MaxFileSizeBytes", Value = "104857600", Description = "Maximum file size limit in bytes (100 MB)", UpdatedAt = DateTime.UtcNow },
                    new SettingEntity { Key = "DefaultFont", Value = "Times New Roman", Description = "Default font family for documents", UpdatedAt = DateTime.UtcNow },
                    new SettingEntity { Key = "DefaultFontSize", Value = "12", Description = "Default font size in pt", UpdatedAt = DateTime.UtcNow },
                    new SettingEntity { Key = "RequestTimeoutSeconds", Value = "30", Description = "Client request timeout in seconds", UpdatedAt = DateTime.UtcNow }
                );
            }

            // Seed initial log entry
            if (!await context.Logs.AnyAsync())
            {
                context.Logs.Add(new LogEntryEntity
                {
                    Timestamp = DateTime.UtcNow,
                    Level = "Information",
                    Message = "Database initialized and seeded with default documents, templates, and settings.",
                    Source = "DbInitializer"
                });
            }

            // Seed Users
            if (!await context.Users.AnyAsync())
            {
                context.Users.AddRange(
                    new UserEntity
                    {
                        Username = "admin",
                        FullName = "Lead Architect",
                        Email = "admin@company.local",
                        ApiKey = "secret-internship-api-key-2026",
                        PasswordHash = PasswordHasher.HashPassword("admin123"),
                        Role = "Admin",
                        CreatedAt = DateTime.UtcNow,
                        IsActive = true
                    },
                    new UserEntity
                    {
                        Username = "ivan_intern",
                        FullName = "Ivan Developer",
                        Email = "ivan@internship.local",
                        ApiKey = "internship-key-demo",
                        PasswordHash = PasswordHasher.HashPassword("ivan123"),
                        Role = "Intern",
                        CreatedAt = DateTime.UtcNow,
                        IsActive = true
                    }
                );
            }
            else
            {
                // Backfill password hashes for existing databases upgraded from older schema
                var usersWithoutPassword = await context.Users
                    .Where(u => string.IsNullOrEmpty(u.PasswordHash))
                    .ToListAsync();

                foreach (var u in usersWithoutPassword)
                {
                    u.PasswordHash = u.Username switch
                    {
                        "admin" => PasswordHasher.HashPassword("admin123"),
                        "ivan_intern" => PasswordHasher.HashPassword("ivan123"),
                        _ => PasswordHasher.HashPassword("123456")
                    };
                }
            }

            // Template Categories already seeded above before Templates

            // Seed Document Version for doc-001
            if (!await context.DocumentVersions.AnyAsync(v => v.DocumentId == "doc-001"))
            {
                context.DocumentVersions.Add(new DocumentVersionEntity
                {
                    DocumentId = "doc-001",
                    VersionNumber = 1,
                    Title = "Report v1.0",
                    ContentJson = "{\"status\": \"Initial approved release\"}",
                    Author = "Ivan",
                    ChangeNotes = "Initial generation from internship template",
                    CreatedAt = DateTime.UtcNow
                });
            }

            // Seed Document Comment for doc-001
            if (!await context.DocumentComments.AnyAsync(c => c.DocumentId == "doc-001"))
            {
                context.DocumentComments.Add(new DocumentCommentEntity
                {
                    DocumentId = "doc-001",
                    Author = "Mentor",
                    Comment = "Отличная структура отчета, все стили и таблица проверены.",
                    Resolved = true,
                    CreatedAt = DateTime.UtcNow
                });
            }

            // Seed Processing History
            if (!await context.ProcessingHistory.AnyAsync())
            {
                context.ProcessingHistory.Add(new ProcessingHistoryEntity
                {
                    DocumentId = "doc-001",
                    Operation = "uppercase",
                    InputLength = 120,
                    ResultLength = 120,
                    ClientIp = "127.0.0.1",
                    Timestamp = DateTime.UtcNow
                });
            }

            await context.SaveChangesAsync();
        }

        private static async Task EnsurePasswordHashColumnExistsAsync(AppDbContext context)
        {
            var provider = context.Database.ProviderName ?? string.Empty;
            try
            {
                if (provider.Contains("Sqlite", StringComparison.OrdinalIgnoreCase))
                {
                    var conn = context.Database.GetDbConnection();
                    if (conn.State != System.Data.ConnectionState.Open)
                        await conn.OpenAsync();

                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = "PRAGMA table_info('Users');";
                    bool hasColumn = false;
                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            if (string.Equals(reader["name"]?.ToString(), "PasswordHash", StringComparison.OrdinalIgnoreCase))
                            {
                                hasColumn = true;
                                break;
                            }
                        }
                    }

                    if (!hasColumn)
                    {
                        using var alter = conn.CreateCommand();
                        alter.CommandText = "ALTER TABLE \"Users\" ADD COLUMN \"PasswordHash\" TEXT NOT NULL DEFAULT '';";
                        await alter.ExecuteNonQueryAsync();
                    }
                }
                else if (provider.Contains("Npgsql", StringComparison.OrdinalIgnoreCase))
                {
                    await context.Database.ExecuteSqlRawAsync(
                        "ALTER TABLE \"Users\" ADD COLUMN IF NOT EXISTS \"PasswordHash\" character varying(500) NOT NULL DEFAULT '';");
                }
                else if (provider.Contains("SqlServer", StringComparison.OrdinalIgnoreCase))
                {
                    await context.Database.ExecuteSqlRawAsync(
                        "IF COL_LENGTH('Users', 'PasswordHash') IS NULL ALTER TABLE [Users] ADD [PasswordHash] nvarchar(500) NOT NULL DEFAULT N'';");
                }
            }
            catch
            {
                // Ignore if already exists or non-applicable
            }
        }
    }
}
