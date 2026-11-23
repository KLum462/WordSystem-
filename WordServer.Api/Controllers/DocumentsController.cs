using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using WordClient.Core.Models;
using WordServer.Api.Data;
using WordServer.Api.Data.Entities;
using WordServer.Api.Middleware;
using WordServer.Api.Services;

namespace WordServer.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentsController : ControllerBase
    {
        private const long DefaultMaxAllowedSizeBytes = 104_857_600; // 100 MB
        private readonly AppDbContext _db;
        private readonly TextProcessingService _textService;
        private readonly IWebHostEnvironment _environment;
        private readonly IConfiguration _configuration;
        private readonly ILogger<DocumentsController> _logger;
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        public DocumentsController(
            AppDbContext db,
            TextProcessingService textService,
            IWebHostEnvironment environment,
            IConfiguration configuration,
            ILogger<DocumentsController> logger)
        {
            _db = db;
            _textService = textService;
            _environment = environment;
            _configuration = configuration;
            _logger = logger;
        }

        private long MaxAllowedSizeBytes =>
            _configuration.GetValue<long?>("Storage:MaxFileSizeBytes") ?? DefaultMaxAllowedSizeBytes;

        private string GetUploadsDirectory()
        {
            var configuredPath = _configuration.GetValue<string>("Storage:UploadsPath") ?? Path.Combine("App_Data", "uploads");
            var fullPath = Path.IsPathRooted(configuredPath)
                ? configuredPath
                : Path.Combine(_environment.ContentRootPath, configuredPath);
            Directory.CreateDirectory(fullPath);
            return fullPath;
        }

        [HttpGet]
        public async Task<IActionResult> GetAllDocuments()
        {
            var docs = await _db.Documents
                .OrderByDescending(d => d.UpdatedAt ?? d.CreatedAt)
                .ToListAsync();

            var summaries = docs.Select(d =>
            {
                var meta = JsonSerializer.Deserialize<DocumentMetadataDto>(d.MetadataJson, JsonOptions);
                return new DocumentSummaryDto
                {
                    Id = d.Id,
                    Title = d.Title,
                    Author = meta?.Author ?? "Unknown",
                    Version = meta?.Version ?? 1,
                    CreatedAt = d.CreatedAt,
                    UpdatedAt = d.UpdatedAt
                };
            }).ToList();

            return Ok(summaries);
        }

        [HttpGet("{id}")]
        public async Task<IActionResult> GetDocument(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return BadRequest(new ErrorResponse { Error = "Document ID cannot be empty." });
            }

            var entity = await _db.Documents.FirstOrDefaultAsync(d => d.Id == id);
            if (entity == null)
            {
                return NotFound(new ErrorResponse { Error = $"Document '{id}' was not found." });
            }

            var doc = new DocumentDto
            {
                Id = entity.Id,
                Title = entity.Title,
                Metadata = JsonSerializer.Deserialize<DocumentMetadataDto>(entity.MetadataJson, JsonOptions) ?? new DocumentMetadataDto(),
                Content = JsonSerializer.Deserialize<DocumentContentDto>(entity.ContentJson, JsonOptions) ?? new DocumentContentDto(),
                Settings = JsonSerializer.Deserialize<DocumentSettingsDto>(entity.SettingsJson, JsonOptions) ?? new DocumentSettingsDto()
            };

            return Ok(doc);
        }

        [HttpPost("process")]
        [RequestSizeLimit(DefaultMaxAllowedSizeBytes)]
        public async Task<IActionResult> ProcessDocument([FromBody] DocumentProcessRequest request)
        {
            if (Request.ContentLength.HasValue && Request.ContentLength.Value > MaxAllowedSizeBytes)
            {
                return StatusCode(StatusCodes.Status413PayloadTooLarge, new ErrorResponse
                {
                    Error = $"Request payload size ({Request.ContentLength.Value / 1024.0 / 1024.0:F2} MB) exceeds maximum allowed limit of {MaxAllowedSizeBytes / 1024.0 / 1024.0:F0} MB."
                });
            }

            if (request == null)
            {
                return BadRequest(new ErrorResponse { Error = "Request body cannot be null." });
            }

            if (string.IsNullOrEmpty(request.Text))
            {
                return BadRequest(new ErrorResponse { Error = "Document text cannot be empty." });
            }

            long byteSize = Encoding.UTF8.GetByteCount(request.Text);
            if (byteSize > MaxAllowedSizeBytes)
            {
                return StatusCode(StatusCodes.Status413PayloadTooLarge, new ErrorResponse
                {
                    Error = $"Document payload size ({byteSize / 1024.0 / 1024.0:F2} MB) exceeds maximum allowed limit of {MaxAllowedSizeBytes / 1024.0 / 1024.0:F0} MB."
                });
            }

            var operation = request.Operation?.Trim().ToLowerInvariant() ?? "uppercase";
            if (!_textService.TryProcess(request.Text, operation, out var processedText, out var error))
            {
                return BadRequest(new ErrorResponse { Error = error });
            }

            var docId = string.IsNullOrWhiteSpace(request.Id) ? "doc-proc" : request.Id.Trim();
            _db.ProcessingHistory.Add(new ProcessingHistoryEntity
            {
                DocumentId = docId,
                Operation = operation,
                InputLength = request.Text.Length,
                ResultLength = processedText.Length,
                ClientIp = HttpContext.Connection.RemoteIpAddress?.ToString(),
                Timestamp = DateTime.UtcNow
            });
            await _db.SaveChangesAsync();

            var paragraphs = request.Text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var words = request.Text.Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);

            return Ok(new DocumentProcessResponse
            {
                Id = docId,
                Result = processedText,
                CharactersCount = request.Text.Length,
                WordsCount = words.Length,
                ParagraphsCount = paragraphs.Length
            });
        }

        [HttpPost("upload")]
        [RequestSizeLimit(DefaultMaxAllowedSizeBytes)]
        [RequestFormLimits(MultipartBodyLengthLimit = DefaultMaxAllowedSizeBytes)]
        public async Task<IActionResult> UploadDocument(IFormFile? file)
        {
            if (Request.ContentLength.HasValue && Request.ContentLength.Value > MaxAllowedSizeBytes)
            {
                return StatusCode(StatusCodes.Status413PayloadTooLarge, new ErrorResponse
                {
                    Error = $"Uploaded file size ({Request.ContentLength.Value / 1024.0 / 1024.0:F2} MB) exceeds maximum allowed limit of {MaxAllowedSizeBytes / 1024.0 / 1024.0:F0} MB."
                });
            }

            if (file == null || file.Length == 0)
            {
                return BadRequest(new ErrorResponse { Error = "A non-empty .docx file is required." });
            }

            if (file.Length > MaxAllowedSizeBytes)
            {
                return StatusCode(StatusCodes.Status413PayloadTooLarge, new ErrorResponse
                {
                    Error = $"Uploaded file '{file.FileName}' ({file.Length / 1024.0 / 1024.0:F2} MB) exceeds maximum allowed limit of {MaxAllowedSizeBytes / 1024.0 / 1024.0:F0} MB."
                });
            }

            var originalFileName = Path.GetFileName(file.FileName);
            var extension = Path.GetExtension(originalFileName);
            if (!string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new ErrorResponse { Error = "Only .docx files are supported for upload." });
            }

            var safeBaseName = new string(Path.GetFileNameWithoutExtension(originalFileName)
                .Select(ch => char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' ? ch : '_')
                .ToArray());
            if (string.IsNullOrWhiteSpace(safeBaseName))
            {
                safeBaseName = "document";
            }

            var uploadId = $"upl-{Guid.NewGuid():N}".Substring(0, 12);
            var storedFileName = $"{DateTime.UtcNow:yyyyMMddHHmmss}_{uploadId}_{safeBaseName}.docx";
            var uploadsDir = GetUploadsDirectory();
            var fullFilePath = Path.Combine(uploadsDir, storedFileName);

            await using (var stream = new FileStream(fullFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                await file.CopyToAsync(stream);
            }

            var now = DateTime.UtcNow;
            _db.ProcessingHistory.Add(new ProcessingHistoryEntity
            {
                DocumentId = uploadId,
                Operation = "upload_docx",
                InputLength = (int)Math.Min(int.MaxValue, file.Length),
                ResultLength = (int)Math.Min(int.MaxValue, file.Length),
                ClientIp = HttpContext.Connection.RemoteIpAddress?.ToString(),
                Timestamp = now
            });

            _db.Logs.Add(new LogEntryEntity
            {
                Timestamp = now,
                Level = "Information",
                Message = $"Uploaded file '{originalFileName}' ({file.Length} bytes) stored as '{storedFileName}'.",
                Source = "DocumentsController.UploadDocument"
            });

            await _db.SaveChangesAsync();

            _logger.LogInformation("Stored uploaded .docx file '{FileName}' ({SizeBytes} bytes) in App_Data/uploads as '{StoredFileName}'.",
                originalFileName, file.Length, storedFileName);

            var response = new DocumentUploadResponse
            {
                Id = uploadId,
                FileName = originalFileName,
                StoredFileName = storedFileName,
                RelativePath = $"App_Data/uploads/{storedFileName}",
                SizeBytes = file.Length,
                UploadedAt = now
            };

            return CreatedAtAction(nameof(DownloadUploadedDocument), new { fileName = storedFileName }, response);
        }

        [HttpGet("uploads")]
        public IActionResult ListUploadedDocuments()
        {
            var uploadsDir = GetUploadsDirectory();
            var files = Directory.GetFiles(uploadsDir, "*.docx")
                .Select(path => new FileInfo(path))
                .OrderByDescending(fi => fi.CreationTimeUtc)
                .Select(fi => new DocumentUploadResponse
                {
                    Id = fi.Name,
                    FileName = fi.Name,
                    StoredFileName = fi.Name,
                    RelativePath = $"App_Data/uploads/{fi.Name}",
                    SizeBytes = fi.Length,
                    UploadedAt = fi.CreationTimeUtc
                })
                .ToList();

            return Ok(files);
        }

        [HttpGet("uploads/{fileName}")]
        public IActionResult DownloadUploadedDocument(string fileName)
        {
            var safeName = Path.GetFileName(fileName);
            if (string.IsNullOrWhiteSpace(safeName) || !string.Equals(safeName, fileName, StringComparison.Ordinal))
            {
                return BadRequest(new ErrorResponse { Error = "Invalid file name." });
            }

            var uploadsDir = GetUploadsDirectory();
            var fullPath = Path.Combine(uploadsDir, safeName);
            if (!System.IO.File.Exists(fullPath))
            {
                return NotFound(new ErrorResponse { Error = $"Uploaded file '{safeName}' was not found." });
            }

            var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", safeName);
        }

        [HttpPost("check")]
        public async Task<IActionResult> CheckDocument([FromBody] DocumentCheckRequest request)
        {
            if (request == null)
            {
                return BadRequest(new ErrorResponse { Error = "Request body cannot be null." });
            }

            var result = _textService.InspectDocument(request);

            _db.ProcessingHistory.Add(new ProcessingHistoryEntity
            {
                DocumentId = string.IsNullOrWhiteSpace(request.DocumentId) ? "active-doc" : request.DocumentId.Trim(),
                Operation = "check",
                InputLength = request.Text?.Length ?? 0,
                ResultLength = result.Summary.Length,
                ClientIp = HttpContext.Connection.RemoteIpAddress?.ToString(),
                Timestamp = DateTime.UtcNow
            });
            await _db.SaveChangesAsync();

            return Ok(result);
        }

        [HttpPost]
        public async Task<IActionResult> CreateDocument([FromBody] DocumentDto doc)
        {
            if (doc == null || string.IsNullOrWhiteSpace(doc.Id) || string.IsNullOrWhiteSpace(doc.Title))
            {
                return BadRequest(new ErrorResponse { Error = "Document must have a valid Id and Title." });
            }

            if (await _db.Documents.AnyAsync(d => d.Id == doc.Id))
            {
                return Conflict(new ErrorResponse { Error = $"Document with ID '{doc.Id}' already exists." });
            }

            var authUser = HttpContext.Items.TryGetValue(ApiKeyMiddleware.AuthenticatedUserItemKey, out var uObj)
                ? uObj as UserEntity
                : null;

            var author = !string.IsNullOrWhiteSpace(doc.Metadata?.Author)
                ? doc.Metadata.Author
                : authUser?.FullName ?? authUser?.Username ?? "Unknown";

            doc.Metadata ??= new DocumentMetadataDto();
            doc.Metadata.Author = author;
            doc.Metadata.Version = doc.Metadata.Version ?? 1;
            if (string.IsNullOrWhiteSpace(doc.Metadata.CreatedAt))
            {
                doc.Metadata.CreatedAt = DateTime.UtcNow.ToString("O");
            }

            var contentJson = JsonSerializer.Serialize(doc.Content, JsonOptions);
            var entity = new DocumentEntity
            {
                Id = doc.Id,
                Title = doc.Title,
                MetadataJson = JsonSerializer.Serialize(doc.Metadata, JsonOptions),
                ContentJson = contentJson,
                SettingsJson = JsonSerializer.Serialize(doc.Settings, JsonOptions),
                CreatedAt = DateTime.UtcNow
            };

            _db.Documents.Add(entity);

            _db.Metadata.Add(new MetadataEntity
            {
                DocumentId = doc.Id,
                Author = author,
                CreatedAt = DateTime.UtcNow,
                Version = doc.Metadata.Version ?? 1,
                Description = doc.Metadata.Description
            });

            _db.DocumentVersions.Add(new DocumentVersionEntity
            {
                DocumentId = doc.Id,
                VersionNumber = doc.Metadata.Version ?? 1,
                Title = doc.Title,
                ContentJson = contentJson,
                Author = author,
                ChangeNotes = doc.Metadata.Description ?? "Initial document creation",
                CreatedAt = DateTime.UtcNow
            });

            await _db.SaveChangesAsync();

            return CreatedAtAction(nameof(GetDocument), new { id = doc.Id }, doc);
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateDocument(string id, [FromBody] DocumentDto doc)
        {
            if (string.IsNullOrWhiteSpace(id) || doc == null)
            {
                return BadRequest(new ErrorResponse { Error = "Document ID and body are required." });
            }

            var entity = await _db.Documents.FirstOrDefaultAsync(d => d.Id == id);
            if (entity == null)
            {
                return NotFound(new ErrorResponse { Error = $"Document '{id}' was not found." });
            }

            var authUser = HttpContext.Items.TryGetValue(ApiKeyMiddleware.AuthenticatedUserItemKey, out var uObj)
                ? uObj as UserEntity
                : null;

            var latestVersion = await _db.DocumentVersions
                .Where(v => v.DocumentId == id)
                .MaxAsync(v => (int?)v.VersionNumber) ?? 1;
            var nextVersion = latestVersion + 1;

            var existingMeta = JsonSerializer.Deserialize<DocumentMetadataDto>(entity.MetadataJson, JsonOptions)
                               ?? new DocumentMetadataDto();

            var author = !string.IsNullOrWhiteSpace(doc.Metadata?.Author)
                ? doc.Metadata.Author
                : authUser?.FullName ?? authUser?.Username ?? existingMeta.Author;

            var updatedMeta = new DocumentMetadataDto
            {
                Author = string.IsNullOrWhiteSpace(author) ? "Unknown" : author,
                CreatedAt = string.IsNullOrWhiteSpace(existingMeta.CreatedAt) ? entity.CreatedAt.ToString("O") : existingMeta.CreatedAt,
                Description = doc.Metadata?.Description ?? existingMeta.Description,
                Version = nextVersion
            };

            entity.Title = string.IsNullOrWhiteSpace(doc.Title) ? entity.Title : doc.Title;
            entity.MetadataJson = JsonSerializer.Serialize(updatedMeta, JsonOptions);
            if (doc.Content != null)
            {
                entity.ContentJson = JsonSerializer.Serialize(doc.Content, JsonOptions);
            }
            if (doc.Settings != null)
            {
                entity.SettingsJson = JsonSerializer.Serialize(doc.Settings, JsonOptions);
            }
            entity.UpdatedAt = DateTime.UtcNow;

            var metadataEntity = await _db.Metadata.FirstOrDefaultAsync(m => m.DocumentId == id);
            if (metadataEntity != null)
            {
                metadataEntity.Version = nextVersion;
                metadataEntity.Author = updatedMeta.Author;
                metadataEntity.Description = updatedMeta.Description;
            }

            _db.DocumentVersions.Add(new DocumentVersionEntity
            {
                DocumentId = id,
                VersionNumber = nextVersion,
                Title = entity.Title,
                ContentJson = entity.ContentJson,
                Author = updatedMeta.Author,
                ChangeNotes = doc.Metadata?.Description ?? $"Updated to version {nextVersion}",
                CreatedAt = DateTime.UtcNow
            });

            await _db.SaveChangesAsync();

            var responseDto = new DocumentDto
            {
                Id = entity.Id,
                Title = entity.Title,
                Metadata = updatedMeta,
                Content = JsonSerializer.Deserialize<DocumentContentDto>(entity.ContentJson, JsonOptions) ?? new DocumentContentDto(), 
                Settings = JsonSerializer.Deserialize<DocumentSettingsDto>(entity.SettingsJson, JsonOptions) ?? new DocumentSettingsDto()
            };

            return Ok(responseDto);
        }
    }
}
