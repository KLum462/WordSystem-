using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WordClient.Core.Models;
using WordServer.Api.Data;
using WordServer.Api.Data.Entities;

namespace WordServer.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentsController : ControllerBase
    {
        private const int MaxAllowedSizeBytes = 104857600; // 100 MB
        private readonly AppDbContext _db;
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        public DocumentsController(AppDbContext db)
        {
            _db = db;
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
        [RequestSizeLimit(MaxAllowedSizeBytes)]
        public IActionResult ProcessDocument([FromBody] DocumentProcessRequest request)
        {
            if (request == null)
            {
                return BadRequest(new ErrorResponse { Error = "Request body cannot be null." });
            }

            if (string.IsNullOrEmpty(request.Text))
            {
                return BadRequest(new ErrorResponse { Error = "Document text cannot be empty." });
            }

            int byteSize = Encoding.UTF8.GetByteCount(request.Text);
            if (byteSize > MaxAllowedSizeBytes)
            {
                return StatusCode(413, new ErrorResponse
                {
                    Error = $"Document payload size ({byteSize / 1024 / 1024} MB) exceeds maximum allowed limit of 100 MB."
                });
            }

            var operation = request.Operation?.Trim().ToLowerInvariant() ?? "uppercase";
            string processedText;

            switch (operation)
            {
                case "uppercase":
                    processedText = request.Text.ToUpperInvariant();
                    break;
                case "lowercase":
                    processedText = request.Text.ToLowerInvariant();
                    break;
                case "trim":
                    processedText = request.Text.Trim();
                    break;
                case "reverse":
                    var chars = request.Text.ToCharArray();
                    Array.Reverse(chars);
                    processedText = new string(chars);
                    break;
                default:
                    processedText = request.Text.ToUpperInvariant();
                    break;
            }

            var paragraphs = request.Text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var words = request.Text.Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);

            return Ok(new DocumentProcessResponse
            {
                Id = request.Id ?? "doc-proc",
                Result = processedText,
                CharactersCount = request.Text.Length,
                WordsCount = words.Length,
                ParagraphsCount = paragraphs.Length
            });
        }

        [HttpPost("check")]
        public IActionResult CheckDocument([FromBody] DocumentCheckRequest request)
        {
            if (request == null)
            {
                return BadRequest(new ErrorResponse { Error = "Request body cannot be null." });
            }

            var text = request.Text ?? string.Empty;
            var paragraphsList = request.Paragraphs ?? text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).ToList();
            var words = text.Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);

            int emptyParagraphs = paragraphsList.Count(p => string.IsNullOrWhiteSpace(p));
            var issues = new List<string>();

            if (emptyParagraphs > 0)
            {
                issues.Add($"Found {emptyParagraphs} empty paragraph(s) in document.");
            }

            if (words.Length == 0)
            {
                issues.Add("Document contains no readable words.");
            }

            var result = new DocumentCheckResult
            {
                WordCount = words.Length,
                ParagraphCount = paragraphsList.Count,
                EmptyParagraphCount = emptyParagraphs,
                TableCount = request.TableCount ?? 0,
                CharacterCount = text.Length,
                Issues = issues,
                Summary = $"Document inspection complete: {words.Length} words, {paragraphsList.Count} paragraphs ({emptyParagraphs} empty), {request.TableCount ?? 0} tables."
            };

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

            var entity = new DocumentEntity
            {
                Id = doc.Id,
                Title = doc.Title,
                MetadataJson = JsonSerializer.Serialize(doc.Metadata, JsonOptions),
                ContentJson = JsonSerializer.Serialize(doc.Content, JsonOptions),
                SettingsJson = JsonSerializer.Serialize(doc.Settings, JsonOptions),
                CreatedAt = DateTime.UtcNow
            };

            _db.Documents.Add(entity);

            _db.Metadata.Add(new MetadataEntity
            {
                DocumentId = doc.Id,
                Author = doc.Metadata?.Author ?? "Unknown",
                CreatedAt = DateTime.UtcNow,
                Version = doc.Metadata?.Version ?? 1,
                Description = doc.Metadata?.Description
            });

            await _db.SaveChangesAsync();

            return CreatedAtAction(nameof(GetDocument), new { id = doc.Id }, doc);
        }
    }
}
