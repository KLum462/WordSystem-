using System;
using System.Collections.Generic;
using System.Linq;
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
    [Route("api/template")]
    public class TemplatesController : ControllerBase
    {
        private readonly AppDbContext _db;
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        public TemplatesController(AppDbContext db)
        {
            _db = db;
        }

        [HttpGet]
        public async Task<IActionResult> GetAllTemplates([FromQuery] int? categoryId = null)
        {
            var query = _db.Templates.AsQueryable();
            if (categoryId.HasValue)
            {
                query = query.Where(t => t.CategoryId == categoryId.Value);
            }

            var entities = await query.OrderByDescending(t => t.CreatedAt).ToListAsync();
            var dtos = entities.Select(MapToDto).ToList();
            return Ok(dtos);
        }

        [HttpGet("{id}")]
        public async Task<IActionResult> GetTemplateById(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return BadRequest(new ErrorResponse { Error = "Template ID cannot be empty." });
            }

            var entity = await _db.Templates.FirstOrDefaultAsync(t => t.Id == id);
            if (entity == null)
            {
                return NotFound(new ErrorResponse { Error = $"Template '{id}' not found." });
            }

            return Ok(MapToDto(entity));
        }

        [HttpPost]
        public async Task<IActionResult> CreateTemplate([FromBody] TemplateDto dto)
        {
            if (dto == null)
            {
                return BadRequest(new ErrorResponse { Error = "Template payload cannot be null." });
            }

            if (string.IsNullOrWhiteSpace(dto.Id))
            {
                dto.Id = $"tpl-{Guid.NewGuid().ToString("N").Substring(0, 8)}";
            }

            if (string.IsNullOrWhiteSpace(dto.Name))
            {
                return BadRequest(new ErrorResponse { Error = "Template Name is required." });
            }

            if (await _db.Templates.AnyAsync(t => t.Id == dto.Id))
            {
                return Conflict(new ErrorResponse { Error = $"Template with ID '{dto.Id}' already exists." });
            }

            if (dto.CategoryId.HasValue && !await _db.TemplateCategories.AnyAsync(c => c.Id == dto.CategoryId.Value))
            {
                return BadRequest(new ErrorResponse { Error = $"Template category #{dto.CategoryId.Value} does not exist." });
            }

            var entity = new TemplateEntity
            {
                Id = dto.Id,
                Name = dto.Name,
                Description = dto.Description,
                CategoryId = dto.CategoryId,
                DocumentStructureJson = JsonSerializer.Serialize(dto.Content, JsonOptions),
                SettingsJson = JsonSerializer.Serialize(dto.Settings, JsonOptions),
                MetadataJson = JsonSerializer.Serialize(dto.Metadata, JsonOptions),
                CreatedAt = DateTime.UtcNow
            };

            _db.Templates.Add(entity);
            await _db.SaveChangesAsync();

            return CreatedAtAction(nameof(GetTemplateById), new { id = dto.Id }, MapToDto(entity));
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateTemplate(string id, [FromBody] TemplateDto dto)
        {
            if (string.IsNullOrWhiteSpace(id) || dto == null)
            {
                return BadRequest(new ErrorResponse { Error = "Invalid template update parameters." });
            }

            var entity = await _db.Templates.FirstOrDefaultAsync(t => t.Id == id);
            if (entity == null)
            {
                return NotFound(new ErrorResponse { Error = $"Template '{id}' not found." });
            }

            entity.Name = string.IsNullOrWhiteSpace(dto.Name) ? entity.Name : dto.Name;
            entity.Description = dto.Description ?? entity.Description;
            if (dto.CategoryId.HasValue)
            {
                if (!await _db.TemplateCategories.AnyAsync(c => c.Id == dto.CategoryId.Value))
                {
                    return BadRequest(new ErrorResponse { Error = $"Template category #{dto.CategoryId.Value} does not exist." });
                }
                entity.CategoryId = dto.CategoryId;
            }
            if (dto.Content != null)
            {
                entity.DocumentStructureJson = JsonSerializer.Serialize(dto.Content, JsonOptions);
            }
            if (dto.Settings != null)
            {
                entity.SettingsJson = JsonSerializer.Serialize(dto.Settings, JsonOptions);
            }
            if (dto.Metadata != null)
            {
                entity.MetadataJson = JsonSerializer.Serialize(dto.Metadata, JsonOptions);
            }
            entity.UpdatedAt = DateTime.UtcNow;

            await _db.SaveChangesAsync();

            return Ok(MapToDto(entity));
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteTemplate(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return BadRequest(new ErrorResponse { Error = "Template ID cannot be empty." });
            }

            var entity = await _db.Templates.FirstOrDefaultAsync(t => t.Id == id);
            if (entity == null)
            {
                return NotFound(new ErrorResponse { Error = $"Template '{id}' not found." });
            }

            _db.Templates.Remove(entity);
            await _db.SaveChangesAsync();

            return NoContent();
        }

        private static TemplateDto MapToDto(TemplateEntity entity)
        {
            return new TemplateDto
            {
                Id = entity.Id,
                Name = entity.Name,
                Description = entity.Description,
                CategoryId = entity.CategoryId,
                Content = JsonSerializer.Deserialize<DocumentContentDto>(entity.DocumentStructureJson, JsonOptions) ?? new DocumentContentDto(),
                Settings = JsonSerializer.Deserialize<DocumentSettingsDto>(entity.SettingsJson, JsonOptions) ?? new DocumentSettingsDto(),
                Metadata = JsonSerializer.Deserialize<DocumentMetadataDto>(entity.MetadataJson, JsonOptions) ?? new DocumentMetadataDto(),
                CreatedAt = entity.CreatedAt,
                UpdatedAt = entity.UpdatedAt
            };
        }
    }
}
