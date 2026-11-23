using System;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WordClient.Core.Models;
using WordServer.Api.Data;
using WordServer.Api.Data.Entities;
using WordServer.Api.Middleware;

namespace WordServer.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ManagementController : ControllerBase
    {
        private readonly AppDbContext _db;

        public ManagementController(AppDbContext db)
        {
            _db = db;
        }

        [HttpGet("users")]
        public async Task<IActionResult> GetUsers()
        {
            var users = await _db.Users
                .OrderBy(u => u.Id)
                .Select(u => new
                {
                    u.Id,
                    u.Username,
                    u.FullName,
                    u.Email,
                    u.Role,
                    u.IsActive,
                    u.CreatedAt
                })
                .ToListAsync();

            return Ok(users);
        }

        [HttpPost("users")]
        public async Task<IActionResult> CreateUser([FromBody] CreateUserRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Username))
            {
                return BadRequest(new ErrorResponse { Error = "Username is required." });
            }

            var username = request.Username.Trim();
            if (await _db.Users.AnyAsync(u => u.Username == username))
            {
                return Conflict(new ErrorResponse { Error = $"User '{username}' already exists." });
            }

            var apiKey = string.IsNullOrWhiteSpace(request.ApiKey)
                ? $"ws-key-{Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant()}"
                : request.ApiKey.Trim();

            if (await _db.Users.AnyAsync(u => u.ApiKey == apiKey))
            {
                return Conflict(new ErrorResponse { Error = "Specified API key is already assigned to another user." });
            }

            var user = new UserEntity
            {
                Username = username,
                FullName = string.IsNullOrWhiteSpace(request.FullName) ? username : request.FullName.Trim(),
                Email = request.Email?.Trim() ?? string.Empty,
                Role = string.IsNullOrWhiteSpace(request.Role) ? "User" : request.Role.Trim(),
                ApiKey = apiKey,
                IsActive = true,
                CreatedAt = DateTime.UtcNow
            };

            _db.Users.Add(user);
            await _db.SaveChangesAsync();

            return CreatedAtAction(nameof(GetUsers), new { id = user.Id }, new
            {
                user.Id,
                user.Username,
                user.FullName,
                user.Email,
                user.Role,
                user.IsActive,
                user.ApiKey,
                user.CreatedAt
            });
        }

        [HttpPatch("users/{id:int}/status")]
        public async Task<IActionResult> UpdateUserStatus(int id, [FromBody] UpdateUserStatusRequest request)
        {
            if (request == null)
            {
                return BadRequest(new ErrorResponse { Error = "Request body is required." });
            }

            var user = await _db.Users.FirstOrDefaultAsync(u => u.Id == id);
            if (user == null)
            {
                return NotFound(new ErrorResponse { Error = $"User #{id} not found." });
            }

            user.IsActive = request.IsActive;
            await _db.SaveChangesAsync();

            return Ok(new
            {
                user.Id,
                user.Username,
                user.IsActive
            });
        }

        [HttpGet("categories")]
        public async Task<IActionResult> GetCategories()
        {
            var categories = await _db.TemplateCategories
                .OrderBy(c => c.Name)
                .ToListAsync();

            return Ok(categories);
        }

        [HttpPost("categories")]
        public async Task<IActionResult> CreateCategory([FromBody] CreateCategoryRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Name))
            {
                return BadRequest(new ErrorResponse { Error = "Category name is required." });
            }

            var name = request.Name.Trim();
            if (await _db.TemplateCategories.AnyAsync(c => c.Name == name))
            {
                return Conflict(new ErrorResponse { Error = $"Category '{name}' already exists." });
            }

            var entity = new TemplateCategoryEntity
            {
                Name = name,
                Description = request.Description?.Trim(),
                CreatedAt = DateTime.UtcNow
            };

            _db.TemplateCategories.Add(entity);
            await _db.SaveChangesAsync();

            return CreatedAtAction(nameof(GetCategories), new { id = entity.Id }, entity);
        }

        [HttpGet("versions/{documentId}")]
        public async Task<IActionResult> GetDocumentVersions(string documentId)
        {
            var versions = await _db.DocumentVersions
                .Where(v => v.DocumentId == documentId)
                .OrderByDescending(v => v.VersionNumber)
                .ToListAsync();

            return Ok(versions);
        }

        [HttpGet("comments/{documentId}")]
        public async Task<IActionResult> GetDocumentComments(string documentId)
        {
            var comments = await _db.DocumentComments
                .Where(c => c.DocumentId == documentId)
                .OrderByDescending(c => c.CreatedAt)
                .ToListAsync();

            return Ok(comments);
        }

        [HttpPost("comments")]
        public async Task<IActionResult> CreateDocumentComment([FromBody] CreateCommentRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.DocumentId) || string.IsNullOrWhiteSpace(request.Comment))
            {
                return BadRequest(new ErrorResponse { Error = "DocumentId and Comment are required." });
            }

            var authUser = HttpContext.Items.TryGetValue(ApiKeyMiddleware.AuthenticatedUserItemKey, out var uObj)
                ? uObj as UserEntity
                : null;

            var comment = new DocumentCommentEntity
            {
                DocumentId = request.DocumentId.Trim(),
                Author = !string.IsNullOrWhiteSpace(request.Author)
                    ? request.Author.Trim()
                    : authUser?.FullName ?? authUser?.Username ?? "Unknown",
                Comment = request.Comment.Trim(),
                Resolved = false,
                CreatedAt = DateTime.UtcNow
            };

            _db.DocumentComments.Add(comment);
            await _db.SaveChangesAsync();

            return CreatedAtAction(nameof(GetDocumentComments), new { documentId = comment.DocumentId }, comment);
        }

        [HttpPatch("comments/{id:int}/resolve")]
        public async Task<IActionResult> ResolveDocumentComment(int id, [FromBody] ResolveCommentRequest? request)
        {
            var comment = await _db.DocumentComments.FirstOrDefaultAsync(c => c.Id == id);
            if (comment == null)
            {
                return NotFound(new ErrorResponse { Error = $"Comment #{id} not found." });
            }

            comment.Resolved = request?.Resolved ?? true;
            await _db.SaveChangesAsync();

            return Ok(comment);
        }

        [HttpGet("history")]
        public async Task<IActionResult> GetProcessingHistory([FromQuery] int limit = 50, [FromQuery] string? documentId = null)
        {
            var query = _db.ProcessingHistory.AsQueryable();
            if (!string.IsNullOrWhiteSpace(documentId))
            {
                query = query.Where(h => h.DocumentId == documentId);
            }

            var history = await query
                .OrderByDescending(h => h.Timestamp)
                .Take(Math.Clamp(limit, 1, 500))
                .ToListAsync();

            return Ok(history);
        }

        [HttpGet("settings")]
        public async Task<IActionResult> GetSettings()
        {
            var settings = await _db.Settings
                .OrderBy(s => s.Key)
                .ToListAsync();

            return Ok(settings);
        }

        [HttpPut("settings/{key}")]
        public async Task<IActionResult> UpsertSetting(string key, [FromBody] UpdateSettingRequest request)
        {
            if (string.IsNullOrWhiteSpace(key) || request == null || request.Value == null)
            {
                return BadRequest(new ErrorResponse { Error = "Setting key and value are required." });
            }

            var normalizedKey = key.Trim();
            var setting = await _db.Settings.FirstOrDefaultAsync(s => s.Key == normalizedKey);
            if (setting == null)
            {
                setting = new SettingEntity
                {
                    Key = normalizedKey,
                    Value = request.Value,
                    Description = request.Description,
                    UpdatedAt = DateTime.UtcNow
                };
                _db.Settings.Add(setting);
            }
            else
            {
                setting.Value = request.Value;
                if (request.Description != null)
                {
                    setting.Description = request.Description;
                }
                setting.UpdatedAt = DateTime.UtcNow;
            }

            await _db.SaveChangesAsync();
            return Ok(setting);
        }
    }
}
