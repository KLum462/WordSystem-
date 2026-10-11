using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WordServer.Api.Data;
using WordServer.Api.Data.Entities;

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

        [HttpGet("categories")]
        public async Task<IActionResult> GetCategories()
        {
            var categories = await _db.TemplateCategories
                .OrderBy(c => c.Name)
                .ToListAsync();

            return Ok(categories);
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

        [HttpGet("history")]
        public async Task<IActionResult> GetProcessingHistory([FromQuery] int limit = 50)
        {
            var history = await _db.ProcessingHistory
                .OrderByDescending(h => h.Timestamp)
                .Take(limit)
                .ToListAsync();

            return Ok(history);
        }
    }
}
