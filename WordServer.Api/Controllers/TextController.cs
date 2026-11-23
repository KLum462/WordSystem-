using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using WordClient.Core.Models;
using WordServer.Api.Data;
using WordServer.Api.Data.Entities;
using WordServer.Api.Services;

namespace WordServer.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class TextController : ControllerBase
    {
        private readonly AppDbContext _db;
        private readonly TextProcessingService _textService;

        public TextController(AppDbContext db, TextProcessingService textService)
        {
            _db = db;
            _textService = textService;
        }

        [HttpGet("operations")]
        public IActionResult GetSupportedOperations()
        {
            return Ok(TextProcessingService.SupportedOperations);
        }

        [HttpPost("process")]
        public async Task<IActionResult> ProcessText([FromBody] TextProcessRequest request)
        {
            if (request == null || request.Text == null)
            {
                return BadRequest(new ErrorResponse { Error = "Text cannot be null." });
            }

            var operation = request.Operation?.Trim().ToLowerInvariant() ?? "uppercase";
            if (!_textService.TryProcess(request.Text, operation, out var result, out var error))
            {
                return BadRequest(new ErrorResponse { Error = error });
            }

            var historyEntry = new ProcessingHistoryEntity
            {
                DocumentId = string.IsNullOrWhiteSpace(request.DocumentId) ? "selected-text" : request.DocumentId.Trim(),
                Operation = operation,
                InputLength = request.Text.Length,
                ResultLength = result.Length,
                ClientIp = HttpContext.Connection.RemoteIpAddress?.ToString(),
                Timestamp = DateTime.UtcNow
            };

            _db.ProcessingHistory.Add(historyEntry);
            await _db.SaveChangesAsync();

            return Ok(new TextProcessResponse
            {
                Id = historyEntry.Id,
                Result = result
            });
        }
    }
}
