using Microsoft.AspNetCore.Mvc;
using WordClient.Core.Models;
using WordClient.Core.Operations;
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
            return Ok(TextOperations.OperationsNames());
        }

        [HttpPost("process")]
        public async Task<IActionResult> ProcessText([FromBody] TextProcessRequest request)
        {
            if (request == null || request.Text == null)
            {
                return BadRequest(new ErrorResponse { Error = "Text cannot be null." });
            }

            var op = TextProcessingService.GetOperation(request);
            var result = op.Process(request.Text);
            var historyEntry = new ProcessingHistoryEntity
            {
                DocumentId = string.IsNullOrWhiteSpace(request.DocumentId) ? "selected-text" : request.DocumentId.Trim(),
                Operation = op.Name(),
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
