using System;
using System.Globalization;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using WordClient.Core.Models;

namespace WordServer.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class TextController : ControllerBase
    {
        [HttpPost("process")]
        public IActionResult ProcessText([FromBody] TextProcessRequest request)
        {
            if (request == null || request.Text == null)
            {
                return BadRequest(new ErrorResponse { Error = "Text cannot be null." });
            }

            var operation = request.Operation?.Trim().ToLowerInvariant() ?? "uppercase";
            string result;

            switch (operation)
            {
                case "uppercase":
                    result = request.Text.ToUpperInvariant();
                    break;

                case "lowercase":
                    result = request.Text.ToLowerInvariant();
                    break;

                case "trim":
                    result = request.Text.Trim();
                    break;

                case "reverse":
                    var chars = request.Text.ToCharArray();
                    Array.Reverse(chars);
                    result = new string(chars);
                    break;

                case "word_count":
                    var words = request.Text.Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    result = $"Word count: {words.Length}";
                    break;

                case "title_case":
                    result = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(request.Text.ToLowerInvariant());
                    break;

                case "capitalize":
                    if (request.Text.Length > 0)
                    {
                        result = char.ToUpperInvariant(request.Text[0]) + request.Text.Substring(1);
                    }
                    else
                    {
                        result = string.Empty;
                    }
                    break;

                default:
                    return BadRequest(new ErrorResponse
                    {
                        Error = $"Unsupported operation '{request.Operation}'. Supported operations: uppercase, lowercase, trim, reverse, word_count, title_case, capitalize."
                    });
            }

            return Ok(new TextProcessResponse
            {
                Id = 0,
                Result = result
            });
        }
    }
}
