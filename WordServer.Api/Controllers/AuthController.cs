using System.Linq;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using WordClient.Core.Models;

namespace WordServer.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AuthController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public AuthController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpPost("validate")]
        public IActionResult ValidateAuth([FromBody] AuthValidateRequest? request)
        {
            // First check header, or fall back to body
            string? key = null;
            if (Request.Headers.TryGetValue("X-API-Key", out var headerKeys))
            {
                key = headerKeys.FirstOrDefault();
            }
            if (string.IsNullOrWhiteSpace(key) && request != null)
            {
                key = request.ApiKey;
            }

            var configuredKeys = _configuration.GetSection("Authentication:ApiKeys").Get<string[]>()
                                ?? new[] { "secret-internship-api-key-2026", "internship-key-demo" };

            if (!string.IsNullOrWhiteSpace(key) && configuredKeys.Contains(key))
            {
                return Ok(new AuthValidateResponse
                {
                    Valid = true,
                    Message = "API Key is valid. Access granted.",
                    User = "InternDeveloper"
                });
            }

            return Unauthorized(new AuthValidateResponse
            {
                Valid = false,
                Message = "Invalid API Key provided.",
                User = null
            });
        }
    }
}
