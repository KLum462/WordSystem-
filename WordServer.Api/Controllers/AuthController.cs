using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using WordClient.Core.Models;
using WordServer.Api.Data;
using WordServer.Api.Data.Entities;
using WordServer.Api.Middleware;

namespace WordServer.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AuthController : ControllerBase
    {
        private readonly AppDbContext _db;
        private readonly IConfiguration _configuration;

        public AuthController(AppDbContext db, IConfiguration configuration)
        {
            _db = db;
            _configuration = configuration;
        }

        [HttpPost("validate")]
        public async Task<IActionResult> ValidateAuth([FromBody] AuthValidateRequest? request)
        {
            // First check authenticated user resolved by ApiKeyMiddleware
            if (HttpContext.Items.TryGetValue(ApiKeyMiddleware.AuthenticatedUserItemKey, out var userObj) &&
                userObj is UserEntity middlewareUser)
            {
                return Ok(new AuthValidateResponse
                {
                    Valid = true,
                    Message = $"API Key is valid. Access granted to {middlewareUser.FullName} ({middlewareUser.Role}).",
                    User = middlewareUser.Username,
                    FullName = middlewareUser.FullName,
                    Role = middlewareUser.Role
                });
            }

            // Fall back to resolving key from header or request body
            string? key = null;
            if (Request.Headers.TryGetValue(ApiKeyMiddleware.ApiKeyHeaderName, out var headerKeys))
            {
                key = headerKeys.FirstOrDefault()?.Trim();
            }
            if (string.IsNullOrWhiteSpace(key) && request != null)
            {
                key = request.ApiKey?.Trim();
            }

            if (!string.IsNullOrWhiteSpace(key))
            {
                var dbUser = await _db.Users
                    .AsNoTracking()
                    .FirstOrDefaultAsync(u => u.ApiKey == key);

                if (dbUser != null && dbUser.IsActive)
                {
                    return Ok(new AuthValidateResponse
                    {
                        Valid = true,
                        Message = $"API Key is valid. Access granted to {dbUser.FullName} ({dbUser.Role}).",
                        User = dbUser.Username,
                        FullName = dbUser.FullName,
                        Role = dbUser.Role
                    });
                }

                var fallbackKeys = _configuration.GetSection("Authentication:FallbackApiKeys").Get<string[]>()
                                   ?? _configuration.GetSection("Authentication:ApiKeys").Get<string[]>()
                                   ?? Array.Empty<string>();

                if (fallbackKeys.Contains(key, StringComparer.Ordinal))
                {
                    return Ok(new AuthValidateResponse
                    {
                        Valid = true,
                        Message = "API Key is valid (configured fallback key).",
                        User = "service-account",
                        FullName = "Configured Service Account",
                        Role = "Service"
                    });
                }
            }

            return Unauthorized(new AuthValidateResponse
            {
                Valid = false,
                Message = "Invalid or inactive API Key provided.",
                User = null
            });
        }
    }
}
