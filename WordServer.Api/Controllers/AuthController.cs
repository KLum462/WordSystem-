using System;
using System.Linq;
using System.Threading.Tasks;
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
    public class AuthController : ControllerBase
    {
        private readonly AppDbContext _db;
        private readonly IConfiguration _configuration;
        private readonly ILogger<AuthController> _logger;

        public AuthController(AppDbContext db, IConfiguration configuration, ILogger<AuthController> logger)
        {
            _db = db;
            _configuration = configuration;
            _logger = logger;
        }

        [HttpPost("login")]
        public async Task<IActionResult> Login([FromBody] LoginRequest? request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Username) || string.IsNullOrEmpty(request.Password))
            {
                return BadRequest(new AuthValidateResponse
                {
                    Valid = false,
                    Message = "Необходимо указать логин и пароль."
                });
            }

            var username = request.Username.Trim();
            var user = await _db.Users.FirstOrDefaultAsync(u => u.Username == username);

            if (user == null || !PasswordHasher.VerifyPassword(request.Password, user.PasswordHash))
            {
                _logger.LogWarning("Failed login attempt for username '{Username}'.", username);
                return Unauthorized(new AuthValidateResponse
                {
                    Valid = false,
                    Message = "Неверный логин или пароль."
                });
            }

            if (!user.IsActive)
            {
                _logger.LogWarning("Deactivated user '{Username}' attempted login.", username);
                return Unauthorized(new AuthValidateResponse
                {
                    Valid = false,
                    Message = "Учётная запись отключена администратором."
                });
            }

            _logger.LogInformation("User '{Username}' ({Role}) logged in via username/password.", user.Username, user.Role);
            return Ok(new AuthValidateResponse
            {
                Valid = true,
                Message = $"Вход выполнен: {user.FullName} ({user.Role}).",
                User = user.Username,
                FullName = user.FullName,
                Role = user.Role,
                ApiKey = user.ApiKey
            });
        }

        [HttpPost("change-password")]
        public async Task<IActionResult> ChangePassword([FromBody] ChangePasswordRequest? request)
        {
            if (request == null ||
                string.IsNullOrWhiteSpace(request.Username) ||
                string.IsNullOrEmpty(request.CurrentPassword) ||
                string.IsNullOrWhiteSpace(request.NewPassword))
            {
                return BadRequest(new ErrorResponse { Error = "Необходимо указать логин, текущий пароль и новый пароль." });
            }

            if (request.NewPassword.Length < 4)
            {
                return BadRequest(new ErrorResponse { Error = "Новый пароль должен содержать не менее 4 символов." });
            }

            var username = request.Username.Trim();
            var user = await _db.Users.FirstOrDefaultAsync(u => u.Username == username);
            if (user == null || !PasswordHasher.VerifyPassword(request.CurrentPassword, user.PasswordHash))
            {
                _logger.LogWarning("Failed password change attempt for username '{Username}'.", username);
                return Unauthorized(new ErrorResponse { Error = "Неверный логин или текущий пароль." });
            }

            if (!user.IsActive)
            {
                return Unauthorized(new ErrorResponse { Error = "Учётная запись отключена администратором." });
            }

            user.PasswordHash = PasswordHasher.HashPassword(request.NewPassword);
            await _db.SaveChangesAsync();

            _logger.LogInformation("Password changed successfully for user '{Username}'.", username);
            return Ok(new { message = $"Пароль для пользователя '{user.Username}' успешно изменён." });
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
