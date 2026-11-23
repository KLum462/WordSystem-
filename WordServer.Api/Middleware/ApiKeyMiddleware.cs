using System;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using WordServer.Api.Data;

namespace WordServer.Api.Middleware
{
    public class ApiKeyMiddleware
    {
        public const string ApiKeyHeaderName = "X-API-Key";
        public const string AuthenticatedUserItemKey = "AuthenticatedUser";

        private readonly RequestDelegate _next;
        private readonly ILogger<ApiKeyMiddleware> _logger;

        public ApiKeyMiddleware(RequestDelegate next, ILogger<ApiKeyMiddleware> logger)
        {
            _next = next;
            _logger = logger;
        }

        public async Task InvokeAsync(HttpContext context, IConfiguration configuration, AppDbContext db)
        {
            var path = context.Request.Path.Value?.ToLowerInvariant() ?? string.Empty;

            // Allow public paths without API key
            if (path.StartsWith("/api/health") ||
                path.StartsWith("/swagger") ||
                path.Equals("/") ||
                path.StartsWith("/favicon.ico"))
            {
                await _next(context);
                return;
            }

            if (!context.Request.Headers.TryGetValue(ApiKeyHeaderName, out var extractedApiKeyValues))
            {
                _logger.LogWarning("Missing {HeaderName} header for path {Path}", ApiKeyHeaderName, path);
                await WriteUnauthorizedAsync(context, "Unauthorized. Missing X-API-Key header.");
                return;
            }

            var extractedApiKey = extractedApiKeyValues.FirstOrDefault()?.Trim();
            if (string.IsNullOrWhiteSpace(extractedApiKey))
            {
                _logger.LogWarning("Empty {HeaderName} header for {Method} {Path}", ApiKeyHeaderName, context.Request.Method, path);
                await WriteUnauthorizedAsync(context, "Unauthorized. Empty X-API-Key header.");
                return;
            }

            // Primary authentication: validate against Users table in database
            var useDatabaseUsers = configuration.GetValue<bool>("Authentication:UseDatabaseUsers", true);
            if (useDatabaseUsers)
            {
                var user = await db.Users
                    .AsNoTracking()
                    .FirstOrDefaultAsync(u => u.ApiKey == extractedApiKey);

                if (user != null)
                {
                    if (!user.IsActive)
                    {
                        _logger.LogWarning("Deactivated user '{Username}' attempted access to {Method} {Path}", user.Username, context.Request.Method, path);
                        await WriteUnauthorizedAsync(context, "Unauthorized. User account is deactivated.");
                        return;
                    }

                    _logger.LogInformation("Authenticated user '{Username}' ({Role}) for {Method} {Path}", user.Username, user.Role, context.Request.Method, path);
                    context.Items[AuthenticatedUserItemKey] = user;
                    await _next(context);
                    return;
                }
            }

            // Optional fallback keys from configuration (empty by default)
            var fallbackKeys = configuration.GetSection("Authentication:FallbackApiKeys").Get<string[]>()
                               ?? configuration.GetSection("Authentication:ApiKeys").Get<string[]>()
                               ?? Array.Empty<string>();

            if (fallbackKeys.Contains(extractedApiKey, StringComparer.Ordinal))
            {
                _logger.LogInformation("Authenticated configured fallback API key ({MaskedKey}) for {Method} {Path}", MaskKey(extractedApiKey), context.Request.Method, path);
                await _next(context);
                return;
            }

            _logger.LogWarning("Invalid {HeaderName} ({MaskedKey}) for path {Path}", ApiKeyHeaderName, MaskKey(extractedApiKey), path);
            await WriteUnauthorizedAsync(context, "Unauthorized. Invalid X-API-Key header.");
        }

        private static string MaskKey(string key)
        {
            if (key.Length <= 4) return "****";
            return key.Substring(0, 4) + new string('*', Math.Min(8, key.Length - 4));
        }

        private static async Task WriteUnauthorizedAsync(HttpContext context, string message)
        {
            context.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
            context.Response.ContentType = "application/json";
            await context.Response.WriteAsync(JsonSerializer.Serialize(new { error = message }));
        }
    }
}
