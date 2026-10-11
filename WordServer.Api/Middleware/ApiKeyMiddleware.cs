using System;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace WordServer.Api.Middleware
{
    public class ApiKeyMiddleware
    {
        private const string ApiKeyHeaderName = "X-API-Key";
        private readonly RequestDelegate _next;
        private readonly ILogger<ApiKeyMiddleware> _logger;

        public ApiKeyMiddleware(RequestDelegate _next, ILogger<ApiKeyMiddleware> logger)
        {
            this._next = _next;
            _logger = logger;
        }

        public async Task InvokeAsync(HttpContext context, IConfiguration configuration)
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

            var configuredKeys = configuration.GetSection("Authentication:ApiKeys").Get<string[]>()
                                ?? new[] { "secret-internship-api-key-2026", "internship-key-demo" };

            if (!context.Request.Headers.TryGetValue(ApiKeyHeaderName, out var extractedApiKeyValues))
            {
                _logger.LogWarning("Missing {HeaderName} header for path {Path}", ApiKeyHeaderName, path);
                context.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                context.Response.ContentType = "application/json";
                await context.Response.WriteAsync(JsonSerializer.Serialize(new
                {
                    error = "Unauthorized. Missing X-API-Key header."
                }));
                return;
            }

            var extractedApiKey = extractedApiKeyValues.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(extractedApiKey) || !configuredKeys.Contains(extractedApiKey, StringComparer.Ordinal))
            {
                _logger.LogWarning("Invalid {HeaderName} header: {Key} for path {Path}", ApiKeyHeaderName, extractedApiKey, path);
                context.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                context.Response.ContentType = "application/json";
                await context.Response.WriteAsync(JsonSerializer.Serialize(new
                {
                    error = "Unauthorized. Invalid X-API-Key header."
                }));
                return;
            }

            await _next(context);
        }
    }
}
