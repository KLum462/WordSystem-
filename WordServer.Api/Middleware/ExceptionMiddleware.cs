using System;
using System.IO;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using WordServer.Api.Data;
using WordServer.Api.Data.Entities;

namespace WordServer.Api.Middleware
{
    public class ExceptionMiddleware
    {
        private readonly RequestDelegate _next;
        private readonly ILogger<ExceptionMiddleware> _logger;

        public ExceptionMiddleware(RequestDelegate next, ILogger<ExceptionMiddleware> logger)
        {
            _next = next;
            _logger = logger;
        }

        public async Task InvokeAsync(HttpContext context)
        {
            try
            {
                await _next(context);
            }
            catch (BadHttpRequestException badHttpEx)
            {
                _logger.LogWarning(badHttpEx, "Bad HTTP request: {Message}", badHttpEx.Message);

                int statusCode = badHttpEx.StatusCode != 0 ? badHttpEx.StatusCode : (int)HttpStatusCode.BadRequest;
                if (badHttpEx.Message.Contains("exceeds the maximum allow", StringComparison.OrdinalIgnoreCase) ||
                    badHttpEx.StatusCode == (int)HttpStatusCode.RequestEntityTooLarge)
                {
                    statusCode = (int)HttpStatusCode.RequestEntityTooLarge; // 413
                }

                await WriteErrorResponseAsync(context, statusCode, badHttpEx.Message);
            }
            catch (JsonException jsonEx)
            {
                _logger.LogWarning(jsonEx, "Malformed JSON received: {Message}", jsonEx.Message);
                await WriteErrorResponseAsync(context, (int)HttpStatusCode.BadRequest, $"Malformed JSON payload: {jsonEx.Message}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unhandled server error: {Message}", ex.Message);

                // Save to database Logs table if possible
                try
                {
                    using var scope = context.RequestServices.CreateScope();
                    var db = scope.ServiceProvider.GetService<AppDbContext>();
                    if (db != null)
                    {
                        db.Logs.Add(new LogEntryEntity
                        {
                            Timestamp = DateTime.UtcNow,
                            Level = "Error",
                            Message = ex.Message,
                            Source = context.Request.Path.Value,
                            Exception = ex.ToString()
                        });
                        await db.SaveChangesAsync();
                    }
                }
                catch
                {
                    // Prevent logging failure from masking the error
                }

                await WriteErrorResponseAsync(context, (int)HttpStatusCode.InternalServerError, "An internal server error occurred.");
            }
        }

        private static async Task WriteErrorResponseAsync(HttpContext context, int statusCode, string message)
        {
            if (context.Response.HasStarted)
                return;

            context.Response.StatusCode = statusCode;
            context.Response.ContentType = "application/json";

            var response = new { error = message };
            await context.Response.WriteAsync(JsonSerializer.Serialize(response));
        }
    }
}
