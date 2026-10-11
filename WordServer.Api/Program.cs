using System;
using System.IO;
using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.OpenApi.Models;
using Serilog;
using WordServer.Api.Data;
using WordServer.Api.Middleware;

var builder = WebApplication.CreateBuilder(args);

// Configure Serilog for file logging and console
var logsDir = Path.Combine(AppContext.BaseDirectory, "logs");
if (!Directory.Exists(logsDir))
{
    Directory.CreateDirectory(logsDir);
}

Log.Logger = new LoggerConfiguration()
    .ReadFrom.Configuration(builder.Configuration)
    .Enrich.FromLogContext()
    .WriteTo.Console()
    .WriteTo.File(Path.Combine(logsDir, "server-.log"),
        rollingInterval: RollingInterval.Day,
        retainedFileCountLimit: 14)
    .CreateLogger();

builder.Host.UseSerilog();

// Configure 100 MB max request body size
const long maxBodySizeBytes = 104_857_600; // 100 MB
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = maxBodySizeBytes;
});
builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = maxBodySizeBytes;
});

// Configure Database (SQL Server 2022 default, SQLite supported)
var dbProvider = builder.Configuration.GetValue<string>("DatabaseProvider") ?? "SqlServer";
var sqlServerConnection = builder.Configuration.GetConnectionString("DefaultConnection");
var sqliteConnection = builder.Configuration.GetConnectionString("SqliteConnection") ?? "Data Source=WordSystem.db";

builder.Services.AddDbContext<AppDbContext>(options =>
{
    if (dbProvider.Equals("Sqlite", StringComparison.OrdinalIgnoreCase))
    {
        options.UseSqlite(sqliteConnection);
    }
    else
    {
        options.UseSqlServer(sqlServerConnection, sqlOptions =>
        {
            sqlOptions.EnableRetryOnFailure(
                maxRetryCount: 3,
                maxRetryDelay: TimeSpan.FromSeconds(5),
                errorNumbersToAdd: null);
        });
    }
});

// Configure Controllers
builder.Services.AddControllers()
    .AddJsonOptions(options =>
    {
        options.JsonSerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
        options.JsonSerializerOptions.PropertyNameCaseInsensitive = true;
    });

// Configure Swagger with API Key authentication
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "Word Automation & Processing API",
        Version = "v1",
        Description = "ASP.NET Core 8 Web API for Microsoft Word Document Generation and Processing"
    });

    c.AddSecurityDefinition("ApiKey", new OpenApiSecurityScheme
    {
        Description = "API Key authentication using X-API-Key header. Example: 'secret-internship-api-key-2026'",
        Type = SecuritySchemeType.ApiKey,
        Name = "X-API-Key",
        In = ParameterLocation.Header,
        Scheme = "ApiKeyScheme"
    });

    c.AddSecurityRequirement(new OpenApiSecurityRequirement
    {
        {
            new OpenApiSecurityScheme
            {
                Reference = new OpenApiReference
                {
                    Type = ReferenceType.SecurityScheme,
                    Id = "ApiKey"
                }
            },
            Array.Empty<string>()
        }
    });
});

var app = builder.Build();

// Enable Swagger in Development and Production
app.UseSwagger();
app.UseSwaggerUI(c =>
{
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "Word System API v1");
    c.RoutePrefix = "swagger";
});

// Global exception handling
app.UseMiddleware<ExceptionMiddleware>();

app.UseRouting();

// API Key Authentication middleware
app.UseMiddleware<ApiKeyMiddleware>();

app.MapControllers();

// Initialize and seed database
using (var scope = app.Services.CreateScope())
{
    var logger = scope.ServiceProvider.GetRequiredService<Microsoft.Extensions.Logging.ILogger<Program>>();
    try
    {
        var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
        logger.LogInformation("Initializing database ({Provider})...", db.Database.ProviderName);
        await DbInitializer.InitializeAsync(db);
        logger.LogInformation("Database initialized and verified successfully.");
    }
    catch (Exception ex)
    {
        logger.LogError(ex, "Database initialization warning: {Message}", ex.Message);
    }
}

app.Run();

// Required for WebApplicationFactory in integration tests
public partial class Program { }
