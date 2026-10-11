using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using WordClient.Core.Models;
using WordServer.Api.Data;
using Xunit;

namespace WordServer.Tests
{
    public class CustomWebApplicationFactory : WebApplicationFactory<Program>
    {
        private readonly string _testDbPath;

        public CustomWebApplicationFactory()
        {
            _testDbPath = Path.Combine(Path.GetTempPath(), $"WordTestDb_{Guid.NewGuid():N}.db");
        }

        protected override void ConfigureWebHost(IWebHostBuilder builder)
        {
            builder.ConfigureServices(services =>
            {
                // Remove existing AppDbContext descriptor
                var descriptor = services.SingleOrDefault(d => d.ServiceType == typeof(DbContextOptions<AppDbContext>));
                if (descriptor != null)
                {
                    services.Remove(descriptor);
                }

                // Use isolated SQLite database file for tests
                services.AddDbContext<AppDbContext>(options =>
                {
                    options.UseSqlite($"Data Source={_testDbPath}");
                });

                // Build service provider and initialize schema & seed data
                var sp = services.BuildServiceProvider();
                using var scope = sp.CreateScope();
                var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
                DbInitializer.InitializeAsync(db).GetAwaiter().GetResult();
            });
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            try
            {
                if (File.Exists(_testDbPath))
                {
                    File.Delete(_testDbPath);
                }
            }
            catch
            {
                // Ignore cleanup error
            }
        }
    }

    public class ServerIntegrationTests : IClassFixture<CustomWebApplicationFactory>
    {
        private readonly HttpClient _client;
        private const string ValidApiKey = "secret-internship-api-key-2026";
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        public ServerIntegrationTests(CustomWebApplicationFactory factory)
        {
            _client = factory.CreateClient();
        }

        [Fact]
        public async Task HealthCheck_ReturnsOkAndHealthyStatus()
        {
            var response = await _client.GetAsync("/api/health");
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);

            var content = await response.Content.ReadAsStringAsync();
            Assert.Contains("Healthy", content);
        }

        [Fact]
        public async Task AuthValidate_WithValidKey_ReturnsSuccess()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/auth/validate");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new AuthValidateRequest { ApiKey = ValidApiKey }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);

            var content = await response.Content.ReadAsStringAsync();
            var auth = JsonSerializer.Deserialize<AuthValidateResponse>(content, JsonOptions);
            Assert.NotNull(auth);
            Assert.True(auth.Valid);
        }

        [Fact]
        public async Task AuthValidate_WithInvalidKey_ReturnsUnauthorized()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/auth/validate");
            request.Headers.Add("X-API-Key", "invalid-key-999");
            request.Content = new StringContent(
                JsonSerializer.Serialize(new AuthValidateRequest { ApiKey = "invalid-key-999" }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
        }

        [Fact]
        public async Task ProtectedEndpoint_WithoutApiKey_Returns401Unauthorized()
        {
            var response = await _client.GetAsync("/api/templates");
            Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
        }

        [Fact]
        public async Task GetDocument_ExistingDoc001_ReturnsExpectedStructure()
        {
            var request = new HttpRequestMessage(HttpMethod.Get, "/api/documents/doc-001");
            request.Headers.Add("X-API-Key", ValidApiKey);

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);

            var content = await response.Content.ReadAsStringAsync();
            var doc = JsonSerializer.Deserialize<DocumentDto>(content, JsonOptions);

            Assert.NotNull(doc);
            Assert.Equal("doc-001", doc.Id);
            Assert.Equal("Report", doc.Title);
            Assert.Equal("Ivan", doc.Metadata.Author);
            Assert.Equal("Times New Roman", doc.Settings.Font);
            Assert.Equal(12, doc.Settings.FontSize);
            Assert.NotEmpty(doc.Content.Paragraphs);
            Assert.NotEmpty(doc.Content.Tables);
        }

        [Fact]
        public async Task GetDocument_NonExistent_Returns404NotFound()
        {
            var request = new HttpRequestMessage(HttpMethod.Get, "/api/documents/non-existent-doc-999");
            request.Headers.Add("X-API-Key", ValidApiKey);

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
        }

        [Fact]
        public async Task ProcessText_UppercaseOperation_ReturnsUpperResult()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/text/process");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new TextProcessRequest { Text = "hello world", Operation = "uppercase" }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);

            var content = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<TextProcessResponse>(content, JsonOptions);
            Assert.NotNull(result);
            Assert.Equal("HELLO WORLD", result.Result);
        }

        [Fact]
        public async Task ProcessText_ReverseOperation_ReturnsReversedText()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/text/process");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new TextProcessRequest { Text = "Internship 2026", Operation = "reverse" }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);

            var content = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<TextProcessResponse>(content, JsonOptions);
            Assert.NotNull(result);
            Assert.Equal("6202 pihsnretnI", result.Result);
        }

        [Fact]
        public async Task ProcessText_WordCountOperation_ReturnsWordCount()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/text/process");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new TextProcessRequest { Text = "one two three four", Operation = "word_count" }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);

            var content = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<TextProcessResponse>(content, JsonOptions);
            Assert.NotNull(result);
            Assert.Equal("Word count: 4", result.Result);
        }

        [Fact]
        public async Task ProcessText_InvalidOperation_ReturnsBadRequest()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/text/process");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new TextProcessRequest { Text = "test", Operation = "invalid_op_xyz" }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        }

        [Fact]
        public async Task ProcessDocument_ValidText_ReturnsProcessedStats()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/documents/process");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new DocumentProcessRequest
                {
                    Id = "doc-001",
                    Text = "Paragraph 1\r\nParagraph 2\r\nParagraph 3",
                    Operation = "uppercase"
                }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);

            var content = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<DocumentProcessResponse>(content, JsonOptions);
            Assert.NotNull(result);
            Assert.Equal("PARAGRAPH 1\r\nPARAGRAPH 2\r\nPARAGRAPH 3", result.Result);
            Assert.Equal(3, result.ParagraphsCount);
        }

        [Fact]
        public async Task CheckDocument_DetectsEmptyParagraphsAndStats()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/documents/check");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new DocumentCheckRequest
                {
                    DocumentId = "doc-001",
                    Text = "First paragraph\r\n\r\nThird paragraph",
                    Paragraphs = new List<string> { "First paragraph", "", "Third paragraph" },
                    TableCount = 1
                }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);

            var content = await response.Content.ReadAsStringAsync();
            var checkResult = JsonSerializer.Deserialize<DocumentCheckResult>(content, JsonOptions);
            Assert.NotNull(checkResult);
            Assert.Equal(1, checkResult.EmptyParagraphCount);
            Assert.Equal(3, checkResult.ParagraphCount);
            Assert.Equal(1, checkResult.TableCount);
            Assert.NotEmpty(checkResult.Issues);
        }

        [Fact]
        public async Task TemplateCrud_FullLifecycle_Succeeds()
        {
            var newTemplate = new TemplateDto
            {
                Id = $"test-tpl-{Guid.NewGuid():N}",
                Name = "Integration Test Template",
                Description = "Created during test run",
                Content = new DocumentContentDto
                {
                    Paragraphs = new List<ParagraphDto>
                    {
                        new ParagraphDto { Style = "Heading1", Text = "Test Heading" }
                    }
                },
                Settings = new DocumentSettingsDto { Font = "Arial", FontSize = 11 }
            };

            // 1. CREATE
            var createReq = new HttpRequestMessage(HttpMethod.Post, "/api/templates");
            createReq.Headers.Add("X-API-Key", ValidApiKey);
            createReq.Content = new StringContent(JsonSerializer.Serialize(newTemplate), Encoding.UTF8, "application/json");

            var createResp = await _client.SendAsync(createReq);
            Assert.Equal(HttpStatusCode.Created, createResp.StatusCode);

            // 2. READ by ID
            var getReq = new HttpRequestMessage(HttpMethod.Get, $"/api/templates/{newTemplate.Id}");
            getReq.Headers.Add("X-API-Key", ValidApiKey);
            var getResp = await _client.SendAsync(getReq);
            Assert.Equal(HttpStatusCode.OK, getResp.StatusCode);

            var fetched = JsonSerializer.Deserialize<TemplateDto>(await getResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(fetched);
            Assert.Equal(newTemplate.Name, fetched.Name);

            // 3. READ ALL
            var getAllReq = new HttpRequestMessage(HttpMethod.Get, "/api/templates");
            getAllReq.Headers.Add("X-API-Key", ValidApiKey);
            var getAllResp = await _client.SendAsync(getAllReq);
            Assert.Equal(HttpStatusCode.OK, getAllResp.StatusCode);

            var allList = JsonSerializer.Deserialize<List<TemplateDto>>(await getAllResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(allList);
            Assert.Contains(allList, t => t.Id == newTemplate.Id);

            // 4. UPDATE
            newTemplate.Name = "Updated Integration Test Template";
            var updateReq = new HttpRequestMessage(HttpMethod.Put, $"/api/templates/{newTemplate.Id}");
            updateReq.Headers.Add("X-API-Key", ValidApiKey);
            updateReq.Content = new StringContent(JsonSerializer.Serialize(newTemplate), Encoding.UTF8, "application/json");
            var updateResp = await _client.SendAsync(updateReq);
            Assert.Equal(HttpStatusCode.OK, updateResp.StatusCode);

            // 5. DELETE
            var deleteReq = new HttpRequestMessage(HttpMethod.Delete, $"/api/templates/{newTemplate.Id}");
            deleteReq.Headers.Add("X-API-Key", ValidApiKey);
            var deleteResp = await _client.SendAsync(deleteReq);
            Assert.Equal(HttpStatusCode.NoContent, deleteResp.StatusCode);

            // 6. CONFIRM 404
            var confirmReq = new HttpRequestMessage(HttpMethod.Get, $"/api/templates/{newTemplate.Id}");
            confirmReq.Headers.Add("X-API-Key", ValidApiKey);
            var confirmResp = await _client.SendAsync(confirmReq);
            Assert.Equal(HttpStatusCode.NotFound, confirmResp.StatusCode);
        }

        [Fact]
        public async Task TextProcess_MalformedJson_ReturnsBadRequest()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/text/process");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent("{ malformed_json: true, missing_quotes }", Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        }

        [Fact]
        public async Task CreateDocument_ValidDocument_ReturnsCreated()
        {
            var newDoc = new DocumentDto
            {
                Id = $"new-doc-{Guid.NewGuid():N}",
                Title = "Newly Created Document",
                Metadata = new DocumentMetadataDto { Author = "Tester", CreatedAt = DateTime.UtcNow.ToString("O") },
                Content = new DocumentContentDto
                {
                    Paragraphs = new List<ParagraphDto> { new ParagraphDto { Style = "Normal", Text = "Sample text" } }
                },
                Settings = new DocumentSettingsDto { Font = "Calibri", FontSize = 11 }
            };

            var request = new HttpRequestMessage(HttpMethod.Post, "/api/documents");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(JsonSerializer.Serialize(newDoc), Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.Created, response.StatusCode);

            // Fetch to verify persistence
            var getReq = new HttpRequestMessage(HttpMethod.Get, $"/api/documents/{newDoc.Id}");
            getReq.Headers.Add("X-API-Key", ValidApiKey);
            var getResp = await _client.SendAsync(getReq);
            Assert.Equal(HttpStatusCode.OK, getResp.StatusCode);
        }

        [Fact]
        public async Task CreateDocument_DuplicateId_ReturnsConflict()
        {
            var duplicateDoc = new DocumentDto
            {
                Id = "doc-001", // Already seeded in DB
                Title = "Duplicate Document"
            };

            var request = new HttpRequestMessage(HttpMethod.Post, "/api/documents");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(JsonSerializer.Serialize(duplicateDoc), Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.Conflict, response.StatusCode);
        }

        [Theory]
        [InlineData("lowercase", "HELLO WORLD", "hello world")]
        [InlineData("trim", "   padded text   ", "padded text")]
        [InlineData("capitalize", "capitalized", "Capitalized")]
        public async Task ProcessText_Operations_ReturnExpectedOutput(string operation, string input, string expected)
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/text/process");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new TextProcessRequest { Text = input, Operation = operation }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);

            var content = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<TextProcessResponse>(content, JsonOptions);
            Assert.NotNull(result);
            Assert.Equal(expected, result.Result);
        }

        [Fact]
        public async Task ProcessDocument_EmptyText_ReturnsBadRequest()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/documents/process");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new DocumentProcessRequest { Text = "" }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        }

        [Fact]
        public async Task Template_GetNonExistent_ReturnsNotFound()
        {
            var request = new HttpRequestMessage(HttpMethod.Get, "/api/templates/non-existent-tpl-xyz");
            request.Headers.Add("X-API-Key", ValidApiKey);

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
        }

        [Fact]
        public async Task Template_UpdateNonExistent_ReturnsNotFound()
        {
            var request = new HttpRequestMessage(HttpMethod.Put, "/api/templates/non-existent-tpl-xyz");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new TemplateDto { Name = "Updated" }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
        }

        [Fact]
        public async Task Template_DeleteNonExistent_ReturnsNotFound()
        {
            var request = new HttpRequestMessage(HttpMethod.Delete, "/api/templates/non-existent-tpl-xyz");
            request.Headers.Add("X-API-Key", ValidApiKey);

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
        }

        [Fact]
        public async Task GetDocument_EmptyId_ReturnsBadRequest()
        {
            var request = new HttpRequestMessage(HttpMethod.Get, "/api/documents/%20");
            request.Headers.Add("X-API-Key", ValidApiKey);

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        }

        [Fact]
        public async Task ProcessText_NullText_ReturnsBadRequest()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/text/process");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new TextProcessRequest { Text = null!, Operation = "uppercase" }),
                Encoding.UTF8,
                "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        }

        [Fact]
        public async Task CheckDocument_NullRequestBody_ReturnsBadRequest()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/documents/check");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent("null", Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        }

        [Fact]
        public async Task CreateTemplate_MissingName_ReturnsBadRequest()
        {
            var invalidTpl = new TemplateDto
            {
                Id = $"test-tpl-{Guid.NewGuid():N}",
                Name = "" // Empty name
            };

            var request = new HttpRequestMessage(HttpMethod.Post, "/api/templates");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(JsonSerializer.Serialize(invalidTpl), Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        }

        [Fact]
        public async Task CreateTemplate_DuplicateId_ReturnsConflict()
        {
            var duplicateTpl = new TemplateDto
            {
                Id = "tpl-001", // Already seeded in DB
                Name = "Duplicate Template"
            };

            var request = new HttpRequestMessage(HttpMethod.Post, "/api/templates");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(JsonSerializer.Serialize(duplicateTpl), Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.Conflict, response.StatusCode);
        }
    }
}
