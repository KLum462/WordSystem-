using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging.Abstractions;
using WordClient.Core.Configuration;
using WordClient.Core.Models;
using WordClient.Core.Services;
using WordServer.Api.Data;
using WordServer.Api.Middleware;
using Xunit;

namespace WordServer.Tests
{
    public class CustomWebApplicationFactory : WebApplicationFactory<Program>
    {
        private readonly SqliteConnection _inMemoryConnection;
        public string UploadsTestDir { get; }

        public CustomWebApplicationFactory()
        {
            // True SQLite in-memory database kept open for the lifetime of TestServer
            _inMemoryConnection = new SqliteConnection("DataSource=:memory:");
            _inMemoryConnection.Open();

            UploadsTestDir = Path.Combine(Path.GetTempPath(), $"WordUploadsTest_{Guid.NewGuid():N}");
            Directory.CreateDirectory(UploadsTestDir);
        }

        protected override void ConfigureWebHost(IWebHostBuilder builder)
        {
            builder.UseSetting("DatabaseProvider", "Sqlite");
            builder.UseSetting("Storage:UploadsPath", UploadsTestDir);

            builder.ConfigureServices(services =>
            {
                // Remove existing AppDbContext descriptor
                var descriptor = services.SingleOrDefault(d => d.ServiceType == typeof(DbContextOptions<AppDbContext>));
                if (descriptor != null)
                {
                    services.Remove(descriptor);
                }

                // Use shared SQLite in-memory connection for TestServer
                services.AddDbContext<AppDbContext>(options =>
                {
                    options.UseSqlite(_inMemoryConnection);
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
            if (disposing)
            {
                _inMemoryConnection.Close();
                _inMemoryConnection.Dispose();
                try
                {
                    if (Directory.Exists(UploadsTestDir))
                    {
                        Directory.Delete(UploadsTestDir, recursive: true);
                    }
                }
                catch
                {
                    // Ignore cleanup error
                }
            }
        }
    }

    public class ServerIntegrationTests : IClassFixture<CustomWebApplicationFactory>
    {
        private readonly CustomWebApplicationFactory _factory;
        private readonly HttpClient _client;
        private const string ValidApiKey = "secret-internship-api-key-2026";
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };

        public ServerIntegrationTests(CustomWebApplicationFactory factory)
        {
            _factory = factory;
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

        [Fact]
        public async Task AuthValidate_ReturnsRealDatabaseUserAndRole()
        {
            var adminReq = new HttpRequestMessage(HttpMethod.Post, "/api/auth/validate");
            adminReq.Headers.Add("X-API-Key", ValidApiKey);
            adminReq.Content = new StringContent("{}", Encoding.UTF8, "application/json");

            var adminResp = await _client.SendAsync(adminReq);
            Assert.Equal(HttpStatusCode.OK, adminResp.StatusCode);
            var adminAuth = JsonSerializer.Deserialize<AuthValidateResponse>(await adminResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(adminAuth);
            Assert.True(adminAuth.Valid);
            Assert.Equal("admin", adminAuth.User);
            Assert.Equal("Admin", adminAuth.Role);

            var internReq = new HttpRequestMessage(HttpMethod.Post, "/api/auth/validate");
            internReq.Headers.Add("X-API-Key", "internship-key-demo");
            internReq.Content = new StringContent("{}", Encoding.UTF8, "application/json");

            var internResp = await _client.SendAsync(internReq);
            Assert.Equal(HttpStatusCode.OK, internResp.StatusCode);
            var internAuth = JsonSerializer.Deserialize<AuthValidateResponse>(await internResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(internAuth);
            Assert.Equal("ivan_intern", internAuth.User);
            Assert.Equal("Intern", internAuth.Role);
        }

        [Fact]
        public async Task UserManagement_CreateAndDeactivateUser_ControlsApiAccess()
        {
            var uniqueUser = $"tester_{Guid.NewGuid():N}".Substring(0, 16);
            var createReq = new HttpRequestMessage(HttpMethod.Post, "/api/management/users");
            createReq.Headers.Add("X-API-Key", ValidApiKey);
            createReq.Content = new StringContent(
                JsonSerializer.Serialize(new CreateUserRequest
                {
                    Username = uniqueUser,
                    FullName = "QA Engineer",
                    Email = $"{uniqueUser}@company.local",
                    Role = "Reviewer"
                }),
                Encoding.UTF8,
                "application/json");

            var createResp = await _client.SendAsync(createReq);
            Assert.Equal(HttpStatusCode.Created, createResp.StatusCode);

            using var docJson = JsonDocument.Parse(await createResp.Content.ReadAsStringAsync());
            int userId = docJson.RootElement.GetProperty("id").GetInt32();
            string newApiKey = docJson.RootElement.GetProperty("apiKey").GetString()!;
            Assert.False(string.IsNullOrWhiteSpace(newApiKey));

            // New key works immediately
            var pingReq = new HttpRequestMessage(HttpMethod.Get, "/api/documents");
            pingReq.Headers.Add("X-API-Key", newApiKey);
            var pingResp = await _client.SendAsync(pingReq);
            Assert.Equal(HttpStatusCode.OK, pingResp.StatusCode);

            // Deactivate user
            var deactivateReq = new HttpRequestMessage(HttpMethod.Patch, $"/api/management/users/{userId}/status");
            deactivateReq.Headers.Add("X-API-Key", ValidApiKey);
            deactivateReq.Content = new StringContent(
                JsonSerializer.Serialize(new UpdateUserStatusRequest { IsActive = false }),
                Encoding.UTF8,
                "application/json");
            var deactivateResp = await _client.SendAsync(deactivateReq);
            Assert.Equal(HttpStatusCode.OK, deactivateResp.StatusCode);

            // Deactivated user's key is immediately rejected
            var blockedReq = new HttpRequestMessage(HttpMethod.Get, "/api/documents");
            blockedReq.Headers.Add("X-API-Key", newApiKey);
            var blockedResp = await _client.SendAsync(blockedReq);
            Assert.Equal(HttpStatusCode.Unauthorized, blockedResp.StatusCode);
        }

        [Fact]
        public async Task DocumentVersioning_CreateAndUpdateDocument_CreatesVersions()
        {
            var docId = $"ver-doc-{Guid.NewGuid():N}";
            var createReq = new HttpRequestMessage(HttpMethod.Post, "/api/documents");
            createReq.Headers.Add("X-API-Key", ValidApiKey);
            createReq.Content = new StringContent(JsonSerializer.Serialize(new DocumentDto
            {
                Id = docId,
                Title = "Contract Draft",
                Content = new DocumentContentDto
                {
                    Paragraphs = new List<ParagraphDto> { new ParagraphDto { Style = "Normal", Text = "Version 1 text" } }
                }
            }), Encoding.UTF8, "application/json");

            var createResp = await _client.SendAsync(createReq);
            Assert.Equal(HttpStatusCode.Created, createResp.StatusCode);

            var updateReq = new HttpRequestMessage(HttpMethod.Put, $"/api/documents/{docId}");
            updateReq.Headers.Add("X-API-Key", ValidApiKey);
            updateReq.Content = new StringContent(JsonSerializer.Serialize(new DocumentDto
            {
                Id = docId,
                Title = "Contract Final",
                Metadata = new DocumentMetadataDto { Description = "Approved by legal" },
                Content = new DocumentContentDto
                {
                    Paragraphs = new List<ParagraphDto> { new ParagraphDto { Style = "Normal", Text = "Version 2 text" } }
                }
            }), Encoding.UTF8, "application/json");

            var updateResp = await _client.SendAsync(updateReq);
            Assert.Equal(HttpStatusCode.OK, updateResp.StatusCode);

            var versionsReq = new HttpRequestMessage(HttpMethod.Get, $"/api/management/versions/{docId}");
            versionsReq.Headers.Add("X-API-Key", ValidApiKey);
            var versionsResp = await _client.SendAsync(versionsReq);
            Assert.Equal(HttpStatusCode.OK, versionsResp.StatusCode);

            using var versionsJson = JsonDocument.Parse(await versionsResp.Content.ReadAsStringAsync());
            Assert.Equal(2, versionsJson.RootElement.GetArrayLength());
            Assert.Equal(2, versionsJson.RootElement[0].GetProperty("versionNumber").GetInt32());
            Assert.Equal(1, versionsJson.RootElement[1].GetProperty("versionNumber").GetInt32());
        }

        [Fact]
        public async Task ProcessingHistory_RecordedAutomaticallyOnTextAndDocumentProcessing()
        {
            var markerDocId = $"hist-{Guid.NewGuid():N}";

            var textReq = new HttpRequestMessage(HttpMethod.Post, "/api/text/process");
            textReq.Headers.Add("X-API-Key", ValidApiKey);
            textReq.Content = new StringContent(JsonSerializer.Serialize(new TextProcessRequest
            {
                DocumentId = markerDocId,
                Text = "  Hello   world  ,  test! ",
                Operation = "normalize"
            }), Encoding.UTF8, "application/json");

            var textResp = await _client.SendAsync(textReq);
            Assert.Equal(HttpStatusCode.OK, textResp.StatusCode);
            var textResult = JsonSerializer.Deserialize<TextProcessResponse>(await textResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(textResult);
            Assert.True(textResult.Id > 0);
            Assert.Equal("Hello world, test!", textResult.Result);

            var histReq = new HttpRequestMessage(HttpMethod.Get, $"/api/management/history?documentId={markerDocId}");
            histReq.Headers.Add("X-API-Key", ValidApiKey);
            var histResp = await _client.SendAsync(histReq);
            Assert.Equal(HttpStatusCode.OK, histResp.StatusCode);

            var history = JsonSerializer.Deserialize<List<ProcessingHistoryDto>>(await histResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(history);
            Assert.Contains(history, h => h.DocumentId == markerDocId && h.Operation == "normalize");
        }

        [Fact]
        public async Task SmartTextProcessing_AnonymizeExtractSummarizeStats_WorkCorrectly()
        {
            // Anonymize PII
            var anonReq = new HttpRequestMessage(HttpMethod.Post, "/api/text/process");
            anonReq.Headers.Add("X-API-Key", ValidApiKey);
            anonReq.Content = new StringContent(JsonSerializer.Serialize(new TextProcessRequest
            {
                Text = "Contact ivan@example.com or +7 (999) 123-45-67, INN: 7701234567, card 4276 1234 5678 9012.",
                Operation = "anonymize"
            }), Encoding.UTF8, "application/json");
            var anonResp = await _client.SendAsync(anonReq);
            Assert.Equal(HttpStatusCode.OK, anonResp.StatusCode);
            var anonResult = JsonSerializer.Deserialize<TextProcessResponse>(await anonResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.Contains("[EMAIL]", anonResult!.Result);
            Assert.Contains("[PHONE]", anonResult.Result);
            Assert.Contains("[INN]", anonResult.Result);
            Assert.Contains("[CARD]", anonResult.Result);

            // Extract entities
            var extReq = new HttpRequestMessage(HttpMethod.Post, "/api/text/process");
            extReq.Headers.Add("X-API-Key", ValidApiKey);
            extReq.Content = new StringContent(JsonSerializer.Serialize(new TextProcessRequest
            {
                Text = "Invoice dated 15.10.2026 for 125 000 руб. sent to billing@acme.org see https://acme.org/inv",
                Operation = "extract_entities"
            }), Encoding.UTF8, "application/json");
            var extResp = await _client.SendAsync(extReq);
            Assert.Equal(HttpStatusCode.OK, extResp.StatusCode);
            var extResult = JsonSerializer.Deserialize<TextProcessResponse>(await extResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.Contains("15.10.2026", extResult!.Result);
            Assert.Contains("billing@acme.org", extResult.Result);
            Assert.Contains("125 000 руб.", extResult.Result);
            Assert.Contains("https://acme.org/inv", extResult.Result);

            // Summarize & Stats
            var sumReq = new HttpRequestMessage(HttpMethod.Post, "/api/documents/process");
            sumReq.Headers.Add("X-API-Key", ValidApiKey);
            sumReq.Content = new StringContent(JsonSerializer.Serialize(new DocumentProcessRequest
            {
                Id = "doc-summary-test",
                Text = "Enterprise document automation reduces manual errors. Teams generate contracts from structured server templates. Server templates enforce consistent typography and tables. Automated checks detect formatting issues and personal data leaks.",
                Operation = "summarize"
            }), Encoding.UTF8, "application/json");
            var sumResp = await _client.SendAsync(sumReq);
            Assert.Equal(HttpStatusCode.OK, sumResp.StatusCode);
            var sumResult = JsonSerializer.Deserialize<DocumentProcessResponse>(await sumResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.False(string.IsNullOrWhiteSpace(sumResult!.Result));

            var statsReq = new HttpRequestMessage(HttpMethod.Post, "/api/text/process");
            statsReq.Headers.Add("X-API-Key", ValidApiKey);
            statsReq.Content = new StringContent(JsonSerializer.Serialize(new TextProcessRequest
            {
                Text = "Server automation improves document quality. Automation speeds up reporting.",
                Operation = "stats"
            }), Encoding.UTF8, "application/json");
            var statsResp = await _client.SendAsync(statsReq);
            Assert.Equal(HttpStatusCode.OK, statsResp.StatusCode);
            var statsResult = JsonSerializer.Deserialize<TextProcessResponse>(await statsResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.Contains("Text Analytics", statsResult!.Result);
            Assert.Contains("automation", statsResult.Result);
        }

        [Fact]
        public async Task CheckDocument_DetectsPiiAndDoubleSpaces()
        {
            var request = new HttpRequestMessage(HttpMethod.Post, "/api/documents/check");
            request.Headers.Add("X-API-Key", ValidApiKey);
            request.Content = new StringContent(JsonSerializer.Serialize(new DocumentCheckRequest
            {
                DocumentId = "doc-pii",
                Text = "Hello  world! Reach us at secret@corp.local.",
                TableCount = 0
            }), Encoding.UTF8, "application/json");

            var response = await _client.SendAsync(request);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            var checkResult = JsonSerializer.Deserialize<DocumentCheckResult>(await response.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(checkResult);
            Assert.Contains(checkResult.Issues, i => i.Contains("multiple consecutive spaces"));
            Assert.Contains(checkResult.Issues, i => i.Contains("Potential PII"));
        }

        [Fact]
        public async Task Management_CategoriesCommentsAndSettings_FullFlow()
        {
            // Create category
            var catReq = new HttpRequestMessage(HttpMethod.Post, "/api/management/categories");
            catReq.Headers.Add("X-API-Key", ValidApiKey);
            catReq.Content = new StringContent(JsonSerializer.Serialize(new CreateCategoryRequest
            {
                Name = $"Legal-{Guid.NewGuid():N}".Substring(0, 14),
                Description = "Legal templates"
            }), Encoding.UTF8, "application/json");
            var catResp = await _client.SendAsync(catReq);
            Assert.Equal(HttpStatusCode.Created, catResp.StatusCode);
            var createdCat = JsonSerializer.Deserialize<TemplateCategoryDto>(await catResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(createdCat);

            // Create template linked to category and filter by categoryId
            var tplId = $"cat-tpl-{Guid.NewGuid():N}";
            var tplReq = new HttpRequestMessage(HttpMethod.Post, "/api/templates");
            tplReq.Headers.Add("X-API-Key", ValidApiKey);
            tplReq.Content = new StringContent(JsonSerializer.Serialize(new TemplateDto
            {
                Id = tplId,
                Name = "NDA Template",
                CategoryId = createdCat.Id
            }), Encoding.UTF8, "application/json");
            Assert.Equal(HttpStatusCode.Created, (await _client.SendAsync(tplReq)).StatusCode);

            var filterReq = new HttpRequestMessage(HttpMethod.Get, $"/api/templates?categoryId={createdCat.Id}");
            filterReq.Headers.Add("X-API-Key", ValidApiKey);
            var filterResp = await _client.SendAsync(filterReq);
            var filteredTemplates = JsonSerializer.Deserialize<List<TemplateDto>>(await filterResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.Single(filteredTemplates!);
            Assert.Equal(tplId, filteredTemplates![0].Id);

            // Create and resolve comment
            var cmtReq = new HttpRequestMessage(HttpMethod.Post, "/api/management/comments");
            cmtReq.Headers.Add("X-API-Key", ValidApiKey);
            cmtReq.Content = new StringContent(JsonSerializer.Serialize(new CreateCommentRequest
            {
                DocumentId = "doc-001",
                Comment = "Please verify Q4 figures."
            }), Encoding.UTF8, "application/json");
            var cmtResp = await _client.SendAsync(cmtReq);
            Assert.Equal(HttpStatusCode.Created, cmtResp.StatusCode);
            using var cmtJson = JsonDocument.Parse(await cmtResp.Content.ReadAsStringAsync());
            int commentId = cmtJson.RootElement.GetProperty("id").GetInt32();

            var resolveReq = new HttpRequestMessage(HttpMethod.Patch, $"/api/management/comments/{commentId}/resolve");
            resolveReq.Headers.Add("X-API-Key", ValidApiKey);
            resolveReq.Content = new StringContent(JsonSerializer.Serialize(new ResolveCommentRequest { Resolved = true }), Encoding.UTF8, "application/json");
            var resolveResp = await _client.SendAsync(resolveReq);
            Assert.Equal(HttpStatusCode.OK, resolveResp.StatusCode);

            // Upsert setting
            var setReq = new HttpRequestMessage(HttpMethod.Put, "/api/management/settings/DemoMode");
            setReq.Headers.Add("X-API-Key", ValidApiKey);
            setReq.Content = new StringContent(JsonSerializer.Serialize(new UpdateSettingRequest { Value = "false", Description = "Production mode" }), Encoding.UTF8, "application/json");
            Assert.Equal(HttpStatusCode.OK, (await _client.SendAsync(setReq)).StatusCode);
        }

        [Fact]
        public async Task GetAllDocuments_ReturnsSeededDoc001()
        {
            var req = new HttpRequestMessage(HttpMethod.Get, "/api/documents");
            req.Headers.Add("X-API-Key", ValidApiKey);

            var resp = await _client.SendAsync(req);
            Assert.Equal(HttpStatusCode.OK, resp.StatusCode);

            var docs = JsonSerializer.Deserialize<List<DocumentSummaryDto>>(await resp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(docs);
            Assert.Contains(docs, d => d.Id == "doc-001" && d.Title == "Report");
        }

        [Fact]
        public async Task UploadDocument_ValidDocx_SavesToUploadsAndCanBeListedAndDownloaded()
        {
            var sampleBytes = Encoding.UTF8.GetBytes("PK\x03\x04MockDocxContentForIntegrationTest");
            using var multipart = new MultipartFormDataContent();
            var fileContent = new ByteArrayContent(sampleBytes);
            fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            multipart.Add(fileContent, "file", "quarterly-report.docx");

            var uploadReq = new HttpRequestMessage(HttpMethod.Post, "/api/documents/upload")
            {
                Content = multipart
            };
            uploadReq.Headers.Add("X-API-Key", ValidApiKey);

            var uploadResp = await _client.SendAsync(uploadReq);
            Assert.Equal(HttpStatusCode.Created, uploadResp.StatusCode);

            var uploaded = JsonSerializer.Deserialize<DocumentUploadResponse>(await uploadResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(uploaded);
            Assert.Equal("quarterly-report.docx", uploaded.FileName);
            Assert.Equal(sampleBytes.Length, uploaded.SizeBytes);
            Assert.StartsWith("App_Data/uploads/", uploaded.RelativePath);
            Assert.True(File.Exists(Path.Combine(_factory.UploadsTestDir, uploaded.StoredFileName)));

            // List uploaded files
            var listReq = new HttpRequestMessage(HttpMethod.Get, "/api/documents/uploads");
            listReq.Headers.Add("X-API-Key", ValidApiKey);
            var listResp = await _client.SendAsync(listReq);
            Assert.Equal(HttpStatusCode.OK, listResp.StatusCode);
            var list = JsonSerializer.Deserialize<List<DocumentUploadResponse>>(await listResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(list);
            Assert.Contains(list, f => f.StoredFileName == uploaded.StoredFileName);

            // Download uploaded file
            var downloadReq = new HttpRequestMessage(HttpMethod.Get, $"/api/documents/uploads/{uploaded.StoredFileName}");
            downloadReq.Headers.Add("X-API-Key", ValidApiKey);
            var downloadResp = await _client.SendAsync(downloadReq);
            Assert.Equal(HttpStatusCode.OK, downloadResp.StatusCode);
            var downloadedBytes = await downloadResp.Content.ReadAsByteArrayAsync();
            Assert.Equal(sampleBytes, downloadedBytes);

            // Non-existent upload returns 404
            var missingReq = new HttpRequestMessage(HttpMethod.Get, "/api/documents/uploads/does-not-exist.docx");
            missingReq.Headers.Add("X-API-Key", ValidApiKey);
            Assert.Equal(HttpStatusCode.NotFound, (await _client.SendAsync(missingReq)).StatusCode);
        }

        [Fact]
        public async Task UploadDocument_NonDocxOrEmpty_ReturnsBadRequest()
        {
            using var multipart = new MultipartFormDataContent();
            multipart.Add(new ByteArrayContent(Encoding.UTF8.GetBytes("plain text")), "file", "notes.txt");

            var req = new HttpRequestMessage(HttpMethod.Post, "/api/documents/upload") { Content = multipart };
            req.Headers.Add("X-API-Key", ValidApiKey);

            var resp = await _client.SendAsync(req);
            Assert.Equal(HttpStatusCode.BadRequest, resp.StatusCode);
        }

        [Fact]
        public async Task PayloadExceedingConfiguredLimit_Returns413PayloadTooLarge()
        {
            using var limitedFactory = _factory.WithWebHostBuilder(builder =>
            {
                builder.UseSetting("Storage:MaxFileSizeBytes", "512");
            });
            using var limitedClient = limitedFactory.CreateClient();

            // 1. Oversized text processing payload -> 413
            var bigText = new string('A', 1024);
            var procReq = new HttpRequestMessage(HttpMethod.Post, "/api/documents/process");
            procReq.Headers.Add("X-API-Key", ValidApiKey);
            procReq.Content = new StringContent(
                JsonSerializer.Serialize(new DocumentProcessRequest { Id = "big-doc", Text = bigText, Operation = "uppercase" }),
                Encoding.UTF8,
                "application/json");

            var procResp = await limitedClient.SendAsync(procReq);
            Assert.Equal(HttpStatusCode.RequestEntityTooLarge, procResp.StatusCode);

            // 2. Oversized .docx upload -> 413
            using var multipart = new MultipartFormDataContent();
            multipart.Add(new ByteArrayContent(new byte[2048]), "file", "oversized.docx");

            var uploadReq = new HttpRequestMessage(HttpMethod.Post, "/api/documents/upload") { Content = multipart };
            uploadReq.Headers.Add("X-API-Key", ValidApiKey);

            var uploadResp = await limitedClient.SendAsync(uploadReq);
            Assert.Equal(HttpStatusCode.RequestEntityTooLarge, uploadResp.StatusCode);
        }

        [Fact]
        public async Task TemplateSingularRoute_CrudLifecycle_Succeeds()
        {
            var id = $"singular-tpl-{Guid.NewGuid():N}".Substring(0, 20);
            var dto = new TemplateDto
            {
                Id = id,
                Name = "Singular Endpoint Template",
                Content = new DocumentContentDto
                {
                    Paragraphs = new List<ParagraphDto> { new ParagraphDto { Style = "Normal", Text = "Body" } }
                }
            };

            var postReq = new HttpRequestMessage(HttpMethod.Post, "/api/template");
            postReq.Headers.Add("X-API-Key", ValidApiKey);
            postReq.Content = new StringContent(JsonSerializer.Serialize(dto), Encoding.UTF8, "application/json");
            Assert.Equal(HttpStatusCode.Created, (await _client.SendAsync(postReq)).StatusCode);

            var getReq = new HttpRequestMessage(HttpMethod.Get, $"/api/template/{id}");
            getReq.Headers.Add("X-API-Key", ValidApiKey);
            Assert.Equal(HttpStatusCode.OK, (await _client.SendAsync(getReq)).StatusCode);

            dto.Name = "Updated Singular Template";
            var putReq = new HttpRequestMessage(HttpMethod.Put, $"/api/template/{id}");
            putReq.Headers.Add("X-API-Key", ValidApiKey);
            putReq.Content = new StringContent(JsonSerializer.Serialize(dto), Encoding.UTF8, "application/json");
            Assert.Equal(HttpStatusCode.OK, (await _client.SendAsync(putReq)).StatusCode);

            var delReq = new HttpRequestMessage(HttpMethod.Delete, $"/api/template/{id}");
            delReq.Headers.Add("X-API-Key", ValidApiKey);
            Assert.Equal(HttpStatusCode.NoContent, (await _client.SendAsync(delReq)).StatusCode);
        }

        [Fact]
        public async Task ManagementAndOperationsGetEndpoints_ReturnOk()
        {
            string[] paths =
            {
                "/api/text/operations",
                "/api/management/users",
                "/api/management/categories",
                "/api/management/comments/doc-001",
                "/api/management/settings"
            };

            foreach (var path in paths)
            {
                var req = new HttpRequestMessage(HttpMethod.Get, path);
                req.Headers.Add("X-API-Key", ValidApiKey);
                var resp = await _client.SendAsync(req);
                Assert.Equal(HttpStatusCode.OK, resp.StatusCode);
            }
        }

        [Fact]
        public async Task ExceptionMiddleware_Handles413And400And500Gracefully()
        {
            // 413 from BadHttpRequestException
            var ctx413 = new DefaultHttpContext();
            ctx413.Response.Body = new MemoryStream();
            var mw413 = new ExceptionMiddleware(
                _ => throw new BadHttpRequestException("Request body too large", StatusCodes.Status413PayloadTooLarge),
                NullLogger<ExceptionMiddleware>.Instance);
            await mw413.InvokeAsync(ctx413);
            Assert.Equal(StatusCodes.Status413PayloadTooLarge, ctx413.Response.StatusCode);

            // 413 from multipart InvalidDataException
            var ctxMultipart = new DefaultHttpContext();
            ctxMultipart.Response.Body = new MemoryStream();
            var mwMultipart = new ExceptionMiddleware(
                _ => throw new InvalidDataException("Multipart body length limit 104857600 exceeded."),
                NullLogger<ExceptionMiddleware>.Instance);
            await mwMultipart.InvokeAsync(ctxMultipart);
            Assert.Equal(StatusCodes.Status413PayloadTooLarge, ctxMultipart.Response.StatusCode);

            // 400 from JsonException
            var ctxJson = new DefaultHttpContext();
            ctxJson.Response.Body = new MemoryStream();
            var mwJson = new ExceptionMiddleware(
                _ => throw new JsonException("Unexpected token"),
                NullLogger<ExceptionMiddleware>.Instance);
            await mwJson.InvokeAsync(ctxJson);
            Assert.Equal(StatusCodes.Status400BadRequest, ctxJson.Response.StatusCode);

            // 500 from unhandled Exception (also persists error log to DB)
            var ctx500 = new DefaultHttpContext
            {
                RequestServices = _factory.Services
            };
            ctx500.Request.Path = "/api/test/crash";
            ctx500.Response.Body = new MemoryStream();
            var mw500 = new ExceptionMiddleware(
                _ => throw new InvalidOperationException("Simulated unhandled server error"),
                NullLogger<ExceptionMiddleware>.Instance);
            await mw500.InvokeAsync(ctx500);
            Assert.Equal(StatusCodes.Status500InternalServerError, ctx500.Response.StatusCode);
        }

        [Fact]
        public async Task ApiClient_Validates100MbLimitAndReadsAppConfigFile()
        {
            var tempCfgPath = Path.Combine(Path.GetTempPath(), $"test_addin_{Guid.NewGuid():N}.config");
            try
            {
                File.WriteAllText(tempCfgPath,
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?><configuration><appSettings>" +
                    "<add key=\"ServerUrl\" value=\"http://localhost:5055\" />" +
                    "<add key=\"ApiKey\" value=\"custom-key-test\" />" +
                    "<add key=\"TimeoutSeconds\" value=\"45\" />" +
                    "<add key=\"MaxRetries\" value=\"4\" />" +
                    "<add key=\"MaxFileSizeBytes\" value=\"1024\" />" +
                    "</appSettings></configuration>");

                var cfg = ClientConfig.LoadFromFile(tempCfgPath);
                Assert.Equal("http://localhost:5055", cfg.ServerUrl);
                Assert.Equal("custom-key-test", cfg.ApiKey);
                Assert.Equal(45, cfg.TimeoutSeconds);
                Assert.Equal(4, cfg.MaxRetries);
                Assert.Equal(1024, cfg.MaxFileSizeBytes);

                using var api = new ApiClient(cfg, _factory.CreateClient());
                await Assert.ThrowsAsync<InvalidOperationException>(() =>
                    api.ProcessDocumentAsync("doc-big", new string('X', 2048), "uppercase"));
            }
            finally
            {
                if (File.Exists(tempCfgPath)) File.Delete(tempCfgPath);
            }
        }

        [Fact]
        public async Task AuthLoginAndChangePassword_WorksAndRejectsWrongPassword()
        {
            // Valid login for seeded admin
            var loginReq = new HttpRequestMessage(HttpMethod.Post, "/api/auth/login")
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(new LoginRequest { Username = "admin", Password = "admin123" }),
                    Encoding.UTF8,
                    "application/json")
            };
            var loginResp = await _client.SendAsync(loginReq);
            Assert.Equal(HttpStatusCode.OK, loginResp.StatusCode);

            var auth = JsonSerializer.Deserialize<AuthValidateResponse>(await loginResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.NotNull(auth);
            Assert.True(auth!.Valid);
            Assert.Equal("admin", auth.User);
            Assert.Equal("Admin", auth.Role);
            Assert.Equal(ValidApiKey, auth.ApiKey);

            // Wrong password -> 401
            var badReq = new HttpRequestMessage(HttpMethod.Post, "/api/auth/login")
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(new LoginRequest { Username = "admin", Password = "wrong-password" }),
                    Encoding.UTF8,
                    "application/json")
            };
            Assert.Equal(HttpStatusCode.Unauthorized, (await _client.SendAsync(badReq)).StatusCode);

            // Change password for ivan_intern
            var cpReq = new HttpRequestMessage(HttpMethod.Post, "/api/auth/change-password")
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(new ChangePasswordRequest
                    {
                        Username = "ivan_intern",
                        CurrentPassword = "ivan123",
                        NewPassword = "newpass456"
                    }),
                    Encoding.UTF8,
                    "application/json")
            };
            Assert.Equal(HttpStatusCode.OK, (await _client.SendAsync(cpReq)).StatusCode);

            // Old password now fails -> 401
            var oldLoginReq = new HttpRequestMessage(HttpMethod.Post, "/api/auth/login")
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(new LoginRequest { Username = "ivan_intern", Password = "ivan123" }),
                    Encoding.UTF8,
                    "application/json")
            };
            Assert.Equal(HttpStatusCode.Unauthorized, (await _client.SendAsync(oldLoginReq)).StatusCode);

            // New password succeeds -> 200
            var newLoginReq = new HttpRequestMessage(HttpMethod.Post, "/api/auth/login")
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(new LoginRequest { Username = "ivan_intern", Password = "newpass456" }),
                    Encoding.UTF8,
                    "application/json")
            };
            var newLoginResp = await _client.SendAsync(newLoginReq);
            Assert.Equal(HttpStatusCode.OK, newLoginResp.StatusCode);
            var ivanAuth = JsonSerializer.Deserialize<AuthValidateResponse>(await newLoginResp.Content.ReadAsStringAsync(), JsonOptions);
            Assert.Equal("internship-key-demo", ivanAuth!.ApiKey);
        }
    }
}
