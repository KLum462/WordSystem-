using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using WordClient.Core.Configuration;
using WordClient.Core.Models;

namespace WordClient.Core.Services
{
    public class ApiClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly ClientConfig _config;
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        public ApiClient(ClientConfig? config = null, HttpClient? httpClient = null)
        {
            _config = config ?? ClientConfig.LoadFromConfiguration(Assembly.GetCallingAssembly());
            _httpClient = httpClient ?? new HttpClient();

            _httpClient.BaseAddress = new Uri(_config.ServerUrl.TrimEnd('/') + "/");
            _httpClient.Timeout = TimeSpan.FromSeconds(_config.TimeoutSeconds);
        }

        public ClientConfig Config => _config;

        public async Task<bool> ValidateAuthAsync()
        {
            var response = await ValidateAuthDetailedAsync();
            return response?.Valid ?? false;
        }

        public async Task<AuthValidateResponse?> ValidateAuthDetailedAsync()
        {
            ClientLogger.LogInfo($"Validating API key against {_httpClient.BaseAddress}api/auth/validate...");

            var request = new HttpRequestMessage(HttpMethod.Post, "api/auth/validate");
            AddApiKeyHeader(request);
            request.Content = new StringContent(
                JsonSerializer.Serialize(new AuthValidateRequest { ApiKey = _config.ApiKey }),
                Encoding.UTF8,
                "application/json");

            var response = await SendWithRetryAsync(request);
            if (response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsStringAsync();
                var result = JsonSerializer.Deserialize<AuthValidateResponse>(content, JsonOptions);
                ClientLogger.LogInfo($"Authentication successful: {result?.Message}");
                return result;
            }

            ClientLogger.LogWarning($"Authentication failed with HTTP {(int)response.StatusCode}: {response.ReasonPhrase}");
            return new AuthValidateResponse
            {
                Valid = false,
                Message = $"Authentication failed (HTTP {(int)response.StatusCode})."
            };
        }

        public async Task<List<DocumentSummaryDto>> GetDocumentsAsync()
        {
            ClientLogger.LogInfo("Fetching documents list from server...");

            var request = new HttpRequestMessage(HttpMethod.Get, "api/documents");
            AddApiKeyHeader(request);

            var response = await SendWithRetryAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Failed to list documents (HTTP {(int)response.StatusCode}): {err}");
            }

            var content = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<List<DocumentSummaryDto>>(content, JsonOptions) ?? new List<DocumentSummaryDto>();
        }

        public async Task<List<TemplateCategoryDto>> GetCategoriesAsync()
        {
            ClientLogger.LogInfo("Fetching template categories from server...");

            var request = new HttpRequestMessage(HttpMethod.Get, "api/management/categories");
            AddApiKeyHeader(request);

            var response = await SendWithRetryAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Failed to fetch categories (HTTP {(int)response.StatusCode}): {err}");
            }

            var content = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<List<TemplateCategoryDto>>(content, JsonOptions) ?? new List<TemplateCategoryDto>();
        }

        public async Task<List<string>> GetSupportedOperationsAsync()
        {
            var request = new HttpRequestMessage(HttpMethod.Get, "api/text/operations");
            AddApiKeyHeader(request);

            var response = await SendWithRetryAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                return new List<string>
                {
                    "uppercase", "lowercase", "trim", "reverse", "word_count",
                    "title_case", "capitalize", "normalize", "anonymize",
                    "extract_entities", "summarize", "stats"
                };
            }

            var content = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<List<string>>(content, JsonOptions) ?? new List<string>();
        }

        public async Task<DocumentDto> GetDocumentAsync(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                throw new ArgumentException("Document ID cannot be empty.", nameof(id));

            ClientLogger.LogInfo($"Requesting document '{id}' from server...");

            var request = new HttpRequestMessage(HttpMethod.Get, $"api/documents/{Uri.EscapeDataString(id)}");
            AddApiKeyHeader(request);

            var response = await SendWithRetryAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Failed to get document '{id}'. Server returned {(int)response.StatusCode}: {err}");
            }

            var content = await response.Content.ReadAsStringAsync();
            var doc = JsonSerializer.Deserialize<DocumentDto>(content, JsonOptions);
            if (doc == null)
            {
                throw new InvalidOperationException($"Unable to parse document JSON for '{id}'.");
            }

            ClientLogger.LogInfo($"Document '{id}' retrieved successfully ({doc.Title}).");
            return doc;
        }

        public async Task<string> ProcessSelectedTextAsync(string text, string operation = "uppercase", string? documentId = null)
        {
            if (text == null)
                throw new ArgumentNullException(nameof(text));

            ClientLogger.LogInfo($"Sending selected text ({text.Length} chars) for '{operation}' operation...");

            var requestPayload = new TextProcessRequest
            {
                DocumentId = documentId,
                Text = text,
                Operation = operation
            };

            var request = new HttpRequestMessage(HttpMethod.Post, "api/text/process");
            AddApiKeyHeader(request);
            request.Content = new StringContent(JsonSerializer.Serialize(requestPayload, JsonOptions), Encoding.UTF8, "application/json");

            var response = await SendWithRetryAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Text processing failed (HTTP {(int)response.StatusCode}): {err}");
            }

            var content = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<TextProcessResponse>(content, JsonOptions);
            return result?.Result ?? string.Empty;
        }

        public async Task<DocumentProcessResponse> ProcessDocumentAsync(string docId, string text, string operation = "uppercase")
        {
            if (string.IsNullOrEmpty(text))
                throw new ArgumentException("Document text cannot be empty.", nameof(text));

            // Enforce 100 MB limit before sending
            long byteCount = Encoding.UTF8.GetByteCount(text);
            if (byteCount > _config.MaxFileSizeBytes)
            {
                throw new InvalidOperationException(
                    $"Document size ({byteCount / 1024.0 / 1024.0:F2} MB) exceeds maximum allowed limit of {_config.MaxFileSizeBytes / 1024.0 / 1024.0:F0} MB.");
            }

            ClientLogger.LogInfo($"Sending entire document '{docId}' ({byteCount / 1024.0:F1} KB) for '{operation}' processing...");

            var requestPayload = new DocumentProcessRequest
            {
                Id = docId,
                Text = text,
                Operation = operation
            };

            var request = new HttpRequestMessage(HttpMethod.Post, "api/documents/process");
            AddApiKeyHeader(request);
            request.Content = new StringContent(JsonSerializer.Serialize(requestPayload, JsonOptions), Encoding.UTF8, "application/json");

            var response = await SendWithRetryAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Document processing failed (HTTP {(int)response.StatusCode}): {err}");
            }

            var content = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<DocumentProcessResponse>(content, JsonOptions)
                   ?? new DocumentProcessResponse();
        }

        public async Task<DocumentUploadResponse> UploadDocumentFileAsync(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException($"File not found for upload: {filePath}", filePath);

            var fileInfo = new FileInfo(filePath);
            if (!string.Equals(fileInfo.Extension, ".docx", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Only .docx files can be uploaded to the server.");
            }

            if (fileInfo.Length > _config.MaxFileSizeBytes)
            {
                throw new InvalidOperationException(
                    $"File '{fileInfo.Name}' ({fileInfo.Length / 1024.0 / 1024.0:F2} MB) exceeds maximum allowed limit of {_config.MaxFileSizeBytes / 1024.0 / 1024.0:F0} MB.");
            }

            ClientLogger.LogInfo($"Uploading .docx file '{fileInfo.Name}' ({fileInfo.Length / 1024.0:F1} KB) to server...");

            byte[] fileBytes;
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var ms = new MemoryStream())
            {
                await fs.CopyToAsync(ms);
                fileBytes = ms.ToArray();
            }

            var multipart = new MultipartFormDataContent();
            var fileContent = new ByteArrayContent(fileBytes);
            fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            multipart.Add(fileContent, "file", fileInfo.Name);

            var request = new HttpRequestMessage(HttpMethod.Post, "api/documents/upload")
            {
                Content = multipart
            };
            AddApiKeyHeader(request);

            var response = await SendWithRetryAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Document upload failed (HTTP {(int)response.StatusCode}): {err}");
            }

            var content = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<DocumentUploadResponse>(content, JsonOptions)
                   ?? new DocumentUploadResponse();
        }

        public async Task<DocumentCheckResult> CheckDocumentAsync(DocumentCheckRequest checkRequest)
        {
            ClientLogger.LogInfo("Sending document inspection request to server...");

            var request = new HttpRequestMessage(HttpMethod.Post, "api/documents/check");
            AddApiKeyHeader(request);
            request.Content = new StringContent(JsonSerializer.Serialize(checkRequest, JsonOptions), Encoding.UTF8, "application/json");

            var response = await SendWithRetryAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Document check failed (HTTP {(int)response.StatusCode}): {err}");
            }

            var content = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<DocumentCheckResult>(content, JsonOptions)
                   ?? new DocumentCheckResult();
        }

        public async Task<List<TemplateDto>> GetTemplatesAsync()
        {
            ClientLogger.LogInfo("Fetching templates from server...");

            var request = new HttpRequestMessage(HttpMethod.Get, "api/templates");
            AddApiKeyHeader(request);

            var response = await SendWithRetryAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Failed to get templates (HTTP {(int)response.StatusCode}): {err}");
            }

            var content = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<List<TemplateDto>>(content, JsonOptions) ?? new List<TemplateDto>();
        }

        public async Task<TemplateDto?> GetTemplateAsync(string id)
        {
            var request = new HttpRequestMessage(HttpMethod.Get, $"api/templates/{Uri.EscapeDataString(id)}");
            AddApiKeyHeader(request);

            var response = await SendWithRetryAsync(request);
            if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                return null;

            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Failed to get template '{id}' (HTTP {(int)response.StatusCode}): {err}");
            }

            var content = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<TemplateDto>(content, JsonOptions);
        }

        public async Task<TemplateDto> SaveTemplateAsync(TemplateDto template)
        {
            ClientLogger.LogInfo($"Saving template '{template.Name}' to server...");

            var request = new HttpRequestMessage(HttpMethod.Post, "api/templates");
            AddApiKeyHeader(request);
            request.Content = new StringContent(JsonSerializer.Serialize(template, JsonOptions), Encoding.UTF8, "application/json");

            var response = await SendWithRetryAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Failed to save template (HTTP {(int)response.StatusCode}): {err}");
            }

            var content = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<TemplateDto>(content, JsonOptions) ?? template;
        }

        public async Task<TemplateDto> UpdateTemplateAsync(string id, TemplateDto template)
        {
            var request = new HttpRequestMessage(HttpMethod.Put, $"api/templates/{Uri.EscapeDataString(id)}");
            AddApiKeyHeader(request);
            request.Content = new StringContent(JsonSerializer.Serialize(template, JsonOptions), Encoding.UTF8, "application/json");

            var response = await SendWithRetryAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException($"Failed to update template '{id}' (HTTP {(int)response.StatusCode}): {err}");
            }

            var content = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<TemplateDto>(content, JsonOptions) ?? template;
        }

        public async Task<bool> DeleteTemplateAsync(string id)
        {
            var request = new HttpRequestMessage(HttpMethod.Delete, $"api/templates/{Uri.EscapeDataString(id)}");
            AddApiKeyHeader(request);

            var response = await SendWithRetryAsync(request);
            return response.IsSuccessStatusCode;
        }

        private void AddApiKeyHeader(HttpRequestMessage request)
        {
            if (!string.IsNullOrWhiteSpace(_config.ApiKey))
            {
                request.Headers.Remove("X-API-Key");
                request.Headers.Add("X-API-Key", _config.ApiKey);
            }
        }

        private async Task<HttpResponseMessage> SendWithRetryAsync(HttpRequestMessage originalRequest)
        {
            int maxAttempts = Math.Max(1, _config.MaxRetries);
            Exception? lastException = null;
            int? lastStatusCode = null;

            // Clone request content if any, so we can re-send in retry
            byte[]? contentBytes = null;
            string? contentType = null;
            if (originalRequest.Content != null)
            {
                contentBytes = await originalRequest.Content.ReadAsByteArrayAsync();
                contentType = originalRequest.Content.Headers.ContentType?.ToString();
            }

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                HttpRequestMessage request;
                if (attempt == 1)
                {
                    request = originalRequest;
                }
                else
                {
                    request = new HttpRequestMessage(originalRequest.Method, originalRequest.RequestUri);
                    foreach (var header in originalRequest.Headers)
                    {
                        request.Headers.TryAddWithoutValidation(header.Key, header.Value);
                    }
                    if (contentBytes != null)
                    {
                        request.Content = new ByteArrayContent(contentBytes);
                        if (!string.IsNullOrEmpty(contentType))
                        {
                            request.Content.Headers.TryAddWithoutValidation("Content-Type", contentType);
                        }
                    }
                }

                try
                {
                    ClientLogger.LogInfo($"Sending {request.Method} {request.RequestUri} (Attempt {attempt}/{maxAttempts})...");
                    var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);

                    // Don't retry client errors (400, 401, 404, 413)
                    if (response.IsSuccessStatusCode || ((int)response.StatusCode >= 400 && (int)response.StatusCode < 500))
                    {
                        return response;
                    }

                    lastStatusCode = (int)response.StatusCode;
                    ClientLogger.LogWarning($"Server returned {lastStatusCode}. Attempt {attempt}/{maxAttempts} failed.");
                    response.Dispose();
                }
                catch (HttpRequestException ex)
                {
                    lastException = ex;
                    ClientLogger.LogWarning($"Network/Server request error (Attempt {attempt}/{maxAttempts}): {ex.Message}");
                }
                catch (TaskCanceledException ex)
                {
                    lastException = ex;
                    ClientLogger.LogWarning($"Request timed out after {_config.TimeoutSeconds} seconds (Attempt {attempt}/{maxAttempts}).");
                }

                if (attempt < maxAttempts)
                {
                    int delayMs = attempt * 1000;
                    ClientLogger.LogInfo($"Waiting {delayMs}ms before retry attempt {attempt + 1}...");
                    await Task.Delay(delayMs);
                }
            }

            string reason = lastException != null
                ? $"Error: {lastException.Message}"
                : (lastStatusCode.HasValue ? $"Last HTTP status: {lastStatusCode.Value}." : "Please check your network and server status.");

            throw new InvalidOperationException(
                $"Server is unavailable or unresponsive after {maxAttempts} attempts. {reason}",
                lastException);
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}
