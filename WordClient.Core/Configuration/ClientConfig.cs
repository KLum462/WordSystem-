using System;
using System.Configuration;

namespace WordClient.Core.Configuration
{
    public class ClientConfig
    {
        public string ServerUrl { get; set; } = "http://localhost:5000";
        public string ApiKey { get; set; } = "secret-internship-api-key-2026";
        public int TimeoutSeconds { get; set; } = 30;
        public int MaxRetries { get; set; } = 3;
        public long MaxFileSizeBytes { get; set; } = 104857600; // 100 MB

        public static ClientConfig LoadFromConfiguration()
        {
            var config = new ClientConfig();

            try
            {
                var serverUrl = ConfigurationManager.AppSettings["ServerUrl"];
                if (!string.IsNullOrWhiteSpace(serverUrl))
                    config.ServerUrl = serverUrl.Trim();

                var apiKey = ConfigurationManager.AppSettings["ApiKey"];
                if (!string.IsNullOrWhiteSpace(apiKey))
                    config.ApiKey = apiKey.Trim();

                var timeoutStr = ConfigurationManager.AppSettings["TimeoutSeconds"];
                if (int.TryParse(timeoutStr, out int timeout) && timeout > 0)
                    config.TimeoutSeconds = timeout;

                var retriesStr = ConfigurationManager.AppSettings["MaxRetries"];
                if (int.TryParse(retriesStr, out int retries) && retries >= 0)
                    config.MaxRetries = retries;

                var maxFileSizeStr = ConfigurationManager.AppSettings["MaxFileSizeBytes"];
                if (long.TryParse(maxFileSizeStr, out long maxFileSize) && maxFileSize > 0)
                    config.MaxFileSizeBytes = maxFileSize;
            }
            catch
            {
                // Fallback to default values if app.config is unavailable or cannot be read
            }

            return config;
        }
    }
}
