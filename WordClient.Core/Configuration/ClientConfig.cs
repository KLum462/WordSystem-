using System;
using System.Configuration;
using System.IO;
using System.Reflection;

namespace WordClient.Core.Configuration
{
    public class ClientConfig
    {
        public string ServerUrl { get; set; } = "http://localhost:5000";
        public string ApiKey { get; set; } = "secret-internship-api-key-2026";
        public int TimeoutSeconds { get; set; } = 30;
        public int MaxRetries { get; set; } = 3;
        public long MaxFileSizeBytes { get; set; } = 104857600; // 100 MB

        public static ClientConfig LoadFromConfiguration(Assembly? sourceAssembly = null)
        {
            var config = new ClientConfig();

            try
            {
                // 1. Prefer config file located next to the add-in / calling assembly
                //    (critical when hosted inside WINWORD.EXE where AppDomain config is WINWORD.EXE.config)
                var candidatePaths = GetCandidateConfigPaths(sourceAssembly ?? Assembly.GetCallingAssembly());
                foreach (var configPath in candidatePaths)
                {
                    if (!string.IsNullOrWhiteSpace(configPath) && File.Exists(configPath))
                    {
                        if (TryApplyFromMappedConfig(config, configPath))
                        {
                            return config;
                        }
                    }
                }

                // 2. Fallback to standard process App.config (e.g., WordClient.TestApp.dll.config)
                ApplySettings(config, key => ConfigurationManager.AppSettings[key]);
            }
            catch
            {
                // Fallback to default values if app.config is unavailable or cannot be read
            }

            return config;
        }

        public static ClientConfig LoadFromFile(string configFilePath)
        {
            var config = new ClientConfig();
            if (!string.IsNullOrWhiteSpace(configFilePath) && File.Exists(configFilePath))
            {
                TryApplyFromMappedConfig(config, configFilePath);
            }
            return config;
        }

        private static string[] GetCandidateConfigPaths(Assembly? callingAssembly)
        {
            var executingAssembly = Assembly.GetExecutingAssembly();
            var paths = new System.Collections.Generic.List<string>();

            void AddAssemblyCandidates(Assembly? asm)
            {
                if (asm == null || string.IsNullOrWhiteSpace(asm.Location)) return;
                paths.Add(asm.Location + ".config");
                var dir = Path.GetDirectoryName(asm.Location);
                if (!string.IsNullOrEmpty(dir))
                {
                    paths.Add(Path.Combine(dir, "App.config"));
                }
            }

            AddAssemblyCandidates(callingAssembly);
            if (callingAssembly != executingAssembly)
            {
                AddAssemblyCandidates(executingAssembly);
            }

            return paths.ToArray();
        }

        private static bool TryApplyFromMappedConfig(ClientConfig config, string configPath)
        {
            try
            {
                var fileMap = new ExeConfigurationFileMap { ExeConfigFilename = configPath };
                var mappedConfig = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);
                var appSettings = mappedConfig.AppSettings?.Settings;
                if (appSettings == null || appSettings.Count == 0)
                {
                    return false;
                }

                ApplySettings(config, key => appSettings[key]?.Value);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void ApplySettings(ClientConfig config, Func<string, string?> getValue)
        {
            var serverUrl = getValue("ServerUrl");
            if (!string.IsNullOrWhiteSpace(serverUrl))
                config.ServerUrl = serverUrl!.Trim();

            var apiKey = getValue("ApiKey");
            if (!string.IsNullOrWhiteSpace(apiKey))
                config.ApiKey = apiKey!.Trim();

            var timeoutStr = getValue("TimeoutSeconds");
            if (int.TryParse(timeoutStr, out int timeout) && timeout > 0)
                config.TimeoutSeconds = timeout;

            var retriesStr = getValue("MaxRetries");
            if (int.TryParse(retriesStr, out int retries) && retries >= 0)
                config.MaxRetries = retries;

            var maxFileSizeStr = getValue("MaxFileSizeBytes");
            if (long.TryParse(maxFileSizeStr, out long maxFileSize) && maxFileSize > 0)
                config.MaxFileSizeBytes = maxFileSize;
        }
    }
}
