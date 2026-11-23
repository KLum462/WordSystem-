using System;
using System.IO;
using System.Reflection;

namespace WordClient.Core.Services
{
    public static class ClientLogger
    {
        private static readonly object LockObj = new object();
        private static string? _resolvedLogPath;

        public static event Action<string>? OnLog;

        public static string LogFilePath => _resolvedLogPath ??= ResolveLogFilePath();

        public static void LogInfo(string message) => Log("INFO", message);
        public static void LogWarning(string message) => Log("WARN", message);
        public static void LogError(string message, Exception? ex = null)
        {
            var msg = ex != null ? $"{message} | Exception: {ex.Message}" : message;
            Log("ERROR", msg);
        }

        private static string ResolveLogFilePath()
        {
            // Prefer directory next to WordClient.Core.dll unless it's under Program Files (read-only inside WINWORD.EXE)
            var assemblyLocation = Assembly.GetExecutingAssembly().Location;
            var assemblyDir = !string.IsNullOrWhiteSpace(assemblyLocation)
                ? Path.GetDirectoryName(assemblyLocation)
                : AppDomain.CurrentDomain.BaseDirectory;

            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            bool isProtectedDir =
                string.IsNullOrWhiteSpace(assemblyDir) ||
                (!string.IsNullOrEmpty(programFiles) && assemblyDir.StartsWith(programFiles, StringComparison.OrdinalIgnoreCase)) ||
                (!string.IsNullOrEmpty(programFilesX86) && assemblyDir.StartsWith(programFilesX86, StringComparison.OrdinalIgnoreCase));

            if (!isProtectedDir && TryEnsureWritableDirectory(Path.Combine(assemblyDir!, "logs"), out var localLogDir))
            {
                return Path.Combine(localLogDir, "client.log");
            }

            var appDataDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WordSystem",
                "logs");

            if (TryEnsureWritableDirectory(appDataDir, out var fallbackDir))
            {
                return Path.Combine(fallbackDir, "client.log");
            }

            return Path.Combine(Path.GetTempPath(), "WordSystem_client.log");
        }

        private static bool TryEnsureWritableDirectory(string dir, out string verifiedDir)
        {
            verifiedDir = dir;
            try
            {
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var probePath = Path.Combine(dir, $".write_test_{Guid.NewGuid():N}.tmp");
                File.WriteAllText(probePath, string.Empty);
                File.Delete(probePath);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void Log(string level, string message)
        {
            var line = $"[{DateTime.UtcNow:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";

            try
            {
                Console.WriteLine(line);
                OnLog?.Invoke(line);

                var logPath = LogFilePath;
                var dir = Path.GetDirectoryName(logPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                lock (LockObj)
                {
                    File.AppendAllText(logPath, line + Environment.NewLine);
                }
            }
            catch
            {
                // Avoid crashing on log failure
            }
        }
    }
}
