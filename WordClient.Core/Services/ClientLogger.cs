using System;
using System.IO;

namespace WordClient.Core.Services
{
    public static class ClientLogger
    {
        private static readonly object LockObj = new object();
        public static event Action<string>? OnLog;

        public static void LogInfo(string message) => Log("INFO", message);
        public static void LogWarning(string message) => Log("WARN", message);
        public static void LogError(string message, Exception? ex = null)
        {
            var msg = ex != null ? $"{message} | Exception: {ex.Message}" : message;
            Log("ERROR", msg);
        }

        private static void Log(string level, string message)
        {
            var line = $"[{DateTime.UtcNow:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";

            try
            {
                Console.WriteLine(line);
                OnLog?.Invoke(line);

                var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var logPath = Path.Combine(dir, "client.log");
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
