using System;
using System.IO;
using System.Diagnostics;

namespace NetworkLabeler
{
    public static class DebugLogger
    {
        private static string _logPath;
        private static bool _isInitialized;

        public static void Initialize(string logPath = null)
        {
            if (_isInitialized) return;

            if (string.IsNullOrEmpty(logPath))
            {
                string appDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "NetworkLabeler"
                );
                Directory.CreateDirectory(appDataPath);
                _logPath = Path.Combine(appDataPath, "debug.log");
            }
            else
            {
                _logPath = logPath;
            }

            _isInitialized = true;
        }

        public static void Log(string message, LogLevel level = LogLevel.Info)
        {
            if (!_isInitialized)
            {
                Initialize();
            }

            try
            {
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}";
                File.AppendAllText(_logPath, logMessage + Environment.NewLine);
                Debug.WriteLine(logMessage);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error writing to log file: {ex.Message}");
            }
        }

        public static void LogError(string message, Exception ex = null)
        {
            string errorMessage = message;
            if (ex != null)
            {
                errorMessage += $"\nException: {ex.Message}\nStack Trace: {ex.StackTrace}";
            }
            Log(errorMessage, LogLevel.Error);
        }
    }

    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }
} 