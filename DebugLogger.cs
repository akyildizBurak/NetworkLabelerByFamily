using System;
using System.Diagnostics;

namespace NetworkLabeler
{
    public static class DebugLogger
    {
        public static void Initialize(string logPath = null)
        {
            // No initialization needed
        }

        public static void Log(string message, LogLevel level = LogLevel.Info)
        {
            string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}";
            Debug.WriteLine(logMessage);
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