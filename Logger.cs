using System;
using System.Diagnostics;

namespace NetworkLabeler
{
    public static class Logger
    {
        public static void Log(string message)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            var logMessage = $"{timestamp}: {message}";
            Debug.WriteLine(logMessage);
        }

        public static void LogError(string message, Exception ex)
        {
            Log($"ERROR - {message}");
            Log($"Exception: {ex.Message}");
            Log($"Stack Trace: {ex.StackTrace}");
        }
    }
} 