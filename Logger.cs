using System;
using System.IO;
using System.Diagnostics;

namespace NetworkLabeler
{
    public static class Logger
    {
        private static readonly string LogFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "NetworkLabeler.log");

        static Logger()
        {
            try
            {
                // Create or clear the log file
                File.WriteAllText(LogFile, $"=== Log started at {DateTime.Now} ===\n");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to initialize log file: {ex.Message}");
            }
        }

        public static void Log(string message)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            var logMessage = $"{timestamp}: {message}";
            
            // Write to Debug output
            Debug.WriteLine(logMessage);
            Trace.WriteLine(logMessage);
            
            // Write to file
            try
            {
                File.AppendAllText(LogFile, logMessage + "\n");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to write to log file: {ex.Message}");
            }
        }

        public static void LogError(string message, Exception ex)
        {
            Log($"ERROR - {message}");
            Log($"Exception: {ex.Message}");
            Log($"Stack Trace: {ex.StackTrace}");
        }
    }
} 