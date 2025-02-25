using System;
using System.IO;
using System.Text;
using System.Configuration;

namespace PriceMacro.Utilities
{
    public class Logger
    {
        private static readonly object lockObj = new object();
        private static string LogDirectory => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

        public static void LogInfo(string message)
        {
            WriteLog("INFO", message);
        }

        public static void LogError(string message, Exception ex = null)
        {
            string errorMessage = ex != null
                ? $"{message}\nException: {ex.Message}\nStack Trace: {ex.StackTrace}"
                : message;
            WriteLog("ERROR", errorMessage);
        }

        public static void LogWarning(string message)
        {
            WriteLog("WARNING", message);
        }

        private static void WriteLog(string level, string message)
        {
            try
            {
                if (!Directory.Exists(LogDirectory))
                {
                    Directory.CreateDirectory(LogDirectory);
                }

                string logFile = Path.Combine(LogDirectory, $"Log_{DateTime.Now:yyyy-MM-dd}.txt");
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}{Environment.NewLine}";

                lock (lockObj)
                {
                    File.AppendAllText(logFile, logMessage, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                // If logging itself fails, write to Windows Event Log as fallback
                try
                {
                    System.Diagnostics.EventLog.WriteEntry("PriceMacro",
                        $"Failed to write to log file: {ex.Message}\nOriginal Message: {message}",
                        System.Diagnostics.EventLogEntryType.Error);
                }
                catch
                {
                    // At this point, we can't do much more than silently fail
                }
            }
        }
    }
}