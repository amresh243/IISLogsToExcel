using System.IO;
using System.Windows;

namespace IISLogsToExcel;

public static class Logger
{
    private static bool _loggingEnabled = true;
    private static string _logFilePath = string.Empty;

    public static bool DisableLogging
    {
        set
        {
            _loggingEnabled = !value;
        }
    }

    public static void Create(string logFile)
    {
        try
        {
            if (!_loggingEnabled)
                return;

            var logParts = logFile.Split(LogTokens.ExtensionSplitMarker);
            var extension = logParts.LastOrDefault();
            var firstPart = logFile.Replace(extension ?? "", string.Empty);

            _logFilePath = $"{firstPart}{DateTime.Now:yyyyMMdd}.{extension}";

            if (!File.Exists(_logFilePath))
            {
                using (File.Create(_logFilePath)) {}
            }
        }
        catch
        {
            _loggingEnabled = false;
            MessageBox.Show(Messages.LoggingError, Captions.LoggingError, MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    public static void LogHeader()
    {
        if (!_loggingEnabled)
            return;

        File.AppendAllText(_logFilePath, Constants.LogHeader + Environment.NewLine);
    }

    public static void LogMarker(long processingCount)
    {
        if (!_loggingEnabled)
            return;

        var marker = string.Format(Constants.LogMarker, processingCount);
        File.AppendAllText(_logFilePath, marker + Environment.NewLine);
    }

    public static void LogInfo(string message)
    {
        Log("INFO", message);
    }

    public static void LogWarning(string message)
    {
        Log("WARNING", message);
    }

    public static void LogError(string message)
    {
        Log("ERROR", message);
    }

    public static void LogException(string message, Exception ex)
    {
        string fullMessage = $"{message}\nException: {ex.Message}\nStack Trace: {ex.StackTrace}";
        Log("EXCEPTION", fullMessage);
    }

    private static void Log(string level, string message)
    {
        if (!_loggingEnabled)
            return;

        string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}";
        File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
    }
}
