// Author: Amresh Kumar (July 2025)

using System.IO;
using System.Windows;

namespace IISLogsToExcel.tools;

public static class Logger
{
    private static bool _loggingEnabled = true;
    private static string _logFilePath = string.Empty;

    public static string LogFilePath
    {
        get
        {
            if (!string.IsNullOrEmpty(_logFilePath))
                return _logFilePath;

            return GetComputedLogFile(Constants.LogFile);
        }
    }

    public static bool DisableLogging
    {
        set => _loggingEnabled = !value;
    }

    private static string GetComputedLogFile(string logFile)
    {
        var logParts = logFile.Split(LogTokens.ExtensionSplitMarker);
        var extension = logParts.LastOrDefault();
        var firstPart = logFile.Replace(extension ?? string.Empty, string.Empty);

        return $"{firstPart}{DateTime.Now:yyyyMMdd}.{extension}";
    }

    private static void Initialize(string logFile)
    {
        if (!_loggingEnabled)
            return;

        _logFilePath = GetComputedLogFile(logFile);
        if (!File.Exists(_logFilePath))
            using (File.Create(_logFilePath)) { }
    }

    public static void Create(string logFile, IISLogExporter? app = null)
    {
        try
        {
            Initialize(logFile);
        }
        catch
        {
            _loggingEnabled = false;
            if (app == null)
                MessageBox.Show(Messages.LoggingError, Captions.LoggingError, MessageBoxButton.OK, MessageBoxImage.Warning);
            else
                app.MessageBox.Show(Messages.LoggingError, Captions.LoggingError, DialogTypes.Warning);
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

    public static void LogInfo(string message) =>
        Log("INFO", message);

    public static void LogWarning(string message) =>
        Log("WARNING", message);

    public static void LogError(string message) =>
        Log("ERROR", message);

    public static void LogException(string message, Exception ex) =>
        Log("EXCEPTION", $"{message}\nException: {ex?.Message}\nStack Trace: {ex?.StackTrace}");

    private static void Log(string level, string message)
    {
        if (!_loggingEnabled)
            return;

        string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}";
        File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
    }
}
