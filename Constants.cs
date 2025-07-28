// Author: Amresh Kumar (July 2025)

namespace IISLogsToExcel;

internal static class Constants
{
    public static string[] NumberColumns = { "s-port", "sc-status", "sc-substatus", "sc-win32-status", "sc-bytes", "cs-bytes", "time-taken" };

    public const string LogHeader = "=============================================================";
    public const string LogMarker = "#########~{0}~#########";
    public const string ApplicationName = "IISLogsToExcel";
    public const string IniFile = "IISLogsToExcel.ini";
    public const string LogFile = "IISLogsToExcel.log";
    public const string SettingsSection = "Settings";
    public const string SingleWorkbook = "SingleWorkbook";
    public const string CreatePivot = "CreatePivot";
    public const string EnableLogging = "EnableLogging";
    public const string DarkMode = "DarkMode";
    public const string FolderPath = "FolderPath";
    public const string ThemeKey = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
    public const string ThemeValue = "AppsUseLightTheme";
    public const string ExplorerApp = "explorer.exe";
    public const string False = "false";
}

internal static class LogTokens
{
    public const string LogMarker = "#Fields:";
    public const string HourFormulae = "=TEXT(B{0}, \"hh:mm\")";
    public const string DefaultLogSheet = "IIS_Logs";
    public const string PivotMarker = "Pivot_";
    public const string PivotTable = "PivotTable";
    public const string ExcelExtension = ".xlsx";

    public const char LineSplitMarker = ' ';
    public const char PathSplitMarker = '\\';
    public const char FileSplitMarker = '-';
    public const char ExtensionSplitMarker = '.';
    public const char CommentMarker = '#';
}

internal static class Messages
{
    public static string GetLogDetails(int fileCount, string folderName) =>
        $"Found {fileCount} log file{(fileCount > 1 ? "s" : string.Empty)} in the folder '{folderName}'.";

    public const string InstanceWarning = "One instance of application IISLogsToExcel.exe is already running.";
    public const string InvalidInput = "Please select a valid folder.";
    public const string NoLogs = "No log file found in the selected folder.";
    public const string AppError = "Error occurred! Message: {0}";
    public const string CreateSheet = "Creating IIS log sheet - {0}...";
    public const string CreatePivot = "Creating pivot table for sheet - {0}...";
    public const string LogError = "An error occurred at line {0} while processing log file {1}." +
                                   "\n\nExported IIS log sheet and respective pivot sheet may have incomplete and corrupt data.";
    public const string PivotError = "An error occurred while processing pivot data for sheet {0}.";
    public const string LogFileExporting = "Exporting data to excel file - {0}...";
    public const string LogFileProcessing = "Processing data for file ??{0}...";
    public const string ProcessingStarted = "Processing...";
    public const string ProcessingCompleted = "Processing complete.";
    public const string LoggingError = "Failed to initialize logging. Please check file permissions or path.";
}

internal static class Captions
{
    public const string InstanceWarning = "IIS Logs to Excel Converter";
    public const string InvalidInput = "Invalid Input";
    public const string NoLogs = "No Logs Found!";
    public const string AppError = "Application Error!";
    public const string LogError = "Log Export Error!";
    public const string PivotError = "Pivot Error!";
    public const string LoggingError = "Logging Error!";
}

internal static class Headers
{
    public const string Time = "time";
    public const string Hour = "hour";
    public const string Date = "date";
    public const string UriStem = "cs-uri-stem";
    public const string UriStemCount = "cs-uri-stem[count]";
    public const string TimeTaken = "time-taken";
    public const string TimeTakenAvg = "time-taken[avg]";
}
