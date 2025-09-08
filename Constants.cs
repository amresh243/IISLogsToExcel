// Author: Amresh Kumar (July 2025)

namespace IISLogsToExcel;

internal static class Constants
{
    public static string[] NumberColumns = ["s-port", "sc-status", "sc-substatus", "sc-win32-status", "sc-bytes", "cs-bytes", "time-taken"];
    public static string[] validHandlers = ["Border", "MenuItem"];

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
    public const string ZeroPercent = "0%";
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
    public const string SettingError = "Failed to save settings. Please check file permissions or path.";
    public const string ExitWarning = "Application is processing data, are you sure you want to quit?";
    public const string ConfirmReset = "Are you sure you want to reset settings to default?";
    public const string NoOldLogs = "No old log files found to delete.";
    public const string LogCleanupError = "Error encountered while cleaning old log files!";
    public const string IniWarning = "Setting file {0} doesn't exist!";
    public const string LogWarning = "Log file {0} doesn't exist!";
    public const string IISLogWarning = "IIS log file {0} doesn't exist!";
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
    public const string SettingError = "Settings Error!";
    public const string ExitWarning = "Confirm Exit";
    public const string ConfirmReset = "Confirm Reset";
    public const string LogCleanup = "Log Cleanup Summary";
    public const string IniWarning = "App Settings Missing!";
    public const string LogWarning = "App Log Missing!";
    public const string IISLogWarning = "IIS Log Missing!";
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

internal static class MenuEntry
{
    public const string InputLocation = "Open Input Location";
    public const string OpenAppLog = "Open App Log";
    public const string OpenAppSettings = "Open App Settings";
    public const string ProcessLogs = "Process Logs";
    public const string CleanOldLogs = "Clean Old Logs";
    public const string ResetApplication = "Reset Application";
    public const string ExitApplication = "Exit Application";
    public const string AboutApplication = "About IISLogsToExcel";
}

internal static class Icons
{
    public const string Info = "pack://application:,,,/res/info.png";
    public const string Warning = "pack://application:,,,/res/warning.png";
    public const string Error = "pack://application:,,,/res/error.png";
    public const string Question = "pack://application:,,,/res/question.png";
    public const string App = "pack://application:,,,/app-icon.ico";
    public const string Folder = "pack://application:,,,/res/folder.png";
    public const string Process = "pack://application:,,,/res/process.png";
    public const string CleanLogs = "pack://application:,,,/res/cleanlog.png";
    public const string Reset = "pack://application:,,,/res/reset.png";
    public const string Exit = "pack://application:,,,/res/exit.png";
    public const string AppLog = "pack://application:,,,/res/log-file.png";
    public const string AppSettings = "pack://application:,,,/res/ini-file.png";
}
