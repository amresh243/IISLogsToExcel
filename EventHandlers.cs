// Author: Amresh Kumar (July 2025)

using Microsoft.Win32;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace IISLogsToExcel;

public partial class IISLogExporter : Window
{
    #region Event Handlers

    // Change the Window_Closing method signature to accept nullable sender
    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (_isProcessing)
        {
            DialogResults result = _messageBox.Show(Messages.ExitWarning, Captions.ExitWarning, DialogTypes.Question);
            if (result == DialogResults.No)
            {
                e.Cancel = true;
                return;
            }

            Logger.LogWarning("Application shutdown initiated while processing data.");
        }

        Logger.LogInfo("Saving settings before closing the application...");
        _iniFile.SetValue(Constants.SettingsSection, Constants.SingleWorkbook, _isSingleBook.ToString());
        _iniFile.SetValue(Constants.SettingsSection, Constants.CreatePivot, _createPivot.ToString());
        _iniFile.SetValue(Constants.SettingsSection, Constants.EnableLogging, _enableLogging.ToString());
        _iniFile.SetValue(Constants.SettingsSection, Constants.DarkMode, systemTheme.IsChecked?.ToString() ?? Constants.False);
        _iniFile.SetValue(Constants.SettingsSection, Constants.FolderPath, _folderPath);
        _iniFile.Save();

        Logger.LogInfo("Settings saved successfully.");
        Logger.LogInfo("Application shutting down.");
        Logger.LogHeader();
        _messageBox.Close();
    }

    /// <summary> Opens appliction folder in explorer. </summary>
    private void Application_DblClick(object sender, RoutedEventArgs e)
    {
        if (e != null && !Constants.validHandlers.Contains(e.OriginalSource.GetType().Name))
            return;

        string appDirectory = AppContext.BaseDirectory;
        Logger.LogInfo($"Opening application folder path in explorer: {appDirectory}.");
        var logFile = Logger.LogFilePath;
        var command = File.Exists(logFile) ? $"/select,\"{logFile}\"" : appDirectory;
        Process.Start(Constants.ExplorerApp, command);
    }

    /// <summary> DragOver event handler, only allows folder to be dropped. </summary>
    private void FolderPath_DragOver(object sender, DragEventArgs e)
    {
        if (_isProcessing)
        {
            e.Effects = DragDropEffects.None;
            e.Handled = true;
            return;
        }

        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            // Only allow if the first item is a directory
            e.Effects = (paths.Length > 0 && Directory.Exists(paths[0])) ? DragDropEffects.Copy : DragDropEffects.None;
        }
        else
            e.Effects = DragDropEffects.None;

        e.Handled = true;
    }

    /// <summary> Drop event handler, sets the folder path with the dropped folder path. </summary>
    private void FolderPath_Drop(object sender, DragEventArgs e)
    {
        if (_isProcessing)
        {
            Logger.LogWarning("Drag and drop operation is not allowed while processing!");
            e.Handled = true;
            return;
        }

        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (paths.Length > 0 && Directory.Exists(paths[0]))
            {
                Logger.LogInfo($"Folder {paths[0]} dropped onto the application.");
                InitializeVariables(paths[0]);
            }
        }
    }

    /// <summary> Single workbook Checkbox click handler </summary>
    private void SingleWorkbook_Click(object sender, RoutedEventArgs e)
    {
        Logger.LogInfo($"Single workbook option changed to: {(isSingleWorkBook.IsChecked == true ? "Enabled" : "Disabled")}");
        _isSingleBook = (isSingleWorkBook.IsChecked == true);
    }

    /// <summary> Create pivot Checkbox click handler </summary>
    private void PivotTable_Click(object sender, RoutedEventArgs e)
    {
        Logger.LogInfo($"Create pivot table option changed to: {(createPivotTable.IsChecked == true ? "Enabled" : "Disabled")}");
        _createPivot = (createPivotTable.IsChecked == true);
    }

    /// <summary> Delete source files Checkbox click handler </summary>
    private void EnableLogging_Click(object sender, RoutedEventArgs e)
    {
        _enableLogging = (enableLogging.IsChecked == true);
        if (!_enableLogging)
        {
            Logger.LogWarning("Logging option disabled.");
            Logger.DisableLogging = true;
        }
        else
        {
            Logger.DisableLogging = false;
            Logger.Create(Constants.LogFile, this);
            Logger.LogInfo("Logging option enabled.");
        }
    }

    /// <summary> Applies system theme if the checkbox is checked, otherwise applies light theme. </summary>
    private void SystemTheme_Click(object sender, RoutedEventArgs e)
    {
        Logger.LogInfo($"Dark mode theme option changed to: {(systemTheme.IsChecked == true ? "Enabled" : "Disabled")}");
        _isDarkMode = (systemTheme.IsChecked == true);
        InitializeTheme(_isDarkMode);
    }

    /// <summary> Opens folder selector dialog if no selection else opens selected folder in explorer. </summary>
    private void FolderPathTextBox_DblClick(object sender, RoutedEventArgs e)
    {
        //DependencyObject source = e.OriginalSource as TextBox;
        if (!Directory.Exists(_folderPath))
            SelectFolderButton_Click(sender, e);
        else
        {
            Logger.LogInfo($"Opening selected folder path in explorer: {_folderPath}.");
            Process.Start(Constants.ExplorerApp, _folderPath);
        }
    }

    /// <summary> Select folder button click handler </summary>
    private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
    {
        Logger.LogInfo("Folder selection initiated...");
        var dialog = new OpenFolderDialog();
        if (dialog.ShowDialog() == true)
            InitializeVariables(dialog.FolderName);
    }

    /// <summary> List item double click event handler </summary>
    private void ListBoxItem_DoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is ListBoxItem item)
        {
            var logFileItem = item.Content as LogFileItem;
            var file = logFileItem?.FullPath;
            if (File.Exists(file))
            {
                Logger.LogInfo($"Opening file in notepad: {file}.");
                Process.Start(Constants.NotepadApp, file);
            }
            else
                Logger.LogWarning($"File {file} doesn't exist.");
        }
    }

    /// <summary> Process log button handler </summary>
    private async void ProcessButton_Click(object sender, RoutedEventArgs e)
    {
        var stopwatch = Stopwatch.StartNew();
        if (string.IsNullOrWhiteSpace(_folderPath) || !Directory.Exists(_folderPath))
        {
            Logger.LogWarning("Invalid folder path selected!");
            _messageBox.Show(Messages.InvalidInput, Captions.InvalidInput, DialogTypes.Warning);
            return;
        }

        var logFiles = Utility.GetLogFiles(_folderPath);
        if (logFiles.Length == 0)
        {
            Logger.LogWarning($"No log files found in the selected folder {_folderPath}!");
            _messageBox.Show(Messages.NoLogs, Captions.NoLogs, DialogTypes.Warning);
            return;
        }

        Logger.LogInfo($"Processing started for {_folderPath} with {logFiles.Length} log files.");
        ChangeControlState(false);
        InitializeList(logFiles);
        statusText.Text = Messages.ProcessingStarted;

        try
        {
            _isProcessing = true;

            if (!_isSingleBook)
                await Task.Run(() => CreateSeperateFiles());
            else
                await Task.Run(() => CreateSingleFile());

            _isProcessing = false;
        }
        catch (Exception ex)
        {
            _messageBox.Show(string.Format(Messages.AppError, ex.Message), Captions.AppError, DialogTypes.Error);
            Logger.LogException("Error while processing log files!", ex);
        }

        Dispatcher.Invoke(() =>
        {
            statusText.Text = Messages.ProcessingCompleted;
            ChangeControlState(true);
        });

        stopwatch.Stop();

        if(!_isProcessing)
            _messageBox.Hide();
        else
            _isProcessing = false;

        Logger.LogInfo($"Processing completed successfully in {stopwatch.Elapsed.TotalSeconds} seconds.");
        Logger.LogMarker(++_processingCount);
    }

    /// <summary> Right click app event handler, shows context menu. </summary>
    private void Window_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        Point position = e.GetPosition(this);
        
        _contextMenu.PlacementTarget = this;
        _contextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Relative;
        _contextMenu.HorizontalOffset = position.X;
        _contextMenu.VerticalOffset = position.Y;
        _contextMenu.IsOpen = true;
    }

    /// <summary> Cleans old log files except today's log file. </summary>
    private void CleanLogHistory_Click(object sender, RoutedEventArgs e)
    {
        Logger.LogInfo("Cleaning old log files...");
        string appDirectory = AppContext.BaseDirectory;
        string today = DateTime.Now.ToString("yyyyMMdd");
        var logFiles = Utility.GetLogFiles(appDirectory).Where(x => x.Contains(Constants.ApplicationName) && !x.Contains(today)).ToList();

        if (logFiles.Count == 0)
        {
            Logger.LogInfo(Messages.NoOldLogs);
            _messageBox.Show(Messages.NoOldLogs, Captions.LogCleanup, DialogTypes.Info);
            return;
        }

        try
        {
            foreach (var file in logFiles)
            {
                File.Delete(file);
                Logger.LogInfo($"Deleted old log file: {file}");
            }

            var msg = $"Old log files cleanup completed. Total files deleted: {logFiles.Count}";
            _messageBox.Show(msg, Captions.LogCleanup, DialogTypes.Info);
            Logger.LogInfo(msg);
        }
        catch (Exception ex)
        {
            Logger.LogException("Error encountered while cleaning old log files!", ex);
            _messageBox.Show(string.Format(Messages.LogCleanupError, ex.Message), Captions.LogCleanup, DialogTypes.Error);
        }
    }

    /// <summary> Resets application settings to default values. </summary>
    private void ResetApplication_Click(object sender, RoutedEventArgs e)
    {
        Logger.LogInfo("Resetting application...");

        DialogResults result = _messageBox.Show(Messages.ConfirmReset, Captions.ConfirmReset, DialogTypes.Question);
        if (result == DialogResults.No)
        {
            Logger.LogInfo("Application reset canceled.");
            return;
        }

        InitializeVariables(string.Empty);
        InitializeTheme(false);

        isSingleWorkBook.IsChecked = false;
        enableLogging.IsChecked = false;
        createPivotTable.IsChecked = false;
        systemTheme.IsChecked = false;
        _isSingleBook = false;
        _enableLogging = false;
        _createPivot = false;
        _isDarkMode = false;

        UpdateStatus("Application reset completed.");
        Logger.LogInfo("Application reset completed.");
    }

    /// <summary> Shows about application dialog. </summary>
    private void AboutApplication_Click(object sender, RoutedEventArgs e)
    {

        string version = Assembly.GetExecutingAssembly()?.GetName()?.Version?.ToString() ?? "";
        string message = "About IIS Logs To Excel Converter...\n\n";
        message += $"IISLogsToExcel Version: {version}\n";
        message += $"Copyright © {DateTime.Now.Year} Amresh Kumar\n";
        message += $"Write to kumar.anirudha@gmail.com";

        _messageBox.Show(message, "About IISLogsToExcel", DialogTypes.Info, GetIcon("/app-icon.ico", 48, 48));
    }

    /// <summary> Menu item exit event handler, closes application. </summary>
    private void MenuItemExit_Click(object sender, RoutedEventArgs e)
    {
        this.Close();
    }

    #endregion Event Handlers
}
