// Author: Amresh Kumar (July 2025)

using ClosedXML.Excel;
using Microsoft.Win32;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace IISLogsToExcel;

public partial class IISLogExporter : Window
{
    private readonly ExcelSheetProcessor _processor;
    private readonly IniFile _iniFile = new(Constants.IniFile);
    private readonly List<LogFile> _logFiles = [];

    private string _folderName = string.Empty;
    private string _folderPath = string.Empty;

    private bool _isSingleBook = false;
    private bool _createPivot = false;
    private bool _enableLogging = true;
    private bool _isDarkMode = false;
    private bool _isProcessing = false;

    private long _totalSize = 0;
    private long _processedSize = 0;

    public List<LogFile> LogFiles => _logFiles;

    public IISLogExporter(string folderPath = "")
    {
        InitializeComponent();

        _processor = new ExcelSheetProcessor(this);

        LoadSettings(folderPath);

        if (!string.IsNullOrEmpty(folderPath))
            InitializeVariables(folderPath);
    }


    #region Control State Modifiers

    /// <summary> Loads settings from the INI file and initializes controls. </summary>
    /// <param name="folderPath">folder path to handle, if received from command line.</param>
    private void LoadSettings(string folderPath)
    {
        _isSingleBook = bool.Parse(_iniFile.GetValue(Constants.SettingsSection, Constants.SingleWorkbook) ?? Constants.False);
        _createPivot = bool.Parse(_iniFile.GetValue(Constants.SettingsSection, Constants.CreatePivot) ?? Constants.False);
        _enableLogging = bool.Parse(_iniFile.GetValue(Constants.SettingsSection, Constants.EnableLogging) ?? Constants.False);
        _isDarkMode = bool.Parse(_iniFile.GetValue(Constants.SettingsSection, Constants.DarkMode) ?? Constants.False);
        _folderPath = _iniFile.GetValue(Constants.SettingsSection, Constants.FolderPath) ?? string.Empty;

        isSingleWorkBook.IsChecked = _isSingleBook;
        createPivotTable.IsChecked = _createPivot;
        enableLogging.IsChecked = _enableLogging;
        systemTheme.IsChecked = _isDarkMode;

        if (_enableLogging)
        {
            Logger.Create(Constants.LogFile);
            Logger.LogInfo("Settings loaded successfully.");
        }
        else
            Logger.DisableLogging = true;

        InitializeTheme(_isDarkMode);

        if (!string.IsNullOrEmpty(folderPath))
            InitializeVariables(folderPath);
        else if (!string.IsNullOrEmpty(_folderPath))
            InitializeVariables(_folderPath);
        else
            _folderPath = string.Empty;

    }

    /// <summary> Changes controls background and foreground based on system theme. </summary>
    private void InitializeTheme(bool isDarkMode)
    {
        Logger.LogInfo($"Initializing theme: {(isDarkMode ? "Dark Mode" : "Light Mode")}...");
        var foreColor = (isDarkMode) ? Brushes.White : Brushes.Black;
        var backColor = (isDarkMode) ? Brushes.Black : Brushes.White;

        this.Background = backColor;
        lbLogFiles.Background = backColor;
        progressBar.Background = backColor;
        folderPathTextBox.Background = backColor;
        progressText.Foreground = foreColor;
        folderPathTextBox.Foreground = foreColor;
        lbLogFiles.Foreground = foreColor;
        folderPathTextBox.Foreground = foreColor;
        isSingleWorkBook.Foreground = foreColor;
        enableLogging.Foreground = foreColor;
        createPivotTable.Foreground = foreColor;
        systemTheme.Foreground = foreColor;

        foreach (var item in _logFiles)
            item.Color = foreColor;

        lbLogFiles.Items.Refresh();
        Logger.LogInfo("Theme initialized successfully.");
    }

    /// <summary> Changes the state of controls based on the enable parameter. </summary>
    /// <param name="enable"> true=enalbe/false=disable </param>
    private void ChangeControlState(bool enable)
    {
        Logger.LogInfo($"Changing control state to {(enable ? "Enabled" : "Disabled")}...");
        selectFolderButton.IsEnabled = enable;
        processButton.IsEnabled = enable;
        isSingleWorkBook.IsEnabled = enable;
        createPivotTable.IsEnabled = enable;
        enableLogging.IsEnabled = enable;

        if (enable)
            _totalSize = _processedSize = 0;

        Logger.LogInfo("Control state updated.");
    }

    /// <summary> Updates status bar with the given message. </summary>
    /// <param name="message"> Message to be displayed </param>
    public void UpdateStatus(string message)
    {
        Dispatcher.Invoke(() =>
        {
            statusText.Text = message;
        });
    }

    /// <summary> Updates progress status on the progress bar. </summary>
    public void UpdateProgress(long progressedSize, bool addProgress = true)
    {
        if (addProgress)
            _processedSize += progressedSize;

        Dispatcher.Invoke(() =>
        {
            var progressValue = (_processedSize * 100) / _totalSize;
            progressBar.Value = progressValue;
            progressText.Text = $"{progressValue}%";
        });
    }

    #endregion Control State Modifiers


    #region Event Handlers

    // Change the Window_Closing method signature to accept nullable sender
    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        Logger.LogInfo("Saving settings before closing the application...");
        _iniFile.SetValue(Constants.SettingsSection, Constants.SingleWorkbook, _isSingleBook.ToString());
        _iniFile.SetValue(Constants.SettingsSection, Constants.CreatePivot, _createPivot.ToString());
        _iniFile.SetValue(Constants.SettingsSection, Constants.EnableLogging, _enableLogging.ToString());
        _iniFile.SetValue(Constants.SettingsSection, Constants.DarkMode, systemTheme.IsChecked?.ToString() ?? Constants.False);
        _iniFile.SetValue(Constants.SettingsSection, Constants.FolderPath, _folderPath);
        _iniFile.Save();
        Logger.LogInfo("Settings saved successfully.");
        Logger.LogInfo("Application shutting down.");
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
    /// <param name="sender"></param>
    /// <param name="e"></param>
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
            Logger.LogWarning("Logging option disabled");
            Logger.DisableLogging = true;
        }
        else
        {
            Logger.DisableLogging = false;
            Logger.Create(Constants.LogFile);
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
        if (!Directory.Exists(_folderPath))
            SelectFolderButton_Click(sender, e);
        else
        {
            Logger.LogInfo($"Opening folder in explorer: {_folderPath}.");
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

    /// <summary> Process log button handler </summary>
    private async void ProcessButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_folderPath) || !Directory.Exists(_folderPath))
        {
            Logger.LogWarning("Invalid folder path selected!");
            MessageBox.Show(this, Messages.InvalidInput, Captions.InvalidInput);
            return;
        }

        var logFiles = Utility.GetLogFiles(_folderPath);
        if (logFiles.Length == 0)
        {
            Logger.LogWarning($"No log files found in the selected folder {_folderPath}!");
            MessageBox.Show(this, Messages.NoLogs, Captions.NoLogs);
            return;
        }

        Logger.LogInfo($"Processing started for {_folderPath} with {logFiles.Length} log files.");
        ChangeControlState(false);
        InitializeList(logFiles);
        statusText.Text = Messages.ProcessingStarted;
        
        try
        {
            if (!_isSingleBook)
                await Task.Run(() => CreateSeperateFiles());
            else
                await Task.Run(() => CreateSingleFile());
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, string.Format(Messages.AppError, ex.Message), Captions.AppError);
            Logger.LogException("Error while processing log files!", ex);
        }

        Dispatcher.Invoke(() =>
        {
            statusText.Text = Messages.ProcessingCompleted;
            ChangeControlState(true);
        });

        Logger.LogInfo("Processing completed successfully.");
        Logger.LogHeader();
    }

    #endregion Event Handlers


    #region Utility Methods

    /// <summary> Initializes variables with the given folder path. </summary>
    /// <param name="folderPath">Source folder location.</param>
    private void InitializeVariables(string folderPath)
    {
        Logger.LogInfo("Initializing application...");
        progressBar.Maximum = 100;
        progressBar.Value = 0;
        _totalSize = _processedSize = 0;
        progressText.Text = $"0%";

        if (Directory.Exists(folderPath))
        {
            Logger.LogInfo($"Log folder selected: {folderPath}");
            _folderPath = folderPath;
            folderPathTextBox.Text = _folderPath;
            _folderName = _folderPath.Split(LogTokens.PathSplitMarker, StringSplitOptions.None).Last();
            var logFiles = Utility.GetLogFiles(_folderPath);
            var logFileCount = logFiles.Length;
            InitializeList(logFiles);
            UpdateStatus(Messages.GetLogDetails(logFileCount, _folderName));
        }
    }

    /// <summary> Initiates list with log files found in the selected folder. </summary>
    /// <param name="logFiles">list of log files</param>
    private void InitializeList(string[] logFiles)
    {
        Logger.LogInfo($"Initializing log list with {logFiles.Length} log files...");
        var foreColor = _isDarkMode ? Brushes.White : Brushes.Black;
        _logFiles.Clear();
        lbLogFiles.Items.Clear();
        int id = 1;

        foreach (var file in logFiles)
        {
            var fileName = ExcelSheetProcessor.GetSheetName(file, true);
            var listItem = new LogFile { Name = file, ID = id++.ToString(), Color = foreColor };
            _logFiles.Add(listItem);
            lbLogFiles.Items.Add(listItem);
        }

        Logger.LogInfo("Log list initialized.");
    }

    /// <summary> Updates the list item color for the given file. </summary>
    /// <param name="file">file name, to find list item</param>
    /// <param name="color">forecolor to be set</param>
    public void UpdateList(string file, Brush color)
    {
        var item = _logFiles.FirstOrDefault(x => x.Name == file);
        if (item != null)
        {
            item.Color = color;
            Dispatcher.Invoke(() =>
            {
                lbLogFiles.Items.Refresh();
            });
        }
    }

    /// <summary> Saves workbook object into excel file. </summary>
    /// <param name="workbook">Workbook object, excel file object</param>
    /// <param name="xlsFile">Excel file name to be saved</param>
    private bool SaveExcelFile(XLWorkbook workbook, string xlsFile)
    {
        if (workbook == null)
            return false;

        try
        {
            if (File.Exists(xlsFile))
            {
                Logger.LogInfo($"File {xlsFile} already exists. Deleting it before saving new data.");
                File.Delete(xlsFile);
            }

            workbook.SaveAs(xlsFile);
            Logger.LogInfo($"Excel file saved successfully: {xlsFile}");
            workbook.Dispose();

            return true;
        }
        catch (Exception ex)
        {
            Dispatcher.Invoke(() =>
            {
                MessageBox.Show(this, string.Format(Messages.AppError, ex.Message), Captions.AppError);
            });

            Logger.LogException("Error while saving Excel file!", ex);
            return false;
        }
    }

    #endregion Utility Methods


    #region Thread Methods

    /// <summary> Creates seperate excel file for each file under folder. </summary>
    private void CreateSeperateFiles()
    {
        Logger.LogInfo("Creating separate Excel files for each log file...");
        _isProcessing = true;
        var logFiles = Utility.GetLogFiles(_folderPath);
        Logger.LogInfo($"Found {logFiles.Length} log files in the folder '{_folderPath}'.");

        _totalSize = logFiles.Sum(file => new FileInfo(file).Length);
        Logger.LogInfo($"Total size of log data to be processed: {_totalSize} bytes.");
        _processedSize = 0;

        foreach (var file in logFiles)
        {
            var msg = string.Format(Messages.LogFileProcessing, ExcelSheetProcessor.GetSheetName(file, true));
            Logger.LogInfo(msg);
            UpdateStatus(msg);
            UpdateList(file, Brushes.LimeGreen);
            
            var workbook = new XLWorkbook();
            var sheetName = (!_isSingleBook) ? LogTokens.DefaultLogSheet : file;
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Creating log sheet
            _processor.SetupLogData(worksheet, file);

            // Creating pivot sheet, if option enabled
            if (_createPivot)
                _processor.SetupPivotData(workbook, worksheet, sheetName, file);

            // Saving the workbook seperate excel files
            var excelFile = $"{ExcelSheetProcessor.GetSheetName(file)}{LogTokens.ExcelExtension}";
            Logger.LogInfo($"Log file processed successfully: {file}");
            msg = string.Format(Messages.LogFileExporting, excelFile);
            Logger.LogInfo(msg);
            UpdateStatus(msg);
            bool isSuccess = SaveExcelFile(workbook, Path.Combine(_folderPath, excelFile));
        }

        _processedSize = _totalSize;
        UpdateProgress(_processedSize, false);
        _isProcessing = false;
    }

    /// <summary> Creates single excel file with sheets as multiple files under folder. </summary>
    private void CreateSingleFile()
    {
        _isProcessing = true;
        var sheetCount = 0;
        var workbook = new XLWorkbook();
        var logFiles = Utility.GetLogFiles(_folderPath);
        Logger.LogInfo($"Found {logFiles.Length} log files in the folder '{_folderPath}'.");

        _totalSize = logFiles.Sum(file => new FileInfo(file).Length);
        Logger.LogInfo($"Total size of log data to be processed: {_totalSize} bytes.");
        _processedSize = 0;

        foreach (var file in logFiles)
        {
            var message = string.Format(Messages.LogFileProcessing, ExcelSheetProcessor.GetSheetName(file, true));
            UpdateStatus(message);
            Logger.LogInfo(message);
            UpdateList(file, Brushes.LimeGreen);

            sheetCount++;
            var sheetName = ExcelSheetProcessor.GetSheetName(file);
            var sheetNames = workbook.Worksheets.Select(ws => ws.Name).ToList();
            var existingCount = sheetNames.Count(name => name == sheetName);
            if (existingCount > 0)
                sheetName += $"{LogTokens.FileSplitMarker}{existingCount + 1}";

            var worksheet = workbook.Worksheets.Add(sheetName);

            // Creating log sheet
            _processor.SetupLogData(worksheet, file);

            // Creating pivot sheet, if option enabled
            if (_createPivot)
                _processor.SetupPivotData(workbook, worksheet, sheetName, file);

            Logger.LogInfo($"Log file processed successfully: {file}");
        }

        // Saving the workbook to a single excel file
        var excelFile = $"{_folderName}{LogTokens.ExcelExtension}";
        var msg = string.Format(Messages.LogFileExporting, excelFile);
        UpdateStatus(msg);
        Logger.LogInfo(msg);
        bool isSucess = SaveExcelFile(workbook, Path.Combine(_folderPath, excelFile));

        _processedSize = _totalSize;
        UpdateProgress(_processedSize, false);
        _isProcessing = false;
    }

    #endregion Thread Methods
}
