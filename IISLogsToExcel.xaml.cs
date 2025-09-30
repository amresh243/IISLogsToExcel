// Author: Amresh Kumar (July 2025)

using ClosedXML.Excel;
using IISLogsToExcel.tools;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace IISLogsToExcel;

public partial class IISLogExporter : Window
{
    #region Variables

    private readonly ExcelSheetProcessor _processor;
    private readonly IniFile _iniFile = new(Constants.IniFile);
    private readonly List<LogFileItem> _logFiles = [];
    private readonly List<MenuItem> _stateBasedMenuItems = [];
    private readonly MessageDialog _messageBox;
    private readonly ContextMenu _contextMenu = new();
    private MenuItem? _menuItemProcess, _menuItemReset, _menuItemAbout;

    private string _folderName = string.Empty;
    private string _folderPath = string.Empty;

    private bool _isSingleBook = false;
    private bool _createPivot = false;
    private bool _enableLogging = true;
    private bool _isDarkMode = false;
    private bool _isProcessing = false;
    private bool _isConfirmationDlgOpen = false;

    private long _totalSize = 0;
    private long _processedSize = 0;
    private long _processingCount = 0;

    public List<LogFileItem> LogFiles => _logFiles;
    public MessageDialog MessageBox => _messageBox;

    #endregion Variables


    #region Constructor

    public IISLogExporter(string folderPath = "")
    {
        InitializeComponent();

        _processor = new ExcelSheetProcessor(this);
        _messageBox = new MessageDialog(this);
        systemTheme.IsChecked = _isDarkMode = Utility.IsSystemInDarkMode();

        LoadSettings(folderPath);
    }

    #endregion Constructor


    #region Control State Modifiers

    private bool GetBoolValue(string key) =>
        bool.Parse(_iniFile.GetValue(Constants.SettingsSection, key) ?? Constants.False);

    /// <summary> Loads settings from the INI file and initializes controls. </summary>
    /// <param name="folderPath">folder path to handle, if received from command line.</param>
    private void LoadSettings(string folderPath)
    {
        _folderPath = _iniFile.GetValue(Constants.SettingsSection, Constants.FolderPath) ?? string.Empty;
        isSingleWorkBook.IsChecked = _isSingleBook = GetBoolValue(Constants.SingleWorkbook);
        createPivotTable.IsChecked = _createPivot = GetBoolValue(Constants.CreatePivot);
        enableLogging.IsChecked = _enableLogging = GetBoolValue(Constants.EnableLogging);
        if(File.Exists(Constants.IniFile))
            systemTheme.IsChecked = _isDarkMode = GetBoolValue(Constants.DarkMode);

        if (_enableLogging)
        {
            Logger.Create(Constants.LogFile, this);
            Logger.LogInfo("Settings loaded successfully.");
        }
        else
            Logger.DisableLogging = true;

        InitializeMenu();
        InitializeTheme(_isDarkMode);

        if (!string.IsNullOrEmpty(folderPath))
            InitializeVariables(folderPath);
        else if (!string.IsNullOrEmpty(_folderPath))
            InitializeVariables(_folderPath);
        else
            _folderPath = string.Empty;
    }

    /// <summary> Updates special menu items foreground color with specific theme colors. </summary>
    private static void UpdateSepcialMenuTheme(MenuItem? menuItem, Brush foreColor)
    {
        if (menuItem == null)
            return;

        menuItem.Foreground = foreColor;
        menuItem.FontWeight = FontWeights.DemiBold;
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
        _contextMenu.Background = backColor;
        progressText.Foreground = foreColor;
        folderPathTextBox.Foreground = foreColor;
        lbLogFiles.Foreground = foreColor;
        folderPathTextBox.Foreground = foreColor;
        isSingleWorkBook.Foreground = foreColor;
        enableLogging.Foreground = foreColor;
        createPivotTable.Foreground = foreColor;
        systemTheme.Foreground = foreColor;
        groupOptions.Foreground = foreColor;
        _contextMenu.Foreground = foreColor;

        UpdateSepcialMenuTheme(_menuItemProcess, Brushes.LimeGreen);
        UpdateSepcialMenuTheme(_menuItemReset, Brushes.Goldenrod);
        UpdateSepcialMenuTheme(_menuItemAbout, appborder.BorderBrush);
        foreach (var item in _logFiles)
            item.Color = foreColor;

        lbLogFiles.Items.Refresh();
        _messageBox.ApplyTheme(backColor, foreColor);
        Logger.LogInfo("Theme initialized successfully.");
    }

    /// <summary> Returns an Image control with the specified resource image path. </summary>
    private static Image GetIcon(string iconPath, double width = 16, double height = 16) =>
        new() { Source = new BitmapImage(new Uri(iconPath)), Width = width, Height = height };

    /// <summary> Creates and adds a menu item to the context menu with specified properties. </summary>
    private MenuItem CreateMenuItem(string header, string iconPath, RoutedEventHandler clickHandler,
        bool isStateBased = false, bool isDemiBold = false, Brush? foreColor = null)
    {
        var menuItem = new MenuItem {Header = header, Icon = GetIcon(iconPath)};
        if(isDemiBold && foreColor != null)
        {
            menuItem.FontWeight = FontWeights.DemiBold;
            menuItem.Foreground = foreColor;
        }

        if(isStateBased)
            _stateBasedMenuItems.Add(menuItem);

        menuItem.Click += clickHandler;
        _contextMenu.Items.Add(menuItem);

        return menuItem;
    }

    /// <summary> Initializes context menu with required menu items and their event handlers. </summary>
    private void InitializeMenu()
    {
        Logger.LogInfo("Initializing context menu...");
        _stateBasedMenuItems.Clear();

        CreateMenuItem(MenuEntry.InputLocation, Icons.Folder, FolderPathTextBox_DblClick);
        CreateMenuItem(MenuEntry.OpenAppLog, Icons.AppLog, OpenLog_Click);
        CreateMenuItem(MenuEntry.OpenAppSettings, Icons.AppSettings, OpenSettings_Click);
        _contextMenu.Items.Add(new Separator());
        _menuItemProcess = CreateMenuItem(MenuEntry.ProcessLogs, Icons.Process, ProcessButton_Click, true, true, Brushes.LimeGreen);
        _contextMenu.Items.Add(new Separator());
        CreateMenuItem(MenuEntry.CleanOldLogs, Icons.CleanLogs, CleanLogHistory_Click);
        _menuItemReset = CreateMenuItem(MenuEntry.ResetApplication, Icons.Reset, ResetApplication_Click, true, true, Brushes.Goldenrod);
        CreateMenuItem(MenuEntry.ExitApplication, Icons.Exit, MenuItemExit_Click);
        _contextMenu.Items.Add(new Separator());
        _menuItemAbout = CreateMenuItem(MenuEntry.AboutApplication, Icons.App, AboutApplication_Click, false, true, appborder.BorderBrush);

        this.ContextMenu = _contextMenu;
        Logger.LogInfo("Context menu initialized.");
    }

    /// <summary> Changes the state of controls based on the enable parameter. </summary>
    private void ChangeControlState(bool enable)
    {
        Logger.LogInfo($"Changing control state to {(enable ? "Enabled" : "Disabled")}...");
        selectFolderButton.IsEnabled = processButton.IsEnabled = enable;
        isSingleWorkBook.IsEnabled = enableLogging.IsEnabled = enable;
        createPivotTable.IsEnabled = systemTheme.IsEnabled = enable;

        foreach (var menuItem in _stateBasedMenuItems)
            menuItem.IsEnabled = enable;

        if (enable)
            _totalSize = _processedSize = 0;

        Logger.LogInfo("Control state updated.");
    }

    /// <summary> Updates status bar with the given message. </summary>
    public void UpdateStatus(string message) =>
        Dispatcher.Invoke(() => { statusText.Text = message; });

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


    #region Utility Methods

    /// <summary> Initializes variables with the given folder path. </summary>
    private void InitializeVariables(string folderPath)
    {
        Logger.LogInfo("Initializing application...");
        progressBar.Maximum = 100;
        progressBar.Value = 0;
        _totalSize = _processedSize = 0;
        progressText.Text = Constants.ZeroPercent;
        _folderPath = _folderName = string.Empty;
        folderPathTextBox.Text = string.Empty;
        lbLogFiles.Items.Clear();
        _logFiles.Clear();

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
            var listItem = new LogFileItem
            {
                Name = Utility.GetFileNameWithoutRoot(file, _folderPath),
                ID = id++.ToString(),
                FullPath = file,
                Color = foreColor
            };
            var fileInfo = new FileInfo(file);
            listItem.ToolTip = $"{file}\nSize: {Utility.GetFormattedSize(fileInfo.Length)}\nCreated: {fileInfo.CreationTime}";
            _logFiles.Add(listItem);
            lbLogFiles.Items.Add(listItem);
        }

        Logger.LogInfo("Log list initialized.");
    }

    /// <summary> Updates the list item color for the given file. </summary>
    public void UpdateList(string file, Brush color)
    {
        var fileName = Utility.GetFileNameWithoutRoot(file, _folderPath);
        var item = _logFiles.FirstOrDefault(x => x.Name == fileName);
        if (item != null)
        {
            item.Color = color;
            Dispatcher.Invoke(() => { lbLogFiles.Items.Refresh(); });
        }
    }

    /// <summary> Saves workbook object into excel file. </summary>
    private bool SaveExcelFile(XLWorkbook workbook, string xlsFile)
    {
        if (workbook == null)
            return false;

        try
        {
            if (File.Exists(xlsFile))
            {
                Logger.LogWarning($"File {xlsFile} already exists. Deleting it before saving new data.");
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
                _messageBox.Show(string.Format(Messages.AppError, ex.Message), Captions.AppError, DialogTypes.Error);
            });

            Logger.LogException("Error while saving Excel file!", ex);
            return false;
        }
    }

    #endregion Utility Methods


    #region Thread Methods

    /// <summary> Creates separate excel file for each file under folder. </summary>
    private void CreateSeperateFiles()
    {
        Logger.LogInfo("Creating separate Excel files for each log file...");
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
            _processor.SetupLogSheet(worksheet, file);
            // Creating pivot sheet, if option enabled
            if (_createPivot)
                _processor.SetupPivotSheet(workbook, worksheet, sheetName, file);

            // Saving the workbook separate excel files
            var excelFile = $"{ExcelSheetProcessor.GetSheetName(file)}{LogTokens.ExcelExtension}";
            Logger.LogInfo($"Log file processed successfully: {file}");
            msg = string.Format(Messages.LogFileExporting, excelFile);
            Logger.LogInfo(msg);
            UpdateStatus(msg);
            bool isSuccess = SaveExcelFile(workbook, Path.Combine(_folderPath, excelFile));
        }

        _processedSize = _totalSize;
        UpdateProgress(_processedSize, false);
    }

    /// <summary> Creates single excel file with sheets as multiple files under folder. </summary>
    private void CreateSingleFile()
    {
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
            var sheetName = ExcelSheetProcessor.GetSheetName(file, false, [.. workbook.Worksheets.Select(ws => ws.Name)]);
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Creating log sheet
            _processor.SetupLogSheet(worksheet, file);
            // Creating pivot sheet, if option enabled
            if (_createPivot)
                _processor.SetupPivotSheet(workbook, worksheet, sheetName, file);

            Logger.LogInfo($"Log file processed successfully: {file}");
        }

        // Saving the workbook to a single excel file
        var excelFile = $"{_folderName}{LogTokens.ExcelExtension}";
        var msg = string.Format(Messages.LogFileExporting, excelFile);
        UpdateStatus(msg);
        Logger.LogInfo(msg);
        SaveExcelFile(workbook, Path.Combine(_folderPath, excelFile));

        _processedSize = _totalSize;
        UpdateProgress(_processedSize, false);
    }

    #endregion Thread Methods
}
