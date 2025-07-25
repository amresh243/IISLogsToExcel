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

namespace IISLogsToExcel
{
    public partial class IISLogExporter : Window
    {
        private readonly ExcelSheetProcessor _processor;
        private readonly IniFile _iniFile = new(Constants.IniFile);
        private readonly List<LogFile> _logFiles = [];

        private bool _isSingleBook = false;
        private bool _createPivot = false;
        private bool _deleteSources = false;
        private string _folderName = string.Empty;
        private string _folderPath = string.Empty;
        private long _totalSize = 0;
        private long _processedSize = 0;
        private bool _isDarkMode = false;

        public List<LogFile> LogFiles => _logFiles;

        public IISLogExporter(string folderPath = "")
        {
            InitializeComponent();
            LoadSettings(folderPath);

            _processor = new ExcelSheetProcessor(this);

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
            _deleteSources = bool.Parse(_iniFile.GetValue(Constants.SettingsSection, Constants.DeleteSources) ?? Constants.False);
            _isDarkMode = bool.Parse(_iniFile.GetValue(Constants.SettingsSection, Constants.DarkMode) ?? Constants.False);
            _folderPath = _iniFile.GetValue(Constants.SettingsSection, Constants.FolderPath) ?? string.Empty;

            isSingleWorkBook.IsChecked = _isSingleBook;
            createPivotTable.IsChecked = _createPivot;
            deleteSourceFiles.IsChecked = _deleteSources;
            systemTheme.IsChecked = _isDarkMode;

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
            deleteSourceFiles.Foreground = foreColor;
            createPivotTable.Foreground = foreColor;
            systemTheme.Foreground = foreColor;

            foreach (var item in _logFiles)
                item.Color = foreColor;

            lbLogFiles.Items.Refresh();
        }

        /// <summary> Changes the state of controls based on the enable parameter. </summary>
        /// <param name="enable"> true=enalbe/false=disable </param>
        private void ChangeControlState(bool enable)
        {
            selectFolderButton.IsEnabled = enable;
            processButton.IsEnabled = enable;
            isSingleWorkBook.IsEnabled = enable;
            createPivotTable.IsEnabled = enable;
            deleteSourceFiles.IsEnabled = enable;

            if (enable)
                _totalSize = _processedSize = 0;
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

        // Change the Window_Closing method signature to accept nullable sender
        private void Window_Closing(object? sender, CancelEventArgs e)
        {
            _iniFile.SetValue(Constants.SettingsSection, Constants.SingleWorkbook, _isSingleBook.ToString());
            _iniFile.SetValue(Constants.SettingsSection, Constants.CreatePivot, _createPivot.ToString());
            _iniFile.SetValue(Constants.SettingsSection, Constants.DeleteSources, _deleteSources.ToString());
            _iniFile.SetValue(Constants.SettingsSection, Constants.DarkMode, systemTheme.IsChecked?.ToString() ?? Constants.False);
            _iniFile.SetValue(Constants.SettingsSection, Constants.FolderPath, _folderPath);
            _iniFile.Save();
        }

        /// <summary> DragOver event handler, only allows folder to be dropped. </summary>
        private void FolderPath_DragOver(object sender, DragEventArgs e)
        {
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
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (paths.Length > 0 && Directory.Exists(paths[0]))
                    InitializeVariables(paths[0]);
            }
        }

        /// <summary> Single workbook Checkbox click handler </summary>
        private void SingleWorkbook_Click(object sender, RoutedEventArgs e) =>
            _isSingleBook = (isSingleWorkBook.IsChecked == true);

        /// <summary> Create pivot Checkbox click handler </summary>
        private void PivotTable_Click(object sender, RoutedEventArgs e) =>
            _createPivot = (createPivotTable.IsChecked == true);

        /// <summary> Delete source files Checkbox click handler </summary>
        private void DeleteSources_Click(object sender, RoutedEventArgs e) =>
            _deleteSources = (deleteSourceFiles.IsChecked == true);

        /// <summary> Applies system theme if the checkbox is checked, otherwise applies light theme. </summary>
        private void SystemTheme_Click(object sender, RoutedEventArgs e)
        {
            _isDarkMode = (systemTheme.IsChecked == true);
            InitializeTheme(_isDarkMode);
        }

        /// <summary> Opens folder selector dialog if no selection else opens selected folder in explorer. </summary>
        private void FolderPathTextBox_DblClick(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(_folderPath))
                SelectFolderButton_Click(sender, e);
            else
                Process.Start(Constants.ExplorerApp, _folderPath);
        }

        /// <summary> Select folder button click handler </summary>
        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
                InitializeVariables(dialog.FolderName);
        }

        /// <summary> Process log button handler </summary>
        private async void ProcessButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_folderPath) || !Directory.Exists(_folderPath))
            {
                MessageBox.Show(this, Messages.InvalidInput, Captions.InvalidInput);
                return;
            }

            var logFiles = Utility.GetLogFiles(_folderPath);
            if (logFiles.Length == 0)
            {
                MessageBox.Show(this, Messages.NoLogs, Captions.NoLogs);
                return;
            }

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
            }

            Dispatcher.Invoke(() =>
            {
                statusText.Text = Messages.ProcessingCompleted;
                ChangeControlState(true);
            });
        }

        #endregion Event Handlers


        #region Utility Methods

        /// <summary> Initializes variables with the given folder path. </summary>
        /// <param name="folderPath">Source folder location.</param>
        private void InitializeVariables(string folderPath)
        {
            progressBar.Maximum = 100;
            progressBar.Value = 0;
            _totalSize = _processedSize = 0;

            if (Directory.Exists(folderPath))
            {
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
                    File.Delete(xlsFile);

                workbook.SaveAs(xlsFile);
                workbook.Dispose();

                return true;
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show(this, string.Format(Messages.AppError, ex.Message), Captions.AppError);
                });

                return false;
            }
        }

        /// <summary> Deletes the specified file if it exists and updates the list item color to indicate deletion. </summary>
        /// <param name="file">file to be deleted.</param>
        private void DeleteFile(string file)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
                UpdateList(file, Brushes.LightGray);
            }
        }

        /// <summary> Deletes  log file/s under the source folder path. </summary>
        /// <param name="file">Source file path.</param>
        private void DeleteLogFiles(string file, bool allFiles = false)
        {
            try
            {
                if (allFiles)
                {
                    var files = Utility.GetLogFiles(_folderPath);
                    foreach (var logFile in files)
                        DeleteFile(logFile);

                    return;
                }

                DeleteFile(file);
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show(this, string.Format(Messages.AppError, ex.Message), Captions.AppError);
                });
            }
        }

        #endregion Utility Methods


        #region Thread Methods

        /// <summary> Creates seperate excel file for each file under folder. </summary>
        private void CreateSeperateFiles()
        {
            var logFiles = Utility.GetLogFiles(_folderPath);

            _totalSize = logFiles.Sum(file => new FileInfo(file).Length);
            _processedSize = 0;

            foreach (var file in logFiles)
            {
                UpdateStatus(string.Format(Messages.LogFileProcessing, ExcelSheetProcessor.GetSheetName(file, true)));
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
                UpdateStatus(string.Format(Messages.LogFileExporting, excelFile));
                bool isSuccess = SaveExcelFile(workbook, Path.Combine(_folderPath, excelFile));

                // Deleting source file, if option enabled and save was successful
                if (_deleteSources && isSuccess)
                    DeleteLogFiles(file);
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

            _totalSize = logFiles.Sum(file => new FileInfo(file).Length);
            _processedSize = 0;

            foreach (var file in logFiles)
            {
                UpdateStatus(string.Format(Messages.LogFileProcessing, ExcelSheetProcessor.GetSheetName(file, true)));
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
            }

            // Saving the workbook to a single excel file
            var excelFile = $"{_folderName}{LogTokens.ExcelExtension}";
            UpdateStatus(string.Format(Messages.LogFileExporting, excelFile));
            bool isSucess = SaveExcelFile(workbook, Path.Combine(_folderPath, excelFile));

            // Deleting all source files, if option enabled and save was successful
            if (_deleteSources && isSucess)
                DeleteLogFiles(string.Empty, true);

            _processedSize = _totalSize;
            UpdateProgress(_processedSize, false);
        }

        #endregion Thread Methods
    }
}
