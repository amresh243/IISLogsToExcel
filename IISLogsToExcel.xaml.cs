// Author: Amresh Kumar

using ClosedXML.Excel;
using IISLogsToExcel;
using Microsoft.Win32;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace IISLogToExcelConverter
{
    public partial class IISLogExporter : Window
    {
        private readonly ExcelSheetProcessor _processor;

        private bool _isSingleBook = false;
        private bool _createPivot = false;
        private bool _deleteSources = false;
        private string _folderName = string.Empty;
        private string _folderPath = string.Empty;
        private long _totalSize = 0;
        private long _processedSize = 0;

        public IISLogExporter(string folderPath = "")
        {
            InitializeComponent();

            _processor = new ExcelSheetProcessor(this);
            if (!string.IsNullOrEmpty(folderPath))
                InitVariables(folderPath);
        }

        #region Control State Modifiers

        /// <summary> Changes the state of controls based on the enable parameter. </summary>
        /// <param name="enable"> true=enalbe/false=disable </param>
        private void ChangeControlState(bool enable)
        {
            selectFolderButton.IsEnabled = enable;
            processButton.IsEnabled = enable;
            isSingleWorkBook.IsEnabled = enable;
            createPivotTable.IsEnabled = enable;
            deleteSourceFiles.IsEnabled = enable;

            if(enable)
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
            if(addProgress)
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

        /// <summary> DragOver event handler, only allows folder to be dropped. </summary>
        private void FolderPath_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                // Only allow if the first item is a directory
                e.Effects = (paths.Length > 0 && Directory.Exists(paths[0]))
                    ? DragDropEffects.Copy
                    : DragDropEffects.None;
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
                    InitVariables(paths[0]);
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

        /// <summary> Opens folder selector dialog if no selection else opens selected folder in explorer. </summary>
        private void FolderPathTextBox_DblClick(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(_folderPath))
                SelectFolderButton_Click(sender, e);
            else
                Process.Start("explorer.exe", _folderPath);
        }

        /// <summary> Select folder button click handler </summary>
        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
                InitVariables(dialog.FolderName);
        }

        /// <summary> Process log button handler </summary>
        private async void ProcessButton_Click(object sender, RoutedEventArgs e)
        {
            _folderPath = folderPathTextBox.Text;
            if (string.IsNullOrWhiteSpace(_folderPath) || !Directory.Exists(_folderPath))
            {
                MessageBox.Show(this, "Please select a valid folder.", "Invalid Input!");
                return;
            }

            if(!GetLogFiles(_folderPath).Any())
            {
                MessageBox.Show(this, "No log files found in the selected folder.", "No Logs Found!");
                return;
            }

            ChangeControlState(false);
            statusText.Text = "Processing...";

            try
            {
                if (!_isSingleBook)
                    await Task.Run(() => CreateSeperateFiles());
                else
                    await Task.Run(() => CreateSingleFile());
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Error occurred! Message: {ex.Message}", "Application Error");
            }

            Dispatcher.Invoke(() =>
            {
                statusText.Text = "Processing complete.";
                ChangeControlState(true);
            });
        }

        #endregion Event Handlers


        #region Utility Methods

        /// <summary> Initializes variables with the given folder path. </summary>
        /// <param name="folderPath">Source folder location.</param>
        private void InitVariables(string folderPath)
        {
            progressBar.Maximum = 100;
            progressBar.Value = 0;
            _totalSize = _processedSize = 0;
            _folderPath = folderPath;
            folderPathTextBox.Text = _folderPath;
            _folderName = _folderPath.Split('\\', StringSplitOptions.None).Last();
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
                    MessageBox.Show(this, $"Error occurred! Message: {ex.Message}", "Application Error");
                });

                return false;
            }
        }

        /// <summary> Deletes  log file/s under the source folder path. </summary>
        /// <param name="file">Source file path.</param>
        private void DeleteLogFiles(string file, bool allFiles = false)
        {
            try
            {
                if(allFiles)
                {
                    var files = GetLogFiles(_folderPath);
                    foreach (var logFile in files)
                    {
                        if (File.Exists(logFile))
                            File.Delete(logFile);
                    }
                }
                else if (File.Exists(file))
                    File.Delete(file);
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show(this, $"Error occurred! Message: {ex.Message}", "Application Error");
                });
            }
        }

        /// <summary> Returns all log files under the given folder path. </summary>
        /// <param name="folderPath">Log folder path.</param>
        /// <returns>Array of list file paths.</returns>
        private static string[] GetLogFiles(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                return [];

            return Directory.GetFiles(folderPath, "*.log", SearchOption.AllDirectories);
        }

        /// <summary> Returns sheet name from file name. </summary>
        /// <param name="file">file with path</param>
        /// <returns>sheet name</returns>
        private static string GetSheetName(string file, bool isFile = false)
        {
            if (string.IsNullOrEmpty(file))
                return file;

            if(isFile)
            {
                var fileName = file.Split('\\').LastOrDefault()?.Split('-').LastOrDefault() ?? "";
                var fileNameLength = fileName.Length;

                return (fileNameLength > 10) ? fileName[(fileNameLength - 10)..] : fileName;
            }

            var sheetName = file.Split('\\').LastOrDefault()?.Split('-').LastOrDefault()?.Split('.').FirstOrDefault();
            if (string.IsNullOrEmpty(sheetName))
                return file;

            var sheetNameLength = sheetName.Length;
            return (sheetNameLength > 6) ? sheetName[(sheetNameLength - 6)..] : sheetName;
        }

        #endregion Utility Methods


        #region Thread Methods

        /// <summary> Creates seperate excel file for each file under folder. </summary>
        private void CreateSeperateFiles()
        {
            var logFiles = GetLogFiles(_folderPath);

            _totalSize = logFiles.Sum(file => new FileInfo(file).Length);
            _processedSize = 0;

            foreach (var file in logFiles)
            {
                UpdateStatus($"Processing data for file ??{GetSheetName(file, true)}...");
                var workbook = new XLWorkbook();
                var sheetName = (!_isSingleBook) ? "IIS_Logs" : file;
                var worksheet = workbook.Worksheets.Add(sheetName);

                // Creating log sheet
                _processor.SetupLogData(worksheet, file);

                // Creating pivot sheet, if option enabled
                if (_createPivot)
                    _processor.SetupPivotData(workbook, worksheet, sheetName);

                // Saving the workbook seperate excel files
                var excelFile = $"{GetSheetName(file)}.xlsx";
                UpdateStatus($"Exporting data to excel file - {excelFile}...");
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
            var logFiles = GetLogFiles(_folderPath);

            _totalSize = logFiles.Sum(file => new FileInfo(file).Length);
            _processedSize = 0;

            foreach (var file in logFiles)
            {
                UpdateStatus($"Processing data for file ??{GetSheetName(file, true)}...");

                sheetCount++;
                var sheetName = GetSheetName(file);
                var sheetNames = workbook.Worksheets.Select(ws => ws.Name).ToList();
                if (sheetNames.Contains(sheetName))
                    sheetName += $"-{sheetCount}";

                var worksheet = workbook.Worksheets.Add(sheetName);

                // Creating log sheet
                _processor.SetupLogData(worksheet, file);

                // Creating pivot sheet, if option enabled
                if (_createPivot)
                    _processor.SetupPivotData(workbook, worksheet, sheetName);
            }

            // Saving the workbook to a single excel file
            var excelFile = $"{_folderName}.xlsx";
            UpdateStatus($"Exporting data to excel file - {excelFile}...");
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
