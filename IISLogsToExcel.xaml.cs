// Author: Amresh Kumar

using ClosedXML.Excel;
using IISLogsToExcel;
using Microsoft.Win32;
using System.Data;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Threading;

namespace IISLogToExcelConverter
{
    public partial class IISLogExporter : Window
    {
        private const int MaxSheetRows = 1048576;
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
        private void UpdateStatus(string message)
        {
            Dispatcher.Invoke(() =>
            {
                statusText.Text = message;
            });
        }

        /// <summary> Updates progress status on the progress bar. </summary>
        private void UpdateProgress()
        {
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

        /// <summary> Returns a set of indexes for columns that contain numeric values. </summary>
        /// <param name="headers">list of headers</param>
        /// <returns>index list</returns>
        private static HashSet<int> GetNumberColumnIndexes(List<string> headers)
        {
            try
            {
                string[] numberColumnHeader = { "s-port", "sc-status", "sc-substatus", "sc-win32-status", "sc-bytes", "cs-bytes", "time-taken" };
                return [.. numberColumnHeader.Select(header => Array.IndexOf([.. headers], header))];
            }
            catch
            {
                return [];
            }
        }

        /// <summary> Removes invalid XML characters from the given text. </summary>
        /// <param name="text">Input text</param>
        /// <returns>Cleaned text</returns>
        public static string RemoveInvalidXmlChars(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            return new string([.. text.Where(ch =>
                (ch == 0x9 || ch == 0xA || ch == 0xD ||
                (ch >= 0x20 && ch <= 0xD7FF) ||
                (ch >= 0xE000 && ch <= 0xFFFD) ||
                (ch >= 0x10000 && ch <= 0x10FFFF))
                )]);
        }

        private static void UpdatePreviousCells(IXLWorksheet worksheet, int currentRow, int columnIndex, string value)
        {
            var wronglyUpdatedCell = worksheet.Cell(currentRow, columnIndex - 1);
            var prevCell = worksheet.Cell(currentRow, columnIndex);
            wronglyUpdatedCell.Value = $"{wronglyUpdatedCell.Value} {prevCell.Value}";
            prevCell.Value = value;
        }

        #endregion Utility Methods


        #region Excel Data Processing Methods

        /// <summary> Logic to process excel sheet. </summary>
        /// <param name="worksheet">Worksheet object, excel sheet object</param>
        /// <param name="file">Source log file</param>
        private void SetupLogData(IXLWorksheet worksheet, string file)
        {
            int currentRow = 1;
            try
            {
                var lines = File.ReadAllLines(file, Encoding.UTF8).Where(l => !l.StartsWith('#') || l.StartsWith("#Fields:")).ToList();
                if (lines.Count == 0)
                    return;

                if (lines[0].StartsWith("#Fields:"))
                    lines[0] = lines[0].Replace("#Fields:", string.Empty).Trim();

                var headers = lines[0].Split(' ').Select(x => RemoveInvalidXmlChars(x).ToLowerInvariant()).ToList();
                if (!headers.Contains("date") || !headers.Contains("time"))
                    return;

                // Setup headers and first row
                if (currentRow == 1)
                {
                    _processedSize += Encoding.UTF8.GetByteCount(lines[0]);
                    headers.Insert(2, "hour");
                    for (int i = 0; i < headers.Count; i++)
                        worksheet.Cell(currentRow, i + 1).Value = headers[i];

                    worksheet.SheetView.Freeze(currentRow, 0);
                    worksheet.Row(currentRow).Style.Font.Bold = true;
                    UpdateProgress();
                    currentRow++;
                }

                var specialIndices = GetNumberColumnIndexes(headers);

                // Process each line of the log file and fill the worksheet
                foreach (var line in lines.Skip(1))
                {
                    var values = line.Split(' ').Select(x => RemoveInvalidXmlChars(x)).ToArray();

                    worksheet.Cell(currentRow, 1).Value = values[0];
                    worksheet.Cell(currentRow, 2).Value = values[1];
                    worksheet.Cell(currentRow, 3).FormulaA1 = $"=TEXT(B{currentRow}, \"hh:mm\")";

                    for (int i = 3; i <= values.Length; i++)
                    {
                        var cell = worksheet.Cell(currentRow, i + 1);
                        var value = values[i - 1];
                        var isNumericCell = specialIndices.Contains(i);

                        // In rare cases spacially with special chars in urls, url contains space.
                        // This will cause incorrect update of later cells, so we need to handle it.
                        if (isNumericCell && !value.IsNumeric())
                        {
                            UpdatePreviousCells(worksheet, currentRow, i, value);
                            values = values.Where(x => x != value).ToArray();
                            i--;
                            continue;
                        }

                        cell.Value = isNumericCell ? value.GetValidNumber() : value;
                    }

                    _processedSize += Encoding.UTF8.GetByteCount(line);
                    UpdateProgress();
                    currentRow++;
                }

                // Unfortunately excel has static row count of 1048576
                UpdateStatus($"Creating IIS log sheet - {worksheet.Name}...");
                worksheet.Rows(currentRow, MaxSheetRows).Hide();
                worksheet.SetAutoFilter();
            }
            catch
            {
                MessageBox.Show($"An error occurred at line {currentRow} while processing log file {file}." +
                    $"\n\nExported IIS log sheet and respective pivot sheet may have incomplete and corrupt data.", 
                    "Log Export Error!", MessageBoxButton.OK, MessageBoxImage.Warning);
                worksheet.Rows(currentRow, MaxSheetRows).Hide();
                worksheet.SetAutoFilter();
            }
        }

        /// <summary>
        /// Logic to process pivot sheet. Setups pivot with hour as filter, time as row label, 
        /// cs-uri-stem as value with count and time-taken as value with average.
        /// </summary>
        /// <param name="workbook">Workbook object, excel workbook object</param>
        /// <param name="worksheet">Worksheet object, excel sheet object</param>
        /// <param name="sheetName">sheet against which pivot to be created</param>
        private void SetupPivotData(XLWorkbook workbook, IXLWorksheet worksheet, string sheetName)
        {
            UpdateStatus($"Creating pivot table for sheet - {sheetName}...");
            var dataRange = worksheet.RangeUsed();
            var pivotSheet = workbook.Worksheets.Add($"Pivot_{sheetName}");
            var pt = pivotSheet.PivotTables.Add("PivotTable", pivotSheet.Cell(1, 1), dataRange);
            pt.RowLabels.Add("time");
            pt.ReportFilters.Add("hour");
            pt.Values.Add("cs-uri-stem", "cs-uri-stem[count]").SetSummaryFormula(XLPivotSummary.Count);
            pt.Values.Add("time-taken", "time-taken[avg]").SetSummaryFormula(XLPivotSummary.Average);
            pt.Values.Last().NumberFormat.Format = "0";
            pivotSheet.Cell(3, 1).SetValue("time");
            pivotSheet.Column(2).Width = 16;
            pivotSheet.Column(3).Width = 13;
            pivotSheet.SheetView.Freeze(3, 0);
        }

        #endregion Excel Data Processing Methods


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
                SetupLogData(worksheet, file);

                // Creating pivot sheet, if option enabled
                if (_createPivot)
                    SetupPivotData(workbook, worksheet, sheetName);

                // Saving the workbook seperate excel files
                var excelFile = $"{GetSheetName(file)}.xlsx";
                UpdateStatus($"Exporting data to excel file - {excelFile}...");
                bool isSuccess = SaveExcelFile(workbook, Path.Combine(_folderPath, excelFile));

                // Deleting source file, if option enabled and save was successful
                if (_deleteSources && isSuccess)
                    DeleteLogFiles(file);
            }

            _processedSize = _totalSize;
            UpdateProgress();
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
                SetupLogData(worksheet, file);

                // Creating pivot sheet, if option enabled
                if (_createPivot)
                    SetupPivotData(workbook, worksheet, sheetName);
            }

            // Saving the workbook to a single excel file
            var excelFile = $"{_folderName}.xlsx";
            UpdateStatus($"Exporting data to excel file - {excelFile}...");
            bool isSucess = SaveExcelFile(workbook, Path.Combine(_folderPath, excelFile));

            // Deleting all source files, if option enabled and save was successful
            if (_deleteSources && isSucess)
                DeleteLogFiles(string.Empty, true);

            _processedSize = _totalSize;
            UpdateProgress();
        }

        #endregion Thread Methods
    }
}
