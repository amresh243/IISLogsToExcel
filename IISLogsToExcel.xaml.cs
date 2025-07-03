// Author: Amresh Kumar

using ClosedXML.Excel;
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
        private bool _isSingleBook = false;
        private bool _createPivot = false;
        private string _folderName = string.Empty;

        public IISLogExporter()
        {
            InitializeComponent();
        }

        #region Control State Modifiers
        /// <summary> Changes the state of controls based on the isEnabled parameter. </summary>
        /// <param name="isEnabled"> true=enalbe/false=disable </param>
        private void ChangeControlState(bool isEnabled)
        {
            selectFolderButton.IsEnabled = isEnabled;
            processButton.IsEnabled = isEnabled;
            isSingleWorkBook.IsEnabled = isEnabled;
            createPivotTable.IsEnabled = isEnabled;
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

        #endregion Control State Modifiers


        #region Event Handlers

        /// <summary> DragOver event handler, only allows folder to be dropped. </summary>
        private void FolderPath_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                // Only allow if the first item is a directory
                if (paths.Length > 0 && Directory.Exists(paths[0]) && GetLogFiles(paths[0]).Length != 0)
                    e.Effects = DragDropEffects.Copy;
                else
                    e.Effects = DragDropEffects.None;
            }
            else
                e.Effects = DragDropEffects.None;

            e.Handled = true;
        }

        /// <summary>
        /// Drop event handler, sets the folder path with the dropped folder path.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FolderPath_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (paths.Length > 0 && Directory.Exists(paths[0]))
                {
                    string folderPath = paths[0];
                    folderPathTextBox.Text = folderPath;
                    _folderName = _folderName = folderPath.Split('\\', StringSplitOptions.None).Last();
                }
            }
        }

        /// <summary> Single workbook Checkbox click handler </summary>
        private void SingleWorkbook_Click(object sender, RoutedEventArgs e) =>
            _isSingleBook = (isSingleWorkBook.IsChecked == true);

        /// <summary> Create pivot Checkbox click handler </summary>
        private void PivotTable_Click(object sender, RoutedEventArgs e) =>
            _createPivot = (createPivotTable.IsChecked == true);

        /// <summary> Select folder button click handler </summary>
        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                folderPathTextBox.Text = dialog.FolderName;
                _folderName = dialog.FolderName.Split('\\', StringSplitOptions.None).Last();
            }
        }

        /// <summary> Process log button handler </summary>
        private async void ProcessButton_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = folderPathTextBox.Text;
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                MessageBox.Show(this, "Please select a valid folder.", "Invalid Input");
                return;
            }

            ChangeControlState(false);
            statusText.Text = "Processing...";

            try
            {
                if (!_isSingleBook)
                    await Task.Run(() => CreateSeperateFiles(folderPath));
                else
                    await Task.Run(() => CreateSingleFile(folderPath));
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

        /// <summary>
        /// Saves workbook object into excel file
        /// </summary>
        /// <param name="workbook">Workbook object, excel file object</param>
        /// <param name="xlsFile">Excel file name to be saved</param>
        private void SaveExcelFile(XLWorkbook workbook, string xlsFile)
        {
            if (workbook == null)
                return;

            try
            {
                if (File.Exists(xlsFile))
                    File.Delete(xlsFile);

                workbook.SaveAs(xlsFile);
                workbook.Dispose();
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

        /// <summary>
        /// Returns sheet name from file name
        /// </summary>
        /// <param name="file">file with path</param>
        /// <returns>sheet name</returns>
        private static string GetSheetName(string file)
        {
            var sheetName = file.Split('\\').LastOrDefault()?
                                .Split('-').LastOrDefault()?
                                .Split('.').FirstOrDefault();

            if (string.IsNullOrEmpty(sheetName))
                return file;

            return sheetName;
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
            if (string.IsNullOrEmpty(text)) return text;

            return new string([.. text.Where(ch =>
                (ch == 0x9 || ch == 0xA || ch == 0xD ||
                (ch >= 0x20 && ch <= 0xD7FF) ||
                (ch >= 0xE000 && ch <= 0xFFFD) ||
                (ch >= 0x10000 && ch <= 0x10FFFF))
                )]);
        }


        #endregion Utility Methods


        #region Excel Data Processing Methods

        /// <summary>
        /// Logic to process excel sheet
        /// </summary>
        /// <param name="worksheet">Worksheet object, excel sheet object</param>
        /// <param name="file">Source log file</param>
        private static void SetupLogData(IXLWorksheet worksheet, string file)
        {
            int currentRow = 1;
            var lines = File.ReadAllLines(file).Where(l => !l.StartsWith('#') || l.StartsWith("#Fields:")).ToList();
            if (lines.Count == 0) return;

            if (lines[0].StartsWith("#Fields:"))
                lines[0] = lines[0].Replace("#Fields:", string.Empty).Trim();

            var headers = lines[0].Split(' ').Select(x => RemoveInvalidXmlChars(x).ToLowerInvariant()).ToList();
            if (!headers.Contains("date") || !headers.Contains("time")) return;

            // Setup headers and first row
            if (currentRow == 1)
            {
                headers.Insert(2, "hour");
                for (int i = 0; i < headers.Count; i++)
                    worksheet.Cell(currentRow, i + 1).Value = headers[i];

                worksheet.SheetView.Freeze(currentRow, 0);
                worksheet.Row(currentRow).Style.Font.Bold = true;
                currentRow++;
            }

            var specialIndices = GetNumberColumnIndexes(headers);
            bool hasSpcialIndices = specialIndices.Count != 0;

            // Process each line of the log file and fill the worksheet
            foreach (var line in lines.Skip(1))
            {
                var values = line.Split(' ').Select(x => RemoveInvalidXmlChars(x)).ToArray();
                int valuesLength = values.Length;

                worksheet.Cell(currentRow, 1).Value = values[0];
                worksheet.Cell(currentRow, 2).Value = values[1];
                worksheet.Cell(currentRow, 3).FormulaA1 = $"=TEXT(B{currentRow}, \"hh:mm\")";

                for (int i = 3; i <= valuesLength; i++)
                {
                    var cell = worksheet.Cell(currentRow, i + 1);
                    var value = values[i - 1];

                    cell.Value = hasSpcialIndices && specialIndices.Contains(i) ? int.Parse(value) : value;
                }

                currentRow++;
            }

            // Unfortunately excel has static row count of 1048576
            worksheet.Rows(currentRow, 1048576).Hide();
            worksheet.SetAutoFilter();
        }

        /// <summary>
        /// Logic to process pivot sheet. Setups pivot with hour as filter, time as row label, 
        /// cs-uri-stem as value with count and time-taken as value with average.
        /// </summary>
        /// <param name="workbook">Workbook object, excel workbook object</param>
        /// <param name="worksheet">Worksheet object, excel sheet object</param>
        /// <param name="sheetName">sheet against which pivot to be created</param>
        private static void SetupPivotData(XLWorkbook workbook, IXLWorksheet worksheet, string sheetName)
        {
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

        /// <summary>
        /// Creates seperate excel file for each file under folder
        /// </summary>
        /// <param name="folderPath">Root folder path</param>
        private void CreateSeperateFiles(string folderPath)
        {
            var logFiles = GetLogFiles(folderPath);

            foreach (var file in logFiles)
            {
                UpdateStatus($"Processing data for file {file.Split('\\').LastOrDefault() ?? string.Empty}...");
                var workbook = new XLWorkbook();
                var sheetName = (!_isSingleBook) ? "IIS Logs" : file;
                var worksheet = workbook.Worksheets.Add(sheetName);

                // Creating log sheet
                SetupLogData(worksheet, file);

                // Creating pivot sheet, if option enabled
                if (_createPivot)
                    SetupPivotData(workbook, worksheet, sheetName);

                // Saving excel file
                SaveExcelFile(workbook, Path.Combine(folderPath, $"{file}.xlsx"));
            }
        }

        /// <summary>
        /// Creates single excel file with sheets as multiple files under folder
        /// </summary>
        /// <param name="folderPath">Root folder path</param>
        private void CreateSingleFile(string folderPath)
        {
            var sheetCount = 0;
            var logFiles = GetLogFiles(folderPath);
            var workbook = new XLWorkbook();

            foreach (var file in logFiles)
            {
                UpdateStatus($"Processing data for file {file.Split('\\').LastOrDefault() ?? string.Empty}...");
                sheetCount++;
                var sheetName = GetSheetName(file);
                var sheetNames = workbook.Worksheets.Select(ws => ws.Name).ToList();
                if (sheetNames.Contains(sheetName))
                    sheetName += $"-{sheetCount}";

                var worksheet = workbook.Worksheets.Add(sheetName);

                SetupLogData(worksheet, file);

                if (_createPivot)
                    SetupPivotData(workbook, worksheet, sheetName);
            }

            SaveExcelFile(workbook, Path.Combine(folderPath, $"{_folderName}.xlsx"));
        }

        #endregion Thread Methods
    }
}
