// Author: Amresh Kumar

using ClosedXML.Excel;
using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace IISLogToExcelConverter
{
    public partial class MainWindow : Window
    {
        private bool _isSingleBook = false;
        private bool _createPivot = false;
        private string _folderName = "";

        public MainWindow()
        {
            InitializeComponent();
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

        private void ChangeControlState(bool isEnabled)
        {
            selectFolderButton.IsEnabled = isEnabled;
            processButton.IsEnabled = isEnabled;
            isSingleWorkBook.IsEnabled = isEnabled;
            createPivotTable.IsEnabled = isEnabled;
        }

        /// <summary> Process log button handler </summary>
        private async void ProcessButton_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = folderPathTextBox.Text;
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                MessageBox.Show("Please select a valid folder.");
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
            catch
            {
                Dispatcher.Invoke(() =>
                {
                    statusText.Text = "Error! Something went wrong.";
                    ChangeControlState(true);
                });
            }

            Dispatcher.Invoke(() =>
                {
                    statusText.Text = "Processing complete.";
                    ChangeControlState(true);
                });
        }

        /// <summary>
        /// Saves workbook object into excel file
        /// </summary>
        /// <param name="workbook">Workbook object, excel file object</param>
        /// <param name="xlsFile">Excel file name to be saved</param>
        private static void SaveExcelFile(XLWorkbook workbook, string xlsFile)
        {
            if (workbook == null)
                return;

            try
            {
                if (File.Exists(xlsFile))
                    File.Delete(xlsFile);

                workbook.SaveAs(xlsFile);
            }
            catch { /* We don't want anything if delete and save fails. */ }
        }

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
                lines[0] = lines[0].Replace("#Fields:", "").Trim();

            var headers = lines[0].Split(' ').ToList();
            if (!headers.Contains("date") || !headers.Contains("time")) return;

            // Setup headers and first row
            if (currentRow == 1)
            {
                headers.Insert(2, "hour");

                for (int i = 0; i < headers.Count; i++)
                    worksheet.Cell(currentRow, i + 1).Value = headers[i];

                worksheet.SheetView.Freeze(currentRow, 0);
                var headerRow = worksheet.Row(currentRow);
                headerRow.Style.Font.Bold = true;
                currentRow++;
            }

            // Process each line of the log file and fill the worksheet
            foreach (var line in lines.Skip(1))
            {
                var values = line.Split(' ');
                int valuesLength = values.Length;
                
                worksheet.Cell(currentRow, 1).Value = values[0];
                worksheet.Cell(currentRow, 2).Value = values[1];
                
                worksheet.Cell(currentRow, 3).FormulaA1 = $"=TEXT(B{currentRow}, \"hh:mm\")";
                
                for (int i = 3; i < valuesLength; i++)
                    worksheet.Cell(currentRow, i + 1).Value = values[i - 1];

                int.TryParse(values[valuesLength - 1], out int timeTaken);
                worksheet.Cell(currentRow, headers.Count).Value = timeTaken;
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

        /// <summary>
        /// Creates seperate excel file for each file under folder
        /// </summary>
        /// <param name="folderPath">Root folder path</param>
        private void CreateSeperateFiles(string folderPath)
        {
            var logFiles = Directory.GetFiles(folderPath, "*.log", SearchOption.AllDirectories);

            foreach (var file in logFiles)
            {
                var workbook = new XLWorkbook();
                var sheetName = (!_isSingleBook) ? "IIS Logs" : file;
                var worksheet = workbook.Worksheets.Add(sheetName);

                SetupLogData(worksheet, file);

                if (_createPivot)
                    SetupPivotData(workbook, worksheet, sheetName);

                SaveExcelFile(workbook, Path.Combine(folderPath, $"{file}.xlsx"));
            }
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

            if(string.IsNullOrEmpty(sheetName))
                return file;

            return sheetName;
        }

        /// <summary>
        /// Creates single excel file with sheets as multiple files under folder
        /// </summary>
        /// <param name="folderPath">Root folder path</param>
        private void CreateSingleFile(string folderPath)
        {
            var sheetCount = 0;
            var logFiles = Directory.GetFiles(folderPath, "*.log", SearchOption.AllDirectories);
            var workbook = new XLWorkbook();

            foreach (var file in logFiles)
            {
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
    }
}
