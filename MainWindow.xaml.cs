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
        private string _folderName = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Checkbox handler
        /// </summary>
        /// <param name="sender">sender object</param>
        /// <param name="e">event arg</param>
        private void SingleWorkbook_Click(object sender, RoutedEventArgs e)
        {
            if (isSingleWorkBook.IsChecked == true)
                _isSingleBook = true;
            else
                _isSingleBook = false;
        }

        /// <summary>
        /// Select folder button handler
        /// </summary>
        /// <param name="sender">sender object</param>
        /// <param name="e">event arg</param>
        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                folderPathTextBox.Text = dialog.FolderName;
                _folderName = dialog.FolderName.Split('\\', StringSplitOptions.None).Last();
            }
        }

        /// <summary>
        /// Process log button handler
        /// </summary>
        /// <param name="sender">sender object</param>
        /// <param name="e">event arg</param>
        private async void ProcessButton_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = folderPathTextBox.Text;
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                MessageBox.Show("Please select a valid folder.");
                return;
            }

            selectFolderButton.IsEnabled = false;
            processButton.IsEnabled = false;
            isSingleWorkBook.IsEnabled = false;
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
                    selectFolderButton.IsEnabled = true;
                    processButton.IsEnabled = true;
                    isSingleWorkBook.IsEnabled = true;
                });
            }

            Dispatcher.Invoke(() =>
                {
                    statusText.Text = "Processing complete.";
                    selectFolderButton.IsEnabled = true;
                    processButton.IsEnabled = true;
                    isSingleWorkBook.IsEnabled = true;
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

            if (File.Exists(xlsFile))
                File.Delete(xlsFile);

            workbook.SaveAs(xlsFile);
        }

        /// <summary>
        /// Logic to process excel sheet
        /// </summary>
        /// <param name="worksheet">Worksheet object, excel sheet object</param>
        /// <param name="file">Source log file</param>
        private static void ProcessSheetData(IXLWorksheet worksheet, string file)
        {
            int currentRow = 1;
            var lines = File.ReadAllLines(file).Where(l => !l.StartsWith("#") || l.StartsWith("#Fields:")).ToList();
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
                
                worksheet.Cell(currentRow, headers.Count).Value = values[valuesLength - 1];
                currentRow++;
            }

            // Unfortunately excel has static row count of 1048576
            worksheet.Rows(currentRow, 1048576).Hide();
            worksheet.SetAutoFilter();
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

                ProcessSheetData(worksheet, file);
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

                ProcessSheetData(worksheet, file);
            }

            SaveExcelFile(workbook, Path.Combine(folderPath, $"{_folderName}.xlsx"));
        }
    }
}
