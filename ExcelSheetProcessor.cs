// Author: Amresh Kumar (July 2025)

using ClosedXML.Excel;
using System.IO;
using System.Text;
using System.Windows;

namespace IISLogToExcel
{
    internal class ExcelSheetProcessor(IISLogExporter handler)
    {
        private const int MaxSheetRows = 1048576;
        private readonly IISLogExporter _handler = handler;


        #region Utility Methods

        /// <summary> Removes invalid XML characters from the given text. </summary>
        /// <param name="text">Input text</param>
        /// <returns>Cleaned text</returns>
        private static string RemoveInvalidXmlChars(string text)
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

        /// <summary> Updates previous cells in the row when a cell is wrongly updated. </summary>
        /// <param name="worksheet">Current worksheet</param>
        /// <param name="currentRow">Current row</param>
        /// <param name="columnIndex">Current column</param>
        /// <param name="value">Value to be updated</param>
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
        public void SetupLogData(IXLWorksheet worksheet, string file)
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
                    headers.Insert(2, "hour");
                    for (int i = 0; i < headers.Count; i++)
                        worksheet.Cell(currentRow, i + 1).Value = headers[i];

                    worksheet.SheetView.Freeze(currentRow, 0);
                    worksheet.Row(currentRow).Style.Font.Bold = true;
                    _handler.UpdateProgress(Encoding.UTF8.GetByteCount(lines[0]));
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

                    _handler.UpdateProgress(Encoding.UTF8.GetByteCount(line));
                    currentRow++;
                }

                // Unfortunately excel has static row count of 1048576
                _handler.UpdateStatus($"Creating IIS log sheet - {worksheet.Name}...");
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
        public void SetupPivotData(XLWorkbook workbook, IXLWorksheet worksheet, string sheetName)
        {
            _handler.UpdateStatus($"Creating pivot table for sheet - {sheetName}...");
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
    }
}
