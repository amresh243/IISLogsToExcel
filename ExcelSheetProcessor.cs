// Author: Amresh Kumar (July 2025)

using ClosedXML.Excel;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Media;

namespace IISLogsToExcel;

internal class ExcelSheetProcessor(IISLogExporter handler)
{
    private const int MaxSheetRows = 1048576;
    private readonly IISLogExporter _handler = handler;

    #region Utility Methods

    /// <summary> Returns a set of indexes for columns that contain numeric values. </summary>
    /// <param name="headers">list of headers</param>
    /// <returns>index list</returns>
    private static HashSet<int> GetNumberColumnIndexes(List<string> headers)
    {
        try
        {
            return [.. Constants.NumberColumns.Select(header => Array.IndexOf([.. headers], header))];
        }
        catch
        {
            return [];
        }
    }

    /// <summary> Returns sheet name from file name. </summary>
    /// <param name="file">file with path</param>
    /// <returns>sheet name</returns>
    public static string GetSheetName(string file, bool isFile = false)
    {
        if (string.IsNullOrEmpty(file))
            return file;

        if (isFile)
        {
            var fileName = file.Split(LogTokens.PathSplitMarker).LastOrDefault()?.Split(LogTokens.FileSplitMarker).LastOrDefault() ?? "";
            var fileNameLength = fileName.Length;

            return (fileNameLength > 10) ? fileName[(fileNameLength - 10)..] : fileName;
        }

        var sheetName = file.Split(LogTokens.PathSplitMarker).LastOrDefault()?
            .Split(LogTokens.FileSplitMarker).LastOrDefault()?
            .Split(LogTokens.ExtensionSplitMarker).FirstOrDefault();
        if (string.IsNullOrEmpty(sheetName))
            return file;

        var sheetNameLength = sheetName.Length;
        return (sheetNameLength > 6) ? sheetName[(sheetNameLength - 6)..] : sheetName;
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
    public void SetupLogSheet(IXLWorksheet worksheet, string file)
    {
        int currentRow = 1;
        try
        {
            Logger.LogInfo($"Creating sheet {worksheet.Name} against file {file}...");

            var lines = File.ReadAllLines(file, Encoding.UTF8)
                .Where(l => !l.StartsWith(LogTokens.CommentMarker) || l.StartsWith(LogTokens.LogMarker)).ToList();
            if (lines.Count == 0)
            {
                _handler.UpdateList(file, Brushes.Tomato);
                Logger.LogError($"{file} is empty!");
                return;
            }

            if (lines[0].StartsWith(LogTokens.LogMarker))
                lines[0] = lines[0].Replace(LogTokens.LogMarker, string.Empty).Trim();

            var headers = lines[0].Split(LogTokens.LineSplitMarker).Select(x => x.RemoveInvalidXmlChars().ToLowerInvariant()).ToList();
            if (!headers.Contains(Headers.Date) || !headers.Contains(Headers.Time))
            {
                _handler.UpdateList(file, Brushes.Tomato);
                Logger.LogError($"{file} is not a valid IIS log file!");
                return;
            }

            // Setup headers and first row
            if (currentRow == 1)
            {
                headers.Insert(2, Headers.Hour);
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
                var values = line.Split(' ').Select(x => x.RemoveInvalidXmlChars()).ToArray();

                worksheet.Cell(currentRow, 1).Value = values[0];
                worksheet.Cell(currentRow, 2).Value = values[1];
                worksheet.Cell(currentRow, 3).FormulaA1 = string.Format(LogTokens.HourFormulae, currentRow);

                for (int i = 3; i <= values.Length; i++)
                {
                    var cell = worksheet.Cell(currentRow, i + 1);
                    var value = values[i - 1];
                    var isNumericCell = specialIndices.Contains(i);

                    // In rare cases spacially with special chars in urls, url contains space.
                    // This will cause incorrect update of later cells, so we need to handle it.
                    if (isNumericCell && !value.IsNumeric())
                    {
                        Logger.LogWarning($"Broken or invalid data at line {currentRow} in file {file}, output repair attempted.");
                        UpdatePreviousCells(worksheet, currentRow, i, value);
                        values = [.. values.Where(x => x != value)];
                        i--;
                        continue;
                    }

                    cell.Value = isNumericCell ? value.GetValidNumber() : value;
                }

                _handler.UpdateProgress(Encoding.UTF8.GetByteCount(line));
                currentRow++;
            }

            Logger.LogInfo($"Processed {currentRow - 1} lines from file {file}.");
            // Unfortunately excel has static row count of 1048576
            _handler.UpdateStatus(string.Format(Messages.CreateSheet, worksheet.Name));
            worksheet.Rows(currentRow, MaxSheetRows).Hide();
            worksheet.SetAutoFilter();
            Logger.LogInfo($"Excel sheet {worksheet.Name} created for file {file} with {currentRow - 1} rows.");
        }
        catch
        {
            var message = string.Format(Messages.LogError, currentRow, file);

            MessageBox.Show(message, Captions.LogError, MessageBoxButton.OK, MessageBoxImage.Warning);
            worksheet.Rows(currentRow, MaxSheetRows).Hide();
            worksheet.SetAutoFilter();
            _handler.UpdateList(file, Brushes.Tomato);
            Logger.LogException(message, new Exception($"Error enountered while processing line {currentRow} in file {file}"));
        }
    }

    /// <summary>
    /// Logic to process pivot sheet. Setups pivot with hour as filter, time as row label, 
    /// cs-uri-stem as value with count and time-taken as value with average.
    /// </summary>
    /// <param name="workbook">Workbook object, excel workbook object</param>
    /// <param name="worksheet">Worksheet object, excel sheet object</param>
    /// <param name="sheetName">sheet against which pivot to be created</param>
    public void SetupPivotSheet(XLWorkbook workbook, IXLWorksheet worksheet, string sheetName, string file)
    {
        try
        {
            var msg = string.Format(Messages.CreatePivot, worksheet.Name);

            _handler.UpdateStatus(msg);
            Logger.LogInfo(msg);
            var dataRange = worksheet.RangeUsed();

            if (dataRange == null)
            {
                _handler.UpdateList(file, Brushes.Tomato);
                _handler.UpdateStatus(string.Format(Messages.PivotError, sheetName));
                return;
            }

            var pivotSheet = workbook.Worksheets.Add($"{LogTokens.PivotMarker}{sheetName}");
            var pt = pivotSheet.PivotTables.Add(LogTokens.PivotTable, pivotSheet.Cell(1, 1), dataRange);
            pt.RowLabels.Add(Headers.Time);
            pt.ReportFilters.Add(Headers.Hour);
            pt.Values.Add(Headers.UriStem, Headers.UriStemCount).SetSummaryFormula(XLPivotSummary.Count);
            pt.Values.Add(Headers.TimeTaken, Headers.TimeTakenAvg).SetSummaryFormula(XLPivotSummary.Average);
            pt.Values.Last().NumberFormat.Format = "0";
            pivotSheet.Cell(3, 1).SetValue(Headers.Time);
            pivotSheet.Column(2).Width = 16;
            pivotSheet.Column(3).Width = 13;
            pivotSheet.SheetView.Freeze(3, 0);
            Logger.LogInfo($"Pivot table {sheetName} created for sheet {worksheet.Name}.");
        }
        catch
        {
            var message = string.Format(Messages.PivotError, sheetName);

            MessageBox.Show(message, Captions.PivotError, MessageBoxButton.OK, MessageBoxImage.Warning);
            _handler.UpdateList(file, Brushes.Tomato);
            Logger.LogException(message, new Exception($"Error encountered while processing pivot data for sheet {sheetName}"));
        }
    }

    #endregion Excel Data Processing Methods
}
