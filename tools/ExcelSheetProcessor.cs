// Author: Amresh Kumar (July 2025)

using ClosedXML.Excel;
using System.IO;
using System.Text;
using System.Windows.Media;

namespace IISLogsToExcel.tools;

internal class ExcelSheetProcessor(IISLogExporter handler)
{
    private const int MaxSheetRows = 1048576;
    private readonly IISLogExporter _handler = handler;

    #region Utility Methods

    /// <summary> Returns a set of indexes for columns that contain numeric values. </summary>
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
    public static string GetSheetName(string file, bool isFile = false, string[]? sheets = null)
    {
        if (string.IsNullOrEmpty(file))
            return file;

        if (isFile)
        {
            var fileName = file.Split(LogTokens.PathSplitMarker).LastOrDefault()?
                               .Split(LogTokens.FileSplitMarker).LastOrDefault() ?? string.Empty;
            var fileNameLength = fileName.Length;

            return fileNameLength > 10 ? fileName[(fileNameLength - 10)..] : fileName;
        }

        var sheetName = file.Split(LogTokens.PathSplitMarker).LastOrDefault()?
                            .Split(LogTokens.FileSplitMarker).LastOrDefault()?
                            .Split(LogTokens.ExtensionSplitMarker).FirstOrDefault();
        if (string.IsNullOrEmpty(sheetName))
            return file;

        var sheetNameLength = sheetName.Length;
        var sheet = sheetNameLength > 6 ? sheetName[(sheetNameLength - 6)..] : sheetName;
        if (sheets != null && sheets.Length > 0)
        {
            var existingCount = sheets.Count(name => name == sheet);
            int sheetCount = existingCount;
            while (existingCount != 0)
            {
                sheet = sheet.Split(LogTokens.FileSplitMarker).FirstOrDefault() + $"{LogTokens.FileSplitMarker}{sheetCount}";
                existingCount = sheets.Count(name => name == sheet);
                sheetCount++;
            }
        }

        return sheet;
    }

    #endregion Utility Methods


    #region Excel Data Processing Methods

    /// <summary> Processes a row of data and adds it to the worksheet. </summary>
    private static void AddRowData(IXLWorksheet worksheet, HashSet<int> specialIndices, string[] values, int currentRow)
    {
        worksheet.Cell(currentRow, 1).Value = values[0];
        worksheet.Cell(currentRow, 2).Value = values[1];
        worksheet.Cell(currentRow, 3).FormulaA1 = string.Format(LogTokens.HourFormulae, currentRow);

        for (int i = 3; i <= values.Length; i++)
        {
            var cell = worksheet.Cell(currentRow, i + 1);
            var value = values[i - 1];
            var isNumericCell = specialIndices.Contains(i);

            cell.Value = isNumericCell ? value.GetValidNumber() : value;
        }
    }

    private string[] GetFixedStemData(string[] source)
    {
        if(Constants.portIds.Contains(source[6]))
            return source;

        source[4] = source[4] + source[5];
        var sourceList = source.ToList();
        sourceList.RemoveAt(5);
        return [.. sourceList];
    }

    /// <summary> Logic to process excel sheet. </summary>
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
                var headerRow = worksheet.Range(1, 1, 1, headers.Count);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Fill.BackgroundColor = XLColor.AshGrey;
                _handler.UpdateProgress(Encoding.UTF8.GetByteCount(lines[0]));
                currentRow++;
            }

            var specialIndices = GetNumberColumnIndexes(headers);
            var incompleteCellData = new List<string>();
            // Process each line of the log file and fill the worksheet
            foreach (var line in lines.Skip(1))
            {
                var values = line.Split(' ').Select(x => x.RemoveInvalidXmlChars()).ToArray();
                if (values.Length >= headers.Count)
                    values = GetFixedStemData(values);

                if(values.Length >= headers.Count)
                {
                    Logger.LogWarning($"Skipped data at line {currentRow} in file {file}, output repair attempted but failed.");
                    currentRow++;
                    continue;
                }

                // Handling broken iis log row data
                if (values.Length < headers.Count - 1)
                {
                    var prevDataCount = incompleteCellData.Count;
                    if (prevDataCount == 0)
                        incompleteCellData.AddRange(values);
                    else if (values.Length == 1)
                        incompleteCellData[prevDataCount - 1] += values[0];
                    else
                    {
                        incompleteCellData[prevDataCount - 1] += values[0];
                        incompleteCellData.AddRange(values.Skip(1));
                    }

                    lines = [.. lines.Where(x => x != line)];
                    if (incompleteCellData.Count < headers.Count - 1)
                        continue;
                }

                var progressValue = Encoding.UTF8.GetByteCount(line);

                // Create row if there is broken row computed data available
                if (incompleteCellData.Count > 0)
                {
                    Logger.LogWarning($"Broken or invalid data at line {currentRow} in file {file}, output repair attempted.");
                    _handler.UpdateList(file, Brushes.Tomato);
                    values = [..incompleteCellData];
                    progressValue = Encoding.UTF8.GetByteCount(string.Join(' ', values));
                    incompleteCellData.Clear();
                }

                AddRowData(worksheet, specialIndices, values, currentRow);
                _handler.UpdateProgress(progressValue);
                currentRow++;
            }

            Logger.LogInfo($"Processed {currentRow - 1} lines from file {file}.");
            _handler.UpdateStatus(string.Format(Messages.CreateSheet, worksheet.Name));
            // Unfortunately excel has static row count of 1048576
            worksheet.Rows(currentRow, MaxSheetRows).Hide();
            worksheet.SetAutoFilter();
            Logger.LogInfo($"Excel sheet {worksheet.Name} created for file {file} with {currentRow - 1} rows.");
        }
        catch
        {
            var message = string.Format(Messages.LogError, currentRow, file);

            _handler.MessageBox.Show(message, Captions.LogError, DialogTypes.Warning);
            worksheet.Rows(currentRow, MaxSheetRows).Hide();
            worksheet.SetAutoFilter();
            _handler.UpdateList(file, Brushes.Tomato);
            Logger.LogException(message, new Exception($"Error encountered while processing line {currentRow} in file {file}"));
        }
    }

    /// <summary>
    /// Logic to process pivot sheet. Setups pivot with hour as filter, time as row label, 
    /// cs-uri-stem as value with count and time-taken as value with average.
    /// </summary>
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

            _handler.MessageBox.Show(message, Captions.PivotError, DialogTypes.Warning);
            _handler.UpdateList(file, Brushes.Tomato);
            Logger.LogException(message, new Exception($"Error encountered while processing pivot data for sheet {sheetName}"));
        }
    }

    #endregion Excel Data Processing Methods
}
