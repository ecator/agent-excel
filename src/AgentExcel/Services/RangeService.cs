using System.Diagnostics;
using System.Text.Json;

using AgentExcel.Models;
using AgentExcel.Providers;
using AgentExcel.Utils;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class RangeService : ExcelServiceBase
{
    public RangeService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public Dictionary<string, object> ReadRange(string workbookName, string sheetName, string? rangeAddress)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);

                if (string.IsNullOrWhiteSpace(rangeAddress))
                {
                    range = ws.UsedRange;
                }
                else
                {
                    range = GetRange(ws, rangeAddress);
                }

                var results = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                if (range == null)
                {
                    return results;
                }

                int startRow = range.Row;
                int startCol = range.Column;
                object rawValue = range.Value2;

                if (rawValue == null)
                {
                    return results;
                }

                if (rawValue is object[,] matrix)
                {
                    int rows = matrix.GetLength(0);
                    int cols = matrix.GetLength(1);

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            object? val = matrix[r, c];
                            if (val != null)
                            {
                                string strVal = val.ToString() ?? "";
                                if (!string.IsNullOrEmpty(strVal))
                                {
                                    string cellAddress = $"{GetColumnLetter(startCol + c - 1)}{startRow + r - 1}";
                                    results[cellAddress] = val;
                                }
                            }
                        }
                    }
                }
                else
                {
                    string strVal = rawValue.ToString() ?? "";
                    if (!string.IsNullOrEmpty(strVal))
                    {
                        string cellAddress = $"{GetColumnLetter(startCol)}{startRow}";
                        results[cellAddress] = rawValue;
                    }
                }

                return results;
            }
            finally
            {
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void WriteRange(string workbookName, string sheetName, string rangeAddress, object value)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                range = GetRange(ws, rangeAddress);

                var convertedValue = ConvertValueForExcel(value);
                ValidateRangeAndValueDimensions(range, convertedValue);

                range.Value2 = convertedValue;
            }
            finally
            {
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    private object ConvertValueForExcel(object value)
    {
        if (value == null)
        {
            return null!;
        }

        if (value is JsonElement element)
        {
            return ConvertJsonElement(element);
        }

        // Handle nested arrays/lists (jagged array or list of lists)
        if (value is System.Collections.IEnumerable enumerable && !(value is string))
        {
            // First, let's check if it is a 2D array already (like object[,])
            if (value.GetType().IsArray && value.GetType().GetArrayRank() == 2)
            {
                return value;
            }

            // Convert to a List<List<object?>> first
            var list2D = new List<List<object?>>();
            foreach (var rowObj in enumerable)
            {
                if (rowObj is System.Collections.IEnumerable rowEnumerable && !(rowObj is string))
                {
                    var rowList = new List<object?>();
                    foreach (var cellObj in rowEnumerable)
                    {
                        rowList.Add(cellObj);
                    }
                    list2D.Add(rowList);
                }
                else
                {
                    // If it's a 1D collection (like a flat list/array), treat it as a single row
                    var rowList = new List<object?>();
                    foreach (var cellObj in enumerable)
                    {
                        rowList.Add(cellObj);
                    }
                    list2D.Add(rowList);
                    break; // break because we processed all elements as a single row
                }
            }

            if (list2D.Count > 0)
            {
                int rows = list2D.Count;
                int cols = list2D.Max(r => r.Count);
                object[,] matrix = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < list2D[r].Count; c++)
                    {
                        var cellVal = list2D[r][c];
                        if (cellVal != null)
                        {
                            matrix[r, c] = cellVal is JsonElement el ? ConvertJsonElement(el) : cellVal;
                        }
                    }
                }
                return matrix;
            }
        }

        return value;
    }

    private object ConvertJsonElement(JsonElement element)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.String:
                return element.GetString() ?? string.Empty;
            case JsonValueKind.Number:
                if (element.TryGetInt32(out int i)) return i;
                if (element.TryGetInt64(out long l)) return l;
                if (element.TryGetDouble(out double d)) return d;
                return element.GetRawText();
            case JsonValueKind.True:
                return true;
            case JsonValueKind.False:
                return false;
            case JsonValueKind.Null:
                return null!;
            case JsonValueKind.Array:
                int rowCount = element.GetArrayLength();
                if (rowCount == 0) return new object[0, 0];

                var firstRow = element[0];
                if (firstRow.ValueKind == JsonValueKind.Array)
                {
                    int colCount = 0;
                    for (int idx = 0; idx < rowCount; idx++)
                    {
                        if (element[idx].ValueKind == JsonValueKind.Array)
                        {
                            colCount = Math.Max(colCount, element[idx].GetArrayLength());
                        }
                    }

                    object[,] matrix = new object[rowCount, colCount];
                    for (int r = 0; r < rowCount; r++)
                    {
                        var rowEl = element[r];
                        if (rowEl.ValueKind == JsonValueKind.Array)
                        {
                            int rowLength = rowEl.GetArrayLength();
                            for (int c = 0; c < rowLength; c++)
                            {
                                matrix[r, c] = ConvertJsonElement(rowEl[c]);
                            }
                        }
                    }
                    return matrix;
                }
                else
                {
                    // 1D array: convert to a 1-row 2D array
                    object[,] matrix = new object[1, rowCount];
                    for (int c = 0; c < rowCount; c++)
                    {
                        matrix[0, c] = ConvertJsonElement(element[c]);
                    }
                    return matrix;
                }
            default:
                return element.GetRawText();
        }
    }



    public void WriteFormula(string workbookName, string sheetName, string rangeAddress, object formula)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                range = GetRange(ws, rangeAddress);

                var convertedFormula = ConvertValueForExcel(formula);
                ValidateRangeAndValueDimensions(range, convertedFormula);
                ValidateFormulaValues(convertedFormula);

                try
                {
                    ((dynamic)range).Formula2 = convertedFormula;
                }
                catch
                {
                    range.Formula = convertedFormula;
                }
            }
            finally
            {
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    private void ValidateFormulaValues(object? convertedFormula)
    {
        if (convertedFormula is null)
        {
            throw new ArgumentException("Formula cannot be null.");
        }

        if (convertedFormula is object[,] matrix)
        {
            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    object? cellVal = matrix[r, c];
                    if (cellVal is null)
                    {
                        throw new ArgumentException("Formula cell value cannot be null.");
                    }
                    string formulaStr = cellVal.ToString() ?? "";
                    if (!formulaStr.StartsWith('='))
                    {
                        throw new ArgumentException($"Formula must start with '='. Found: '{formulaStr}'");
                    }
                }
            }
        }
        else
        {
            string formulaStr = convertedFormula.ToString() ?? "";
            if (!formulaStr.StartsWith('='))
            {
                throw new ArgumentException($"Formula must start with '='. Found: '{formulaStr}'");
            }
        }
    }

    public Dictionary<string, string> ReadFormula(string workbookName, string sheetName, string? rangeAddress)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);

                if (string.IsNullOrWhiteSpace(rangeAddress))
                {
                    range = ws.UsedRange;
                }
                else
                {
                    range = GetRange(ws, rangeAddress);
                }

                var results = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (range == null)
                {
                    return results;
                }

                int startRow = range.Row;
                int startCol = range.Column;
                object rawFormula = range.Formula;

                if (rawFormula == null)
                {
                    return results;
                }

                if (rawFormula is object[,] matrix)
                {
                    int rows = matrix.GetLength(0);
                    int cols = matrix.GetLength(1);

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            object? val = matrix[r, c];
                            if (val != null)
                            {
                                string strVal = val.ToString() ?? "";
                                if (strVal.StartsWith('='))
                                {
                                    string cellAddress = $"{GetColumnLetter(startCol + c - 1)}{startRow + r - 1}";
                                    results[cellAddress] = strVal;
                                }
                            }
                        }
                    }
                }
                else
                {
                    string strVal = rawFormula.ToString() ?? "";
                    if (strVal.StartsWith('='))
                    {
                        string cellAddress = $"{GetColumnLetter(startCol)}{startRow}";
                        results[cellAddress] = strVal;
                    }
                }

                return results;
            }
            finally
            {
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public object?[,] ReadTable(string workbookName, string sheetName, string tableName)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.ListObject? table = null;
            Excel.Range? range = null;
            try
            {
                table = GetTable(workbookName, sheetName, tableName);
                range = table.Range;
                object rawValue = range.Value2;

                if (rawValue is object[,] matrix)
                {
                    return matrix;
                }
                else
                {
                    return new object[1, 1] { { rawValue } };
                }
            }
            finally
            {
                SafeReleaseComObject(range);
                SafeReleaseComObject(table);
            }
        });
    }

    public string ReadTableAsMarkdown(string workbookName, string sheetName, string tableName)
    {
        var matrix = ReadTable(workbookName, sheetName, tableName);
        if (matrix == null)
        {
            return string.Empty;
        }

        int rowStart = matrix.GetLowerBound(0);
        int rowEnd = matrix.GetUpperBound(0);
        int colStart = matrix.GetLowerBound(1);
        int colEnd = matrix.GetUpperBound(1);

        int rowCount = rowEnd - rowStart + 1;
        int colCount = colEnd - colStart + 1;

        if (rowCount == 0 || colCount == 0)
        {
            return string.Empty;
        }

        var headers = new string[colCount];
        for (int c = 0; c < colCount; c++)
        {
            headers[c] = matrix[rowStart, colStart + c]?.ToString() ?? "";
        }

        int bodyRowCount = rowCount - 1;
        if (bodyRowCount < 0) bodyRowCount = 0;
        var body = new string[bodyRowCount, colCount];
        for (int r = 0; r < bodyRowCount; r++)
        {
            for (int c = 0; c < colCount; c++)
            {
                body[r, c] = matrix[rowStart + 1 + r, colStart + c]?.ToString() ?? "";
            }
        }

        return MarkdownTableHelper.ToMarkdownTable(headers, body);
    }

    public Dictionary<string, List<string>> ListTables(string workbookName, string? sheetName)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Sheets? sheets = null;
            Excel.Worksheet? ws = null;
            Excel.ListObjects? tables = null;
            try
            {
                wb = GetWorkbook(workbookName);
                var results = new Dictionary<string, List<string>>();

                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    ws = GetWorksheet(wb, sheetName);
                    tables = ws.ListObjects;
                    var tableNames = new List<string>();
                    foreach (Excel.ListObject t in tables)
                    {
                        tableNames.Add(t.Name);
                        SafeReleaseComObject(t);
                    }
                    if (tableNames.Count > 0)
                    {
                        results[ws.Name] = tableNames;
                    }
                }
                else
                {
                    sheets = wb.Sheets;
                    foreach (object s in sheets)
                    {
                        if (s is Excel.Worksheet worksheet)
                        {
                            var sheetTables = worksheet.ListObjects;
                            var tableNames = new List<string>();
                            foreach (Excel.ListObject t in sheetTables)
                            {
                                tableNames.Add(t.Name);
                                SafeReleaseComObject(t);
                            }
                            SafeReleaseComObject(sheetTables);

                            if (tableNames.Count > 0)
                            {
                                results[worksheet.Name] = tableNames;
                            }
                        }
                        SafeReleaseComObject(s);
                    }
                }

                return results;
            }
            finally
            {
                SafeReleaseComObject(tables);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(sheets);
                SafeReleaseComObject(wb);
            }
        });
    }


    public string ConvertToTable(string workbookName, string sheetName, string rangeAddress, string? table, bool hasHeaders)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? excelRange = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? listObj = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);

                if (!string.IsNullOrEmpty(table))
                {
                    try
                    {
                        Excel.ListObject? existingTable = GetTable(workbookName, null, table);
                        SafeReleaseComObject(existingTable);
                        throw new ArgumentException($"Table name '{table}' is already in use by another table in the workbook.");
                    }
                    catch (Exception ex) when (ex is not ArgumentException)
                    {
                        // Table not found means no duplicate exists, which is normal.
                    }
                }

                ws = GetWorksheet(wb, sheetName);
                excelRange = GetRange(ws, rangeAddress);
                tables = ws.ListObjects;

                listObj = tables.Add(Excel.XlListObjectSourceType.xlSrcRange, excelRange,
                    Type.Missing, hasHeaders ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo);

                if (!string.IsNullOrEmpty(table))
                {
                    listObj.Name = table;
                }

                return listObj.Name;
            }
            finally
            {
                SafeReleaseComObject(listObj);
                SafeReleaseComObject(tables);
                SafeReleaseComObject(excelRange);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public string ConvertToRange(string workbookName, string sheetName, string table)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.ListObject? listObj = null;
            Excel.Range? excelRange = null;
            try
            {
                listObj = GetTable(workbookName, sheetName, table);
                excelRange = listObj.Range;
                string address = excelRange.get_Address();
                listObj.Unlist();
                return address;
            }
            finally
            {
                SafeReleaseComObject(excelRange);
                SafeReleaseComObject(listObj);
            }
        });
    }

    public List<FindResult> Find(string workbookName, string? sheetName, string? rangeAddress, string what, bool matchCase, bool wholeWord)
    {
        return ExecuteWithRetry(() =>
        {
            var results = new List<FindResult>();
            Excel.Workbook? wb = null;
            try
            {
                wb = GetWorkbook(workbookName);
                if (!string.IsNullOrEmpty(sheetName))
                {
                    Excel.Worksheet? ws = null;
                    try
                    {
                        ws = GetWorksheet(wb, sheetName);
                        FindInSheet(ws, rangeAddress, what, matchCase, wholeWord, results);
                    }
                    finally
                    {
                        SafeReleaseComObject(ws);
                    }
                }
                else
                {
                    Excel.Sheets? sheets = null;
                    try
                    {
                        sheets = wb.Sheets;
                        foreach (object s in sheets)
                        {
                            if (s is Excel.Worksheet worksheet)
                            {
                                try
                                {
                                    FindInSheet(worksheet, rangeAddress, what, matchCase, wholeWord, results);
                                }
                                finally
                                {
                                    SafeReleaseComObject(worksheet);
                                }
                            }
                            else
                            {
                                SafeReleaseComObject(s);
                            }
                        }
                    }
                    finally
                    {
                        SafeReleaseComObject(sheets);
                    }
                }
                return results;
            }
            finally
            {
                SafeReleaseComObject(wb);
            }
        });
    }

    private void FindInSheet(Excel.Worksheet ws, string? rangeAddress, string what, bool matchCase, bool wholeWord, List<FindResult> results)
    {
        Excel.Range? searchRange = null;
        Excel.Range? firstMatch = null;
        Excel.Range? currentMatch = null;
        try
        {
            searchRange = string.IsNullOrEmpty(rangeAddress) ? ws.UsedRange : GetRange(ws, rangeAddress);
            if (searchRange == null)
            {
                return;
            }

            object lookAt = wholeWord ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart;

            firstMatch = searchRange.Find(what, Type.Missing, Excel.XlFindLookIn.xlValues,
                lookAt, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext,
                matchCase, Type.Missing, Type.Missing);

            if (firstMatch != null)
            {
                string firstAddress = firstMatch.get_Address();
                currentMatch = firstMatch;

                while (currentMatch != null)
                {
                    results.Add(new FindResult(ws.Name, currentMatch.get_Address(), currentMatch.Value2?.ToString() ?? ""));

                    Excel.Range? nextMatch = searchRange.FindNext(currentMatch);

                    if (nextMatch != null && nextMatch.get_Address() == firstAddress)
                    {
                        SafeReleaseComObject(nextMatch);
                        break;
                    }

                    if (currentMatch != firstMatch)
                    {
                        SafeReleaseComObject(currentMatch);
                    }
                    currentMatch = nextMatch;
                }
            }
        }
        finally
        {
            if (currentMatch != null && currentMatch != firstMatch) SafeReleaseComObject(currentMatch);
            SafeReleaseComObject(firstMatch);
            SafeReleaseComObject(searchRange);
        }
    }

    public int Replace(string workbookName, string? sheetName, string? rangeAddress, string what, string replacement, bool matchCase, bool wholeWord)
    {
        return ExecuteWithRetry(() =>
        {
            var findResults = Find(workbookName, sheetName, rangeAddress, what, matchCase, wholeWord);
            if (findResults.Count == 0)
            {
                return 0;
            }

            Excel.Workbook? wb = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                if (!string.IsNullOrEmpty(sheetName))
                {
                    Excel.Worksheet? ws = null;
                    Excel.Range? searchRange = null;
                    try
                    {
                        ws = GetWorksheet(wb, sheetName);
                        searchRange = string.IsNullOrEmpty(rangeAddress) ? ws.UsedRange : GetRange(ws, rangeAddress);
                        if (searchRange != null)
                        {
                            object lookAt = wholeWord ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart;
                            searchRange.Replace(what, replacement, lookAt, Excel.XlSearchOrder.xlByRows,
                                matchCase, Type.Missing, Type.Missing, Type.Missing);
                        }
                    }
                    finally
                    {
                        SafeReleaseComObject(searchRange);
                        SafeReleaseComObject(ws);
                    }
                }
                else
                {
                    Excel.Sheets? sheets = null;
                    try
                    {
                        sheets = wb.Sheets;
                        foreach (object s in sheets)
                        {
                            if (s is Excel.Worksheet worksheet)
                            {
                                Excel.Range? searchRange = null;
                                try
                                {
                                    searchRange = string.IsNullOrEmpty(rangeAddress) ? worksheet.UsedRange : GetRange(worksheet, rangeAddress);
                                    if (searchRange != null)
                                    {
                                        object lookAt = wholeWord ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart;
                                        searchRange.Replace(what, replacement, lookAt, Excel.XlSearchOrder.xlByRows,
                                            matchCase, Type.Missing, Type.Missing, Type.Missing);
                                    }
                                }
                                finally
                                {
                                    SafeReleaseComObject(searchRange);
                                    SafeReleaseComObject(worksheet);
                                }
                            }
                            else
                            {
                                SafeReleaseComObject(s);
                            }
                        }
                    }
                    finally
                    {
                        SafeReleaseComObject(sheets);
                    }
                }
            }
            finally
            {
                SafeReleaseComObject(wb);
            }

            return findResults.Count;
        });
    }

    public string Clear(string workbookName, string sheetName, string? rangeAddress, string? clearType)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? excelRange = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);

                if (string.IsNullOrWhiteSpace(rangeAddress))
                {
                    excelRange = ws.UsedRange;
                }
                else
                {
                    excelRange = GetRange(ws, rangeAddress);
                }

                if (excelRange == null)
                {
                    throw new InvalidOperationException("Target range could not be determined.");
                }

                string address = excelRange.get_Address();
                string type = (clearType ?? "all").Trim().ToLowerInvariant();

                switch (type)
                {
                    case "all":
                        excelRange.Clear();
                        break;
                    case "formats":
                        excelRange.ClearFormats();
                        break;
                    case "contents":
                        excelRange.ClearContents();
                        break;
                    case "comments":
                        excelRange.ClearComments();
                        break;
                    case "hyperlinks":
                        excelRange.ClearHyperlinks();
                        break;
                    default:
                        throw new ArgumentException($"Invalid clear type '{clearType}'. Supported types are: all, formats, contents, comments, hyperlinks.");
                }

                return address;
            }
            finally
            {
                SafeReleaseComObject(excelRange);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public string Delete(string workbookName, string sheetName, string? rangeAddress, string? shiftOption)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? excelRange = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);

                if (string.IsNullOrWhiteSpace(rangeAddress))
                {
                    excelRange = ws.UsedRange;
                }
                else
                {
                    excelRange = GetRange(ws, rangeAddress);
                }

                if (excelRange == null)
                {
                    throw new InvalidOperationException("Target range could not be determined.");
                }

                string address = excelRange.get_Address();
                string shift = (shiftOption ?? "").Trim().ToLowerInvariant();

                switch (shift)
                {
                    case "":
                    case "default":
                        excelRange.Delete();
                        break;
                    case "shift_left":
                    case "shiftleft":
                    case "left":
                    case "xlshiftoleft":
                        excelRange.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
                        break;
                    case "shift_up":
                    case "shiftup":
                    case "up":
                    case "xlshiftup":
                        excelRange.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                        break;
                    case "entire_row":
                    case "entirerow":
                    case "row":
                        Excel.Range? entireRow = null;
                        try
                        {
                            entireRow = excelRange.EntireRow;
                            address = entireRow.get_Address();
                            entireRow.Delete();
                        }
                        finally
                        {
                            SafeReleaseComObject(entireRow);
                        }
                        break;
                    case "entire_column":
                    case "entire_col":
                    case "entirecolumn":
                    case "column":
                    case "col":
                        Excel.Range? entireCol = null;
                        try
                        {
                            entireCol = excelRange.EntireColumn;
                            address = entireCol.get_Address();
                            entireCol.Delete();
                        }
                        finally
                        {
                            SafeReleaseComObject(entireCol);
                        }
                        break;
                    default:
                        throw new ArgumentException($"Invalid shift option '{shiftOption}'. Supported options are: shift_left, shift_up, entire_row, entire_column.");
                }

                return address;
            }
            finally
            {
                SafeReleaseComObject(excelRange);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }


    public void SetStyle(string workbookName, string sheetName, string rangeAddress, CellStyle style)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? excelRange = null;
            Excel.Font? font = null;
            Excel.Interior? interior = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                excelRange = GetRange(ws, rangeAddress);

                if (style.FontName != null || style.FontSize != null || style.Bold != null || style.Italic != null || style.Color != null)
                {
                    font = excelRange.Font;
                    if (style.FontName != null) font.Name = style.FontName;
                    if (style.FontSize != null) font.Size = style.FontSize;
                    if (style.Bold != null) font.Bold = style.Bold;
                    if (style.Italic != null) font.Italic = style.Italic;
                    if (style.Color != null) font.Color = ColorHelper.HexToOleColor(style.Color);
                }

                if (style.BackgroundColor != null)
                {
                    interior = excelRange.Interior;
                    interior.Color = ColorHelper.HexToOleColor(style.BackgroundColor);
                }

                if (style.HorizontalAlignment != null)
                {
                    excelRange.HorizontalAlignment = style.HorizontalAlignment switch
                    {
                        "Left" => Excel.XlHAlign.xlHAlignLeft,
                        "Center" => Excel.XlHAlign.xlHAlignCenter,
                        "Right" => Excel.XlHAlign.xlHAlignRight,
                        _ => excelRange.HorizontalAlignment
                    };
                }

                if (style.VerticalAlignment != null)
                {
                    excelRange.VerticalAlignment = style.VerticalAlignment switch
                    {
                        "Top" => Excel.XlVAlign.xlVAlignTop,
                        "Center" => Excel.XlVAlign.xlVAlignCenter,
                        "Bottom" => Excel.XlVAlign.xlVAlignBottom,
                        _ => excelRange.VerticalAlignment
                    };
                }
            }
            finally
            {
                SafeReleaseComObject(interior);
                SafeReleaseComObject(font);
                SafeReleaseComObject(excelRange);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public CellStyle GetStyle(string workbookName, string sheetName, string? rangeAddress)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? excelRange = null;
            Excel.Range? cells = null;
            Excel.Range? firstCell = null;
            Excel.Font? font = null;
            Excel.Interior? interior = null;
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);

                if (string.IsNullOrWhiteSpace(rangeAddress))
                {
                    excelRange = ws.UsedRange;
                }
                else
                {
                    excelRange = GetRange(ws, rangeAddress);
                }

                if (excelRange == null)
                {
                    throw new Exception("Target range could not be determined.");
                }

                cells = excelRange.Cells;
                firstCell = (Excel.Range)cells[1, 1];

                font = firstCell.Font;
                interior = firstCell.Interior;

                string? fontName = font.Name?.ToString();

                double? fontSize = null;
                if (font.Size is double dVal)
                {
                    fontSize = dVal;
                }
                else if (font.Size is float fVal)
                {
                    fontSize = fVal;
                }
                else if (font.Size is not null)
                {
                    try { fontSize = Convert.ToDouble(font.Size); } catch { }
                }

                bool? bold = null;
                if (font.Bold is bool bVal)
                {
                    bold = bVal;
                }
                else if (font.Bold is not null)
                {
                    try { bold = Convert.ToBoolean(font.Bold); } catch { }
                }

                bool? italic = null;
                if (font.Italic is bool iVal)
                {
                    italic = iVal;
                }
                else if (font.Italic is not null)
                {
                    try { italic = Convert.ToBoolean(font.Italic); } catch { }
                }

                string? color = null;
                if (font.Color is not null)
                {
                    try
                    {
                        long oleColor = Convert.ToInt64(font.Color);
                        color = ColorHelper.OleColorToHex(oleColor);
                    }
                    catch { }
                }

                string? backgroundColor = null;
                if (interior.Color is not null)
                {
                    try
                    {
                        long oleColor = Convert.ToInt64(interior.Color);
                        backgroundColor = ColorHelper.OleColorToHex(oleColor);
                    }
                    catch { }
                }

                string? horizontalAlignment = null;
                if (firstCell.HorizontalAlignment is int alignVal)
                {
                    horizontalAlignment = alignVal switch
                    {
                        (int)Excel.XlHAlign.xlHAlignLeft => "Left",
                        (int)Excel.XlHAlign.xlHAlignCenter => "Center",
                        (int)Excel.XlHAlign.xlHAlignRight => "Right",
                        _ => null
                    };
                }
                else if (firstCell.HorizontalAlignment is not null)
                {
                    try
                    {
                        int alignInt = Convert.ToInt32(firstCell.HorizontalAlignment);
                        horizontalAlignment = alignInt switch
                        {
                            (int)Excel.XlHAlign.xlHAlignLeft => "Left",
                            (int)Excel.XlHAlign.xlHAlignCenter => "Center",
                            (int)Excel.XlHAlign.xlHAlignRight => "Right",
                            _ => null
                        };
                    }
                    catch { }
                }

                string? verticalAlignment = null;
                if (firstCell.VerticalAlignment is int vAlignVal)
                {
                    verticalAlignment = vAlignVal switch
                    {
                        (int)Excel.XlVAlign.xlVAlignTop => "Top",
                        (int)Excel.XlVAlign.xlVAlignCenter => "Center",
                        (int)Excel.XlVAlign.xlVAlignBottom => "Bottom",
                        _ => null
                    };
                }
                else if (firstCell.VerticalAlignment is not null)
                {
                    try
                    {
                        int vAlignInt = Convert.ToInt32(firstCell.VerticalAlignment);
                        verticalAlignment = vAlignInt switch
                        {
                            (int)Excel.XlVAlign.xlVAlignTop => "Top",
                            (int)Excel.XlVAlign.xlVAlignCenter => "Center",
                            (int)Excel.XlVAlign.xlVAlignBottom => "Bottom",
                            _ => null
                        };
                    }
                    catch { }
                }

                return new CellStyle(
                    FontName: fontName,
                    FontSize: fontSize,
                    Bold: bold,
                    Italic: italic,
                    Color: color,
                    BackgroundColor: backgroundColor,
                    HorizontalAlignment: horizontalAlignment,
                    VerticalAlignment: verticalAlignment
                );
            }
            finally
            {
                SafeReleaseComObject(interior);
                SafeReleaseComObject(font);
                SafeReleaseComObject(firstCell);
                SafeReleaseComObject(cells);
                SafeReleaseComObject(excelRange);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public string GetSelection(string workbookName, string sheetName)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Application? app = null;
            object? selectionObj = null;
            Excel.Range? range = null;
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                app = GetApp(createNew: false);
                if (app == null)
                {
                    throw new Exception("Excel is not running. Please open Excel first.");
                }

                wb.Activate();
                ws.Activate();

                selectionObj = app.Selection;
                if (selectionObj is Excel.Range r)
                {
                    range = r;
                    return range.get_Address();
                }
                else
                {
                    throw new InvalidOperationException("The current selection is not a cell range.");
                }
            }
            finally
            {
                SafeReleaseComObject(range);
                if (selectionObj != null && selectionObj != range)
                {
                    SafeReleaseComObject(selectionObj);
                }
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public string SetSelection(string workbookName, string sheetName, string rangeAddress)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? excelRange = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                excelRange = GetRange(ws, rangeAddress);

                wb.Activate();
                ws.Activate();

                excelRange.Select();
                excelRange.Activate();

                return excelRange.get_Address();
            }
            finally
            {
                SafeReleaseComObject(excelRange);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void RenameTable(string workbookName, string sheetName, string tableName, string newTableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("Table name cannot be null or empty.", nameof(tableName));
        }
        if (string.IsNullOrWhiteSpace(newTableName))
        {
            throw new ArgumentException("New table name cannot be null or empty.", nameof(newTableName));
        }

        ExecuteWithRetry(() =>
        {
            Excel.ListObject? targetTable = null;
            try
            {
                targetTable = GetTable(workbookName, sheetName, tableName);

                // Verify duplicate table name workbook-wide (sheetName is empty/null)
                try
                {
                    Excel.ListObject? existingTable = GetTable(workbookName, null, newTableName);
                    try
                    {
                        if (!existingTable.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                        {
                            throw new ArgumentException($"Table name '{newTableName}' is already in use by another table in the workbook.");
                        }
                    }
                    finally
                    {
                        SafeReleaseComObject(existingTable);
                    }
                }
                catch (Exception ex) when (ex is not ArgumentException)
                {
                    // Table not found means no duplicate exists, which is normal.
                }

                if (tableName.Equals(newTableName, StringComparison.OrdinalIgnoreCase) && tableName != newTableName)
                {
                    string tempName = newTableName + "_temp_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                    targetTable.Name = tempName;
                }

                targetTable.Name = newTableName;
            }
            finally
            {
                SafeReleaseComObject(targetTable);
            }
        });
    }

    public void SetListValidation(string workbookName, string sheetName, string rangeAddress, string formula)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            Excel.Validation? validation = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                range = GetRange(ws, rangeAddress);
                validation = range.Validation;
                validation.Delete();
                validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
                    Excel.XlFormatConditionOperator.xlBetween, formula);
                validation.IgnoreBlank = true;
                validation.InCellDropdown = true;
            }
            finally
            {
                SafeReleaseComObject(validation);
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    private Excel.ListObject GetTable(string workbookName, string? sheetName, string tableName)
    {
        Excel.Workbook? wb = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? ws = null;
        Excel.ListObjects? tables = null;
        Excel.ListObject? targetTable = null;
        try
        {
            wb = GetWorkbook(workbookName);
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                ws = GetWorksheet(wb, sheetName);
                tables = ws.ListObjects;
                foreach (Excel.ListObject t in tables)
                {
                    if (t.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetTable = t;
                        break;
                    }
                    SafeReleaseComObject(t);
                }
            }
            else
            {
                sheets = wb.Sheets;
                foreach (object s in sheets)
                {
                    if (s is Excel.Worksheet worksheet)
                    {
                        Excel.ListObjects? sheetTables = null;
                        try
                        {
                            sheetTables = worksheet.ListObjects;
                            foreach (Excel.ListObject t in sheetTables)
                            {
                                if (t.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                                {
                                    targetTable = t;
                                    break;
                                }
                                SafeReleaseComObject(t);
                            }
                        }
                        finally
                        {
                            SafeReleaseComObject(sheetTables);
                            if (targetTable == null)
                            {
                                SafeReleaseComObject(worksheet);
                            }
                        }
                    }
                    else
                    {
                        SafeReleaseComObject(s);
                    }

                    if (targetTable != null)
                    {
                        break;
                    }
                }
            }

            if (targetTable == null)
            {
                string scope = string.IsNullOrWhiteSpace(sheetName) ? "workbook" : $"sheet '{sheetName}'";
                throw new Exception($"Table '{tableName}' not found in {scope}.");
            }

            return targetTable;
        }
        finally
        {
            SafeReleaseComObject(tables);
            SafeReleaseComObject(ws);
            SafeReleaseComObject(sheets);
            SafeReleaseComObject(wb);
        }
    }

    public List<CommentInfo> GetComments(string workbookName, string sheetName)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Comments? comments = null;
            var results = new List<CommentInfo>();
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                comments = ws.Comments;

                foreach (Excel.Comment comment in comments)
                {
                    Excel.Range? parentCell = null;
                    try
                    {
                        parentCell = (Excel.Range)comment.Parent;
                        string address = parentCell.get_Address(false, false);
                        bool visible = comment.Visible;
                        string author = comment.Author;
                        string text = comment.Text();

                        results.Add(new CommentInfo(address, visible, author, text));
                    }
                    finally
                    {
                        SafeReleaseComObject(parentCell);
                        SafeReleaseComObject(comment);
                    }
                }

                return results;
            }
            finally
            {
                SafeReleaseComObject(comments);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void SetComment(string workbookName, string sheetName, string rangeAddress, string text, bool? visible = null)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? excelRange = null;
            Excel.Comment? comment = null;
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                excelRange = GetRange(ws, rangeAddress);

                if (excelRange.Count > 1)
                {
                    throw new ArgumentException($"Setting a comment is only supported for a single cell, but range '{rangeAddress}' contains multiple cells.");
                }

                comment = excelRange.Comment;
                if (comment != null)
                {
                    comment.Delete();
                    SafeReleaseComObject(comment);
                    comment = null;
                }

                comment = excelRange.AddComment(text);

                if (visible.HasValue)
                {
                    comment.Visible = visible.Value;
                }
            }
            finally
            {
                SafeReleaseComObject(comment);
                SafeReleaseComObject(excelRange);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public string AutoFit(string workbookName, string sheetName, string rangeAddress, string? target = null)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            Excel.Range? cols = null;
            Excel.Range? rows = null;
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                range = GetRange(ws, rangeAddress);

                string normTarget = (target ?? "columns").Trim().ToLowerInvariant();

                bool fixColumns = normTarget switch
                {
                    "column" or "columns" or "col" or "width" => true,
                    "both" or "all" or "row_and_column" or "column_and_row" or "rows_and_columns" or "columns_and_rows" or "row_column" or "rowcolumn" => true,
                    "row" or "rows" or "height" => false,
                    _ => throw new ArgumentException($"Invalid autofit target '{target}'. Supported targets are 'columns', 'rows', or 'both'.")
                };

                bool fixRows = normTarget switch
                {
                    "row" or "rows" or "height" => true,
                    "both" or "all" or "row_and_column" or "column_and_row" or "rows_and_columns" or "columns_and_rows" or "row_column" or "rowcolumn" => true,
                    "column" or "columns" or "col" or "width" => false,
                    _ => throw new ArgumentException($"Invalid autofit target '{target}'. Supported targets are 'columns', 'rows', or 'both'.")
                };

                string? colAddress = null;
                string? rowAddress = null;

                if (fixColumns)
                {
                    cols = range.EntireColumn;
                    cols.AutoFit();
                    colAddress = cols.get_Address(false, false);
                }

                if (fixRows)
                {
                    rows = range.EntireRow;
                    rows.AutoFit();
                    rowAddress = rows.get_Address(false, false);
                }

                if (fixColumns && fixRows)
                {
                    return $"{colAddress},{rowAddress}";
                }
                return colAddress ?? rowAddress ?? range.get_Address(false, false);
            }
            finally
            {
                SafeReleaseComObject(rows);
                SafeReleaseComObject(cols);
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public string SetColumnWidth(string workbookName, string sheetName, string rangeAddress, double width)
    {
        if (width < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(width), "Column width cannot be negative.");
        }

        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            Excel.Range? cols = null;
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                range = GetRange(ws, rangeAddress);
                cols = range.EntireColumn;
                cols.ColumnWidth = width;
                return cols.get_Address(false, false);
            }
            finally
            {
                SafeReleaseComObject(cols);
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public string SetRowHeight(string workbookName, string sheetName, string rangeAddress, double height)
    {
        if (height < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(height), "Row height cannot be negative.");
        }

        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            Excel.Range? rows = null;
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                range = GetRange(ws, rangeAddress);
                rows = range.EntireRow;
                rows.RowHeight = height;
                return rows.get_Address(false, false);
            }
            finally
            {
                SafeReleaseComObject(rows);
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void MergeRange(string workbookName, string sheetName, string rangeAddress, bool across = false)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            var app = GetApp(createNew: false);
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                range = GetRange(ws, rangeAddress);

                if (app != null)
                {
                    app.DisplayAlerts = false;
                }

                range.Merge(across);
            }
            finally
            {
                if (app != null)
                {
                    try
                    {
                        app.DisplayAlerts = true;
                    }
                    catch
                    {
                        // Ignore
                    }
                }
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void UnmergeRange(string workbookName, string sheetName, string rangeAddress)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                range = GetRange(ws, rangeAddress);
                range.UnMerge();
            }
            finally
            {
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    private static string GetColumnLetter(int columnNumber)
    {
        int temp;
        string columnName = string.Empty;
        while (columnNumber > 0)
        {
            temp = (columnNumber - 1) % 26;
            columnName = (char)(65 + temp) + columnName;
            columnNumber = (columnNumber - temp - 1) / 26;
        }
        return columnName;
    }

    private void ValidateRangeAndValueDimensions(Excel.Range range, object? convertedValue)
    {
        if (convertedValue is null)
        {
            return;
        }

        Excel.Range? rows = null;
        Excel.Range? cols = null;
        try
        {
            rows = range.Rows;
            cols = range.Columns;
            int rangeRows = rows.Count;
            int rangeCols = cols.Count;

            int valRows = 1;
            int valCols = 1;

            if (convertedValue is object[,] matrix)
            {
                valRows = matrix.GetLength(0);
                valCols = matrix.GetLength(1);
            }

            if (rangeRows != valRows || rangeCols != valCols)
            {
                throw new ArgumentException($"The size of the target range ({rangeRows}x{rangeCols}) does not match the size of the data to be written ({valRows}x{valCols}).");
            }
        }
        finally
        {
            SafeReleaseComObject(cols);
            SafeReleaseComObject(rows);
        }
    }
}
