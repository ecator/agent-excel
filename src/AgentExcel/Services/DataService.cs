using System.Diagnostics;

using AgentExcel.Models;
using AgentExcel.Providers;
using AgentExcel.Utils;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class DataService : ExcelServiceBase
{
    public DataService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
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
                range.Value2 = value;
            }
            finally
            {
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }



    public void WriteFormula(string workbookName, string sheetName, string rangeAddress, string formula)
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
                try
                {
                    ((dynamic)range).Formula2 = formula;
                }
                catch
                {
                    range.Formula = formula;
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

    public object? ReadFormula(string workbookName, string sheetName, string rangeAddress)
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
                range = GetRange(ws, rangeAddress);
                return range.Formula;
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


    public string ConvertToTable(string workbookName, string sheetName, string range, string? table, bool hasHeaders)
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
                excelRange = GetRange(ws, range);
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

    public List<FindResult> Find(string workbookName, string sheetName, string? rangeAddress, string what, bool matchCase, bool wholeWord)
    {
        return ExecuteWithRetry(() =>
        {
            var results = new List<FindResult>();
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? searchRange = null;
            Excel.Range? firstMatch = null;
            Excel.Range? currentMatch = null;

            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                searchRange = string.IsNullOrEmpty(rangeAddress) ? ws.UsedRange : GetRange(ws, rangeAddress);

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
                return results;
            }
            finally
            {
                if (currentMatch != null && currentMatch != firstMatch) SafeReleaseComObject(currentMatch);
                SafeReleaseComObject(firstMatch);
                SafeReleaseComObject(searchRange);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public bool Replace(string workbookName, string sheetName, string? rangeAddress, string what, string replacement, bool matchCase, bool wholeWord)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? searchRange = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                searchRange = string.IsNullOrEmpty(rangeAddress) ? ws.UsedRange : GetRange(ws, rangeAddress);

                object lookAt = wholeWord ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart;

                return searchRange.Replace(what, replacement, lookAt, Excel.XlSearchOrder.xlByRows,
                    matchCase, Type.Missing, Type.Missing, Type.Missing);
            }
            finally
            {
                SafeReleaseComObject(searchRange);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public List<GrepResult> SearchInFolder(string folderPath, string pattern)
    {
        var results = new List<GrepResult>();
        if (!Directory.Exists(folderPath))
        {
            throw new Exception($"Folder '{folderPath}' does not exist.");
        }

        string[] extensions = { "*.xlsx", "*.xls", "*.xlsm", "*.xlsb" };
        var files = extensions.SelectMany(ext => Directory.GetFiles(folderPath, ext)).ToArray();
        string fullFolderPath = Path.GetFullPath(folderPath);

        foreach (var file in files)
        {
            string fullPath = Path.GetFullPath(file);
            Excel.Workbooks? wbs = null;
            Excel.Workbook? wb = null;
            Excel.Sheets? sheets = null;
            bool wasAlreadyOpen = false;

            try
            {
                var app = GetApp(createNew: true);
                if (app == null)
                {
                    throw new Exception("Failed to start Excel. Please verify that Microsoft Office 2016 or later is installed.");
                }
                wbs = app.Workbooks;

                foreach (Excel.Workbook openWb in wbs)
                {
                    try
                    {
                        if (openWb.FullName.Equals(fullPath, StringComparison.OrdinalIgnoreCase))
                        {
                            wb = openWb;
                            wasAlreadyOpen = true;
                            break;
                        }
                    }
                    catch
                    {
                        // Some workbooks might be in a state where FullName is inaccessible
                    }

                    if (wb == null)
                    {
                        SafeReleaseComObject(openWb);
                    }
                }

                if (wb == null)
                {
                    wb = ExecuteWithRetry(() => wbs.Open(fullPath, ReadOnly: true));
                }

                sheets = wb.Sheets;
                foreach (object s in sheets)
                {
                    if (s is Excel.Worksheet ws)
                    {
                        // 1. Search in Cells (UsedRange)
                        Excel.Range? usedRange = null;
                        try
                        {
                            usedRange = ws.UsedRange;
                            object[,]? values = usedRange.Value2 as object[,];
                            if (values != null)
                            {
                                int rows = values.GetLength(0);
                                int cols = values.GetLength(1);
                                for (int r = 1; r <= rows; r++)
                                {
                                    for (int c = 1; c <= cols; c++)
                                    {
                                        string cellText = values[r, c]?.ToString() ?? "";
                                        if (cellText.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                                        {
                                            Excel.Range? cell = ws.Cells[r, c] as Excel.Range;
                                            results.Add(new GrepResult(Path.GetFileName(file), ws.Name, "Cell", cell?.get_Address() ?? $"R{r}C{c}", cellText));
                                            SafeReleaseComObject(cell);
                                        }
                                    }
                                }
                            }
                        }
                        finally
                        {
                            SafeReleaseComObject(usedRange);
                        }

                        // 2. Search in Shapes (Textboxes, etc.)
                        Excel.Shapes? shapes = null;
                        try
                        {
                            shapes = ws.Shapes;
                            foreach (Excel.Shape shape in shapes)
                            {
                                try
                                {
                                    var textFrame = shape.TextFrame;
                                    var characters = textFrame.Characters();
                                    string shapeText = characters.Text;
                                    if (!string.IsNullOrEmpty(shapeText) && shapeText.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                                    {
                                        results.Add(new GrepResult(Path.GetFileName(file), ws.Name, "Shape", shape.Name, shapeText));
                                    }
                                    SafeReleaseComObject(characters);
                                    SafeReleaseComObject(textFrame);
                                }
                                catch
                                {
                                    // Some shapes don't support text
                                }
                                finally
                                {
                                    SafeReleaseComObject(shape);
                                }
                            }
                        }
                        finally
                        {
                            SafeReleaseComObject(shapes);
                        }

                        SafeReleaseComObject(ws);
                    }
                    else if (s != null)
                    {
                        SafeReleaseComObject(s);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to grep file {file}: {ex.Message}");
            }
            finally
            {
                if (wb != null)
                {
                    if (!wasAlreadyOpen)
                    {
                        wb.Close(SaveChanges: false);
                    }
                    SafeReleaseComObject(wb);
                }
                SafeReleaseComObject(wbs);
            }
        }

        return results;
    }

    public void SetStyle(string workbookName, string sheetName, string rangeAddress, CellStyle style)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            Excel.Font? font = null;
            Excel.Interior? interior = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                range = GetRange(ws, rangeAddress);

                if (style.FontName != null || style.FontSize != null || style.Bold != null || style.Italic != null || style.Color != null)
                {
                    font = range.Font;
                    if (style.FontName != null) font.Name = style.FontName;
                    if (style.FontSize != null) font.Size = style.FontSize;
                    if (style.Bold != null) font.Bold = style.Bold;
                    if (style.Italic != null) font.Italic = style.Italic;
                    if (style.Color != null) font.Color = ColorHelper.HexToOleColor(style.Color);
                }

                if (style.BackgroundColor != null)
                {
                    interior = range.Interior;
                    interior.Color = ColorHelper.HexToOleColor(style.BackgroundColor);
                }

                if (style.HorizontalAlignment != null)
                {
                    range.HorizontalAlignment = style.HorizontalAlignment switch
                    {
                        "Left" => Excel.XlHAlign.xlHAlignLeft,
                        "Center" => Excel.XlHAlign.xlHAlignCenter,
                        "Right" => Excel.XlHAlign.xlHAlignRight,
                        _ => range.HorizontalAlignment
                    };
                }

                if (style.VerticalAlignment != null)
                {
                    range.VerticalAlignment = style.VerticalAlignment switch
                    {
                        "Top" => Excel.XlVAlign.xlVAlignTop,
                        "Center" => Excel.XlVAlign.xlVAlignCenter,
                        "Bottom" => Excel.XlVAlign.xlVAlignBottom,
                        _ => range.VerticalAlignment
                    };
                }
            }
            finally
            {
                SafeReleaseComObject(interior);
                SafeReleaseComObject(font);
                SafeReleaseComObject(range);
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

    private static Excel.Range GetRange(Excel.Worksheet ws, string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
        {
            throw new ArgumentException("Range address cannot be null or empty.", nameof(rangeAddress));
        }

        try
        {
            return ws.Range[rangeAddress];
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            throw new ArgumentException($"Invalid Excel range address: '{rangeAddress}'. Ensure it follows a valid format (e.g., 'A1', 'A1:B2', 'A:B').", ex);
        }
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
}
