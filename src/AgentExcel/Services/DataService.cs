using System.Diagnostics;
using System.Runtime.InteropServices;

using AgentExcel.Models;
using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class DataService : ExcelServiceBase
{
    public DataService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public object?[,] ReadRange(string workbookName, string sheetName, string rangeAddress)
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
                range = ws.Range[rangeAddress];

                object[,] values;
                object rawValue = range.Value2;

                if (rawValue is object[,] matrix)
                {
                    values = matrix;
                }
                else
                {
                    values = new object[1, 1] { { rawValue } };
                }
                return values;
            }
            finally
            {
                if (range != null) Marshal.ReleaseComObject(range);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
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
                range = ws.Range[rangeAddress];
                range.Value2 = value;
            }
            finally
            {
                if (range != null) Marshal.ReleaseComObject(range);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        });
    }

    public string GetUsedRangeAddress(string workbookName, string sheetName)
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
                range = ws.UsedRange;
                return range.get_Address();
            }
            finally
            {
                if (range != null) Marshal.ReleaseComObject(range);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
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
                range = ws.Range[rangeAddress];
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
                if (range != null) Marshal.ReleaseComObject(range);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
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
                range = ws.Range[rangeAddress];
                return range.Formula;
            }
            finally
            {
                if (range != null) Marshal.ReleaseComObject(range);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        });
    }

    public object?[,] ReadTable(string workbookName, string sheetName, string tableName)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            Excel.Range? range = null;
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                tables = ws.ListObjects;

                foreach (Excel.ListObject t in tables)
                {
                    if (t.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        table = t;
                        break;
                    }
                    Marshal.ReleaseComObject(t);
                }

                if (table == null)
                {
                    throw new Exception($"Table '{tableName}' not found in sheet '{ws.Name}'.");
                }

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
                if (range != null) Marshal.ReleaseComObject(range);
                if (table != null) Marshal.ReleaseComObject(table);
                if (tables != null) Marshal.ReleaseComObject(tables);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        });
    }

    public string ConvertToTable(string workbookName, string sheetName, string rangeAddress, string? tableName, bool hasHeaders)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                range = ws.Range[rangeAddress];
                tables = ws.ListObjects;

                table = tables.Add(Excel.XlListObjectSourceType.xlSrcRange, range,
                    Type.Missing, hasHeaders ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo);

                if (!string.IsNullOrEmpty(tableName))
                {
                    table.Name = tableName;
                }

                return table.Name;
            }
            finally
            {
                if (table != null) Marshal.ReleaseComObject(table);
                if (tables != null) Marshal.ReleaseComObject(tables);
                if (range != null) Marshal.ReleaseComObject(range);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        });
    }

    public void ConvertToRange(string workbookName, string sheetName, string tableName)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                tables = ws.ListObjects;

                foreach (Excel.ListObject t in tables)
                {
                    if (t.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        table = t;
                        break;
                    }
                    Marshal.ReleaseComObject(t);
                }

                if (table == null)
                {
                    throw new Exception($"Table '{tableName}' not found in sheet '{ws.Name}'.");
                }

                table.Unlist();
            }
            finally
            {
                if (table != null) Marshal.ReleaseComObject(table);
                if (tables != null) Marshal.ReleaseComObject(tables);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
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
                searchRange = string.IsNullOrEmpty(rangeAddress) ? ws.UsedRange : ws.Range[rangeAddress];

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
                            Marshal.ReleaseComObject(nextMatch);
                            break;
                        }

                        if (currentMatch != firstMatch)
                        {
                            Marshal.ReleaseComObject(currentMatch);
                        }
                        currentMatch = nextMatch;
                    }
                }
                return results;
            }
            finally
            {
                if (currentMatch != null && currentMatch != firstMatch) Marshal.ReleaseComObject(currentMatch);
                if (firstMatch != null) Marshal.ReleaseComObject(firstMatch);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
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
                searchRange = string.IsNullOrEmpty(rangeAddress) ? ws.UsedRange : ws.Range[rangeAddress];

                object lookAt = wholeWord ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart;

                return searchRange.Replace(what, replacement, lookAt, Excel.XlSearchOrder.xlByRows,
                    matchCase, Type.Missing, Type.Missing, Type.Missing);
            }
            finally
            {
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
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
                        Marshal.ReleaseComObject(openWb);
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
                                            if (cell != null) Marshal.ReleaseComObject(cell);
                                        }
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (usedRange != null) Marshal.ReleaseComObject(usedRange);
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
                                    Marshal.ReleaseComObject(characters);
                                    Marshal.ReleaseComObject(textFrame);
                                }
                                catch
                                {
                                    // Some shapes don't support text
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(shape);
                                }
                            }
                        }
                        finally
                        {
                            if (shapes != null) Marshal.ReleaseComObject(shapes);
                        }

                        Marshal.ReleaseComObject(ws);
                    }
                    else if (s != null)
                    {
                        Marshal.ReleaseComObject(s);
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
                    Marshal.ReleaseComObject(wb);
                }
                if (wbs != null) Marshal.ReleaseComObject(wbs);
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
                range = ws.Range[rangeAddress];

                if (style.FontName != null || style.FontSize != null || style.Bold != null || style.Italic != null || style.Color != null)
                {
                    font = range.Font;
                    if (style.FontName != null) font.Name = style.FontName;
                    if (style.FontSize != null) font.Size = style.FontSize;
                    if (style.Bold != null) font.Bold = style.Bold;
                    if (style.Italic != null) font.Italic = style.Italic;
                    if (style.Color != null) font.Color = HexToOleColor(style.Color);
                }

                if (style.BackgroundColor != null)
                {
                    interior = range.Interior;
                    interior.Color = HexToOleColor(style.BackgroundColor);
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
                if (interior != null) Marshal.ReleaseComObject(interior);
                if (font != null) Marshal.ReleaseComObject(font);
                if (range != null) Marshal.ReleaseComObject(range);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        });
    }

    private int HexToOleColor(string hex)
    {
        try
        {
            hex = hex.TrimStart('#');
            if (hex.Length == 6)
            {
                int r = int.Parse(hex.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                int g = int.Parse(hex.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
                int b = int.Parse(hex.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
                return (b << 16) | (g << 8) | r;
            }
        }
        catch { }
        return 0;
    }
}
