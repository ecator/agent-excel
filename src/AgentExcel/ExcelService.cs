using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel
{
    public class ExcelService
    {
        private readonly Excel.Application _app;

        public ExcelService(Excel.Application app)
        {
            _app = app;
        }

        #region Helper Methods
        
        private T ExecuteWithRetry<T>(Func<T> func, int maxRetries = 3)
        {
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    return func();
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x80010001) // RPC_E_CALL_REJECTED
                {
                    if (i == maxRetries - 1) throw new Exception("Excel is busy (Cell Edit Mode?). Please exit edit mode and try again.");
                    Thread.Sleep(500);
                }
            }
            throw new Exception("Unexpected execution failure.");
        }

        private void ExecuteWithRetry(Action action, int maxRetries = 3)
        {
            ExecuteWithRetry<object?>(() => { action(); return null; }, maxRetries);
        }

        private Excel.Workbook GetWorkbook(string? workbookName)
        {
            Excel.Workbooks? wbs = null;
            try
            {
                wbs = _app.Workbooks;
                if (string.IsNullOrEmpty(workbookName))
                {
                    var activeWb = _app.ActiveWorkbook;
                    if (activeWb == null) throw new Exception("No active workbook found.");
                    return activeWb;
                }

                foreach (Excel.Workbook wb in wbs)
                {
                    if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                    {
                        return wb;
                    }
                    // Release intermediate wb if not matched
                    Marshal.ReleaseComObject(wb);
                }
                throw new Exception($"Workbook '{workbookName}' not found.");
            }
            finally
            {
                if (wbs != null) Marshal.ReleaseComObject(wbs);
            }
        }

        private Excel.Worksheet GetWorksheet(Excel.Workbook workbook, string? sheetName)
        {
            Excel.Sheets? sheets = null;
            try
            {
                sheets = workbook.Sheets;
                if (string.IsNullOrEmpty(sheetName))
                {
                    var activeSheet = workbook.ActiveSheet as Excel.Worksheet;
                    if (activeSheet == null) throw new Exception("No active worksheet found.");
                    return activeSheet;
                }

                foreach (object s in sheets)
                {
                    if (s is Excel.Worksheet ws && ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return ws;
                    }
                    if (s != null) Marshal.ReleaseComObject(s);
                }
                throw new Exception($"Worksheet '{sheetName}' not found in workbook '{workbook.Name}'.");
            }
            finally
            {
                if (sheets != null) Marshal.ReleaseComObject(sheets);
            }
        }

        #endregion

        #region Workbook Operations

        public List<string> ListWorkbooks()
        {
            return ExecuteWithRetry(() =>
            {
                var names = new List<string>();
                Excel.Workbooks? wbs = null;
                try
                {
                    wbs = _app.Workbooks;
                    foreach (Excel.Workbook wb in wbs)
                    {
                        names.Add(wb.Name);
                        Marshal.ReleaseComObject(wb);
                    }
                    return names;
                }
                finally
                {
                    if (wbs != null) Marshal.ReleaseComObject(wbs);
                }
            });
        }

        public string OpenWorkbook(string filePath)
        {
            return ExecuteWithRetry(() =>
            {
                Excel.Workbooks? wbs = null;
                Excel.Workbook? wb = null;
                try
                {
                    wbs = _app.Workbooks;
                    wb = wbs.Open(Path.GetFullPath(filePath));
                    return wb.Name;
                }
                finally
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                    if (wbs != null) Marshal.ReleaseComObject(wbs);
                }
            });
        }

        public string AddWorkbook()
        {
            return ExecuteWithRetry(() =>
            {
                Excel.Workbooks? wbs = null;
                Excel.Workbook? wb = null;
                try
                {
                    wbs = _app.Workbooks;
                    wb = wbs.Add();
                    return wb.Name;
                }
                finally
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                    if (wbs != null) Marshal.ReleaseComObject(wbs);
                }
            });
        }

        public void SaveWorkbook(string? workbookName)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    wb.Save();
                }
                finally
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void CloseWorkbook(string? workbookName, bool saveChanges)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    wb.Close(saveChanges);
                }
                finally
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        #endregion

        #region Sheet Operations

        public List<string> ListSheets(string? workbookName)
        {
            return ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Sheets? sheets = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    sheets = wb.Sheets;
                    var names = new List<string>();
                    foreach (object s in sheets)
                    {
                        if (s is Excel.Worksheet ws)
                        {
                            names.Add(ws.Name);
                            Marshal.ReleaseComObject(ws);
                        }
                        else if (s != null)
                        {
                            Marshal.ReleaseComObject(s);
                        }
                    }
                    return names;
                }
                finally
                {
                    if (sheets != null) Marshal.ReleaseComObject(sheets);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void AddSheet(string? workbookName, string? sheetName)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Sheets? sheets = null;
                Excel.Worksheet? ws = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    sheets = wb.Sheets;
                    ws = sheets.Add() as Excel.Worksheet;
                    if (ws != null && !string.IsNullOrEmpty(sheetName))
                    {
                        ws.Name = sheetName;
                    }
                }
                finally
                {
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (sheets != null) Marshal.ReleaseComObject(sheets);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void DeleteSheet(string? workbookName, string sheetName)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    ws = GetWorksheet(wb, sheetName);
                    _app.DisplayAlerts = false; // Prevent confirmation dialog
                    ws.Delete();
                    _app.DisplayAlerts = true;
                }
                finally
                {
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void RenameSheet(string? workbookName, string oldName, string newName)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    ws = GetWorksheet(wb, oldName);
                    ws.Name = newName;
                }
                finally
                {
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void SaveAsWorkbook(string? workbookName, string filePath)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    wb.SaveAs(Path.GetFullPath(filePath));
                }
                finally
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        #endregion

        #region Range Operations

        public object?[,] ReadRange(string? workbookName, string? sheetName, string rangeAddress)
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

        public void WriteRange(string? workbookName, string? sheetName, string rangeAddress, object value)
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

        public string GetUsedRangeAddress(string? workbookName, string? sheetName)
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

        public void WriteFormula(string? workbookName, string? sheetName, string rangeAddress, string formula)
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
                    range = ws.Range[rangeAddress];
                    // Formula2 is preferred for modern Excel (supports dynamic arrays)
                    // If using very old Excel, fallback to .Formula might be needed
                    try { ((dynamic)range).Formula2 = formula; }
                    catch { range.Formula = formula; }
                }
                finally
                {
                    if (range != null) Marshal.ReleaseComObject(range);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public object? ReadFormula(string? workbookName, string? sheetName, string rangeAddress)
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

        public object?[,] ReadTable(string? workbookName, string? sheetName, string tableName)
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
                    
                    // Try to find the table by name
                    foreach (Excel.ListObject t in tables)
                    {
                        if (t.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                        {
                            table = t;
                            break;
                        }
                        Marshal.ReleaseComObject(t);
                    }

                    if (table == null) throw new Exception($"Table '{tableName}' not found in sheet '{ws.Name}'.");

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

        public string ConvertToTable(string? workbookName, string? sheetName, string rangeAddress, string? tableName, bool hasHeaders)
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
                    wb = GetWorkbook(workbookName);
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

        public void ConvertToRange(string? workbookName, string? sheetName, string tableName)
        {
            // ... (existing code)
        }

        public List<FindResult> Find(string? workbookName, string? sheetName, string? rangeAddress, string what, bool matchCase, bool wholeWord)
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
                            
                            // If we looped back to the start, stop
                            if (nextMatch != null && nextMatch.get_Address() == firstAddress)
                            {
                                Marshal.ReleaseComObject(nextMatch);
                                break;
                            }

                            // Release old match before moving to next
                            if (currentMatch != firstMatch) Marshal.ReleaseComObject(currentMatch);
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

        public bool Replace(string? workbookName, string? sheetName, string? rangeAddress, string what, string replacement, bool matchCase, bool wholeWord)
        {
            return ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                Excel.Range? searchRange = null;
                try
                {
                    wb = GetWorkbook(workbookName);
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
            if (!Directory.Exists(folderPath)) throw new Exception($"Folder '{folderPath}' does not exist.");

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
                    wbs = _app.Workbooks;
                    
                    // Check if the file is already open in the current Excel instance
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
                        catch { /* Some workbooks might be in a state where FullName is inaccessible */ }
                        
                        if (wb == null) Marshal.ReleaseComObject(openWb);
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
                                object[,] values = usedRange.Value2 as object[,];
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
                                    catch { /* Some shapes don't support text */ }
                                    finally { Marshal.ReleaseComObject(shape); }
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

        #endregion

        #region Style Operations

        public void SetStyle(string? workbookName, string? sheetName, string rangeAddress, CellStyle style)
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
                    wb = GetWorkbook(workbookName);
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

        #endregion

        #region Advanced Operations (Charts, Export, Pivot, Macros, Validation, Calculate)

        // --- Charts ---
        public List<string> ListCharts(string? workbookName, string? sheetName)
        {
            return ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                Excel.ChartObjects? charts = null;
                var result = new List<string>();
                try
                {
                    wb = GetWorkbook(workbookName);
                    ws = GetWorksheet(wb, sheetName);
                    charts = (Excel.ChartObjects)ws.ChartObjects();
                    foreach (Excel.ChartObject co in charts)
                    {
                        result.Add(co.Name);
                        Marshal.ReleaseComObject(co);
                    }
                    return result;
                }
                finally
                {
                    if (charts != null) Marshal.ReleaseComObject(charts);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public string AddChart(string? workbookName, string? sheetName, string rangeAddress, string chartType, string title)
        {
            return ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                Excel.Range? source = null;
                Excel.ChartObjects? charts = null;
                Excel.ChartObject? co = null;
                Excel.Chart? chart = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    ws = GetWorksheet(wb, sheetName);
                    source = ws.Range[rangeAddress];
                    charts = (Excel.ChartObjects)ws.ChartObjects();
                    co = charts.Add(100, 100, 400, 300);
                    chart = co.Chart;

                    var type = chartType.ToLower() switch
                    {
                        "column" => Excel.XlChartType.xlColumnClustered,
                        "line" => Excel.XlChartType.xlLine,
                        "pie" => Excel.XlChartType.xlPie,
                        "bar" => Excel.XlChartType.xlBarClustered,
                        _ => Excel.XlChartType.xlColumnClustered
                    };

                    chart.SetSourceData(source);
                    chart.ChartType = type;
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = title;

                    return co.Name;
                }
                finally
                {
                    if (chart != null) Marshal.ReleaseComObject(chart);
                    if (co != null) Marshal.ReleaseComObject(co);
                    if (charts != null) Marshal.ReleaseComObject(charts);
                    if (source != null) Marshal.ReleaseComObject(source);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void UpdateChart(string? workbookName, string? sheetName, string chartName, string? rangeAddress, string? chartType, string? title)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                Excel.ChartObjects? charts = null;
                Excel.ChartObject? co = null;
                Excel.Chart? chart = null;
                Excel.Range? source = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    ws = GetWorksheet(wb, sheetName);
                    charts = (Excel.ChartObjects)ws.ChartObjects();
                    co = (Excel.ChartObject)charts.Item(chartName);
                    chart = co.Chart;

                    if (!string.IsNullOrEmpty(rangeAddress))
                    {
                        source = ws.Range[rangeAddress];
                        chart.SetSourceData(source);
                    }

                    if (!string.IsNullOrEmpty(chartType))
                    {
                        chart.ChartType = chartType.ToLower() switch
                        {
                            "column" => Excel.XlChartType.xlColumnClustered,
                            "line" => Excel.XlChartType.xlLine,
                            "pie" => Excel.XlChartType.xlPie,
                            "bar" => Excel.XlChartType.xlBarClustered,
                            _ => chart.ChartType
                        };
                    }

                    if (title != null)
                    {
                        chart.HasTitle = true;
                        chart.ChartTitle.Text = title;
                    }
                }
                finally
                {
                    if (source != null) Marshal.ReleaseComObject(source);
                    if (chart != null) Marshal.ReleaseComObject(chart);
                    if (co != null) Marshal.ReleaseComObject(co);
                    if (charts != null) Marshal.ReleaseComObject(charts);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void DeleteChart(string? workbookName, string? sheetName, string chartName)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                Excel.ChartObjects? charts = null;
                Excel.ChartObject? co = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    ws = GetWorksheet(wb, sheetName);
                    charts = (Excel.ChartObjects)ws.ChartObjects();
                    co = (Excel.ChartObject)charts.Item(chartName);
                    co.Delete();
                }
                finally
                {
                    if (co != null) Marshal.ReleaseComObject(co);
                    if (charts != null) Marshal.ReleaseComObject(charts);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        // --- Export ---
        public void ExportRangeAsImage(string? workbookName, string? sheetName, string rangeAddress, string outputPath)
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
                    range = ws.Range[rangeAddress];
                    range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                    // Use a temporary chart to save the clipboard image
                    Excel.ChartObjects? charts = (Excel.ChartObjects)ws.ChartObjects();
                    Excel.ChartObject? co = charts.Add(0, 0, (double)range.Width, (double)range.Height);
                    Excel.Chart? chart = co.Chart;
                    try
                    {
                        chart.Paste();
                        chart.Export(Path.GetFullPath(outputPath), "PNG");
                    }
                    finally
                    {
                        co.Delete();
                        Marshal.ReleaseComObject(chart);
                        Marshal.ReleaseComObject(co);
                        Marshal.ReleaseComObject(charts);
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

        public void ExportAsPdf(string? workbookName, string outputPath)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    wb.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Path.GetFullPath(outputPath));
                }
                finally
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        // --- Pivot Table ---
        public void CreatePivotTable(string? workbookName, string sourceSheet, string sourceRange, string targetSheet, string targetCell, string tableName)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? wsSource = null;
                Excel.Worksheet? wsTarget = null;
                Excel.Range? srcRange = null;
                Excel.Range? tgtRange = null;
                Excel.PivotCaches? pcaches = null;
                Excel.PivotCache? pcache = null;
                Excel.PivotTables? ptables = null;
                Excel.PivotTable? ptable = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    wsSource = GetWorksheet(wb, sourceSheet);
                    wsTarget = GetWorksheet(wb, targetSheet);
                    srcRange = wsSource.Range[sourceRange];
                    tgtRange = wsTarget.Range[targetCell];

                    pcaches = wb.PivotCaches();
                    pcache = pcaches.Create(Excel.XlPivotTableSourceType.xlDatabase, srcRange);
                    ptable = pcache.CreatePivotTable(tgtRange, tableName);
                }
                finally
                {
                    if (ptable != null) Marshal.ReleaseComObject(ptable);
                    if (ptables != null) Marshal.ReleaseComObject(ptables);
                    if (pcache != null) Marshal.ReleaseComObject(pcache);
                    if (pcaches != null) Marshal.ReleaseComObject(pcaches);
                    if (tgtRange != null) Marshal.ReleaseComObject(tgtRange);
                    if (srcRange != null) Marshal.ReleaseComObject(srcRange);
                    if (wsTarget != null) Marshal.ReleaseComObject(wsTarget);
                    if (wsSource != null) Marshal.ReleaseComObject(wsSource);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        // --- Macros ---
        public object? RunMacro(string macroName, object[]? args)
        {
            return ExecuteWithRetry(() =>
            {
                if (args == null || args.Length == 0)
                    return _app.Run(macroName);
                
                // _app.Run supports up to 30 parameters
                object[] fullArgs = new object[args.Length + 1];
                fullArgs[0] = macroName;
                Array.Copy(args, 0, fullArgs, 1, args.Length);
                
                return typeof(Excel.Application).InvokeMember("Run", 
                    System.Reflection.BindingFlags.InvokeMethod, null, _app, fullArgs);
            });
        }

        // --- Data Validation ---
        public void SetListValidation(string? workbookName, string? sheetName, string rangeAddress, string formula)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                Excel.Range? range = null;
                Excel.Validation? validation = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    ws = GetWorksheet(wb, sheetName);
                    range = ws.Range[rangeAddress];
                    validation = range.Validation;
                    validation.Delete();
                    validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, 
                        Excel.XlFormatConditionOperator.xlBetween, formula);
                    validation.IgnoreBlank = true;
                    validation.InCellDropdown = true;
                }
                finally
                {
                    if (validation != null) Marshal.ReleaseComObject(validation);
                    if (range != null) Marshal.ReleaseComObject(range);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        // --- System ---
        public void Calculate()
        {
            ExecuteWithRetry(() => _app.Calculate());
        }

        #endregion

        #region Power Query & Data Model Operations

        // --- Power Query (M Language) ---
        public List<QueryInfo> ListQueries(string? workbookName)
        {
            return ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                dynamic? queries = null;
                var result = new List<QueryInfo>();
                try
                {
                    wb = GetWorkbook(workbookName);
                    dynamic dynWb = wb;
                    queries = dynWb.Queries;
                    if (queries != null)
                    {
                        foreach (dynamic q in queries)
                        {
                            result.Add(new QueryInfo(q.Name, q.Formula, q.Description));
                            Marshal.ReleaseComObject(q);
                        }
                    }
                    return result;
                }
                finally
                {
                    if (queries != null) Marshal.ReleaseComObject(queries);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void AddOrUpdateQuery(string? workbookName, string name, string formula, string? description)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                dynamic? queries = null;
                dynamic? query = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    dynamic dynWb = wb;
                    queries = dynWb.Queries;
                    if (queries == null) return;

                    // Check if exists
                    bool exists = false;
                    foreach (dynamic q in queries)
                    {
                        string qName = q.Name;
                        if (qName.Equals(name, StringComparison.OrdinalIgnoreCase))
                        {
                            query = q;
                            exists = true;
                            break;
                        }
                        Marshal.ReleaseComObject(q);
                    }

                    if (exists && query != null)
                    {
                        query.Formula = formula;
                        if (description != null) query.Description = description;
                    }
                    else
                    {
                        queries.Add(name, formula, description);
                    }
                }
                finally
                {
                    if (query != null) Marshal.ReleaseComObject(query);
                    if (queries != null) Marshal.ReleaseComObject(queries);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void DeleteQuery(string? workbookName, string name)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                dynamic? queries = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    dynamic dynWb = wb;
                    queries = dynWb.Queries;
                    if (queries == null) return;

                    foreach (dynamic q in queries)
                    {
                        string qName = q.Name;
                        if (qName.Equals(name, StringComparison.OrdinalIgnoreCase))
                        {
                            q.Delete();
                            Marshal.ReleaseComObject(q);
                            break;
                        }
                        Marshal.ReleaseComObject(q);
                    }
                }
                finally
                {
                    if (queries != null) Marshal.ReleaseComObject(queries);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        // --- Data Model & Connections ---
        public void RefreshAllDataConnections(string? workbookName)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    // This refreshes Power Query, Data Model, and external connections
                    wb.RefreshAll();
                }
                finally
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void RefreshModel(string? workbookName)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Model? model = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    model = wb.Model;
                    model.Refresh();
                }
                finally
                {
                    if (model != null) Marshal.ReleaseComObject(model);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        #endregion

        #region Shape Operations

        public List<ShapeInfo> ListShapes(string? workbookName, string? sheetName)
        {
            return ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                Excel.Shapes? shapes = null;
                var result = new List<ShapeInfo>();
                try
                {
                    wb = GetWorkbook(workbookName);
                    ws = GetWorksheet(wb, sheetName);
                    shapes = ws.Shapes;

                    foreach (Excel.Shape shape in shapes)
                    {
                        string? text = null;
                        try
                        {
                            var tf = shape.TextFrame;
                            var chars = tf.Characters();
                            text = chars.Text;
                            Marshal.ReleaseComObject(chars);
                            Marshal.ReleaseComObject(tf);
                        }
                        catch { }

                        result.Add(new ShapeInfo(
                            shape.Name,
                            shape.Type.ToString(),
                            (float)shape.Left,
                            (float)shape.Top,
                            (float)shape.Width,
                            (float)shape.Height,
                            text
                        ));
                        Marshal.ReleaseComObject(shape);
                    }
                    return result;
                }
                finally
                {
                    if (shapes != null) Marshal.ReleaseComObject(shapes);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public string AddShape(string? workbookName, string? sheetName, string type, float left, float top, float width, float height, string? text)
        {
            return ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                Excel.Shapes? shapes = null;
                Excel.Shape? shape = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    ws = GetWorksheet(wb, sheetName);
                    shapes = ws.Shapes;

                    // Support "Textbox" specifically or generic Shapes
                    if (type.Equals("Textbox", StringComparison.OrdinalIgnoreCase))
                    {
                        shape = shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
                    }
                    else
                    {
                        var autoType = type.ToLower() switch
                        {
                            "rectangle" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                            "oval" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
                            "arrow" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRightArrow,
                            _ => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle
                        };
                        shape = shapes.AddShape(autoType, left, top, width, height);
                    }

                    if (text != null)
                    {
                        var tf = shape.TextFrame;
                        var chars = tf.Characters();
                        chars.Text = text;
                        Marshal.ReleaseComObject(chars);
                        Marshal.ReleaseComObject(tf);
                    }

                    return shape.Name;
                }
                finally
                {
                    if (shape != null) Marshal.ReleaseComObject(shape);
                    if (shapes != null) Marshal.ReleaseComObject(shapes);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void UpdateShape(string? workbookName, string? sheetName, string shapeName, float? left, float? top, float? width, float? height, string? text)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                Excel.Shapes? shapes = null;
                Excel.Shape? shape = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    ws = GetWorksheet(wb, sheetName);
                    shapes = ws.Shapes;
                    shape = shapes.Item(shapeName);

                    if (left.HasValue) shape.Left = (float)left.Value;
                    if (top.HasValue) shape.Top = (float)top.Value;
                    if (width.HasValue) shape.Width = (float)width.Value;
                    if (height.HasValue) shape.Height = (float)height.Value;

                    if (text != null)
                    {
                        var tf = shape.TextFrame;
                        var chars = tf.Characters();
                        chars.Text = text;
                        Marshal.ReleaseComObject(chars);
                        Marshal.ReleaseComObject(tf);
                    }
                }
                finally
                {
                    if (shape != null) Marshal.ReleaseComObject(shape);
                    if (shapes != null) Marshal.ReleaseComObject(shapes);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        public void DeleteShape(string? workbookName, string? sheetName, string shapeName)
        {
            ExecuteWithRetry(() =>
            {
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                Excel.Shapes? shapes = null;
                Excel.Shape? shape = null;
                try
                {
                    wb = GetWorkbook(workbookName);
                    ws = GetWorksheet(wb, sheetName);
                    shapes = ws.Shapes;
                    shape = shapes.Item(shapeName);
                    shape.Delete();
                }
                finally
                {
                    if (shape != null) Marshal.ReleaseComObject(shape);
                    if (shapes != null) Marshal.ReleaseComObject(shapes);
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
            });
        }

        #endregion

        public record GrepResult(string File, string Sheet, string Type, string Location, string Content);

        public record ShapeInfo(string Name, string Type, float Left, float Top, float Width, float Height, string? Text);

        public record CellStyle(
            string? FontName = null,
            double? FontSize = null,
            bool? Bold = null,
            bool? Italic = null,
            string? Color = null,           // Hex like "#FF0000"
            string? BackgroundColor = null, // Hex like "#FFFF00"
            string? HorizontalAlignment = null, // "Left", "Center", "Right"
            string? VerticalAlignment = null   // "Top", "Center", "Bottom"
        );

        public record QueryInfo(string Name, string Formula, string Description);

        public record FindResult(string Sheet, string Address, string Value);
    }
}
