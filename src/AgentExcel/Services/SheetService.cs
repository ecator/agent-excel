using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class SheetService : ExcelServiceBase
{
    public SheetService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public List<string> ListSheets(string workbookName)
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
                        SafeReleaseComObject(ws);
                    }
                    else if (s != null)
                    {
                        SafeReleaseComObject(s);
                    }
                }
                return names;
            }
            finally
            {
                SafeReleaseComObject(sheets);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void AddSheet(string workbookName, string sheetName)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Sheets? sheets = null;
            Excel.Worksheet? ws = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                if (HasSheet(wb, sheetName))
                {
                    throw new InvalidOperationException($"A sheet named '{sheetName}' already exists.");
                }
                sheets = wb.Sheets;
                ws = sheets.Add() as Excel.Worksheet;
                if (ws != null)
                {
                    ws.Name = sheetName;
                }
            }
            finally
            {
                SafeReleaseComObject(ws);
                SafeReleaseComObject(sheets);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void DeleteSheet(string workbookName, string sheetName)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Sheets? sheets = null;
            Excel.Worksheet? ws = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                sheets = wb.Sheets;
                if (sheets.Count <= 1)
                {
                    throw new InvalidOperationException("Cannot delete the only sheet in the workbook.");
                }

                ws = GetWorksheet(wb, sheetName);
                var app = GetApp(createNew: false);
                if (app != null)
                {
                    app.DisplayAlerts = false; // Prevent confirmation dialog
                }
                ws.Delete();
                if (app != null)
                {
                    app.DisplayAlerts = true;
                }
            }
            finally
            {
                SafeReleaseComObject(ws);
                SafeReleaseComObject(sheets);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void RenameSheet(string workbookName, string oldName, string newName)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                if (!newName.Equals(oldName, StringComparison.OrdinalIgnoreCase) && HasSheet(wb, newName))
                {
                    throw new InvalidOperationException($"A sheet named '{newName}' already exists.");
                }
                ws = GetWorksheet(wb, oldName);
                ws.Name = newName;
            }
            finally
            {
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    /// <summary>
    /// Copies a worksheet to the target workbook. If the target workbook is omitted, the source workbook is used.
    /// Checks for sheet name collisions in the target workbook before copying.
    /// Supports target position: 0 for first, -1 for last. If null, defaults to immediately after the source sheet (same workbook) or last position (different workbook).
    /// </summary>
    /// <param name="sourceWorkbookName">The source workbook containing the worksheet.</param>
    /// <param name="oldName">The name of the worksheet to copy.</param>
    /// <param name="newName">The new name of the copied worksheet (optional). If null or empty, it defaults to oldName.</param>
    /// <param name="targetWorkbookName">The target workbook (optional). If null or empty, defaults to the source workbook.</param>
    /// <param name="position">The target position (optional).</param>
    public void CopySheet(string sourceWorkbookName, string oldName, string? newName, string? targetWorkbookName, int? position = null)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? sourceWb = null;
            Excel.Worksheet? sourceWs = null;
            Excel.Workbook? targetWb = null;
            Excel.Sheets? targetSheets = null;
            object? relativeSheet = null;
            bool isBefore = false;
            bool isSameWorkbook = false;
            try
            {
                sourceWb = GetWorkbook(sourceWorkbookName, createNew: true);
                sourceWs = GetWorksheet(sourceWb, oldName);

                isSameWorkbook = string.IsNullOrEmpty(targetWorkbookName) ||
                                 targetWorkbookName.Equals(sourceWorkbookName, StringComparison.OrdinalIgnoreCase);

                if (isSameWorkbook)
                {
                    targetWb = sourceWb;
                }
                else
                {
                    targetWb = GetWorkbook(targetWorkbookName!, createNew: true);
                }

                string finalNewName = string.IsNullOrEmpty(newName) ? oldName : newName;

                if (HasSheet(targetWb, finalNewName))
                {
                    string targetName = isSameWorkbook ? sourceWorkbookName : targetWorkbookName!;
                    throw new Exception($"A sheet named '{finalNewName}' already exists in target workbook '{targetName}'.");
                }

                targetSheets = targetWb.Sheets;
                var beforeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (object s in targetSheets)
                {
                    if (s is Excel.Worksheet ws)
                    {
                        beforeNames.Add(ws.Name);
                        SafeReleaseComObject(ws);
                    }
                    else if (s != null)
                    {
                        SafeReleaseComObject(s);
                    }
                }

                int count = targetSheets.Count;

                if (position == null)
                {
                    if (isSameWorkbook)
                    {
                        relativeSheet = sourceWs;
                        isBefore = false;
                    }
                    else
                    {
                        relativeSheet = targetSheets[count];
                        isBefore = false;
                    }
                }
                else
                {
                    int targetIndex = position.Value >= 0 ? position.Value : count + position.Value + 1;
                    targetIndex = Math.Max(0, Math.Min(targetIndex, count));

                    if (targetIndex < count)
                    {
                        relativeSheet = targetSheets[targetIndex + 1];
                        isBefore = true;
                    }
                    else
                    {
                        relativeSheet = targetSheets[count];
                        isBefore = false;
                    }
                }

                if (isBefore)
                {
                    sourceWs.Copy(relativeSheet, Type.Missing);
                }
                else
                {
                    sourceWs.Copy(Type.Missing, relativeSheet);
                }

                // Retrieve the updated sheets list to find the newly added sheet and rename it
                Excel.Sheets? postSheets = null;
                Excel.Worksheet? newWs = null;
                try
                {
                    postSheets = targetWb.Sheets;
                    string? addedSheetName = null;
                    foreach (object s in postSheets)
                    {
                        if (s is Excel.Worksheet ws)
                        {
                            string wsName = ws.Name;
                            if (!beforeNames.Contains(wsName))
                            {
                                addedSheetName = wsName;
                                newWs = ws; // Hold onto this COM reference to rename it!
                            }
                            else
                            {
                                SafeReleaseComObject(ws);
                            }
                        }
                        else if (s != null)
                        {
                            SafeReleaseComObject(s);
                        }
                    }

                    if (newWs != null && addedSheetName != null)
                    {
                        if (!addedSheetName.Equals(finalNewName, StringComparison.OrdinalIgnoreCase))
                        {
                            newWs.Name = finalNewName;
                        }
                    }
                }
                finally
                {
                    SafeReleaseComObject(newWs);
                    SafeReleaseComObject(postSheets);
                }
            }
            finally
            {
                if (relativeSheet != sourceWs)
                {
                    SafeReleaseComObject(relativeSheet);
                }
                SafeReleaseComObject(targetSheets);
                SafeReleaseComObject(sourceWs);
                if (!isSameWorkbook)
                {
                    SafeReleaseComObject(targetWb);
                }
                SafeReleaseComObject(sourceWb);
            }
        });
    }

    /// <summary>
    /// Moves a worksheet to the target workbook. If the target workbook is omitted, the source workbook is used.
    /// Checks for sheet name collisions in the target workbook if moving to a different workbook.
    /// Supports target position: 0 for first, -1 for last. If null, defaults to the last position (-1).
    /// </summary>
    /// <param name="sourceWorkbookName">The source workbook containing the worksheet.</param>
    /// <param name="sheetName">The name of the worksheet to move.</param>
    /// <param name="targetWorkbookName">The target workbook (optional). If null or empty, defaults to the source workbook.</param>
    /// <param name="position">The target position (optional).</param>
    public void MoveSheet(string sourceWorkbookName, string sheetName, string? targetWorkbookName, int? position = null)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? sourceWb = null;
            Excel.Worksheet? sourceWs = null;
            Excel.Workbook? targetWb = null;
            Excel.Sheets? targetSheets = null;
            object? relativeSheet = null;
            bool isBefore = false;
            bool isSameWorkbook = false;
            try
            {
                sourceWb = GetWorkbook(sourceWorkbookName, createNew: true);

                // Check if the source workbook only has one sheet
                Excel.Sheets? sourceSheets = null;
                try
                {
                    sourceSheets = sourceWb.Sheets;
                    if (sourceSheets.Count <= 1)
                    {
                        throw new InvalidOperationException("Cannot move the only sheet in the workbook.");
                    }
                }
                finally
                {
                    SafeReleaseComObject(sourceSheets);
                }

                sourceWs = GetWorksheet(sourceWb, sheetName);

                isSameWorkbook = string.IsNullOrEmpty(targetWorkbookName) ||
                                 targetWorkbookName.Equals(sourceWorkbookName, StringComparison.OrdinalIgnoreCase);

                if (isSameWorkbook)
                {
                    targetWb = sourceWb;
                }
                else
                {
                    targetWb = GetWorkbook(targetWorkbookName!, createNew: true);
                }

                if (!isSameWorkbook && HasSheet(targetWb, sheetName))
                {
                    throw new Exception($"A sheet named '{sheetName}' already exists in target workbook '{targetWorkbookName}'.");
                }

                targetSheets = targetWb.Sheets;
                int count = targetSheets.Count;

                int targetIndex;
                if (position == null)
                {
                    targetIndex = isSameWorkbook ? count - 1 : count;
                }
                else if (position.Value >= 0)
                {
                    targetIndex = position.Value;
                }
                else
                {
                    targetIndex = isSameWorkbook ? count + position.Value : count + position.Value + 1;
                }

                int maxIndex = isSameWorkbook ? count - 1 : count;
                targetIndex = Math.Max(0, Math.Min(targetIndex, maxIndex));

                if (isSameWorkbook)
                {
                    int srcIndex = sourceWs.Index;
                    int destIndex = targetIndex + 1;

                    if (destIndex == srcIndex)
                    {
                        relativeSheet = null;
                    }
                    else if (destIndex < srcIndex)
                    {
                        relativeSheet = targetSheets[destIndex];
                        isBefore = true;
                    }
                    else // destIndex > srcIndex
                    {
                        relativeSheet = targetSheets[destIndex];
                        isBefore = false;
                    }
                }
                else
                {
                    if (targetIndex < count)
                    {
                        relativeSheet = targetSheets[targetIndex + 1];
                        isBefore = true;
                    }
                    else
                    {
                        relativeSheet = targetSheets[count];
                        isBefore = false;
                    }
                }

                if (relativeSheet != null || !isSameWorkbook)
                {
                    if (isBefore)
                    {
                        sourceWs.Move(relativeSheet, Type.Missing);
                    }
                    else
                    {
                        sourceWs.Move(Type.Missing, relativeSheet);
                    }
                }
            }
            finally
            {
                if (relativeSheet != sourceWs)
                {
                    SafeReleaseComObject(relativeSheet);
                }
                SafeReleaseComObject(targetSheets);
                SafeReleaseComObject(sourceWs);
                if (!isSameWorkbook)
                {
                    SafeReleaseComObject(targetWb);
                }
                SafeReleaseComObject(sourceWb);
            }
        });
    }

    private bool HasSheet(Excel.Workbook workbook, string name)
    {
        Excel.Sheets? sheets = null;
        try
        {
            sheets = workbook.Sheets;
            foreach (object s in sheets)
            {
                if (s is Excel.Worksheet ws)
                {
                    bool match = ws.Name.Equals(name, StringComparison.OrdinalIgnoreCase);
                    SafeReleaseComObject(ws);
                    if (match)
                    {
                        return true;
                    }
                }
                else if (s != null)
                {
                    SafeReleaseComObject(s);
                }
            }
            return false;
        }
        finally
        {
            SafeReleaseComObject(sheets);
        }
    }
}

