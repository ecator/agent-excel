using System.Runtime.InteropServices;

using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public abstract class ExcelServiceBase
{
    protected readonly IExcelConnectionProvider ConnectionProvider;

    protected ExcelServiceBase(IExcelConnectionProvider connectionProvider)
    {
        ConnectionProvider = connectionProvider;
    }

    protected Excel.Application? GetApp(bool createNew = false) => ConnectionProvider.GetApp(createNew);

    public string GetActiveWorkbookName()
    {
        var app = GetApp(createNew: false);
        if (app == null)
        {
            return "None";
        }

        Excel.Workbook? wb = null;
        try
        {
            wb = app.ActiveWorkbook;
            return wb?.Name ?? "None";
        }
        catch
        {
            return "Busy";
        }
        finally
        {
            if (wb != null)
            {
                Marshal.ReleaseComObject(wb);
            }
        }
    }

    protected T ExecuteWithRetry<T>(Func<T> func, int maxRetries = 3)
    {
        for (int i = 0; i < maxRetries; i++)
        {
            try
            {
                return func();
            }
            catch (COMException ex) when ((uint)ex.ErrorCode == 0x80010001) // RPC_E_CALL_REJECTED
            {
                if (i == maxRetries - 1)
                {
                    throw new Exception("Excel is busy (Cell Edit Mode?). Please exit edit mode and try again.");
                }
                Thread.Sleep(500);
            }
        }
        throw new Exception("Unexpected execution failure.");
    }

    protected void ExecuteWithRetry(Action action, int maxRetries = 3)
    {
        ExecuteWithRetry<object?>(() => { action(); return null; }, maxRetries);
    }

    protected Excel.Workbook GetWorkbook(string workbookName, bool createNew = false)
    {
        if (string.IsNullOrEmpty(workbookName))
        {
            throw new ArgumentException("Workbook name cannot be null or empty.", nameof(workbookName));
        }

        var app = GetApp(createNew);
        if (app == null)
        {
            throw new Exception("Excel is not running. Please open Excel first.");
        }

        Excel.Workbooks? wbs = null;
        try
        {
            wbs = app.Workbooks;
            foreach (Excel.Workbook wb in wbs)
            {
                if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                {
                    return wb;
                }
                Marshal.ReleaseComObject(wb);
            }
            throw new Exception($"Workbook '{workbookName}' not found.");
        }
        finally
        {
            if (wbs != null)
            {
                Marshal.ReleaseComObject(wbs);
            }
        }
    }

    protected Excel.Worksheet GetWorksheet(Excel.Workbook workbook, string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
        {
            throw new ArgumentException("Sheet name cannot be null or empty.", nameof(sheetName));
        }

        Excel.Sheets? sheets = null;
        try
        {
            sheets = workbook.Sheets;
            foreach (object s in sheets)
            {
                if (s is Excel.Worksheet ws && ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return ws;
                }
                if (s != null)
                {
                    Marshal.ReleaseComObject(s);
                }
            }
            throw new Exception($"Worksheet '{sheetName}' not found in workbook '{workbook.Name}'.");
        }
        finally
        {
            if (sheets != null)
            {
                Marshal.ReleaseComObject(sheets);
            }
        }
    }
}
