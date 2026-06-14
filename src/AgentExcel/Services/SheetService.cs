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
            Excel.Worksheet? ws = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
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
}
