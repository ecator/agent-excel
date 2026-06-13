using System.Runtime.InteropServices;

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
                if (sheets != null)
                {
                    Marshal.ReleaseComObject(sheets);
                }
                if (wb != null)
                {
                    Marshal.ReleaseComObject(wb);
                }
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
                if (ws != null)
                {
                    Marshal.ReleaseComObject(ws);
                }
                if (sheets != null)
                {
                    Marshal.ReleaseComObject(sheets);
                }
                if (wb != null)
                {
                    Marshal.ReleaseComObject(wb);
                }
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
                if (ws != null)
                {
                    Marshal.ReleaseComObject(ws);
                }
                if (wb != null)
                {
                    Marshal.ReleaseComObject(wb);
                }
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
                if (ws != null)
                {
                    Marshal.ReleaseComObject(ws);
                }
                if (wb != null)
                {
                    Marshal.ReleaseComObject(wb);
                }
            }
        });
    }
}
