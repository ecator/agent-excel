using System.Runtime.InteropServices;

using AgentExcel.Models;
using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class WorkbookService : ExcelServiceBase
{
    public WorkbookService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }


    public WorkbookInfo GetWorkbookInfo(string workbookName)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Sheets? shs = null;
            var sheetNames = new List<string>();
            try
            {
                wb = GetWorkbook(workbookName);
                shs = wb.Sheets;
                foreach (object s in shs)
                {
                    if (s is Excel.Worksheet ws)
                    {
                        sheetNames.Add(ws.Name);
                        Marshal.ReleaseComObject(ws);
                    }
                    else if (s != null)
                    {
                        Marshal.ReleaseComObject(s);
                    }
                }
                return new WorkbookInfo(wb.Name, wb.FullName, sheetNames);
            }
            finally
            {
                if (shs != null)
                {
                    Marshal.ReleaseComObject(shs);
                }
                if (wb != null)
                {
                    Marshal.ReleaseComObject(wb);
                }
            }
        });
    }
    public List<WorkbookInfo> ListWorkbooks()
    {
        return ExecuteWithRetry(() =>
        {
            var workbooks = new List<WorkbookInfo>();
            Excel.Workbooks? wbs = null;
            try
            {
                var app = GetApp(createNew: false);
                if (app == null)
                {
                    return workbooks;
                }
                wbs = app.Workbooks;
                foreach (Excel.Workbook wb in wbs)
                {
                    workbooks.Add(GetWorkbookInfo(wb.Name));
                    Marshal.ReleaseComObject(wb);
                }
                return workbooks;
            }
            finally
            {
                if (wbs != null)
                {
                    Marshal.ReleaseComObject(wbs);
                }
            }
        });
    }

    public WorkbookInfo OpenWorkbook(string filePath)
    {
        ValidateFilePath(filePath, isSave: false);
        return ExecuteWithRetry(() =>
        {
            Excel.Workbooks? wbs = null;
            Excel.Workbook? wb = null;
            try
            {
                var app = GetApp(createNew: true);
                if (app == null)
                {
                    throw new Exception("Failed to start Excel. Please verify that Microsoft Office 2016 or later is installed.");
                }
                wbs = app.Workbooks;
                wb = wbs.Open(Path.GetFullPath(filePath));
                return GetWorkbookInfo(wb.Name);
            }
            finally
            {
                if (wb != null)
                {
                    Marshal.ReleaseComObject(wb);
                }
                if (wbs != null)
                {
                    Marshal.ReleaseComObject(wbs);
                }
            }
        });
    }

    public WorkbookInfo AddWorkbook()
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbooks? wbs = null;
            Excel.Workbook? wb = null;
            try
            {
                var app = GetApp(createNew: true);
                if (app == null)
                {
                    throw new Exception("Failed to start Excel. Please verify that Microsoft Office 2016 or later is installed.");
                }
                wbs = app.Workbooks;
                wb = wbs.Add(Type.Missing);
                return GetWorkbookInfo(wb.Name);
            }
            finally
            {
                if (wb != null)
                {
                    Marshal.ReleaseComObject(wb);
                }
                if (wbs != null)
                {
                    Marshal.ReleaseComObject(wbs);
                }
            }
        });
    }

    public void SaveWorkbook(string workbookName)
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
                if (wb != null)
                {
                    Marshal.ReleaseComObject(wb);
                }
            }
        });
    }

    public void SaveAsWorkbook(string workbookName, string filePath)
    {
        ValidateFilePath(filePath, isSave: true);
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        Excel.XlFileFormat fileFormat = extension switch
        {
            ".xls" => Excel.XlFileFormat.xlExcel8,
            ".xlsm" => Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
            _ => Excel.XlFileFormat.xlOpenXMLWorkbook
        };

        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            try
            {
                wb = GetWorkbook(workbookName);
                wb.SaveAs(Path.GetFullPath(filePath), FileFormat: fileFormat);
            }
            finally
            {
                if (wb != null)
                {
                    Marshal.ReleaseComObject(wb);
                }
            }
        });
    }

    private void ValidateFilePath(string filePath, bool isSave)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
        }

        string extension = Path.GetExtension(filePath);
        if (string.IsNullOrEmpty(extension))
        {
            throw new ArgumentException("File path must have a valid extension.", nameof(filePath));
        }

        string extLower = extension.ToLowerInvariant();
        if (extLower != ".xls" && extLower != ".xlsx" && extLower != ".xlsm")
        {
            throw new ArgumentException($"Invalid file extension '{extension}'. Only .xls, .xlsx, and .xlsm are allowed.", nameof(filePath));
        }

        if (isSave)
        {
            string? directory = Path.GetDirectoryName(Path.GetFullPath(filePath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                throw new DirectoryNotFoundException($"The directory '{directory}' does not exist.");
            }
        }
        else
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"The file '{filePath}' does not exist.", filePath);
            }
        }
    }

    public void CloseWorkbook(string workbookName, bool saveChanges)
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
                if (wb != null)
                {
                    Marshal.ReleaseComObject(wb);
                }
            }
        });
    }
}
