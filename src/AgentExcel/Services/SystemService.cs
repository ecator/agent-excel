using System.Runtime.InteropServices;

using AgentExcel.Models;
using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class SystemService : ExcelServiceBase
{
    public SystemService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public void Shutdown()
    {
        ConnectionProvider.Dispose();
    }

    public string GetActiveWorkbookName()
    {
        Excel.Workbook? wb = null;
        try
        {
            wb = GetActiveWorkbook();
            return wb.Name;
        }
        catch (COMException ex) when ((uint)ex.ErrorCode == 0x80010001)
        {
            return "Busy";
        }
        catch
        {
            return "None";
        }
        finally
        {
            SafeReleaseComObject(wb);
        }
    }

    public void Calculate()
    {
        ExecuteWithRetry(() =>
        {
            var app = GetApp(createNew: true);
            if (app != null)
            {
                app.Calculate();
            }
        });
    }
}
