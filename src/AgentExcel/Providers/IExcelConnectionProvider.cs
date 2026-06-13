using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Providers;

public interface IExcelConnectionProvider : IDisposable
{
    Excel.Application? GetApp(bool createNew = false);
}
