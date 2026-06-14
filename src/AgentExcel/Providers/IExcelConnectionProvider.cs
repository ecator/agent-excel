using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Providers;

public interface IExcelConnectionProvider : IDisposable
{
    Excel.Application? GetApp(bool createNew = false);

    /// <summary>
    /// Safely releases a COM object if it is not null and is a valid COM object.
    /// </summary>
    void SafeReleaseComObject(object? obj);
}
