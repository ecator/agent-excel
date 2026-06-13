using System.Runtime.InteropServices;

using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class PivotTableService : ExcelServiceBase
{
    public PivotTableService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public void CreatePivotTable(string workbookName, string sourceSheet, string sourceRange, string targetSheet, string targetCell, string tableName)
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
                wb = GetWorkbook(workbookName, createNew: true);
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
}
