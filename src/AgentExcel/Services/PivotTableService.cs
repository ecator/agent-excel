using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class PivotTableService : ExcelServiceBase
{
    public PivotTableService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public void CreatePivotTable(string workbookName, string sourceSheetName, string sourceRangeAddress, string targetSheetName, string targetCellAddress, string tableName)
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
                wsSource = GetWorksheet(wb, sourceSheetName);
                wsTarget = GetWorksheet(wb, targetSheetName);
                srcRange = GetRange(wsSource, sourceRangeAddress);
                tgtRange = GetRange(wsTarget, targetCellAddress);

                pcaches = wb.PivotCaches();
                pcache = pcaches.Create(Excel.XlPivotTableSourceType.xlDatabase, srcRange);
                ptable = pcache.CreatePivotTable(tgtRange, tableName);
            }
            finally
            {
                SafeReleaseComObject(ptable);
                SafeReleaseComObject(ptables);
                SafeReleaseComObject(pcache);
                SafeReleaseComObject(pcaches);
                SafeReleaseComObject(tgtRange);
                SafeReleaseComObject(srcRange);
                SafeReleaseComObject(wsTarget);
                SafeReleaseComObject(wsSource);
                SafeReleaseComObject(wb);
            }
        });
    }
}
