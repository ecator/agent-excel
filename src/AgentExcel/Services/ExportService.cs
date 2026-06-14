
using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class ExportService : ExcelServiceBase
{
    public ExportService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public void ExportRangeAsImage(string workbookName, string sheetName, string rangeAddress, string outputPath)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? range = null;
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                range = ws.Range[rangeAddress];
                range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                Excel.ChartObjects? charts = (Excel.ChartObjects)ws.ChartObjects();
                Excel.ChartObject? co = charts.Add(0, 0, (double)range.Width, (double)range.Height);
                Excel.Chart? chart = co.Chart;
                try
                {
                    chart.Paste();
                    chart.Export(Path.GetFullPath(outputPath), "PNG");
                }
                finally
                {
                    co.Delete();
                    SafeReleaseComObject(chart);
                    SafeReleaseComObject(co);
                    SafeReleaseComObject(charts);
                }
            }
            finally
            {
                SafeReleaseComObject(range);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void ExportAsPdf(string workbookName, string outputPath)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            try
            {
                wb = GetWorkbook(workbookName);
                wb.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Path.GetFullPath(outputPath));
            }
            finally
            {
                SafeReleaseComObject(wb);
            }
        });
    }
}
