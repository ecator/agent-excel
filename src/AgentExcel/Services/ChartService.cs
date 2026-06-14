
using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class ChartService : ExcelServiceBase
{
    public ChartService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public List<string> ListCharts(string workbookName, string sheetName)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.ChartObjects? charts = null;
            var result = new List<string>();
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                charts = (Excel.ChartObjects)ws.ChartObjects();
                foreach (Excel.ChartObject co in charts)
                {
                    result.Add(co.Name);
                    SafeReleaseComObject(co);
                }
                return result;
            }
            finally
            {
                SafeReleaseComObject(charts);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public string AddChart(string workbookName, string sheetName, string rangeAddress, string chartType, string title)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? source = null;
            Excel.ChartObjects? charts = null;
            Excel.ChartObject? co = null;
            Excel.Chart? chart = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                source = ws.Range[rangeAddress];
                charts = (Excel.ChartObjects)ws.ChartObjects();
                co = charts.Add(100, 100, 400, 300);
                chart = co.Chart;

                var type = chartType.ToLower() switch
                {
                    "column" => Excel.XlChartType.xlColumnClustered,
                    "line" => Excel.XlChartType.xlLine,
                    "pie" => Excel.XlChartType.xlPie,
                    "bar" => Excel.XlChartType.xlBarClustered,
                    _ => Excel.XlChartType.xlColumnClustered
                };

                chart.SetSourceData(source);
                chart.ChartType = type;
                chart.HasTitle = true;
                chart.ChartTitle.Text = title;

                return co.Name;
            }
            finally
            {
                SafeReleaseComObject(chart);
                SafeReleaseComObject(co);
                SafeReleaseComObject(charts);
                SafeReleaseComObject(source);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void UpdateChart(string workbookName, string sheetName, string chartName, string? rangeAddress, string? chartType, string? title)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.ChartObjects? charts = null;
            Excel.ChartObject? co = null;
            Excel.Chart? chart = null;
            Excel.Range? source = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                charts = (Excel.ChartObjects)ws.ChartObjects();
                co = (Excel.ChartObject)charts.Item(chartName);
                chart = co.Chart;

                if (!string.IsNullOrEmpty(rangeAddress))
                {
                    source = ws.Range[rangeAddress];
                    chart.SetSourceData(source);
                }

                if (!string.IsNullOrEmpty(chartType))
                {
                    chart.ChartType = chartType.ToLower() switch
                    {
                        "column" => Excel.XlChartType.xlColumnClustered,
                        "line" => Excel.XlChartType.xlLine,
                        "pie" => Excel.XlChartType.xlPie,
                        "bar" => Excel.XlChartType.xlBarClustered,
                        _ => chart.ChartType
                    };
                }

                if (title != null)
                {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = title;
                }
            }
            finally
            {
                SafeReleaseComObject(source);
                SafeReleaseComObject(chart);
                SafeReleaseComObject(co);
                SafeReleaseComObject(charts);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void DeleteChart(string workbookName, string sheetName, string chartName)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.ChartObjects? charts = null;
            Excel.ChartObject? co = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                charts = (Excel.ChartObjects)ws.ChartObjects();
                co = (Excel.ChartObject)charts.Item(chartName);
                co.Delete();
            }
            finally
            {
                SafeReleaseComObject(co);
                SafeReleaseComObject(charts);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }
}
