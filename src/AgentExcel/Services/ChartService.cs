using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

using AgentExcel.Models;
using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class ChartService : ExcelServiceBase
{
    public ChartService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public List<ChartInfo> ListCharts(string workbookName, string sheetName)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.ChartObjects? charts = null;
            var result = new List<ChartInfo>();
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                charts = (Excel.ChartObjects)ws.ChartObjects();
                foreach (Excel.ChartObject co in charts)
                {
                    Excel.Chart? chart = null;
                    try
                    {
                        chart = co.Chart;
                        string name = co.Name;
                        string range = GetChartSourceRangeAddress(chart, ws);
                        string type = GetChartTypeString(chart.ChartType);
                        string title = GetChartTitle(chart);
                        result.Add(new ChartInfo(name, range, type, title));
                    }
                    finally
                    {
                        SafeReleaseComObject(chart);
                        SafeReleaseComObject(co);
                    }
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

    public ChartInfo AddChart(string workbookName, string sheetName, string rangeAddress, string chartType, string title)
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
                source = GetRange(ws, rangeAddress);
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

                Excel.ChartTitle? chartTitle = null;
                try
                {
                    chartTitle = chart.ChartTitle;
                    chartTitle.Text = title;
                }
                finally
                {
                    SafeReleaseComObject(chartTitle);
                }

                string name = co.Name;
                string resolvedRange = GetChartSourceRangeAddress(chart, ws);
                string resolvedType = GetChartTypeString(chart.ChartType);
                string resolvedTitle = GetChartTitle(chart);

                return new ChartInfo(name, resolvedRange, resolvedType, resolvedTitle);
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

    public ChartInfo UpdateChart(string workbookName, string sheetName, string chartName, string? rangeAddress, string? chartType, string? title)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.ChartObject? co = null;
            Excel.Chart? chart = null;
            Excel.Range? source = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                co = GetChart(ws, chartName);
                chart = co.Chart;

                if (!string.IsNullOrEmpty(rangeAddress))
                {
                    source = GetRange(ws, rangeAddress);
                    chart.SetSourceData(source);
                }

                if (!string.IsNullOrEmpty(chartType))
                {
                    var type = chartType.ToLower() switch
                    {
                        "column" => Excel.XlChartType.xlColumnClustered,
                        "line" => Excel.XlChartType.xlLine,
                        "pie" => Excel.XlChartType.xlPie,
                        "bar" => Excel.XlChartType.xlBarClustered,
                        _ => chart.ChartType
                    };
                    chart.ChartType = type;
                }

                if (title != null)
                {
                    chart.HasTitle = true;
                    Excel.ChartTitle? chartTitle = null;
                    try
                    {
                        chartTitle = chart.ChartTitle;
                        chartTitle.Text = title;
                    }
                    finally
                    {
                        SafeReleaseComObject(chartTitle);
                    }
                }

                string name = co.Name;
                string resolvedRange = GetChartSourceRangeAddress(chart, ws);
                string resolvedType = GetChartTypeString(chart.ChartType);
                string resolvedTitle = GetChartTitle(chart);

                return new ChartInfo(name, resolvedRange, resolvedType, resolvedTitle);
            }
            finally
            {
                SafeReleaseComObject(source);
                SafeReleaseComObject(chart);
                SafeReleaseComObject(co);
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
            Excel.ChartObject? co = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                co = GetChart(ws, chartName);
                co.Delete();
            }
            finally
            {
                SafeReleaseComObject(co);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    private Excel.ChartObject GetChart(Excel.Worksheet ws, string chartName)
    {
        if (string.IsNullOrWhiteSpace(chartName))
        {
            throw new ArgumentException("Chart name cannot be null or empty.", nameof(chartName));
        }

        Excel.ChartObjects? charts = null;
        try
        {
            charts = (Excel.ChartObjects)ws.ChartObjects();
            try
            {
                return (Excel.ChartObject)charts.Item(chartName);
            }
            catch (Exception ex)
            {
                throw new Exception($"Chart '{chartName}' not found in worksheet '{ws.Name}'.", ex);
            }
        }
        finally
        {
            SafeReleaseComObject(charts);
        }
    }

    private string GetChartSourceRangeAddress(Excel.Chart chart, Excel.Worksheet ws)
    {
        Excel.SeriesCollection? seriesCollection = null;
        Excel.Series? series = null;
        Excel.Range? boundingRange = null;
        try
        {
            seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
            if (seriesCollection.Count == 0)
            {
                return "";
            }

            series = (Excel.Series)seriesCollection.Item(1);
            string formula = series.Formula;
            if (string.IsNullOrEmpty(formula))
            {
                return "";
            }

            // Find all range references (like Sheet1!$A$1:$B$4 or $A$1:$A$4)
            var matches = Regex.Matches(formula, @"(?:'[^']+'|[^',()!]+)!\$?[A-Za-z]+\$?[0-9]+(?::\$?[A-Za-z]+\$?[0-9]+)?");
            var ranges = new List<Excel.Range>();
            try
            {
                foreach (Match match in matches)
                {
                    string matchVal = match.Value;
                    int bangIndex = matchVal.IndexOf('!');
                    string address = bangIndex >= 0 ? matchVal.Substring(bangIndex + 1) : matchVal;
                    address = address.Replace("$", "");
                    try
                    {
                        var r = ws.Range[address];
                        ranges.Add(r);
                    }
                    catch
                    {
                        // Ignore invalid ranges
                    }
                }

                if (ranges.Count == 0)
                {
                    return "";
                }

                // Find the bounding box of all matched ranges
                int minRow = int.MaxValue;
                int minCol = int.MaxValue;
                int maxRow = int.MinValue;
                int maxCol = int.MinValue;

                foreach (var r in ranges)
                {
                    Excel.Range? rows = null;
                    Excel.Range? cols = null;
                    try
                    {
                        rows = r.Rows;
                        cols = r.Columns;
                        minRow = Math.Min(minRow, r.Row);
                        minCol = Math.Min(minCol, r.Column);
                        maxRow = Math.Max(maxRow, r.Row + rows.Count - 1);
                        maxCol = Math.Max(maxCol, r.Column + cols.Count - 1);
                    }
                    finally
                    {
                        SafeReleaseComObject(cols);
                        SafeReleaseComObject(rows);
                    }
                }

                Excel.Range? startCell = null;
                Excel.Range? endCell = null;
                try
                {
                    startCell = (Excel.Range)ws.Cells[minRow, minCol];
                    endCell = (Excel.Range)ws.Cells[maxRow, maxCol];
                    boundingRange = ws.Range[startCell, endCell];
                    return boundingRange.Address[false, false, Excel.XlReferenceStyle.xlA1, false];
                }
                finally
                {
                    SafeReleaseComObject(endCell);
                    SafeReleaseComObject(startCell);
                }
            }
            finally
            {
                foreach (var r in ranges)
                {
                    SafeReleaseComObject(r);
                }
            }
        }
        catch
        {
            return "";
        }
        finally
        {
            SafeReleaseComObject(boundingRange);
            SafeReleaseComObject(series);
            SafeReleaseComObject(seriesCollection);
        }
    }

    private string GetChartTypeString(Excel.XlChartType chartType)
    {
        return chartType switch
        {
            Excel.XlChartType.xlColumnClustered => "column",
            Excel.XlChartType.xlLine => "line",
            Excel.XlChartType.xlPie => "pie",
            Excel.XlChartType.xlBarClustered => "bar",
            _ => chartType.ToString()
        };
    }

    private string GetChartTitle(Excel.Chart chart)
    {
        if (chart.HasTitle)
        {
            Excel.ChartTitle? title = null;
            try
            {
                title = chart.ChartTitle;
                return title.Text;
            }
            finally
            {
                SafeReleaseComObject(title);
            }
        }
        return "";
    }
}
