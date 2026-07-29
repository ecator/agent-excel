
using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class ExportService : ExcelServiceBase
{
    public ExportService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public string ExportRangeAsImage(string workbookName, string sheetName, string? rangeAddress, string outputFile)
    {
        ValidateOutputFile(outputFile, ".png");
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Range? xlRange = null;
            try
            {
                wb = GetWorkbook(workbookName);

                var app = GetApp();
                if (app != null)
                {
                    app.ScreenUpdating = true;
                }

                ws = GetWorksheet(wb, sheetName);

                // Activate workbook and worksheet to ensure they have UI focus
                wb.Activate();
                ws.Activate();

                xlRange = string.IsNullOrEmpty(rangeAddress) ? ws.UsedRange : GetRange(ws, rangeAddress);
                string exportedAddress = xlRange.get_Address(false, false);
                xlRange.Select();
                xlRange.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                // Short delay to allow clipboard to populate
                Thread.Sleep(100);

                Excel.ChartObjects? charts = (Excel.ChartObjects)ws.ChartObjects();
                Excel.ChartObject? co = charts.Add(0, 0, (double)xlRange.Width, (double)xlRange.Height);
                Excel.Chart? chart = co.Chart;
                try
                {
                    co.Activate();
                    chart.Paste();
                    chart.Export(Path.GetFullPath(outputFile), "PNG");
                }
                finally
                {
                    co.Delete();
                    SafeReleaseComObject(chart);
                    SafeReleaseComObject(co);
                    SafeReleaseComObject(charts);
                }

                return exportedAddress;
            }
            finally
            {
                SafeReleaseComObject(xlRange);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void ExportShapeAsImage(string workbookName, string sheetName, string shapeName, string outputFile)
    {
        ValidateOutputFile(outputFile, ".png");
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                wb = GetWorkbook(workbookName);

                var app = GetApp();
                if (app != null)
                {
                    app.ScreenUpdating = true;
                }

                ws = GetWorksheet(wb, sheetName);

                // Activate workbook and worksheet to ensure they have UI focus
                wb.Activate();
                ws.Activate();

                shapes = ws.Shapes;
                try
                {
                    shape = shapes.Item(shapeName);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Shape '{shapeName}' not found in worksheet '{ws.Name}'.", ex);
                }

                shape.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                // Short delay to allow clipboard to populate
                Thread.Sleep(100);

                Excel.ChartObjects? charts = (Excel.ChartObjects)ws.ChartObjects();
                Excel.ChartObject? co = charts.Add(0, 0, (double)shape.Width, (double)shape.Height);
                Excel.Chart? chart = co.Chart;
                try
                {
                    co.Activate();
                    chart.Paste();
                    chart.Export(Path.GetFullPath(outputFile), "PNG");
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
                SafeReleaseComObject(shape);
                SafeReleaseComObject(shapes);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    public void ExportAsPdf(string workbookName, string outputFile)
    {
        ValidateOutputFile(outputFile, ".pdf");
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            try
            {
                wb = GetWorkbook(workbookName);
                wb.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, Path.GetFullPath(outputFile));
            }
            finally
            {
                SafeReleaseComObject(wb);
            }
        });
    }

    private void ValidateOutputFile(string outputFile, string expectedExtension)
    {
        if (string.IsNullOrWhiteSpace(outputFile))
        {
            throw new ArgumentException("File path cannot be null or empty.", nameof(outputFile));
        }

        string extension = Path.GetExtension(outputFile);
        if (string.IsNullOrEmpty(extension))
        {
            throw new ArgumentException("File path must have a valid extension.", nameof(outputFile));
        }

        string extLower = extension.ToLowerInvariant();
        if (extLower != expectedExtension.ToLowerInvariant())
        {
            throw new ArgumentException($"Invalid file extension '{extension}'. Only {expectedExtension} is allowed.", nameof(outputFile));
        }

        string? directory = Path.GetDirectoryName(Path.GetFullPath(outputFile));
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            throw new DirectoryNotFoundException($"The directory '{directory}' does not exist.");
        }
    }
}
