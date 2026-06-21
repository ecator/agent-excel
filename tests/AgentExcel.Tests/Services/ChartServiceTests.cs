using System;

using AgentExcel.Models;
using AgentExcel.Providers;
using AgentExcel.Services;
using AgentExcel.Utils;

using NUnit.Framework;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Tests.Services;

[TestFixture]
[Category("COM")]
public class ChartServiceTests : BaseTests
{
#pragma warning disable NUnit1032
    private static ExcelConnectionProvider? s_provider;
#pragma warning restore NUnit1032
    private ChartService? _service;
    private string _wbName = "";
    private const string SheetName = "Sheet1";

    [OneTimeSetUp]
    public void OneTimeSetup()
    {
        s_provider = new ExcelConnectionProvider();
        // Ensure Excel is running once for the entire fixture
        s_provider.GetApp(createNew: true);
    }

    [OneTimeTearDown]
    public void OneTimeTeardown()
    {
        if (s_provider != null)
        {
            var app = s_provider.GetApp(createNew: false);
            if (app != null)
            {
                Excel.Workbooks? wbs = null;
                try
                {
                    wbs = app.Workbooks;
                    foreach (Excel.Workbook wb in wbs)
                    {
                        wb.Close(false);
                        ExcelConnector.SafeReleaseComObject(wb);
                    }
                }
                catch
                {
                    // Ignore
                }
                finally
                {
                    ExcelConnector.SafeReleaseComObject(wbs);
                }

                ExcelConnector.SafeReleaseComObject(app);
            }
            s_provider = null;
            ExcelConnector.ForceGarbageCollection();
        }
    }

    [SetUp]
    public void Setup()
    {
        _service = new ChartService(s_provider!);

        // Create a new workbook for each test to ensure test isolation
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks? wbs = null;
        Excel.Workbook? wb = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? ws = null;
        Excel.Range? r = null;
        try
        {
            wbs = app.Workbooks;
            wb = wbs.Add(Type.Missing);
            _wbName = wb.Name;

            sheets = wb.Sheets;
            ws = (Excel.Worksheet)sheets[1];
            
            // Populate some test data for chart
            r = ws.Range["A1:B4"];
            r.Value2 = new object[,]
            {
                { "Label", "Value" },
                { "A", 10.0 },
                { "B", 20.0 },
                { "C", 30.0 }
            };
        }
        finally
        {
            ExcelConnector.SafeReleaseComObject(r);
            ExcelConnector.SafeReleaseComObject(ws);
            ExcelConnector.SafeReleaseComObject(sheets);
            ExcelConnector.SafeReleaseComObject(wb);
            ExcelConnector.SafeReleaseComObject(wbs);
        }
    }

    [TearDown]
    public void Teardown()
    {
        if (s_provider != null && !string.IsNullOrEmpty(_wbName))
        {
            var app = s_provider.GetApp(createNew: false);
            if (app != null)
            {
                Excel.Workbooks? wbs = null;
                Excel.Workbook? wb = null;
                try
                {
                    wbs = app.Workbooks;
                    wb = wbs[_wbName];
                    wb.Close(false);
                }
                catch
                {
                    // Ignore
                }
                finally
                {
                    ExcelConnector.SafeReleaseComObject(wb);
                    ExcelConnector.SafeReleaseComObject(wbs);
                }
            }
        }
    }

    [Test]
    public void AddChart_WithValidInputs_CreatesChartAndReturnsName()
    {
        // Act
        ChartInfo chartInfo = _service!.AddChart(_wbName, SheetName, "A1:B4", "column", "Test Chart Title");

        // Assert
        Assert.That(chartInfo, Is.Not.Null);
        Assert.That(chartInfo.Name, Is.Not.Null.And.Not.Empty);
        Assert.That(chartInfo.Range, Is.EqualTo("A1:B4"));
        Assert.That(chartInfo.Type, Is.EqualTo("column"));
        Assert.That(chartInfo.Title, Is.EqualTo("Test Chart Title"));

        var list = _service.ListCharts(_wbName, SheetName);
        Assert.That(list, Has.Count.EqualTo(1));
        
        var listChartInfo = list[0];
        Assert.That(listChartInfo.Name, Is.EqualTo(chartInfo.Name));
        Assert.That(listChartInfo.Range, Is.EqualTo("A1:B4"));
        Assert.That(listChartInfo.Type, Is.EqualTo("column"));
        Assert.That(listChartInfo.Title, Is.EqualTo("Test Chart Title"));
    }

    [Test]
    public void UpdateChart_WithNewProperties_UpdatesChartSuccessfully()
    {
        // Arrange
        ChartInfo initialChart = _service!.AddChart(_wbName, SheetName, "A1:B4", "column", "Test Chart Title");

        // Act
        ChartInfo updatedChart = _service.UpdateChart(_wbName, SheetName, initialChart.Name, "A1:B3", "line", "Updated Chart Title");

        // Assert
        Assert.That(updatedChart, Is.Not.Null);
        Assert.That(updatedChart.Name, Is.EqualTo(initialChart.Name));
        Assert.That(updatedChart.Range, Is.EqualTo("A1:B3"));
        Assert.That(updatedChart.Type, Is.EqualTo("line"));
        Assert.That(updatedChart.Title, Is.EqualTo("Updated Chart Title"));

        var list = _service.ListCharts(_wbName, SheetName);
        Assert.That(list, Has.Count.EqualTo(1));

        var chartInfo = list[0];
        Assert.That(chartInfo.Name, Is.EqualTo(initialChart.Name));
        Assert.That(chartInfo.Range, Is.EqualTo("A1:B3"));
        Assert.That(chartInfo.Type, Is.EqualTo("line"));
        Assert.That(chartInfo.Title, Is.EqualTo("Updated Chart Title"));
    }

    [Test]
    public void DeleteChart_WhenChartExists_DeletesChartSuccessfully()
    {
        // Arrange
        ChartInfo initialChart = _service!.AddChart(_wbName, SheetName, "A1:B4", "column", "Test Chart Title");

        // Act
        _service.DeleteChart(_wbName, SheetName, initialChart.Name);

        // Assert
        var list = _service!.ListCharts(_wbName, SheetName);
        Assert.That(list, Is.Empty);
    }

    [Test]
    public void GetChart_WhenChartDoesNotExist_ThrowsException()
    {
        // Act & Assert
        Assert.That(() => _service!.UpdateChart(_wbName, SheetName, "NonExistentChart", "A1:B4", "column", "Title"),
            Throws.TypeOf<Exception>().And.Message.Contains("not found"));
    }
}
