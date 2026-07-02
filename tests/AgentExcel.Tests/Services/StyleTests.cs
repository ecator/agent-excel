using System;

using AgentExcel.Models;
using AgentExcel.Models.Requests;
using AgentExcel.Providers;
using AgentExcel.Services;
using AgentExcel.Utils;

using NUnit.Framework;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Tests.Services;

[TestFixture]
[Category("COM")]
public class StyleTests : BaseTests
{
#pragma warning disable NUnit1032
    private static ExcelConnectionProvider? s_provider;
#pragma warning restore NUnit1032
    private RangeService? _service;
    private Excel.Workbook? _wb;
    private string? _wbName;

    [OneTimeSetUp]
    public void OneTimeSetup()
    {
        s_provider = new ExcelConnectionProvider();
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
        _service = new RangeService(s_provider!);

        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        _wb = wbs.Add(Type.Missing);
        _wbName = _wb.Name;

        ExcelConnector.SafeReleaseComObject(wbs);
    }

    [TearDown]
    public void Teardown()
    {
        if (_wb != null)
        {
            try
            {
                _wb.Close(false);
            }
            catch
            {
                // Ignore
            }
            ExcelConnector.SafeReleaseComObject(_wb);
            _wb = null;
        }
    }

    [Test]
    public void SetAndGetStyle_WithValidStyle_AppliesAndRetrievesStyle()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var expectedStyle = new CellStyle(
            FontName: "Courier New",
            FontSize: 14,
            Bold: true,
            Italic: true,
            Color: "#FF0000",
            BackgroundColor: "#FFFF00",
            HorizontalAlignment: "Center",
            VerticalAlignment: "Top"
        );

        // Act
        _service!.SetStyle(_wbName, "Sheet1", "A1", expectedStyle);
        var actualStyle = _service.GetStyle(_wbName, "Sheet1", "A1");

        // Assert
        Assert.That(actualStyle, Is.Not.Null);
        Assert.That(actualStyle.FontName, Is.EqualTo("Courier New"));
        Assert.That(actualStyle.FontSize, Is.EqualTo(14));
        Assert.That(actualStyle.Bold, Is.True);
        Assert.That(actualStyle.Italic, Is.True);
        Assert.That(actualStyle.Color, Is.EqualTo("#FF0000"));
        Assert.That(actualStyle.BackgroundColor, Is.EqualTo("#FFFF00"));
        Assert.That(actualStyle.HorizontalAlignment, Is.EqualTo("Center"));
        Assert.That(actualStyle.VerticalAlignment, Is.EqualTo("Top"));
    }

    [Test]
    public void GetStyle_OnMultipleCellsRange_ReturnsStyleOfTopLeftCellOnly()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var styleA1 = new CellStyle(
            FontName: "Arial",
            FontSize: 12,
            Bold: true,
            Italic: false,
            Color: "#0000FF",
            BackgroundColor: "#00FF00",
            HorizontalAlignment: "Left",
            VerticalAlignment: "Bottom"
        );
        var styleB2 = new CellStyle(
            FontName: "Times New Roman",
            FontSize: 16,
            Bold: false,
            Italic: true,
            Color: "#FF0000",
            BackgroundColor: "#FF00FF",
            HorizontalAlignment: "Right",
            VerticalAlignment: "Center"
        );

        _service!.SetStyle(_wbName, "Sheet1", "A1", styleA1);
        _service.SetStyle(_wbName, "Sheet1", "B2", styleB2);

        // Act
        var rangeStyle = _service.GetStyle(_wbName, "Sheet1", "A1:B2");

        // Assert
        Assert.That(rangeStyle, Is.Not.Null);
        Assert.That(rangeStyle.FontName, Is.EqualTo("Arial"));
        Assert.That(rangeStyle.FontSize, Is.EqualTo(12));
        Assert.That(rangeStyle.Bold, Is.True);
        Assert.That(rangeStyle.Italic, Is.False);
        Assert.That(rangeStyle.Color, Is.EqualTo("#0000FF"));
        Assert.That(rangeStyle.BackgroundColor, Is.EqualTo("#00FF00"));
        Assert.That(rangeStyle.HorizontalAlignment, Is.EqualTo("Left"));
        Assert.That(rangeStyle.VerticalAlignment, Is.EqualTo("Bottom"));
    }

    [Test]
    public void GetStyle_WhenRangeIsEmpty_ReturnsStyleOfFirstCellInUsedRange()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        // Write value and set style to B2 to populate UsedRange
        _service!.WriteRange(_wbName, "Sheet1", "B2", "Test");
        var styleB2 = new CellStyle(
            FontName: "Calibri",
            FontSize: 11,
            Bold: true,
            Italic: true,
            Color: "#000000",
            BackgroundColor: "#FFFFFF",
            HorizontalAlignment: "Left",
            VerticalAlignment: "Bottom"
        );
        _service.SetStyle(_wbName, "Sheet1", "B2", styleB2);

        // Act
        var rangeStyle = _service.GetStyle(_wbName, "Sheet1", "");

        // Assert
        Assert.That(rangeStyle, Is.Not.Null);
        Assert.That(rangeStyle.Bold, Is.True);
        Assert.That(rangeStyle.Italic, Is.True);
    }
}
