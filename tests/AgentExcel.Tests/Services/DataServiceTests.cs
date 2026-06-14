using System;
using System.IO;
using System.Linq;

using AgentExcel.Models;
using AgentExcel.Providers;
using AgentExcel.Services;
using AgentExcel.Utils;

using NUnit.Framework;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Tests.Services;

[TestFixture]
[Category("COM")]
public class DataServiceTests : BaseTests
{
#pragma warning disable NUnit1032
    private static ExcelConnectionProvider? s_provider;
#pragma warning restore NUnit1032
    private DataService? _service;
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
        _service = new DataService(s_provider!);

        // Create a temporary workbook for testing
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
    public void ReadRange_WhenRangeIsNotEmpty_ReadsRequestedCells()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "Val1");
        _service.WriteRange(_wbName, "Sheet1", "B2", 123.45);

        // Act
        var results = _service.ReadRange(_wbName, "Sheet1", "A1:B2");

        // Assert
        Assert.That(results, Is.Not.Null);
        Assert.That(results.Count, Is.EqualTo(2));
        Assert.That(results.ContainsKey("A1"), Is.True);
        Assert.That(results["A1"]?.ToString(), Is.EqualTo("Val1"));
        Assert.That(results.ContainsKey("B2"), Is.True);
        Assert.That(Convert.ToDouble(results["B2"]), Is.EqualTo(123.45).Within(0.001));
    }

    [Test]
    public void ReadRange_WhenRangeIsEmpty_ReadsUsedRange()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "B2", "ValB2");
        _service.WriteRange(_wbName, "Sheet1", "C3", "ValC3");

        // Act
        var results = _service.ReadRange(_wbName, "Sheet1", "");

        // Assert
        Assert.That(results, Is.Not.Null);
        Assert.That(results.Count, Is.EqualTo(2));
        Assert.That(results.ContainsKey("B2"), Is.True);
        Assert.That(results["B2"]?.ToString(), Is.EqualTo("ValB2"));
        Assert.That(results.ContainsKey("C3"), Is.True);
        Assert.That(results["C3"]?.ToString(), Is.EqualTo("ValC3"));
        Assert.That(results.ContainsKey("A1"), Is.False);
    }

    [Test]
    public void ReadRange_WhenCellsAreEmpty_OmitsEmptyCells()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "ValA1");
        _service.WriteRange(_wbName, "Sheet1", "A3", "ValA3");

        // Act
        var results = _service.ReadRange(_wbName, "Sheet1", "A1:A3");

        // Assert
        Assert.That(results, Is.Not.Null);
        Assert.That(results.Count, Is.EqualTo(2));
        Assert.That(results.ContainsKey("A1"), Is.True);
        Assert.That(results["A1"]?.ToString(), Is.EqualTo("ValA1"));
        Assert.That(results.ContainsKey("A3"), Is.True);
        Assert.That(results["A3"]?.ToString(), Is.EqualTo("ValA3"));
        Assert.That(results.ContainsKey("A2"), Is.False);
    }

    [Test]
    public void ReadRange_WhenNoCellsHaveData_ReturnsEmptyDictionary()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act
        var results = _service!.ReadRange(_wbName, "Sheet1", "A1:B2");

        // Assert
        Assert.That(results, Is.Not.Null);
        Assert.That(results.Count, Is.EqualTo(0));
    }

    [Test]
    public void ReadRange_WhenCellHasFormattedValue_ReadsRawValue()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", 0.123);

        Excel.Sheets? sheets = null;
        Excel.Worksheet? ws = null;
        Excel.Range? cell = null;
        try
        {
            sheets = _wb!.Sheets;
            ws = (Excel.Worksheet)sheets[1];
            cell = ws.Range["A1"];
            cell.NumberFormat = "0.0%";
        }
        finally
        {
            ExcelConnector.SafeReleaseComObject(cell);
            ExcelConnector.SafeReleaseComObject(ws);
            ExcelConnector.SafeReleaseComObject(sheets);
        }

        // Act
        var results = _service.ReadRange(_wbName, "Sheet1", "A1");

        // Assert
        Assert.That(results, Is.Not.Null);
        Assert.That(results.Count, Is.EqualTo(1));
        Assert.That(results.ContainsKey("A1"), Is.True);
        Assert.That(Convert.ToDouble(results["A1"]), Is.EqualTo(0.123).Within(0.001));
    }

    [Test]
    public void ReadRange_WhenRangeIsInvalid_ThrowsFriendlyException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act & Assert
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            _service!.ReadRange(_wbName, "Sheet1", "InvalidRangeAddress!!!");
        });

        Assert.That(ex!.Message, Contains.Substring("Invalid Excel range address"));
    }

    [Test]
    [Category("COM")]
    public void ReadTableAsMarkdown_WhenTableExists_ReturnsMarkdownFormat()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "Header1");
        _service.WriteRange(_wbName, "Sheet1", "B1", "Header2");
        _service.WriteRange(_wbName, "Sheet1", "A2", "Val\\1|test");
        _service.WriteRange(_wbName, "Sheet1", "B2", "Val2\nLine2");

        _service.ConvertToTable(_wbName, "Sheet1", "A1:B2", "TestTable1", true);

        // Act
        string markdown = _service.ReadTableAsMarkdown(_wbName, "Sheet1", "TestTable1");

        // Assert
        string expected = "| Header1 | Header2 |\r\n| --- | --- |\r\n| Val\\\\1\\|test | Val2<br>Line2 |";
        Assert.That(markdown.Replace("\r\n", "\n"), Is.EqualTo(expected.Replace("\r\n", "\n")));
    }

    [Test]
    [Category("COM")]
    public void ListTables_WhenTablesExist_ReturnsDictionaryWithTables()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "H1");
        _service.WriteRange(_wbName, "Sheet1", "B1", "H2");
        _service.ConvertToTable(_wbName, "Sheet1", "A1:B1", "Table1", true);

        // Act
        var listSpecific = _service.ListTables(_wbName, "Sheet1");
        var listAll = _service.ListTables(_wbName, null);

        // Assert
        Assert.That(listSpecific, Is.Not.Null);
        Assert.That(listSpecific.Count, Is.EqualTo(1));
        Assert.That(listSpecific.ContainsKey("Sheet1"), Is.True);
        Assert.That(listSpecific["Sheet1"], Contains.Item("Table1"));

        Assert.That(listAll, Is.Not.Null);
        Assert.That(listAll.Count, Is.EqualTo(1));
        Assert.That(listAll.ContainsKey("Sheet1"), Is.True);
        Assert.That(listAll["Sheet1"], Contains.Item("Table1"));
    }

    [Test]
    [Category("COM")]
    public void ListTables_WhenNoTablesExist_ReturnsEmptyDictionary()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act
        var listSpecific = _service!.ListTables(_wbName, "Sheet1");
        var listAll = _service.ListTables(_wbName, null);

        // Assert
        Assert.That(listSpecific, Is.Not.Null);
        Assert.That(listSpecific, Is.Empty);
        Assert.That(listAll, Is.Not.Null);
        Assert.That(listAll, Is.Empty);
    }
}

