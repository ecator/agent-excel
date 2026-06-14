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
public class FindReplaceTests : BaseTests
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

        // Create a temporary workbook with multiple sheets
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        _wb = wbs.Add(Type.Missing);
        _wbName = _wb.Name;

        // Ensure we have at least 2 sheets
        Excel.Sheets? sheets = null;
        Excel.Worksheet? newWs = null;
        try
        {
            sheets = _wb.Sheets;
            if (sheets.Count < 2)
            {
                newWs = (Excel.Worksheet)sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
        }
        finally
        {
            ExcelConnector.SafeReleaseComObject(newWs);
            ExcelConnector.SafeReleaseComObject(sheets);
        }

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
    public void Find_WithSpecificSheet_ReturnsCorrectMatches()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "target_value");
        _service.WriteRange(_wbName, "Sheet1", "B2", "other_value");
        _service.WriteRange(_wbName, "Sheet2", "A1", "target_value"); // Same value in Sheet2

        // Act
        var results = _service.Find(_wbName, "Sheet1", null, "target_value", true, true);

        // Assert
        Assert.That(results, Is.Not.Null);
        Assert.That(results.Count, Is.EqualTo(1));
        Assert.That(results[0].Sheet, Is.EqualTo("Sheet1"));
        Assert.That(results[0].Address, Is.EqualTo("$A$1"));
        Assert.That(results[0].Value, Is.EqualTo("target_value"));
    }

    [Test]
    public void Find_WorkbookWide_ReturnsMatchesFromAllSheets()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "hello_world");
        _service.WriteRange(_wbName, "Sheet2", "B2", "hello_world");

        // Act - Pass null for sheet name to search workbook-wide
        var results = _service.Find(_wbName, null, null, "hello_world", true, true);

        // Assert
        Assert.That(results, Is.Not.Null);
        Assert.That(results.Count, Is.EqualTo(2));
        
        var sheet1Match = results.FirstOrDefault(r => r.Sheet == "Sheet1");
        Assert.That(sheet1Match, Is.Not.Null);
        Assert.That(sheet1Match!.Address, Is.EqualTo("$A$1"));
        Assert.That(sheet1Match.Value, Is.EqualTo("hello_world"));

        var sheet2Match = results.FirstOrDefault(r => r.Sheet == "Sheet2");
        Assert.That(sheet2Match, Is.Not.Null);
        Assert.That(sheet2Match!.Address, Is.EqualTo("$B$2"));
        Assert.That(sheet2Match.Value, Is.EqualTo("hello_world"));
    }

    [Test]
    public void Replace_WithSpecificSheet_ReplacesAndReturnsUpdatedCount()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "test_item");
        _service.WriteRange(_wbName, "Sheet1", "A2", "prefix_test_item_suffix");
        _service.WriteRange(_wbName, "Sheet2", "A1", "test_item");

        // Act - Replace only in Sheet1
        var count = _service.Replace(_wbName, "Sheet1", null, "test_item", "new_val", true, false);

        // Assert
        Assert.That(count, Is.EqualTo(2));

        // Verify Excel actually has the replaced values
        var readSheet1 = _service.ReadRange(_wbName, "Sheet1", "A1:A2");
        Assert.That(readSheet1["A1"]?.ToString(), Is.EqualTo("new_val"));
        Assert.That(readSheet1["A2"]?.ToString(), Is.EqualTo("prefix_new_val_suffix"));

        // Verify Sheet2 is unmodified
        var readSheet2 = _service.ReadRange(_wbName, "Sheet2", "A1");
        Assert.That(readSheet2["A1"]?.ToString(), Is.EqualTo("test_item"));
    }

    [Test]
    public void Replace_WorkbookWide_ReplacesAndReturnsUpdatedCount()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "old_text");
        _service.WriteRange(_wbName, "Sheet2", "B2", "old_text");

        // Act - Replace workbook-wide
        var count = _service.Replace(_wbName, null, null, "old_text", "replaced_text", true, true);

        // Assert
        Assert.That(count, Is.EqualTo(2));

        // Verify Excel is updated
        var readSheet1 = _service.ReadRange(_wbName, "Sheet1", "A1");
        Assert.That(readSheet1["A1"]?.ToString(), Is.EqualTo("replaced_text"));

        var readSheet2 = _service.ReadRange(_wbName, "Sheet2", "B2");
        Assert.That(readSheet2["B2"]?.ToString(), Is.EqualTo("replaced_text"));
    }
}
