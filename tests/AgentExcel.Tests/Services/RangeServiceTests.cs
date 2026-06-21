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
public class RangeServiceTests : BaseTests
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
    public void WriteRange_WhenRangeIsInvalid_ThrowsFriendlyException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act & Assert
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            _service!.WriteRange(_wbName, "Sheet1", "InvalidRangeAddress!!!", "Value");
        });

        Assert.That(ex!.Message, Contains.Substring("Invalid Excel range address"));
    }

    [Test]
    public void WriteRange_WithJsonElement2DArray_WritesSuccessfully()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var json = "[[\"Col1\", \"Col2\"], [\"Val1\", 123.45]]";
        using var doc = System.Text.Json.JsonDocument.Parse(json);
        var jsonElement = doc.RootElement;

        // Act
        _service!.WriteRange(_wbName, "Sheet1", "A1:B2", jsonElement);

        // Assert
        var results = _service.ReadRange(_wbName, "Sheet1", "A1:B2");
        Assert.That(results["A1"]?.ToString(), Is.EqualTo("Col1"));
        Assert.That(results["B1"]?.ToString(), Is.EqualTo("Col2"));
        Assert.That(results["A2"]?.ToString(), Is.EqualTo("Val1"));
        Assert.That(Convert.ToDouble(results["B2"]), Is.EqualTo(123.45).Within(0.001));
    }

    [Test]
    public void WriteRange_WithJaggedArray_WritesSuccessfully()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var jagged = new object[][]
        {
            new object[] { "Jagged1", "Jagged2" },
            new object[] { 456, "Jagged3" }
        };

        // Act
        _service!.WriteRange(_wbName, "Sheet1", "A1:B2", jagged);

        // Assert
        var results = _service.ReadRange(_wbName, "Sheet1", "A1:B2");
        Assert.That(results["A1"]?.ToString(), Is.EqualTo("Jagged1"));
        Assert.That(results["B1"]?.ToString(), Is.EqualTo("Jagged2"));
        Assert.That(Convert.ToInt32(results["A2"]), Is.EqualTo(456));
        Assert.That(results["B2"]?.ToString(), Is.EqualTo("Jagged3"));
    }

    [Test]
    public void WriteRange_WithMismatchedScalarSize_ThrowsArgumentException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act & Assert
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            _service!.WriteRange(_wbName, "Sheet1", "A1:B2", "SingleValue");
        });

        Assert.That(ex!.Message, Contains.Substring("does not match the size of the data to be written"));
    }

    [Test]
    public void WriteRange_WithMismatchedMatrixSize_ThrowsArgumentException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var jagged = new object[][]
        {
            new object[] { "Val1", "Val2", "Val3" },
            new object[] { "Val4", "Val5", "Val6" }
        };

        // Act & Assert
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            _service!.WriteRange(_wbName, "Sheet1", "A1:B2", jagged);
        });

        Assert.That(ex!.Message, Contains.Substring("does not match the size of the data to be written"));
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

    [Test]
    [Category("COM")]
    public void RenameTable_WithValidNewName_RenamesSuccessfully()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "H1");
        _service.WriteRange(_wbName, "Sheet1", "B1", "H2");
        _service.ConvertToTable(_wbName, "Sheet1", "A1:B1", "TableToRename", true);

        // Act
        _service.RenameTable(_wbName, "Sheet1", "TableToRename", "NewTableName");

        // Assert
        var tables = _service.ListTables(_wbName, "Sheet1");
        Assert.That(tables, Is.Not.Null);
        Assert.That(tables["Sheet1"], Contains.Item("NewTableName"));
        Assert.That(tables["Sheet1"], Does.Not.Contain("TableToRename"));
    }

    [Test]
    [Category("COM")]
    public void RenameTable_WhenTableDoesNotExist_ThrowsException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act & Assert
        Assert.Throws<Exception>(() =>
        {
            _service!.RenameTable(_wbName, "Sheet1", "NonExistentTable", "NewName");
        });
    }

    [Test]
    [Category("COM")]
    public void RenameTable_WithDuplicateName_ThrowsException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "H1");
        _service.WriteRange(_wbName, "Sheet1", "B1", "H2");
        _service.ConvertToTable(_wbName, "Sheet1", "A1:B1", "Table1", true);

        _service.WriteRange(_wbName, "Sheet1", "D1", "H3");
        _service.WriteRange(_wbName, "Sheet1", "E1", "H4");
        _service.ConvertToTable(_wbName, "Sheet1", "D1:E1", "Table2", true);

        // Act & Assert
        Assert.Throws<ArgumentException>(() =>
        {
            _service.RenameTable(_wbName, "Sheet1", "Table1", "Table2");
        });
    }

    [Test]
    [Category("COM")]
    public void RenameTable_WithCaseOnlyChange_RenamesSuccessfully()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "H1");
        _service.WriteRange(_wbName, "Sheet1", "B1", "H2");
        _service.ConvertToTable(_wbName, "Sheet1", "A1:B1", "Table1", true);

        // Act
        _service.RenameTable(_wbName, "Sheet1", "Table1", "table1");

        // Assert
        var tables = _service.ListTables(_wbName, "Sheet1");
        Assert.That(tables, Is.Not.Null);
        Assert.That(tables["Sheet1"], Contains.Item("table1"));
    }

    [Test]
    [Category("COM")]
    public void ConvertToTable_WithValidInputs_ConvertsAndReturnsTableName()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "Col1");
        _service.WriteRange(_wbName, "Sheet1", "B1", "Col2");
        _service.WriteRange(_wbName, "Sheet1", "A2", "Val1");
        _service.WriteRange(_wbName, "Sheet1", "B2", "Val2");

        // Act
        var tableName = _service.ConvertToTable(_wbName, "Sheet1", "A1:B2", "MyTable", true);

        // Assert
        Assert.That(tableName, Is.EqualTo("MyTable"));
        var tables = _service.ListTables(_wbName, "Sheet1");
        Assert.That(tables, Is.Not.Null);
        Assert.That(tables["Sheet1"], Contains.Item("MyTable"));
    }

    [Test]
    [Category("COM")]
    public void ConvertToTable_WithDuplicateName_ThrowsArgumentException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "Col1");
        _service.WriteRange(_wbName, "Sheet1", "B1", "Col2");
        _service.ConvertToTable(_wbName, "Sheet1", "A1:B1", "DuplicateTable", true);

        _service.WriteRange(_wbName, "Sheet1", "D1", "Col3");
        _service.WriteRange(_wbName, "Sheet1", "E1", "Col4");

        // Act & Assert
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            _service.ConvertToTable(_wbName, "Sheet1", "D1:E1", "DuplicateTable", true);
        });

        Assert.That(ex!.Message, Contains.Substring("already in use by another table"));
    }

    [Test]
    [Category("COM")]
    public void ConvertToRange_WithValidTable_ConvertsAndReturnsAddress()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "Col1");
        _service.WriteRange(_wbName, "Sheet1", "B1", "Col2");
        _service.ConvertToTable(_wbName, "Sheet1", "A1:B1", "TableToRange", true);

        // Act
        var address = _service.ConvertToRange(_wbName, "Sheet1", "TableToRange");

        // Assert
        Assert.That(address, Is.EqualTo("$A$1:$B$2"));
        var tables = _service.ListTables(_wbName, "Sheet1");
        Assert.That(tables.ContainsKey("Sheet1"), Is.False);
    }

    [Test]
    [Category("COM")]
    public void ReadFormula_WhenRangeIsNotEmpty_ReadsRequestedCellFormulas()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", 10);
        _service.WriteRange(_wbName, "Sheet1", "A2", 20);
        _service.WriteFormula(_wbName, "Sheet1", "A3", "=SUM(A1:A2)");

        // Act
        var results = _service.ReadFormula(_wbName, "Sheet1", "A1:A3");

        // Assert
        Assert.That(results, Is.Not.Null);
        Assert.That(results.Count, Is.EqualTo(1));
        Assert.That(results.ContainsKey("A3"), Is.True);
        Assert.That(results["A3"], Is.EqualTo("=SUM(A1:A2)"));
        Assert.That(results.ContainsKey("A1"), Is.False);
    }

    [Test]
    [Category("COM")]
    public void ReadFormula_WhenRangeIsEmpty_ReadsUsedRangeFormulas()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "B1", 5);
        _service.WriteFormula(_wbName, "Sheet1", "B2", "=B1*2");

        // Act
        var results = _service.ReadFormula(_wbName, "Sheet1", "");

        // Assert
        Assert.That(results, Is.Not.Null);
        Assert.That(results.Count, Is.EqualTo(1));
        Assert.That(results.ContainsKey("B2"), Is.True);
        Assert.That(results["B2"], Is.EqualTo("=B1*2"));
    }

    [Test]
    [Category("COM")]
    public void ReadFormula_WhenCellsHaveNoFormula_OmitsThem()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "Plain Value");

        // Act
        var results = _service.ReadFormula(_wbName, "Sheet1", "A1");

        // Assert
        Assert.That(results, Is.Not.Null);
        Assert.That(results, Is.Empty);
    }

    [Test]
    [Category("COM")]
    public void ReadFormula_WhenNoCellsHaveFormula_ReturnsEmptyDictionary()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act
        var results = _service!.ReadFormula(_wbName, "Sheet1", "A1:B2");

        // Assert
        Assert.That(results, Is.Not.Null);
        Assert.That(results, Is.Empty);
    }

    [Test]
    [Category("COM")]
    public void WriteFormula_WithValidSingleFormula_WritesSuccessfully()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act
        _service!.WriteFormula(_wbName, "Sheet1", "A1", "=TODAY()");

        // Assert
        var results = _service.ReadFormula(_wbName, "Sheet1", "A1");
        Assert.That(results.ContainsKey("A1"), Is.True);
        Assert.That(results["A1"], Is.EqualTo("=TODAY()"));
    }

    [Test]
    [Category("COM")]
    public void WriteFormula_WithValidArrayOfFormulas_WritesSuccessfully()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var formulas = new object[,]
        {
            { "=SUM(B1:B2)", "=AVERAGE(C1:C2)" }
        };

        // Act
        _service!.WriteFormula(_wbName, "Sheet1", "A1:B1", formulas);

        // Assert
        var results = _service.ReadFormula(_wbName, "Sheet1", "A1:B1");
        Assert.That(results.Count, Is.EqualTo(2));
        Assert.That(results["A1"], Is.EqualTo("=SUM(B1:B2)"));
        Assert.That(results["B1"], Is.EqualTo("=AVERAGE(C1:C2)"));
    }

    [Test]
    [Category("COM")]
    public void WriteFormula_WithFormulaNotStartingWithEqual_ThrowsArgumentException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act & Assert
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            _service!.WriteFormula(_wbName, "Sheet1", "A1", "NotAFormula");
        });

        Assert.That(ex!.Message, Contains.Substring("Formula must start with '='"));
    }

    [Test]
    [Category("COM")]
    public void WriteFormula_WithMismatchedMatrixSize_ThrowsArgumentException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var formulas = new object[,]
        {
            { "=A1", "=B1" }
        };

        // Act & Assert
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            _service!.WriteFormula(_wbName, "Sheet1", "A1:A2", formulas);
        });

        Assert.That(ex!.Message, Contains.Substring("does not match the size of the data to be written"));
    }

    [Test]
    [Category("COM")]
    public void Clear_WhenRangeIsEmpty_ClearsUsedRange()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "Val1");
        _service.WriteRange(_wbName, "Sheet1", "B2", "Val2");

        // Act
        var address = _service.Clear(_wbName, "Sheet1", null, "all");

        // Assert
        Assert.That(address, Is.EqualTo("$A$1:$B$2"));
        var results = _service.ReadRange(_wbName, "Sheet1", "A1:B2");
        Assert.That(results, Is.Empty);
    }

    [Test]
    [Category("COM")]
    public void Clear_WithFormats_OnlyClearsFormatting()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "Val1");
        _service.SetStyle(_wbName, "Sheet1", "A1", new CellStyle { Bold = true });

        // Act
        _service.Clear(_wbName, "Sheet1", "A1", "formats");

        // Assert
        var results = _service.ReadRange(_wbName, "Sheet1", "A1");
        Assert.That(results["A1"]?.ToString(), Is.EqualTo("Val1"));

        var style = _service.GetStyle(_wbName, "Sheet1", "A1");
        Assert.That(style.Bold, Is.False);
    }

    [Test]
    [Category("COM")]
    public void Clear_WithContents_OnlyClearsContents()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.WriteRange(_wbName, "Sheet1", "A1", "Val1");
        _service.SetStyle(_wbName, "Sheet1", "A1", new CellStyle { Bold = true });

        // Act
        _service.Clear(_wbName, "Sheet1", "A1", "contents");

        // Assert
        var results = _service.ReadRange(_wbName, "Sheet1", "A1");
        Assert.That(results, Is.Empty);

        var style = _service.GetStyle(_wbName, "Sheet1", "A1");
        Assert.That(style.Bold, Is.True);
    }

    [Test]
    [Category("COM")]
    public void Clear_WithInvalidType_ThrowsArgumentException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act & Assert
        Assert.Throws<ArgumentException>(() =>
        {
            _service!.Clear(_wbName, "Sheet1", "A1", "invalid_type");
        });
    }

    [Test]
    [Category("COM")]
    public void SetSelection_WithValidRange_SetsSelectionAndReturnsAddress()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act
        var address = _service!.SetSelection(_wbName, "Sheet1", "B2:C3");

        // Assert
        Assert.That(address, Is.EqualTo("$B$2:$C$3"));
    }

    [Test]
    [Category("COM")]
    public void SetSelection_SingleCell_SetsSelectionAndReturnsAddress()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act
        var address = _service!.SetSelection(_wbName, "Sheet1", "B2");

        // Assert
        Assert.That(address, Is.EqualTo("$B$2"));
    }

    [Test]
    [Category("COM")]
    public void GetSelection_WhenRangeIsSelected_ReturnsCorrectAddress()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        _service!.SetSelection(_wbName, "Sheet1", "A1:B2");

        // Act
        var address = _service.GetSelection(_wbName, "Sheet1");

        // Assert
        Assert.That(address, Is.EqualTo("$A$1:$B$2"));
    }

    [Test]
    [Category("COM")]
    public void SetListValidation_WithValidInputs_AppliesValidationToRange()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act
        _service!.SetListValidation(_wbName, "Sheet1", "A1", "\"Yes,No\"");

        // Assert
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks? wbs = null;
        Excel.Workbook? wb = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? ws = null;
        Excel.Range? range = null;
        Excel.Validation? val = null;
        try
        {
            wbs = app.Workbooks;
            wb = wbs[_wbName];
            sheets = wb.Sheets;
            ws = (Excel.Worksheet)sheets["Sheet1"];
            range = ws.Range["A1"];
            val = range.Validation;

            Assert.That(val.Type, Is.EqualTo((int)Excel.XlDVType.xlValidateList));
            Assert.That(val.Formula1, Is.EqualTo("\"Yes,No\""));
        }
        finally
        {
            ExcelConnector.SafeReleaseComObject(val);
            ExcelConnector.SafeReleaseComObject(range);
            ExcelConnector.SafeReleaseComObject(ws);
            ExcelConnector.SafeReleaseComObject(sheets);
            ExcelConnector.SafeReleaseComObject(wb);
            ExcelConnector.SafeReleaseComObject(wbs);
        }
    }

    [Test]
    [Category("COM")]
    public void SetListValidation_WithInvalidRangeAddress_ThrowsArgumentException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);

        // Act & Assert
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            _service!.SetListValidation(_wbName, "Sheet1", "InvalidRangeAddress!!!", "\"Yes,No\"");
        });

        Assert.That(ex!.Message, Contains.Substring("Invalid Excel range address"));
    }
}


