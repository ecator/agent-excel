using AgentExcel.Providers;
using AgentExcel.Services;

using NUnit.Framework;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Tests.Services;

[TestFixture]
[Category("Unit")]
public class ExcelServiceBaseTests : BaseTests
{
    private FakeExcelConnectionProvider? _provider;
    private TestExcelService? _service;

    [SetUp]
    public void Setup()
    {
        _provider = new FakeExcelConnectionProvider();
        _service = new TestExcelService(_provider);
    }

    [TearDown]
    public void Teardown()
    {
        _provider?.Dispose();
        _provider = null;
    }

    [Test]
    public void GetWorkbook_WithNullOrEmptyName_ThrowsArgumentException()
    {
        // Act & Assert
        Assert.That(() => _service!.PublicGetWorkbook(null!), Throws.ArgumentException);
        Assert.That(() => _service!.PublicGetWorkbook(string.Empty), Throws.ArgumentException);
    }

    [Test]
    public void GetWorksheet_WithNullOrEmptyName_ThrowsArgumentException()
    {
        // Act & Assert
        Assert.That(() => _service!.PublicGetWorksheet(null!, null!), Throws.ArgumentException);
        Assert.That(() => _service!.PublicGetWorksheet(null!, string.Empty), Throws.ArgumentException);
    }

    [Test]
    public void GetActiveWorkbook_WhenExcelNotRunning_ThrowsException()
    {
        // Act & Assert
        Assert.That(() => _service!.PublicGetActiveWorkbook(), Throws.TypeOf<Exception>().And.Message.Contains("Excel is not running"));
    }

    [Test]
    public void GetRange_WithNullOrEmptyAddress_ThrowsArgumentException()
    {
        // Act & Assert
        Assert.That(() => _service!.PublicGetRange(null!, null!), Throws.ArgumentException);
        Assert.That(() => _service!.PublicGetRange(null!, string.Empty), Throws.ArgumentException);
        Assert.That(() => _service!.PublicGetRange(null!, "   "), Throws.ArgumentException);
    }

    private class FakeExcelConnectionProvider : IExcelConnectionProvider
    {
        public Excel.Application? GetApp(bool createNew = false) => null;
        public void SafeReleaseComObject(object? obj) { }
        public void Dispose() { }
    }

    private class TestExcelService : ExcelServiceBase
    {
        public TestExcelService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
        {
        }

        public Excel.Workbook PublicGetWorkbook(string workbookName, bool createNew = false)
        {
            return GetWorkbook(workbookName, createNew);
        }

        public Excel.Worksheet PublicGetWorksheet(Excel.Workbook workbook, string sheetName)
        {
            return GetWorksheet(workbook, sheetName);
        }

        public Excel.Workbook PublicGetActiveWorkbook()
        {
            return GetActiveWorkbook();
        }

        public Excel.Range PublicGetRange(Excel.Worksheet ws, string rangeAddress)
        {
            return GetRange(ws, rangeAddress);
        }
    }
}

