using System;
using System.IO;

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
public class ExportServiceTests : BaseTests
{
#pragma warning disable NUnit1032
    private static ExcelConnectionProvider? s_provider;
#pragma warning restore NUnit1032
    private ExportService? _service;
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
        _service = new ExportService(s_provider!);

        // Create a temporary workbook for testing
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        _wb = wbs.Add(Type.Missing);
        _wbName = _wb.Name;

        Excel.Sheets? sheets = _wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        Excel.Range cell = ws.Range["A1"];
        cell.Value2 = "Export Test Data";

        ExcelConnector.SafeReleaseComObject(cell);
        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
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
    public void ExportRangeAsImage_WithValidParameters_ExportsSuccessfully()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var outputFile = Path.Combine(TestDataTempPath, "range_export_test.png");
        if (File.Exists(outputFile))
        {
            File.Delete(outputFile);
        }

        // Act
        _service!.ExportRangeAsImage(_wbName, "Sheet1", "A1:B2", outputFile);

        // Assert
        Assert.That(File.Exists(outputFile), Is.True);
        File.Delete(outputFile);
    }

    [Test]
    public void ExportRangeAsImage_WithInvalidExtension_ThrowsArgumentException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var outputFile = Path.Combine(TestDataTempPath, "range_export_test.jpg");

        // Act & Assert
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            _service!.ExportRangeAsImage(_wbName, "Sheet1", "A1:B2", outputFile);
        });

        Assert.That(ex!.Message, Contains.Substring("Only .png is allowed"));
    }

    [Test]
    public void ExportRangeAsImage_WithNonExistentDirectory_ThrowsDirectoryNotFoundException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var nonExistentDir = Path.Combine(TestDataTempPath, "non_existent_folder_xyz");
        var outputFile = Path.Combine(nonExistentDir, "range_export_test.png");

        // Act & Assert
        Assert.Throws<DirectoryNotFoundException>(() =>
        {
            _service!.ExportRangeAsImage(_wbName, "Sheet1", "A1:B2", outputFile);
        });
    }

    [Test]
    public void ExportRangeAsImage_WithInvalidRangeAddress_ThrowsArgumentException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var outputFile = Path.Combine(TestDataTempPath, "range_invalid_test.png");

        // Act & Assert
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            _service!.ExportRangeAsImage(_wbName, "Sheet1", "InvalidRangeAddress!!!", outputFile);
        });

        Assert.That(ex!.Message, Contains.Substring("Invalid Excel range address"));
    }

    [Test]
    public void ExportAsPdf_WithValidParameters_ExportsSuccessfully()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var outputFile = Path.Combine(TestDataTempPath, "workbook_export_test.pdf");
        if (File.Exists(outputFile))
        {
            File.Delete(outputFile);
        }

        // Act
        _service!.ExportAsPdf(_wbName, outputFile);

        // Assert
        Assert.That(File.Exists(outputFile), Is.True);
        File.Delete(outputFile);
    }

    [Test]
    public void ExportAsPdf_WithInvalidExtension_ThrowsArgumentException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var outputFile = Path.Combine(TestDataTempPath, "workbook_export_test.docx");

        // Act & Assert
        var ex = Assert.Throws<ArgumentException>(() =>
        {
            _service!.ExportAsPdf(_wbName, outputFile);
        });

        Assert.That(ex!.Message, Contains.Substring("Only .pdf is allowed"));
    }

    [Test]
    public void ExportAsPdf_WithNonExistentDirectory_ThrowsDirectoryNotFoundException()
    {
        // Arrange
        Assert.That(_wbName, Is.Not.Null);
        var nonExistentDir = Path.Combine(TestDataTempPath, "non_existent_folder_abc");
        var outputFile = Path.Combine(nonExistentDir, "workbook_export_test.pdf");

        // Act & Assert
        Assert.Throws<DirectoryNotFoundException>(() =>
        {
            _service!.ExportAsPdf(_wbName, outputFile);
        });
    }
}
