using System.IO;
using System.Runtime.InteropServices;

using AgentExcel.Models;
using AgentExcel.Providers;
using AgentExcel.Services;

using NUnit.Framework;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Tests.Services;

[TestFixture]
[Category("COM")]
public class WorkbookServiceTests : BaseTests
{
#pragma warning disable NUnit1032
    private static ExcelConnectionProvider? s_provider;
#pragma warning restore NUnit1032
    private WorkbookService? _service;

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
                        Marshal.ReleaseComObject(wb);
                    }
                }
                catch
                {
                    // Ignore
                }
                finally
                {
                    if (wbs != null)
                    {
                        Marshal.ReleaseComObject(wbs);
                    }
                }
            }
            s_provider = null;

            AgentExcel.Utils.ExcelConnector.ForceGarbageCollection();
        }
    }

    [SetUp]
    public void Setup()
    {
        _service = new WorkbookService(s_provider!);
    }

    [TearDown]
    public void Teardown()
    {
        // Close any workbooks created during the test to keep a clean state for the next test
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
                        Marshal.ReleaseComObject(wb);
                    }
                }
                catch
                {
                    // Ignore
                }
                finally
                {
                    if (wbs != null)
                    {
                        Marshal.ReleaseComObject(wbs);
                    }
                }
            }
        }
    }

    [Test]
    public void GetWorkbookInfo_WhenWorkbookExists_ReturnsWorkbookInfoWithSheets()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;
        Excel.Sheets sheets = wb.Sheets;

        // Add a second sheet to make sure we have multiple sheets
        Excel.Worksheet newSheet = (Excel.Worksheet)sheets.Add();
        Assert.That(newSheet, Is.Not.Null);
        string sheet1Name = newSheet.Name;

        // Get name of the other default sheet
        Excel.Worksheet defaultSheet = (Excel.Worksheet)sheets[2];
        Assert.That(defaultSheet, Is.Not.Null);
        string sheet2Name = defaultSheet.Name;

        // Release the temporary COM references we created locally in Act/Arrange
        Marshal.ReleaseComObject(defaultSheet);
        Marshal.ReleaseComObject(newSheet);
        Marshal.ReleaseComObject(sheets);
        Marshal.ReleaseComObject(wb);
        Marshal.ReleaseComObject(wbs);

        // Act
        WorkbookInfo info = _service!.GetWorkbookInfo(wbName);

        // Assert
        Assert.That(info, Is.Not.Null);
        Assert.That(info.Name, Is.EqualTo(wbName));
        Assert.That(info.Path, Is.Not.Null);
        Assert.That(info.Sheets, Is.Not.Null);
        Assert.That(info.Sheets, Contains.Item(sheet1Name));
        Assert.That(info.Sheets, Contains.Item(sheet2Name));
    }

    [Test]
    public void GetWorkbookInfo_WhenWorkbookDoesNotExist_ThrowsException()
    {
        // Act & Assert
        Assert.That(() => _service!.GetWorkbookInfo("NonExistentWorkbookName"), Throws.TypeOf<Exception>().And.Message.Contains("not found"));
    }

    [Test]
    public void OpenWorkbook_WithNullOrEmptyPath_ThrowsArgumentException()
    {
        // Arrange
        string path = "";

        // Act & Assert
        Assert.That(() => _service!.OpenWorkbook(path), Throws.TypeOf<ArgumentException>().And.Message.Contains("cannot be null or empty"));
    }

    [Test]
    public void OpenWorkbook_WithInvalidExtension_ThrowsArgumentException()
    {
        // Arrange
        string path = "test_file.txt";

        // Act & Assert
        Assert.That(() => _service!.OpenWorkbook(path), Throws.TypeOf<ArgumentException>().And.Message.Contains("Only .xls, .xlsx, and .xlsm are allowed"));
    }

    [Test]
    public void OpenWorkbook_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string path = "non_existent_file_xyz_123.xlsx";

        // Act & Assert
        Assert.That(() => _service!.OpenWorkbook(path), Throws.TypeOf<System.IO.FileNotFoundException>());
    }

    [Test]
    public void SaveAsWorkbook_WithNullOrEmptyPath_ThrowsArgumentException()
    {
        // Arrange
        string path = "";

        // Act & Assert
        Assert.That(() => _service!.SaveAsWorkbook("Book1", path), Throws.TypeOf<ArgumentException>().And.Message.Contains("cannot be null or empty"));
    }

    [Test]
    public void SaveAsWorkbook_WithInvalidExtension_ThrowsArgumentException()
    {
        // Arrange
        string path = "test_file.txt";

        // Act & Assert
        Assert.That(() => _service!.SaveAsWorkbook("Book1", path), Throws.TypeOf<ArgumentException>().And.Message.Contains("Only .xls, .xlsx, and .xlsm are allowed"));
    }

    [Test]
    public void SaveAsWorkbook_WithNonExistentDirectory_ThrowsDirectoryNotFoundException()
    {
        // Arrange
        string path = Path.Combine("non_existent_folder_xyz_123", "test_file.xlsx");

        // Act & Assert
        Assert.That(() => _service!.SaveAsWorkbook("Book1", path), Throws.TypeOf<DirectoryNotFoundException>());
    }

    [TestCase(".xlsx")]
    [TestCase(".xlsm")]
    [TestCase(".xls")]
    public void SaveAsWorkbook_WithValidExtension_SavesFileSuccessfully(string extension)
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        // Release references immediately to avoid locking issues when doing operations via service
        Marshal.ReleaseComObject(wb);
        Marshal.ReleaseComObject(wbs);

        string tempFileName = $"test_save_{Guid.NewGuid()}{extension}";
        string tempFilePath = Path.Combine(TestDataTempPath, tempFileName);
        string newWbName = Path.GetFileName(tempFilePath);

        try
        {
            // Act
            _service!.SaveAsWorkbook(wbName, tempFilePath);

            // Assert
            Assert.That(File.Exists(tempFilePath), Is.True);

            // Close the workbook from the service side first to release locks
            _service!.CloseWorkbook(newWbName, saveChanges: false);

            // Re-open using Interop directly to verify the physical format of the saved file
            Excel.Workbook? checkWb = null;
            Excel.Workbooks? checkWbs = null;
            try
            {
                checkWbs = app.Workbooks;
                checkWb = checkWbs.Open(tempFilePath);
                Excel.XlFileFormat expectedFormat = extension switch
                {
                    ".xls" => Excel.XlFileFormat.xlExcel8,
                    ".xlsm" => Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
                    _ => Excel.XlFileFormat.xlOpenXMLWorkbook
                };
                Assert.That(checkWb.FileFormat, Is.EqualTo(expectedFormat));
            }
            finally
            {
                if (checkWb != null)
                {
                    checkWb.Close(false);
                    Marshal.ReleaseComObject(checkWb);
                }
                if (checkWbs != null)
                {
                    Marshal.ReleaseComObject(checkWbs);
                }
            }
        }
        finally
        {
            // Close the saved workbook (failsafe in case of exception before it gets closed in the try block)
            try
            {
                _service!.CloseWorkbook(newWbName, saveChanges: false);
            }
            catch
            {
                // Ignore if it's already closed or failed
            }

            // Cleanup
            if (File.Exists(tempFilePath))
            {
                try
                {
                    File.Delete(tempFilePath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
    }

    [Test]
    public void ListWorkbooks_WhenWorkbooksExist_ReturnsListOfWorkbookInfo()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Marshal.ReleaseComObject(wb);
        Marshal.ReleaseComObject(wbs);

        // Act
        var list = _service!.ListWorkbooks();

        // Assert
        Assert.That(list, Is.Not.Null);
        Assert.That(list.Any(w => w.Name == wbName), Is.True);
    }

    [Test]
    public void OpenWorkbook_WithValidExistingFile_OpensFileSuccessfully()
    {
        // Arrange
        string testFile = Path.Combine(TestDataPath, "wk0.xls");
        string expectedName = "wk0.xls";

        // Act
        WorkbookInfo info = _service!.OpenWorkbook(testFile);

        // Assert
        Assert.That(info, Is.Not.Null);
        Assert.That(info.Name, Is.EqualTo(expectedName));
        Assert.That(info.Path, Is.Not.Null.And.Not.Empty);
        Assert.That(info.Sheets, Is.Not.Null.And.Not.Empty);
    }

    [Test]
    public void AddWorkbook_WhenCalled_CreatesNewWorkbookAndReturnsInfo()
    {
        // Act
        WorkbookInfo info = _service!.AddWorkbook();

        // Assert
        Assert.That(info, Is.Not.Null);
        Assert.That(info.Name, Is.Not.Null.And.Not.Empty);
        Assert.That(info.Sheets, Is.Not.Null.And.Not.Empty);
    }

    [Test]
    public void SaveWorkbook_WhenCalled_SavesWorkbookSuccessfully()
    {
        // Arrange
        string sourcePath = Path.Combine(TestDataPath, "wk0.xls");
        string tempPath = Path.Combine(TestDataTempPath, $"temp_save_{Guid.NewGuid()}.xls");
        File.Copy(sourcePath, tempPath, overwrite: true);

        try
        {
            WorkbookInfo info = _service!.OpenWorkbook(tempPath);
            string wbName = info.Name;

            // Act
            _service.SaveWorkbook(wbName);

            // Assert
            Assert.That(File.Exists(tempPath), Is.True);
        }
        finally
        {
            // Cleanup
            if (File.Exists(tempPath))
            {
                try
                {
                    File.Delete(tempPath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
    }

    [Test]
    public void CloseWorkbook_WhenCalled_ClosesWorkbookSuccessfully()
    {
        // Arrange
        WorkbookInfo info = _service!.AddWorkbook();
        string wbName = info.Name;

        // Act
        _service.CloseWorkbook(wbName, saveChanges: false);

        // Assert
        var list = _service.ListWorkbooks();
        Assert.That(list.Any(w => w.Name == wbName), Is.False);
    }
}

