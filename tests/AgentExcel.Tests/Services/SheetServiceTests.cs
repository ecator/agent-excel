using System;
using System.IO;
using System.Linq;

using AgentExcel.Utils;
using AgentExcel.Models;
using AgentExcel.Providers;
using AgentExcel.Services;

using NUnit.Framework;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Tests.Services;

[TestFixture]
[Category("COM")]
public class SheetServiceTests : BaseTests
{
#pragma warning disable NUnit1032
    private static ExcelConnectionProvider? s_provider;
#pragma warning restore NUnit1032
    private SheetService _service = null!;

    private List<string> HelperListSheetNames(string wbName)
    {
        return _service!.ListSheets(wbName).Select(ws => ws.Name).ToList();
    }

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

            AgentExcel.Utils.ExcelConnector.ForceGarbageCollection();
        }
    }

    [SetUp]
    public void Setup()
    {
        _service = new SheetService(s_provider!);
    }

    [TearDown]
    public void Teardown()
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
            }
        }
    }

    [Test]
    public void CopySheet_SameWorkbook_WithCollision_ThrowsException()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string sheetName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act & Assert
        Assert.That(() => _service!.CopySheet(wbName, sheetName, null, wbName),
            Throws.TypeOf<Exception>().And.Message.Contains("already exists"));
    }

    [Test]
    public void CopySheet_DifferentWorkbook_NoCollision_CopiesSuccessfully()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wbSource = wbs.Add(Type.Missing);
        string sourceWbName = wbSource.Name;

        Excel.Sheets sourceSheets = wbSource.Sheets;
        Excel.Worksheet sourceWs = (Excel.Worksheet)sourceSheets[1];
        sourceWs.Name = "UniqueSourceSheet";
        string sheetName = sourceWs.Name;

        Excel.Workbook wbTarget = wbs.Add(Type.Missing);
        string targetWbName = wbTarget.Name;

        ExcelConnector.SafeReleaseComObject(sourceWs);
        ExcelConnector.SafeReleaseComObject(sourceSheets);
        ExcelConnector.SafeReleaseComObject(wbSource);
        ExcelConnector.SafeReleaseComObject(wbTarget);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act
        _service!.CopySheet(sourceWbName, sheetName, null, targetWbName);

        // Assert
        var sheetsList = HelperListSheetNames(targetWbName);
        Assert.That(sheetsList, Contains.Item(sheetName));
    }

    [Test]
    public void CopySheet_DifferentWorkbook_WithCollision_ThrowsException()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wbSource = wbs.Add(Type.Missing);
        string sourceWbName = wbSource.Name;

        Excel.Sheets sourceSheets = wbSource.Sheets;
        Excel.Worksheet sourceWs = (Excel.Worksheet)sourceSheets[1];
        sourceWs.Name = "CollisionSheet";
        string sheetName = sourceWs.Name;

        Excel.Workbook wbTarget = wbs.Add(Type.Missing);
        string targetWbName = wbTarget.Name;

        Excel.Sheets targetSheets = wbTarget.Sheets;
        Excel.Worksheet targetWs = (Excel.Worksheet)targetSheets[1];
        targetWs.Name = "CollisionSheet";

        ExcelConnector.SafeReleaseComObject(targetWs);
        ExcelConnector.SafeReleaseComObject(targetSheets);
        ExcelConnector.SafeReleaseComObject(sourceWs);
        ExcelConnector.SafeReleaseComObject(sourceSheets);
        ExcelConnector.SafeReleaseComObject(wbSource);
        ExcelConnector.SafeReleaseComObject(wbTarget);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act & Assert
        Assert.That(() => _service!.CopySheet(sourceWbName, sheetName, null, targetWbName),
            Throws.TypeOf<Exception>().And.Message.Contains("already exists"));
    }

    [Test]
    public void MoveSheet_SameWorkbook_RepositionsSuccessfully()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        // Make sure we have at least 2 sheets
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        string ws2Name = ws2.Name;

        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[2]; // originally added
        string ws1Name = ws1.Name;

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Originally ws2 is added at index 1 (starts at first position because sheets.Add adds at beginning by default).
        // Let's verify by listing sheets first
        var initialList = HelperListSheetNames(wbName);
        Assert.That(initialList[0], Is.EqualTo(ws2Name));

        // Act: move ws2 to the end (after ws1)
        _service.MoveSheet(wbName, ws2Name, wbName);

        // Assert
        var listAfterMove = HelperListSheetNames(wbName);
        Assert.That(listAfterMove.Last(), Is.EqualTo(ws2Name));
    }

    [Test]
    public void MoveSheet_DifferentWorkbook_NoCollision_MovesSuccessfully()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wbSource = wbs.Add(Type.Missing);
        string sourceWbName = wbSource.Name;

        Excel.Sheets sourceSheets = wbSource.Sheets;
        // Add a second sheet so we can delete/move the first one (Excel needs at least 1 sheet left)
        Excel.Worksheet dummyWs = (Excel.Worksheet)sourceSheets.Add();
        Excel.Worksheet sourceWs = (Excel.Worksheet)sourceSheets[2];
        sourceWs.Name = "MoveMeSheet";
        string sheetName = sourceWs.Name;

        Excel.Workbook wbTarget = wbs.Add(Type.Missing);
        string targetWbName = wbTarget.Name;

        ExcelConnector.SafeReleaseComObject(dummyWs);
        ExcelConnector.SafeReleaseComObject(sourceWs);
        ExcelConnector.SafeReleaseComObject(sourceSheets);
        ExcelConnector.SafeReleaseComObject(wbSource);
        ExcelConnector.SafeReleaseComObject(wbTarget);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act
        _service!.MoveSheet(sourceWbName, sheetName, targetWbName);

        // Assert
        var targetSheetsList = HelperListSheetNames(targetWbName);
        Assert.That(targetSheetsList, Contains.Item(sheetName));

        var sourceSheetsList = HelperListSheetNames(sourceWbName);
        Assert.That(sourceSheetsList, Does.Not.Contain(sheetName));
    }

    [Test]
    public void MoveSheet_DifferentWorkbook_WithCollision_ThrowsException()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wbSource = wbs.Add(Type.Missing);
        string sourceWbName = wbSource.Name;

        Excel.Sheets sourceSheets = wbSource.Sheets;
        Excel.Worksheet dummyWs = (Excel.Worksheet)sourceSheets.Add();
        Excel.Worksheet sourceWs = (Excel.Worksheet)sourceSheets[2];
        sourceWs.Name = "CollisionSheet";
        string sheetName = sourceWs.Name;

        Excel.Workbook wbTarget = wbs.Add(Type.Missing);
        string targetWbName = wbTarget.Name;

        Excel.Sheets targetSheets = wbTarget.Sheets;
        Excel.Worksheet targetWs = (Excel.Worksheet)targetSheets[1];
        targetWs.Name = "CollisionSheet";

        ExcelConnector.SafeReleaseComObject(dummyWs);
        ExcelConnector.SafeReleaseComObject(targetWs);
        ExcelConnector.SafeReleaseComObject(targetSheets);
        ExcelConnector.SafeReleaseComObject(sourceWs);
        ExcelConnector.SafeReleaseComObject(sourceSheets);
        ExcelConnector.SafeReleaseComObject(wbSource);
        ExcelConnector.SafeReleaseComObject(wbTarget);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act & Assert
        Assert.That(() => _service!.MoveSheet(sourceWbName, sheetName, targetWbName),
            Throws.TypeOf<Exception>().And.Message.Contains("already exists"));
    }

    [Test]
    public void CopySheet_DifferentWorkbook_WithPositionZero_CopiesToFirstPosition()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wbSource = wbs.Add(Type.Missing);
        string sourceWbName = wbSource.Name;

        Excel.Sheets sourceSheets = wbSource.Sheets;
        Excel.Worksheet sourceWs = (Excel.Worksheet)sourceSheets[1];
        sourceWs.Name = "SheetToCopy";
        string sheetName = sourceWs.Name;

        Excel.Workbook wbTarget = wbs.Add(Type.Missing);
        string targetWbName = wbTarget.Name;

        ExcelConnector.SafeReleaseComObject(sourceWs);
        ExcelConnector.SafeReleaseComObject(sourceSheets);
        ExcelConnector.SafeReleaseComObject(wbSource);
        ExcelConnector.SafeReleaseComObject(wbTarget);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act
        _service!.CopySheet(sourceWbName, sheetName, null, targetWbName, position: 0);

        // Assert
        var targetSheetsList = HelperListSheetNames(targetWbName);
        Assert.That(targetSheetsList[0], Is.EqualTo(sheetName));
    }

    [Test]
    public void MoveSheet_SameWorkbook_WithPositionZero_MovesToFirstPosition()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        string ws2Name = ws2.Name;

        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[2];
        string ws1Name = ws1.Name;

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act
        _service!.MoveSheet(wbName, ws1Name, wbName, position: 0);

        // Assert
        var listAfterMove = HelperListSheetNames(wbName);
        Assert.That(listAfterMove[0], Is.EqualTo(ws1Name));
    }

    [Test]
    public void MoveSheet_SameWorkbook_WithPositionOne_MovesToSecondPosition()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws3 = (Excel.Worksheet)sheets.Add();
        string ws3Name = ws3.Name;

        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        string ws2Name = ws2.Name;

        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[3];
        string ws1Name = ws1.Name;

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(ws3);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act
        _service!.MoveSheet(wbName, ws2Name, wbName, position: 1);

        // Assert
        var listAfterMove = HelperListSheetNames(wbName);
        Assert.That(listAfterMove[1], Is.EqualTo(ws2Name));
    }

    [Test]
    public void DeleteSheet_WhenMultipleSheetsExist_DeletesSuccessfully()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        string ws2Name = ws2.Name;

        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[2];
        string ws1Name = ws1.Name;

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act
        _service!.DeleteSheet(wbName, ws2Name);

        // Assert
        var remainingSheets = HelperListSheetNames(wbName);
        Assert.That(remainingSheets, Has.Count.EqualTo(1));
        Assert.That(remainingSheets, Does.Not.Contain(ws2Name));
    }

    [Test]
    public void DeleteSheet_WhenOnlyOneSheetExists_ThrowsException()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string sheetName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act & Assert
        Assert.That(() => _service!.DeleteSheet(wbName, sheetName),
            Throws.TypeOf<InvalidOperationException>().And.Message.Contains("Cannot delete the only sheet"));
    }

    [Test]
    public void RenameSheet_WhenNewNameDoesNotExist_RenamesSuccessfully()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string oldName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        string newName = "NewUniqueName";

        // Act
        _service!.RenameSheet(wbName, oldName, newName);

        // Assert
        var remainingSheets = HelperListSheetNames(wbName);
        Assert.That(remainingSheets, Contains.Item(newName));
        Assert.That(remainingSheets, Does.Not.Contain(oldName));
    }

    [Test]
    public void RenameSheet_WhenNewNameAlreadyExists_ThrowsException()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        string name2 = ws2.Name;
        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[2];
        string name1 = ws1.Name;

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act & Assert
        Assert.That(() => _service!.RenameSheet(wbName, name2, name1),
            Throws.TypeOf<InvalidOperationException>().And.Message.Contains("already exists"));
    }

    [Test]
    public void RenameSheet_WhenNewNameIsSameWithDifferentCase_RenamesSuccessfully()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        ws.Name = "tempname";
        string oldName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        string newName = "TempName";

        // Act
        _service!.RenameSheet(wbName, oldName, newName);

        // Assert
        var remainingSheets = HelperListSheetNames(wbName);
        Assert.That(remainingSheets.Contains(newName), Is.True);
    }

    [Test]
    public void AddSheet_WhenNameDoesNotExist_AddsSuccessfully()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        string targetSheetName = "NewUniqueSheet";

        // Act
        _service!.AddSheet(wbName, targetSheetName);

        // Assert
        var sheets = HelperListSheetNames(wbName);
        Assert.That(sheets, Contains.Item(targetSheetName));
    }

    [Test]
    public void AddSheet_WhenNameAlreadyExists_ThrowsInvalidOperationException()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string existingSheetName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act & Assert
        Assert.That(() => _service!.AddSheet(wbName, existingSheetName),
            Throws.TypeOf<InvalidOperationException>().And.Message.Contains("already exists"));
    }

    [Test]
    public void CopySheet_SameWorkbook_WithNewName_CopiesAndRenamesSuccessfully()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string oldName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        string newName = "CopiedSheetNewName";

        // Act
        _service!.CopySheet(wbName, oldName, newName, wbName);

        // Assert
        var sheetsList = HelperListSheetNames(wbName);
        Assert.That(sheetsList, Contains.Item(oldName));
        Assert.That(sheetsList, Contains.Item(newName));
    }

    [Test]
    public void CopySheet_DifferentWorkbook_WithNewName_CopiesAndRenamesSuccessfully()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wbSource = wbs.Add(Type.Missing);
        string sourceWbName = wbSource.Name;

        Excel.Sheets sourceSheets = wbSource.Sheets;
        Excel.Worksheet sourceWs = (Excel.Worksheet)sourceSheets[1];
        sourceWs.Name = "UniqueSourceSheet";
        string oldName = sourceWs.Name;

        Excel.Workbook wbTarget = wbs.Add(Type.Missing);
        string targetWbName = wbTarget.Name;

        ExcelConnector.SafeReleaseComObject(sourceWs);
        ExcelConnector.SafeReleaseComObject(sourceSheets);
        ExcelConnector.SafeReleaseComObject(wbSource);
        ExcelConnector.SafeReleaseComObject(wbTarget);
        ExcelConnector.SafeReleaseComObject(wbs);

        string newName = "CopiedSheetInTarget";

        // Act
        _service!.CopySheet(sourceWbName, oldName, newName, targetWbName);

        // Assert
        var targetSheets = HelperListSheetNames(targetWbName);
        Assert.That(targetSheets, Contains.Item(newName));
        Assert.That(targetSheets, Does.Not.Contain(oldName));
    }

    [Test]
    public void CopySheet_SameWorkbook_VariousPositions_CopiesCorrectly()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws3 = (Excel.Worksheet)sheets.Add();
        ws3.Name = "S3";
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        ws2.Name = "S2";
        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[3];
        ws1.Name = "S1";

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(ws3);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Initial list: S2, S3, S1 (due to sheets.Add default behavior)
        var initialList = HelperListSheetNames(wbName);
        Assert.That(initialList, Is.EqualTo(new[] { "S2", "S3", "S1" }));

        // Act & Assert 1: Copy S1 to "C0" at position 0 (first)
        _service.CopySheet(wbName, "S1", "C0", wbName, position: 0);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "C0", "S2", "S3", "S1" }));

        // Act & Assert 2: Copy S1 to "C1" at position 1 (second)
        _service.CopySheet(wbName, "S1", "C1", wbName, position: 1);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "C0", "C1", "S2", "S3", "S1" }));

        // Act & Assert 3: Copy S1 to "C3" at position 3 (fourth)
        _service.CopySheet(wbName, "S1", "C3", wbName, position: 3);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "C0", "C1", "S2", "C3", "S3", "S1" }));

        // Act & Assert 4: Copy S1 to "CLast" at position -1 (last)
        _service.CopySheet(wbName, "S1", "CLast", wbName, position: -1);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "C0", "C1", "S2", "C3", "S3", "S1", "CLast" }));
    }

    [Test]
    public void CopySheet_DifferentWorkbook_VariousPositions_CopiesCorrectly()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wbSrc = wbs.Add(Type.Missing);
        string srcWbName = wbSrc.Name;
        Excel.Sheets srcSheets = wbSrc.Sheets;
        Excel.Worksheet srcWs = (Excel.Worksheet)srcSheets[1];
        srcWs.Name = "SrcSheet";

        Excel.Workbook wbDest = wbs.Add(Type.Missing);
        string destWbName = wbDest.Name;
        Excel.Sheets destSheets = wbDest.Sheets;
        Excel.Worksheet t3 = (Excel.Worksheet)destSheets.Add();
        t3.Name = "T3";
        Excel.Worksheet t2 = (Excel.Worksheet)destSheets.Add();
        t2.Name = "T2";
        Excel.Worksheet t1 = (Excel.Worksheet)destSheets[3];
        t1.Name = "T1";

        ExcelConnector.SafeReleaseComObject(srcWs);
        ExcelConnector.SafeReleaseComObject(srcSheets);
        ExcelConnector.SafeReleaseComObject(wbSrc);
        ExcelConnector.SafeReleaseComObject(t1);
        ExcelConnector.SafeReleaseComObject(t2);
        ExcelConnector.SafeReleaseComObject(t3);
        ExcelConnector.SafeReleaseComObject(destSheets);
        ExcelConnector.SafeReleaseComObject(wbDest);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Target initial: T2, T3, T1
        Assert.That(HelperListSheetNames(destWbName), Is.EqualTo(new[] { "T2", "T3", "T1" }));

        // Act & Assert 1: Copy SrcSheet to "C0" at target position 0
        _service.CopySheet(srcWbName, "SrcSheet", "C0", destWbName, position: 0);
        Assert.That(HelperListSheetNames(destWbName), Is.EqualTo(new[] { "C0", "T2", "T3", "T1" }));

        // Act & Assert 2: Copy SrcSheet to "C1" at target position 1
        _service.CopySheet(srcWbName, "SrcSheet", "C1", destWbName, position: 1);
        Assert.That(HelperListSheetNames(destWbName), Is.EqualTo(new[] { "C0", "C1", "T2", "T3", "T1" }));

        // Act & Assert 3: Copy SrcSheet to "C3" at target position 3
        _service.CopySheet(srcWbName, "SrcSheet", "C3", destWbName, position: 3);
        Assert.That(HelperListSheetNames(destWbName), Is.EqualTo(new[] { "C0", "C1", "T2", "C3", "T3", "T1" }));

        // Act & Assert 4: Copy SrcSheet to "CLast" at target position -1
        _service.CopySheet(srcWbName, "SrcSheet", "CLast", destWbName, position: -1);
        Assert.That(HelperListSheetNames(destWbName), Is.EqualTo(new[] { "C0", "C1", "T2", "C3", "T3", "T1", "CLast" }));
    }

    [Test]
    public void MoveSheet_SameWorkbook_VariousPositions_MovesCorrectly()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws4 = (Excel.Worksheet)sheets.Add();
        ws4.Name = "S4";
        Excel.Worksheet ws3 = (Excel.Worksheet)sheets.Add();
        ws3.Name = "S3";
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        ws2.Name = "S2";
        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[4];
        ws1.Name = "S1";

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(ws3);
        ExcelConnector.SafeReleaseComObject(ws4);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Initial: S2, S3, S4, S1
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "S2", "S3", "S4", "S1" }));

        // Act & Assert 1: Move S2 (position 0) to position 2 (after S4)
        _service.MoveSheet(wbName, "S2", wbName, position: 2);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "S3", "S4", "S2", "S1" }));

        // Act & Assert 2: Move S1 (position 3) to position 0 (first)
        _service.MoveSheet(wbName, "S1", wbName, position: 0);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "S1", "S3", "S4", "S2" }));

        // Act & Assert 3: Move S4 (position 2) to position -1 (last)
        _service.MoveSheet(wbName, "S4", wbName, position: -1);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "S1", "S3", "S2", "S4" }));
    }

    [Test]
    public void MoveSheet_DifferentWorkbook_VariousPositions_MovesCorrectly()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wbSrc = wbs.Add(Type.Missing);
        string srcWbName = wbSrc.Name;
        Excel.Sheets srcSheets = wbSrc.Sheets;
        Excel.Worksheet dummy = (Excel.Worksheet)srcSheets.Add();
        Excel.Worksheet s2 = (Excel.Worksheet)srcSheets.Add();
        s2.Name = "S2";
        Excel.Worksheet s1 = (Excel.Worksheet)srcSheets[3];
        s1.Name = "S1";

        Excel.Workbook wbDest = wbs.Add(Type.Missing);
        string destWbName = wbDest.Name;
        Excel.Sheets destSheets = wbDest.Sheets;
        Excel.Worksheet t3 = (Excel.Worksheet)destSheets.Add();
        t3.Name = "T3";
        Excel.Worksheet t2 = (Excel.Worksheet)destSheets.Add();
        t2.Name = "T2";
        Excel.Worksheet t1 = (Excel.Worksheet)destSheets[3];
        t1.Name = "T1";

        ExcelConnector.SafeReleaseComObject(s1);
        ExcelConnector.SafeReleaseComObject(s2);
        ExcelConnector.SafeReleaseComObject(dummy);
        ExcelConnector.SafeReleaseComObject(srcSheets);
        ExcelConnector.SafeReleaseComObject(wbSrc);
        ExcelConnector.SafeReleaseComObject(t1);
        ExcelConnector.SafeReleaseComObject(t2);
        ExcelConnector.SafeReleaseComObject(t3);
        ExcelConnector.SafeReleaseComObject(destSheets);
        ExcelConnector.SafeReleaseComObject(wbDest);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Target Initial: T2, T3, T1
        Assert.That(HelperListSheetNames(destWbName), Is.EqualTo(new[] { "T2", "T3", "T1" }));

        // Act & Assert 1: Move S1 to target position 0
        _service.MoveSheet(srcWbName, "S1", destWbName, position: 0);
        Assert.That(HelperListSheetNames(destWbName), Is.EqualTo(new[] { "S1", "T2", "T3", "T1" }));

        // Act & Assert 2: Move S2 to target position 2
        _service.MoveSheet(srcWbName, "S2", destWbName, position: 2);
        Assert.That(HelperListSheetNames(destWbName), Is.EqualTo(new[] { "S1", "T2", "S2", "T3", "T1" }));
    }

    [Test]
    public void CopySheet_SameWorkbook_WithNegativePositions_CopiesCorrectly()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws3 = (Excel.Worksheet)sheets.Add();
        ws3.Name = "S3";
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        ws2.Name = "S2";
        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[3];
        ws1.Name = "S1";

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(ws3);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Initial: S2, S3, S1
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "S2", "S3", "S1" }));

        // Act & Assert 1: Copy S1 to "C_Neg2" at position -2 (second to last)
        // Expected order: S2, S3, C_Neg2, S1
        _service.CopySheet(wbName, "S1", "C_Neg2", wbName, position: -2);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "S2", "S3", "C_Neg2", "S1" }));

        // Act & Assert 2: Copy S1 to "C_Neg4" at position -4 (second sheet)
        // Expected order: S2, C_Neg4, S3, C_Neg2, S1
        _service.CopySheet(wbName, "S1", "C_Neg4", wbName, position: -4);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "S2", "C_Neg4", "S3", "C_Neg2", "S1" }));

        // Act & Assert 3: Copy S1 to "C_Neg5" at position -5 (second sheet)
        // Expected order: S2, C_Neg5, C_Neg4, S3, C_Neg2, S1
        _service.CopySheet(wbName, "S1", "C_Neg5", wbName, position: -5);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "S2", "C_Neg5", "C_Neg4", "S3", "C_Neg2", "S1" }));

        // Act & Assert 4: Copy S1 to "C_Neg6" at position -7 (first sheet due to clamping)
        // Expected order: C_Neg6, S2, C_Neg5, C_Neg4, S3, C_Neg2, S1
        _service.CopySheet(wbName, "S1", "C_Neg6", wbName, position: -7);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "C_Neg6", "S2", "C_Neg5", "C_Neg4", "S3", "C_Neg2", "S1" }));
    }

    [Test]
    public void MoveSheet_SameWorkbook_WithNegativePositions_MovesCorrectly()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws4 = (Excel.Worksheet)sheets.Add();
        ws4.Name = "S4";
        Excel.Worksheet ws3 = (Excel.Worksheet)sheets.Add();
        ws3.Name = "S3";
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        ws2.Name = "S2";
        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[4];
        ws1.Name = "S1";

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(ws3);
        ExcelConnector.SafeReleaseComObject(ws4);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Initial: S2, S3, S4, S1
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "S2", "S3", "S4", "S1" }));

        // Act & Assert 1: Move S2 (position 0) to position -2 (second to last, after S4)
        // Expected order: S3, S4, S2, S1
        _service.MoveSheet(wbName, "S2", wbName, position: -2);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "S3", "S4", "S2", "S1" }));

        // Act & Assert 2: Move S1 (position 3) to position -4 (first sheet)
        // Expected order: S1, S3, S4, S2
        _service.MoveSheet(wbName, "S1", wbName, position: -4);
        Assert.That(HelperListSheetNames(wbName), Is.EqualTo(new[] { "S1", "S3", "S4", "S2" }));
    }

    [Test]
    public void MoveSheet_WhenOnlyOneSheetExists_ThrowsInvalidOperationException()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        while (sheets.Count > 1)
        {
            Excel.Worksheet wsToDelete = (Excel.Worksheet)sheets[1];
            app.DisplayAlerts = false;
            wsToDelete.Delete();
            app.DisplayAlerts = true;
            ExcelConnector.SafeReleaseComObject(wsToDelete);
        }

        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string sheetName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act & Assert
        Assert.That(() => _service!.MoveSheet(wbName, sheetName, "TargetWb"),
            Throws.TypeOf<InvalidOperationException>().And.Message.Contains("Cannot move the only sheet"));
    }

    [Test]
    public void GetWorksheetInfo_Color_WithNoColor_ReturnsNone()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string sheetName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act
        string color = _service!.GetWorksheetInfo(wbName, sheetName).Color;

        // Assert
        Assert.That(color, Is.EqualTo("None"));
    }

    [Test]
    public void SetSheetColor_WithValidHex_SetsColorSuccessfully()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string sheetName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        string targetColor = "#FF0000";

        // Act
        _service!.SetSheetColor(wbName, sheetName, targetColor);
        string color = _service.GetWorksheetInfo(wbName, sheetName).Color;

        // Assert
        Assert.That(color, Is.EqualTo(targetColor));
    }

    [Test]
    public void SetSheetColor_WithNone_ClearsColorSuccessfully()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string sheetName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        _service!.SetSheetColor(wbName, sheetName, "#00FF00");

        // Act
        _service.SetSheetColor(wbName, sheetName, "None");
        string color = _service.GetWorksheetInfo(wbName, sheetName).Color;

        // Assert
        Assert.That(color, Is.EqualTo("None"));
    }

    [Test]
    public void GetWorksheetInfo_Visibility_ReturnsCorrectState()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string sheetName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act
        string visibility = _service!.GetWorksheetInfo(wbName, sheetName).Visibility;

        // Assert
        Assert.That(visibility, Is.EqualTo("Visible"));
    }

    [Test]
    public void SetSheetVisibility_ToHidden_SucceedsWhenMultipleVisibleSheetsExist()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        string ws2Name = ws2.Name;
        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[2];
        string ws1Name = ws1.Name;

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act
        _service!.SetSheetVisibility(wbName, ws2Name, "Hidden");
        string visibility = _service.GetWorksheetInfo(wbName, ws2Name).Visibility;

        // Assert
        Assert.That(visibility, Is.EqualTo("Hidden"));
    }

    [Test]
    public void SetSheetVisibility_ToHidden_ThrowsWhenOnlyOneVisibleSheetExists()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string sheetName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act & Assert
        Assert.That(() => _service!.SetSheetVisibility(wbName, sheetName, "Hidden"),
            Throws.TypeOf<InvalidOperationException>().And.Message.Contains("At least one worksheet must remain visible"));
    }

    [Test]
    public void SetSheetVisibility_ToVeryHidden_SucceedsWhenMultipleVisibleSheetsExist()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        string ws2Name = ws2.Name;
        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[2];
        string ws1Name = ws1.Name;

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act
        _service!.SetSheetVisibility(wbName, ws2Name, "VeryHidden");
        string visibility = _service.GetWorksheetInfo(wbName, ws2Name).Visibility;

        // Assert
        Assert.That(visibility, Is.EqualTo("VeryHidden"));
    }

    [Test]
    public void SetSheetVisibility_ToVeryHidden_ThrowsWhenOnlyOneVisibleSheetExists()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws = (Excel.Worksheet)sheets[1];
        string sheetName = ws.Name;

        ExcelConnector.SafeReleaseComObject(ws);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act & Assert
        Assert.That(() => _service!.SetSheetVisibility(wbName, sheetName, "VeryHidden"),
            Throws.TypeOf<InvalidOperationException>().And.Message.Contains("At least one worksheet must remain visible"));
    }

    [Test]
    public void GetWorksheetInfo_ReturnsCorrectValues()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        string ws2Name = ws2.Name;
        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[2];
        string ws1Name = ws1.Name;

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        _service!.SetSheetColor(wbName, ws2Name, "#FF0000");
        _service.SetSheetVisibility(wbName, ws2Name, "Hidden");

        // Act
        WorksheetInfo info1 = _service.GetWorksheetInfo(wbName, ws1Name);
        WorksheetInfo info2 = _service.GetWorksheetInfo(wbName, ws2Name);

        // Assert
        Assert.That(info1.Name, Is.EqualTo(ws1Name));
        Assert.That(info1.Color, Is.EqualTo("None"));
        Assert.That(info1.Visibility, Is.EqualTo("Visible"));

        Assert.That(info2.Name, Is.EqualTo(ws2Name));
        Assert.That(info2.Color, Is.EqualTo("#FF0000"));
        Assert.That(info2.Visibility, Is.EqualTo("Hidden"));
    }

    [Test]
    public void ListSheets_ReturnsWorksheetInfoArray()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        string ws2Name = ws2.Name;
        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[2];
        string ws1Name = ws1.Name;

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        _service!.SetSheetColor(wbName, ws2Name, "#00FF00");

        // Act
        List<WorksheetInfo> infos = _service.ListSheets(wbName);

        // Assert
        Assert.That(infos, Has.Count.EqualTo(2));
        WorksheetInfo info2 = infos.First(x => x.Name == ws2Name);
        WorksheetInfo info1 = infos.First(x => x.Name == ws1Name);

        Assert.That(info2.Color, Is.EqualTo("#00FF00"));
        Assert.That(info2.Visibility, Is.EqualTo("Visible"));
        Assert.That(info1.Color, Is.EqualTo("None"));
        Assert.That(info1.Visibility, Is.EqualTo("Visible"));
    }

    [Test]
    public void GetActiveSheet_WhenWorkbookExists_ReturnsActiveSheetInfo()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        string ws2Name = ws2.Name;

        // Set ws2 to active
        ws2.Activate();

        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act
        WorksheetInfo info = _service!.GetActiveSheet(wbName);

        // Assert
        Assert.That(info, Is.Not.Null);
        Assert.That(info.Name, Is.EqualTo(ws2Name));
    }

    [Test]
    public void GetActiveSheet_WhenWorkbookDoesNotExist_ThrowsException()
    {
        // Act & Assert
        Assert.That(() => _service!.GetActiveSheet("NonExistentWorkbookName"),
            Throws.TypeOf<Exception>().And.Message.Contains("not found"));
    }

    [Test]
    public void SetActiveSheet_WhenWorkbookAndSheetExist_ActivatesSheet()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        Excel.Sheets sheets = wb.Sheets;
        Excel.Worksheet ws2 = (Excel.Worksheet)sheets.Add();
        string ws2Name = ws2.Name;
        Excel.Worksheet ws1 = (Excel.Worksheet)sheets[2];
        string ws1Name = ws1.Name;

        // Ensure ws1 is currently active (though ws2 is typically active after Add)
        ws1.Activate();

        ExcelConnector.SafeReleaseComObject(ws1);
        ExcelConnector.SafeReleaseComObject(ws2);
        ExcelConnector.SafeReleaseComObject(sheets);
        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Verify active sheet before set
        var initialActive = _service!.GetActiveSheet(wbName);
        Assert.That(initialActive.Name, Is.EqualTo(ws1Name));

        // Act
        _service.SetActiveSheet(wbName, ws2Name);

        // Assert
        var finalActive = _service.GetActiveSheet(wbName);
        Assert.That(finalActive.Name, Is.EqualTo(ws2Name));
    }

    [Test]
    public void SetActiveSheet_WhenWorkbookDoesNotExist_ThrowsException()
    {
        // Act & Assert
        Assert.That(() => _service!.SetActiveSheet("NonExistentWorkbookName", "Sheet1"),
            Throws.TypeOf<Exception>().And.Message.Contains("not found"));
    }

    [Test]
    public void SetActiveSheet_WhenSheetDoesNotExist_ThrowsException()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks wbs = app.Workbooks;
        Excel.Workbook wb = wbs.Add(Type.Missing);
        string wbName = wb.Name;

        ExcelConnector.SafeReleaseComObject(wb);
        ExcelConnector.SafeReleaseComObject(wbs);

        // Act & Assert
        Assert.That(() => _service!.SetActiveSheet(wbName, "NonExistentSheetName"),
            Throws.TypeOf<Exception>().And.Message.Contains("not found"));
    }
}


