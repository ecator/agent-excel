using System;
using System.Collections.Generic;
using System.Linq;

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
public class ShapeServiceTests : BaseTests
{
#pragma warning disable NUnit1032
    private static ExcelConnectionProvider? s_provider;
#pragma warning restore NUnit1032
    private ShapeService? _service;
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
        _service = new ShapeService(s_provider!);

        // Create a new workbook for each test to ensure test isolation
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);

        Excel.Workbooks? wbs = null;
        Excel.Workbook? wb = null;
        try
        {
            wbs = app.Workbooks;
            wb = wbs.Add(Type.Missing);
            _wbName = wb.Name;
        }
        finally
        {
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
    public void AddShape_WithValidInputs_CreatesShapeAndReturnsShapeInfo()
    {
        // Act
        ShapeInfo shapeInfo = _service!.AddShape(_wbName, SheetName, "Rectangle", 10, 20, 100, 50, "Hello Shape");

        // Assert
        Assert.That(shapeInfo, Is.Not.Null);
        Assert.That(shapeInfo.Name, Is.Not.Null.And.Not.Empty);
        Assert.That(shapeInfo.Type, Is.EqualTo("Rectangle"));
        Assert.That(shapeInfo.Left, Is.EqualTo(10f).Within(0.01f));
        Assert.That(shapeInfo.Top, Is.EqualTo(20f).Within(0.01f));
        Assert.That(shapeInfo.Width, Is.EqualTo(100f).Within(0.01f));
        Assert.That(shapeInfo.Height, Is.EqualTo(50f).Within(0.01f));
        Assert.That(shapeInfo.Text, Is.EqualTo("Hello Shape"));
        Assert.That(shapeInfo.Connection, Is.Null);

        var list = _service.ListShapes(_wbName, SheetName);
        Assert.That(list, Has.Count.EqualTo(1));
        Assert.That(list[0].Name, Is.EqualTo(shapeInfo.Name));
        Assert.That(list[0].Text, Is.EqualTo("Hello Shape"));
    }

    [Test]
    public void AddShape_WithFlowchartShapes_CreatesSpecificShapesAndReturnsCorrectTypeNames()
    {
        // Act & Assert flowchart shapes
        ShapeInfo processShape = _service!.AddShape(_wbName, SheetName, "FlowchartProcess", 10, 10, 100, 50, "Process Step");
        Assert.That(processShape.Type, Is.EqualTo("FlowchartProcess"));
        Assert.That(processShape.Text, Is.EqualTo("Process Step"));

        ShapeInfo decisionShape = _service.AddShape(_wbName, SheetName, "decision", 10, 80, 100, 50, "Is Decision?");
        Assert.That(decisionShape.Type, Is.EqualTo("FlowchartDecision"));

        ShapeInfo terminatorShape = _service.AddShape(_wbName, SheetName, "terminator", 10, 150, 100, 50, "End");
        Assert.That(terminatorShape.Type, Is.EqualTo("FlowchartTerminator"));

        // Check listing shapes
        var list = _service.ListShapes(_wbName, SheetName);
        // Should contain 3 shapes
        Assert.That(list, Has.Count.EqualTo(3));

        var foundProcess = list.FirstOrDefault(s => s.Name == processShape.Name);
        Assert.That(foundProcess, Is.Not.Null);
        Assert.That(foundProcess!.Type, Is.EqualTo("FlowchartProcess"));

        var foundDecision = list.FirstOrDefault(s => s.Name == decisionShape.Name);
        Assert.That(foundDecision, Is.Not.Null);
        Assert.That(foundDecision!.Type, Is.EqualTo("FlowchartDecision"));
    }

    [Test]
    public void AddShape_WithCalloutShapes_CreatesSpecificShapesAndReturnsCorrectTypeNames()
    {
        // Act & Assert callout (dialogue bubble) shapes
        ShapeInfo rectCallout = _service!.AddShape(_wbName, SheetName, "RectCallout", 10, 10, 100, 50, "Hello Rect");
        Assert.That(rectCallout.Type, Is.EqualTo("RectangularCallout"));
        Assert.That(rectCallout.Text, Is.EqualTo("Hello Rect"));

        ShapeInfo roundRectCallout = _service.AddShape(_wbName, SheetName, "roundrectcallout", 10, 80, 100, 50, "Hello RoundRect");
        Assert.That(roundRectCallout.Type, Is.EqualTo("RoundedRectangularCallout"));

        ShapeInfo ovalCallout = _service.AddShape(_wbName, SheetName, "OvalCallout", 10, 150, 100, 50, "Hello Oval");
        Assert.That(ovalCallout.Type, Is.EqualTo("OvalCallout"));

        ShapeInfo cloudCallout = _service.AddShape(_wbName, SheetName, "cloudcallout", 10, 220, 100, 50, "Hello Cloud");
        Assert.That(cloudCallout.Type, Is.EqualTo("CloudCallout"));

        // Check listing shapes
        var list = _service.ListShapes(_wbName, SheetName);
        Assert.That(list, Has.Count.EqualTo(4));

        var foundRect = list.FirstOrDefault(s => s.Name == rectCallout.Name);
        Assert.That(foundRect, Is.Not.Null);
        Assert.That(foundRect!.Type, Is.EqualTo("RectangularCallout"));

        var foundRoundRect = list.FirstOrDefault(s => s.Name == roundRectCallout.Name);
        Assert.That(foundRoundRect, Is.Not.Null);
        Assert.That(foundRoundRect!.Type, Is.EqualTo("RoundedRectangularCallout"));
    }

    [Test]
    public void UpdateShape_WithNewProperties_UpdatesShapeSuccessfully()
    {
        // Arrange
        ShapeInfo initialShape = _service!.AddShape(_wbName, SheetName, "Oval", 10, 10, 50, 50, "Initial Text");

        // Act
        ShapeInfo updatedShape = _service.UpdateShape(_wbName, SheetName, initialShape.Name, 15, 25, 60, 70, "Updated Text");

        // Assert
        Assert.That(updatedShape, Is.Not.Null);
        Assert.That(updatedShape.Name, Is.EqualTo(initialShape.Name));
        Assert.That(updatedShape.Left, Is.EqualTo(15f).Within(0.01f));
        Assert.That(updatedShape.Top, Is.EqualTo(25f).Within(0.01f));
        Assert.That(updatedShape.Width, Is.EqualTo(60f).Within(0.01f));
        Assert.That(updatedShape.Height, Is.EqualTo(70f).Within(0.01f));
        Assert.That(updatedShape.Text, Is.EqualTo("Updated Text"));

        var list = _service.ListShapes(_wbName, SheetName);
        Assert.That(list, Has.Count.EqualTo(1));
        Assert.That(list[0].Text, Is.EqualTo("Updated Text"));
    }

    [Test]
    public void DeleteShape_WhenShapeExists_DeletesShapeSuccessfully()
    {
        // Arrange
        ShapeInfo initialShape = _service!.AddShape(_wbName, SheetName, "Rectangle", 10, 10, 50, 50, "To Delete");

        // Act
        _service.DeleteShape(_wbName, SheetName, initialShape.Name);

        // Assert
        var list = _service.ListShapes(_wbName, SheetName);
        Assert.That(list, Is.Empty);
    }

    [Test]
    public void UpdateShape_WhenShapeDoesNotExist_ThrowsException()
    {
        // Act & Assert
        var ex = Assert.Throws<Exception>(() =>
            _service!.UpdateShape(_wbName, SheetName, "NonExistentShape", 15, 25, 60, 70, "Updated Text"));

        Assert.That(ex.Message, Does.Contain("Shape 'NonExistentShape' not found"));
    }

    [Test]
    public void DeleteShape_WhenShapeDoesNotExist_ThrowsException()
    {
        // Act & Assert
        var ex = Assert.Throws<Exception>(() =>
            _service!.DeleteShape(_wbName, SheetName, "NonExistentShape"));

        Assert.That(ex.Message, Does.Contain("Shape 'NonExistentShape' not found"));
    }

    [Test]
    public void ListShapes_WhenFlowchartConnectorExists_ReturnsShapeInfoWithConnectionDetails()
    {
        // Arrange
        var app = s_provider!.GetApp(createNew: false);
        Assert.That(app, Is.Not.Null);
        Excel.Workbooks? wbs = null;
        Excel.Workbook? wb = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? ws = null;
        Excel.Shapes? shapes = null;
        Excel.Shape? shape1 = null;
        Excel.Shape? shape2 = null;
        Excel.Shape? connector = null;
        Excel.ConnectorFormat? connFormat = null;
        try
        {
            wbs = app.Workbooks;
            wb = wbs[_wbName];
            sheets = wb.Sheets;
            ws = (Excel.Worksheet)sheets[SheetName];
            shapes = ws.Shapes;

            // Create two shapes
            shape1 = shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 10, 10, 100, 50);
            shape2 = shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, 10, 150, 100, 50);

            // Create a connector
            connector = shapes.AddConnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight, 0, 0, 0, 0);

            // Connect them
            connFormat = connector.ConnectorFormat;
            connFormat.BeginConnect(shape1, 3); // Site 3
            connFormat.EndConnect(shape2, 1);   // Site 1

            // Act
            var list = _service!.ListShapes(_wbName, SheetName);

            // Assert
            Assert.That(list, Has.Count.EqualTo(3));

            var connInfo = list.FirstOrDefault(s => s.Name == connector.Name);
            Assert.That(connInfo, Is.Not.Null);
            Assert.That(connInfo!.Connection, Is.Not.Null);
            Assert.That(connInfo.Connection!.BeginShapeName, Is.EqualTo(shape1.Name));
            Assert.That(connInfo.Connection.BeginConnectionSite, Is.EqualTo(3));
            Assert.That(connInfo.Connection.EndShapeName, Is.EqualTo(shape2.Name));
            Assert.That(connInfo.Connection.EndConnectionSite, Is.EqualTo(1));
        }
        finally
        {
            ExcelConnector.SafeReleaseComObject(connFormat);
            ExcelConnector.SafeReleaseComObject(connector);
            ExcelConnector.SafeReleaseComObject(shape2);
            ExcelConnector.SafeReleaseComObject(shape1);
            ExcelConnector.SafeReleaseComObject(shapes);
            ExcelConnector.SafeReleaseComObject(ws);
            ExcelConnector.SafeReleaseComObject(sheets);
        }
    }
}
