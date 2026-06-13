using System.Runtime.InteropServices;

using AgentExcel.Models;
using AgentExcel.Providers;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Services;

public class ShapeService : ExcelServiceBase
{
    public ShapeService(IExcelConnectionProvider connectionProvider) : base(connectionProvider)
    {
    }

    public List<ShapeInfo> ListShapes(string workbookName, string sheetName)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Shapes? shapes = null;
            var result = new List<ShapeInfo>();
            try
            {
                wb = GetWorkbook(workbookName);
                ws = GetWorksheet(wb, sheetName);
                shapes = ws.Shapes;

                foreach (Excel.Shape shape in shapes)
                {
                    string? text = null;
                    try
                    {
                        var tf = shape.TextFrame;
                        var chars = tf.Characters();
                        text = chars.Text;
                        Marshal.ReleaseComObject(chars);
                        Marshal.ReleaseComObject(tf);
                    }
                    catch
                    {
                        // Ignore shapes that don't support text frames
                    }

                    result.Add(new ShapeInfo(
                        shape.Name,
                        shape.Type.ToString(),
                        (float)shape.Left,
                        (float)shape.Top,
                        (float)shape.Width,
                        (float)shape.Height,
                        text
                    ));
                    Marshal.ReleaseComObject(shape);
                }
                return result;
            }
            finally
            {
                if (shapes != null) Marshal.ReleaseComObject(shapes);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        });
    }

    public string AddShape(string workbookName, string sheetName, string type, float left, float top, float width, float height, string? text)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                shapes = ws.Shapes;

                if (type.Equals("Textbox", StringComparison.OrdinalIgnoreCase))
                {
                    shape = shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
                }
                else
                {
                    var autoType = type.ToLower() switch
                    {
                        "rectangle" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                        "oval" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
                        "arrow" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRightArrow,
                        _ => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle
                    };
                    shape = shapes.AddShape(autoType, left, top, width, height);
                }

                if (text != null)
                {
                    var tf = shape.TextFrame;
                    var chars = tf.Characters();
                    chars.Text = text;
                    Marshal.ReleaseComObject(chars);
                    Marshal.ReleaseComObject(tf);
                }

                return shape.Name;
            }
            finally
            {
                if (shape != null) Marshal.ReleaseComObject(shape);
                if (shapes != null) Marshal.ReleaseComObject(shapes);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        });
    }

    public void UpdateShape(string workbookName, string sheetName, string shapeName, float? left, float? top, float? width, float? height, string? text)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                shapes = ws.Shapes;
                shape = shapes.Item(shapeName);

                if (left.HasValue) shape.Left = left.Value;
                if (top.HasValue) shape.Top = top.Value;
                if (width.HasValue) shape.Width = width.Value;
                if (height.HasValue) shape.Height = height.Value;

                if (text != null)
                {
                    var tf = shape.TextFrame;
                    var chars = tf.Characters();
                    chars.Text = text;
                    Marshal.ReleaseComObject(chars);
                    Marshal.ReleaseComObject(tf);
                }
            }
            finally
            {
                if (shape != null) Marshal.ReleaseComObject(shape);
                if (shapes != null) Marshal.ReleaseComObject(shapes);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        });
    }

    public void DeleteShape(string workbookName, string sheetName, string shapeName)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                shapes = ws.Shapes;
                shape = shapes.Item(shapeName);
                shape.Delete();
            }
            finally
            {
                if (shape != null) Marshal.ReleaseComObject(shape);
                if (shapes != null) Marshal.ReleaseComObject(shapes);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        });
    }
}
