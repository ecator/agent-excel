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
                        SafeReleaseComObject(chars);
                        SafeReleaseComObject(tf);
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
                    SafeReleaseComObject(shape);
                }
                return result;
            }
            finally
            {
                SafeReleaseComObject(shapes);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
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
                    SafeReleaseComObject(chars);
                    SafeReleaseComObject(tf);
                }

                return shape.Name;
            }
            finally
            {
                SafeReleaseComObject(shape);
                SafeReleaseComObject(shapes);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
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
                    SafeReleaseComObject(chars);
                    SafeReleaseComObject(tf);
                }
            }
            finally
            {
                SafeReleaseComObject(shape);
                SafeReleaseComObject(shapes);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
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
                SafeReleaseComObject(shape);
                SafeReleaseComObject(shapes);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }
}
