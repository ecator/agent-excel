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
                    result.Add(BuildShapeInfo(shape));
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

    public ShapeInfo AddShape(string workbookName, string sheetName, string type, float left, float top, float width, float height, string? text)
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
                    Microsoft.Office.Core.MsoAutoShapeType autoType = ParseAutoShapeType(type);
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

                return BuildShapeInfo(shape);
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

    public ShapeInfo UpdateShape(string workbookName, string sheetName, string shapeName, float? left, float? top, float? width, float? height, string? text)
    {
        return ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Shape? shape = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                shape = GetShape(ws, shapeName);

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

                return BuildShapeInfo(shape);
            }
            finally
            {
                SafeReleaseComObject(shape);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    private ShapeInfo BuildShapeInfo(Excel.Shape shape)
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

        ShapeConnectionInfo? connection = GetShapeConnectionInfo(shape);

        List<ShapeInfo>? children = null;
        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
        {
            Excel.GroupShapes? groupItems = null;
            try
            {
                groupItems = shape.GroupItems;
                int count = groupItems.Count;
                if (count > 0)
                {
                    children = new List<ShapeInfo>();
                    for (int i = 1; i <= count; i++)
                    {
                        Excel.Shape? childShape = null;
                        try
                        {
                            childShape = groupItems.Item(i);
                            children.Add(BuildShapeInfo(childShape));
                        }
                        finally
                        {
                            SafeReleaseComObject(childShape);
                        }
                    }
                }
            }
            catch
            {
                // Ignore shapes where GroupItems cannot be accessed
            }
            finally
            {
                SafeReleaseComObject(groupItems);
            }
        }

        return new ShapeInfo(
            shape.Name,
            GetShapeTypeName(shape),
            (float)shape.Left,
            (float)shape.Top,
            (float)shape.Width,
            (float)shape.Height,
            text,
            connection,
            children
        );
    }

    public void DeleteShape(string workbookName, string sheetName, string shapeName)
    {
        ExecuteWithRetry(() =>
        {
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.Shape? shape = null;
            try
            {
                wb = GetWorkbook(workbookName, createNew: true);
                ws = GetWorksheet(wb, sheetName);
                shape = GetShape(ws, shapeName);
                shape.Delete();
            }
            finally
            {
                SafeReleaseComObject(shape);
                SafeReleaseComObject(ws);
                SafeReleaseComObject(wb);
            }
        });
    }

    private Excel.Shape GetShape(Excel.Worksheet ws, string shapeName)
    {
        if (string.IsNullOrWhiteSpace(shapeName))
        {
            throw new ArgumentException("Shape name cannot be null or empty.", nameof(shapeName));
        }

        Excel.Shapes? shapes = null;
        try
        {
            shapes = ws.Shapes;
            try
            {
                return shapes.Item(shapeName);
            }
            catch (Exception ex)
            {
                throw new Exception($"Shape '{shapeName}' not found in worksheet '{ws.Name}'.", ex);
            }
        }
        finally
        {
            SafeReleaseComObject(shapes);
        }
    }

    private ShapeConnectionInfo? GetShapeConnectionInfo(Excel.Shape shape)
    {
        if (shape.Connector != Microsoft.Office.Core.MsoTriState.msoTrue)
        {
            return null;
        }

        Excel.ConnectorFormat? connFormat = null;
        Excel.Shape? beginShape = null;
        Excel.Shape? endShape = null;
        try
        {
            connFormat = shape.ConnectorFormat;
            string? beginShapeName = null;
            int? beginConnectionSite = null;
            if (connFormat.BeginConnected == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                beginShape = connFormat.BeginConnectedShape;
                beginShapeName = beginShape.Name;
                beginConnectionSite = connFormat.BeginConnectionSite;
            }

            string? endShapeName = null;
            int? endConnectionSite = null;
            if (connFormat.EndConnected == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                endShape = connFormat.EndConnectedShape;
                endShapeName = endShape.Name;
                endConnectionSite = connFormat.EndConnectionSite;
            }

            return new ShapeConnectionInfo(beginShapeName, beginConnectionSite, endShapeName, endConnectionSite);
        }
        catch
        {
            // Ignore connection reading errors
            return null;
        }
        finally
        {
            SafeReleaseComObject(endShape);
            SafeReleaseComObject(beginShape);
            SafeReleaseComObject(connFormat);
        }
    }

    private Microsoft.Office.Core.MsoAutoShapeType ParseAutoShapeType(string typeStr)
    {
        string trimmed = typeStr.Trim();

        // Try parsing directly (case-insensitive)
        if (Enum.TryParse<Microsoft.Office.Core.MsoAutoShapeType>(trimmed, true, out var result))
        {
            return result;
        }

        // Try prepending "msoShape" (case-insensitive)
        if (!trimmed.StartsWith("msoShape", StringComparison.OrdinalIgnoreCase))
        {
            string prefixed = "msoShape" + trimmed;
            if (Enum.TryParse<Microsoft.Office.Core.MsoAutoShapeType>(prefixed, true, out result))
            {
                return result;
            }
        }

        // Fallback switch mapping for common aliases / flowchart shapes
        return trimmed.ToLowerInvariant() switch
        {
            "rectangle" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
            "oval" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
            "arrow" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRightArrow,
            "diamond" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond,
            "parallelogram" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeParallelogram,

            // Flowchart aliases
            "process" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeFlowchartProcess,
            "decision" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeFlowchartDecision,
            "data" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeFlowchartData,
            "document" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeFlowchartDocument,
            "predefinedprocess" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeFlowchartPredefinedProcess,
            "preparation" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeFlowchartPreparation,
            "terminator" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeFlowchartTerminator,

            // Callout aliases (dialogue bubbles)
            "rectangularcallout" or "rectcallout" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangularCallout,
            "roundedrectangularcallout" or "roundrectcallout" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangularCallout,
            "ovalcallout" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeOvalCallout,
            "cloudcallout" => Microsoft.Office.Core.MsoAutoShapeType.msoShapeCloudCallout,

            _ => Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle
        };
    }

    private string GetShapeTypeName(Excel.Shape shape)
    {
        var type = shape.Type;
        if (type == Microsoft.Office.Core.MsoShapeType.msoAutoShape)
        {
            try
            {
                var autoType = shape.AutoShapeType;
                string name = autoType.ToString();
                if (name.StartsWith("msoShape"))
                {
                    return name.Substring("msoShape".Length);
                }
                return name;
            }
            catch
            {
                // Fallback to general type if AutoShapeType read fails
                string typeName = type.ToString();
                if (typeName.StartsWith("mso"))
                {
                    return typeName.Substring(3);
                }
                return typeName;
            }
        }
        else if (type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
        {
            return "TextBox";
        }

        // For other types, return the enum name with "mso" prefix removed if present
        string generalTypeName = type.ToString();
        if (generalTypeName.StartsWith("mso"))
        {
            return generalTypeName.Substring(3);
        }
        return generalTypeName;
    }
}
