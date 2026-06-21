namespace AgentExcel.Models;

/// <summary>
/// Request for adding a shape to a worksheet.
/// </summary>
public record AddShapeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The shape type or geometry name. Supported values include: 'Textbox', 'Rectangle', 'Oval', 'Arrow' (RightArrow),
    /// 'Diamond', 'Parallelogram', and flowchart shapes such as 'FlowchartProcess', 'FlowchartDecision', 'FlowchartData',
    /// 'FlowchartDocument', 'FlowchartPredefinedProcess', 'FlowchartPreparation', 'FlowchartTerminator', and any other MsoAutoShapeType (case-insensitive, defaults to 'Rectangle' if not recognized).
    /// </summary>
    public required string Type { get; init; }

    /// <summary>
    /// Horizontal position in points from the left side of the worksheet.
    /// </summary>
    public required float Left { get; init; }

    /// <summary>
    /// Vertical position in points from the top side of the worksheet.
    /// </summary>
    public required float Top { get; init; }

    /// <summary>
    /// Width of the shape in points.
    /// </summary>
    public required float Width { get; init; }

    /// <summary>
    /// Height of the shape in points.
    /// </summary>
    public required float Height { get; init; }

    /// <summary>
    /// Optional text content to place inside the shape.
    /// </summary>
    public string? Text { get; init; }
}
