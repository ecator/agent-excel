namespace AgentExcel.Models;

/// <summary>
/// Request for updating an existing shape.
/// </summary>
public record UpdateShapeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The name of the shape to update.
    /// </summary>
    public required string ShapeName { get; init; }

    /// <summary>
    /// Optional new horizontal position in points.
    /// </summary>
    public float? Left { get; init; }

    /// <summary>
    /// Optional new vertical position in points.
    /// </summary>
    public float? Top { get; init; }

    /// <summary>
    /// Optional new width in points.
    /// </summary>
    public float? Width { get; init; }

    /// <summary>
    /// Optional new height in points.
    /// </summary>
    public float? Height { get; init; }

    /// <summary>
    /// Optional new text content for the shape.
    /// </summary>
    public string? Text { get; init; }
}
