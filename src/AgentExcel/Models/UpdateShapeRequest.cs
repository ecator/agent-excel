namespace AgentExcel.Models;

/// <summary>
/// Request for updating an existing shape. Only non-null properties will be updated; null properties will keep their current values.
/// </summary>
public record UpdateShapeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The name of the shape to update.
    /// </summary>
    public required string ShapeName { get; init; }

    /// <summary>
    /// Optional new horizontal position in points. If null, the current value is kept.
    /// </summary>
    public float? Left { get; init; }

    /// <summary>
    /// Optional new vertical position in points. If null, the current value is kept.
    /// </summary>
    public float? Top { get; init; }

    /// <summary>
    /// Optional new width in points. If null, the current value is kept.
    /// </summary>
    public float? Width { get; init; }

    /// <summary>
    /// Optional new height in points. If null, the current value is kept.
    /// </summary>
    public float? Height { get; init; }

    /// <summary>
    /// Optional new text content for the shape. If null, the current text is kept.
    /// </summary>
    public string? Text { get; init; }
}
