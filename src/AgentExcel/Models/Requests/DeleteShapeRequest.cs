namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for deleting an existing shape.
/// </summary>
public record DeleteShapeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The name of the shape to delete.
    /// </summary>
    public required string ShapeName { get; init; }
}
