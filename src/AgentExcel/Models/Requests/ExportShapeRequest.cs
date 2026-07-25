namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for exporting a specific shape (e.g., autoshapes, pictures, textboxes, connectors, or charts) in a worksheet as an image file.
/// </summary>
public record ExportShapeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The name of the shape, picture, or textbox to export as an image.
    /// </summary>
    public required string ShapeName { get; init; }

    /// <summary>
    /// The absolute path where the image file will be saved.
    /// </summary>
    public required string OutputFile { get; init; }
}
