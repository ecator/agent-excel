namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for exporting a specific range in a worksheet as an image file.
/// </summary>
public record ExportImageRequest : WorksheetRequestBase
{
    /// <summary>
    /// The range address to export as an image (e.g., 'A1:D10').
    /// </summary>
    public required string Range { get; init; }

    /// <summary>
    /// The absolute path where the image file will be saved.
    /// </summary>
    public required string OutputFile { get; init; }
}
