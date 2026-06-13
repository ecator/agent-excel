namespace AgentExcel.Models;

/// <summary>
/// Request for opening an existing workbook.
/// </summary>
public record OpenWorkbookRequest
{
    /// <summary>
    /// The absolute file path of the workbook to open.
    /// </summary>
    public required string Path { get; init; }
}
