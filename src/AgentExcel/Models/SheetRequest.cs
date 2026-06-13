namespace AgentExcel.Models;

/// <summary>
/// Request for worksheet operations (e.g., adding or deleting).
/// </summary>
public record SheetRequest : WorkbookRequestBase
{
    /// <summary>
    /// The name of the worksheet.
    /// </summary>
    public required string Name { get; init; }
}
