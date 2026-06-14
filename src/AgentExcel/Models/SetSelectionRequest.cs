namespace AgentExcel.Models;

/// <summary>
/// Request for setting the selection in a worksheet.
/// </summary>
public record SetSelectionRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target range address (e.g., 'A1:C10' or 'A1') to select.
    /// </summary>
    public required string Range { get; init; }
}
