namespace AgentExcel.Models;

/// <summary>
/// Request for formatting a range of cells.
/// </summary>
public record SetStyleRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target range address.
    /// </summary>
    public required string Address { get; init; }

    /// <summary>
    /// The styling options to apply to the range.
    /// </summary>
    public required CellStyle Style { get; init; }
}
