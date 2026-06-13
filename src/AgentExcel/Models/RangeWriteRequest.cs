namespace AgentExcel.Models;

/// <summary>
/// Request for writing values to a range in a worksheet.
/// </summary>
public record RangeWriteRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target range address (e.g., 'A1:B2').
    /// </summary>
    public required string Address { get; init; }

    /// <summary>
    /// The value to write. Can be a single value or an array of arrays representing rows.
    /// </summary>
    public required object Value { get; init; }
}
