namespace AgentExcel.Models;

/// <summary>
/// Request for writing values to a range in a worksheet.
/// </summary>
public record RangeWriteRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target range (e.g., 'A1:B2').
    /// </summary>
    public required string Range { get; init; }

    /// <summary>
    /// The value to write. Can be a single value, a 1D array, or a 2D array.
    /// </summary>
    public required object Value { get; init; }
}
