namespace AgentExcel.Models;

/// <summary>
/// Request for reading the content of a specific range in a worksheet.
/// </summary>
public record ReadRangeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The range address (e.g., 'A1:C10' or 'A1').
    /// </summary>
    public required string Address { get; init; }
}
