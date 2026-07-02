namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for converting an Excel Table back to a normal range of cells.
/// </summary>
public record ConvertToRangeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The name of the table to convert.
    /// </summary>
    public required string Table { get; init; }
}
