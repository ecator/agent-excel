namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for reading the content of a specific range in a worksheet.
/// </summary>
public record ReadRangeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The range address (e.g., 'A1:C10' or 'A1'). If null or empty, the used range is read.
    /// </summary>
    public string? Range { get; init; }
}
