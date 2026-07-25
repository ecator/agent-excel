namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for unmerging a range.
/// </summary>
public record UnmergeRangeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target range address (e.g., 'A1:C10'). Required.
    /// </summary>
    public required string Range { get; init; }
}
