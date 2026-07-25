namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for merging a range.
/// </summary>
public record MergeRangeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target range address (e.g., 'A1:C10'). Required.
    /// </summary>
    public required string Range { get; init; }

    /// <summary>
    /// If true, cells in each row of the specified range are merged across separately. Defaults to false.
    /// </summary>
    public bool? Across { get; init; }
}
