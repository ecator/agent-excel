namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for setting column width of a range.
/// </summary>
public record SetWidthRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target range address (e.g., 'A1:C10'). Required.
    /// </summary>
    public required string Range { get; init; }

    /// <summary>
    /// The column width to set. Required.
    /// </summary>
    public required double Width { get; init; }
}
