namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for setting row height of a range.
/// </summary>
public record SetHeightRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target range address (e.g., 'A1:C10'). Required.
    /// </summary>
    public required string Range { get; init; }

    /// <summary>
    /// The row height to set. Required.
    /// </summary>
    public required double Height { get; init; }
}
