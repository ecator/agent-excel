namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for auto-fitting row height or column width (or both) of a range.
/// </summary>
public record AutoFitRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target range address (e.g., 'A1:C10'). Required.
    /// </summary>
    public required string Range { get; init; }

    /// <summary>
    /// The target dimension to auto-fit: 'columns' (or 'column', 'width'), 'rows' (or 'row', 'height'), or 'both'. Defaults to 'columns'.
    /// </summary>
    public string? Target { get; init; }
}
