namespace AgentExcel.Models;

/// <summary>
/// Request for clearing a range in a worksheet.
/// </summary>
public record ClearRequest : WorksheetRequestBase
{
    /// <summary>
    /// The range address (e.g., 'A1:C10' or 'A1'). If null or empty, the used range is cleared.
    /// </summary>
    public string? Range { get; init; }

    /// <summary>
    /// The clear type: 'all', 'formats', 'contents', 'comments', 'hyperlinks'. Defaults to 'all'.
    /// </summary>
    public string? Type { get; init; }
}
