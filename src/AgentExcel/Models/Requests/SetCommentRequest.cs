namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for setting a comment on a single cell range.
/// </summary>
public record SetCommentRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target cell address (e.g., 'A1'). Must be a single cell. Required.
    /// </summary>
    public required string Range { get; init; }

    /// <summary>
    /// The text content of the comment.
    /// </summary>
    public required string Text { get; init; }

    /// <summary>
    /// Optional visibility flag for the comment.
    /// </summary>
    public bool? Visible { get; init; }
}
