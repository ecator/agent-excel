namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for closing a workbook.
/// </summary>
public record CloseWorkbookRequest : WorkbookRequestBase
{
    /// <summary>
    /// Whether to save changes before closing.
    /// </summary>
    public required bool SaveChanges { get; init; }
}
