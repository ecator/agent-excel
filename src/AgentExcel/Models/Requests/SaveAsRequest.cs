namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for saving a workbook as a different file.
/// </summary>
public record SaveAsRequest : WorkbookRequestBase
{
    /// <summary>
    /// The absolute file path where the workbook will be saved.
    /// </summary>
    public required string Path { get; init; }
}
