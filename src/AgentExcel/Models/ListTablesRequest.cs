namespace AgentExcel.Models;

/// <summary>
/// Request for listing Excel Tables (ListObjects) in a workbook or worksheet.
/// </summary>
public record ListTablesRequest : WorkbookRequestBase
{
    /// <summary>
    /// The target worksheet name (optional). If omitted, searches the entire workbook.
    /// </summary>
    public string? Sheet { get; init; }
}
