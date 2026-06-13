namespace AgentExcel.Models;

/// <summary>
/// Base request model for workbook operations.
/// </summary>
public abstract record WorkbookRequestBase
{
    /// <summary>
    /// The target workbook name or path (e.g., 'Book1').
    /// </summary>
    public required string Workbook { get; init; }
}
