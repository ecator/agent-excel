namespace AgentExcel.Models;

/// <summary>
/// Request for writing a formula to a cell or range in a worksheet.
/// </summary>
public record WriteFormulaRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target cell or range address.
    /// </summary>
    public required string Address { get; init; }

    /// <summary>
    /// The Excel formula to write (e.g., '=SUM(A1:A10)').
    /// </summary>
    public required string Formula { get; init; }
}
