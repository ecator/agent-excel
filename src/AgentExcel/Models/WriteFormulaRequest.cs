namespace AgentExcel.Models;

/// <summary>
/// Request for writing a formula to a cell or range in a worksheet.
/// </summary>
public record WriteFormulaRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target cell or range address.
    /// </summary>
    public required string Range { get; init; }

    /// <summary>
    /// The Excel formula to write. Can be a single formula string, a 1D array of formulas, or a 2D array of formulas.
    /// Each formula must start with '='.
    /// </summary>
    public required object Formula { get; init; }
}
