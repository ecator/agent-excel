namespace AgentExcel.Models;

/// <summary>
/// Request for finding text inside a worksheet or range.
/// </summary>
public record FindRequest : WorkbookRequestBase
{
    /// <summary>
    /// Optional target worksheet name (e.g., 'Sheet1'). If null, searches the entire workbook.
    /// </summary>
    public string? Sheet { get; init; }

    /// <summary>
    /// Optional specific range to search. If null, searches the entire sheet/workbook.
    /// </summary>
    public string? Range { get; init; }

    /// <summary>
    /// The text string to search for. Supports Excel wildcard characters (e.g., '*' for multiple characters, '?' for a single character, and '~' to escape them).
    /// </summary>
    public required string What { get; init; }

    /// <summary>
    /// True to perform a case-sensitive search.
    /// </summary>
    public required bool MatchCase { get; init; }

    /// <summary>
    /// True to match the entire cell content.
    /// </summary>
    public required bool WholeWord { get; init; }
}
