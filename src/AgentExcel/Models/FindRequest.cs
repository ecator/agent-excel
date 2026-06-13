namespace AgentExcel.Models;

/// <summary>
/// Request for finding text inside a worksheet or range.
/// </summary>
public record FindRequest : WorksheetRequestBase
{
    /// <summary>
    /// Optional specific range to search. If null, searches the entire sheet.
    /// </summary>
    public string? RangeAddress { get; init; }

    /// <summary>
    /// The text string to search for.
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
