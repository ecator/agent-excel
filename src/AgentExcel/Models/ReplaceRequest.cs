namespace AgentExcel.Models;

/// <summary>
/// Request for replacing text inside a worksheet or range.
/// </summary>
public record ReplaceRequest : WorksheetRequestBase
{
    /// <summary>
    /// Optional specific range to replace within. If null, replaces across the entire sheet.
    /// </summary>
    public string? RangeAddress { get; init; }

    /// <summary>
    /// The text string to search for.
    /// </summary>
    public required string What { get; init; }

    /// <summary>
    /// The replacement text string.
    /// </summary>
    public required string Replacement { get; init; }

    /// <summary>
    /// True to perform a case-sensitive search.
    /// </summary>
    public required bool MatchCase { get; init; }

    /// <summary>
    /// True to match the entire cell content.
    /// </summary>
    public required bool WholeWord { get; init; }
}
