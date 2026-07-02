namespace AgentExcel.Models;

/// <summary>
/// Represents a match from a find operation in a worksheet.
/// </summary>
/// <param name="Sheet">The sheet name where the match was found.</param>
/// <param name="Address">The cell address (e.g., 'A1').</param>
/// <param name="Value">The content of the matching cell.</param>
public record FindResult(string Sheet, string Address, string Value);
