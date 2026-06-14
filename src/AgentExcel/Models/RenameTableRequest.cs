namespace AgentExcel.Models;

/// <summary>
/// Request for renaming an Excel Table.
/// </summary>
public record RenameTableRequest : WorksheetRequestBase
{
    /// <summary>
    /// The current name of the Excel Table.
    /// </summary>
    public required string Table { get; init; }

    /// <summary>
    /// The new name of the Excel Table.
    /// </summary>
    public required string NewName { get; init; }
}
