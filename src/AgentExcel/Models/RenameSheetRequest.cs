namespace AgentExcel.Models;

/// <summary>
/// Request for renaming a worksheet.
/// </summary>
public record RenameSheetRequest : WorkbookRequestBase
{
    /// <summary>
    /// The current name of the worksheet to rename.
    /// </summary>
    public required string OldName { get; init; }

    /// <summary>
    /// The new name of the worksheet.
    /// </summary>
    public required string NewName { get; init; }
}
