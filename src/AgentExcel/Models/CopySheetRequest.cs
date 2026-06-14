namespace AgentExcel.Models;

/// <summary>
/// Request for copying a worksheet.
/// </summary>
public record CopySheetRequest : WorkbookRequestBase
{
    /// <summary>
    /// The current name of the worksheet to copy.
    /// </summary>
    public required string OldName { get; init; }

    /// <summary>
    /// The new name of the worksheet (optional).
    /// If null or empty, it defaults to OldName.
    /// </summary>
    public string? NewName { get; init; }

    /// <summary>
    /// The target workbook name or path (optional).
    /// If null or empty, the source workbook is used.
    /// </summary>
    public string? TargetWorkbook { get; init; }

    /// <summary>
    /// The position to place the copied worksheet (optional).
    /// 0 means the first position (before the first sheet), -1 means the last position (after the last sheet).
    /// If not specified, defaults to immediately after the source sheet (for same workbook) or last position (for different workbook).
    /// </summary>
    public int? Position { get; init; }
}
