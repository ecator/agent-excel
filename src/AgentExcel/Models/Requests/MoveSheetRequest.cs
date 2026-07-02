namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for moving a worksheet.
/// </summary>
public record MoveSheetRequest : WorksheetRequestBase
{
    /// <summary>
    /// The target workbook name or path (optional).
    /// If null or empty, the source workbook is used.
    /// </summary>
    public string? TargetWorkbook { get; init; }

    /// <summary>
    /// The position to place the moved worksheet (optional).
    /// 0 means the first position (before the first sheet), -1 means the last position (after the last sheet).
    /// If not specified, defaults to the last position (-1).
    /// </summary>
    public int? Position { get; init; }
}
