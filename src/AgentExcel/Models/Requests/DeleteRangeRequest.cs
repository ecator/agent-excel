namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for deleting a range in a worksheet.
/// </summary>
public record DeleteRangeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The range address (e.g., 'A1:C10' or 'A1'). If null or empty, the used range is deleted.
    /// </summary>
    public string? Range { get; init; }

    /// <summary>
    /// Optional delete shift option: 'shift_left', 'shift_up', 'entire_row', 'entire_column'.
    /// Also supports synonyms: 'left', 'up', 'row', 'column', 'entireRow', 'entireColumn', 'shiftLeft', 'shiftUp'.
    /// </summary>
    public string? Shift { get; init; }
}
