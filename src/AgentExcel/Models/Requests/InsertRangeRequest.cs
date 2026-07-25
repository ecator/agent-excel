namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for inserting cells, rows, or columns into a range in a worksheet.
/// </summary>
public record InsertRangeRequest : WorksheetRequestBase
{
    /// <summary>
    /// The range address (e.g., 'A1:C10' or 'A1'). If null or empty, defaults to used range.
    /// </summary>
    public string? Range { get; init; }

    /// <summary>
    /// Optional insert shift option: 'shift_right', 'shift_down', 'entire_row', 'entire_column'.
    /// Also supports synonyms: 'right', 'down', 'row', 'column', 'entireRow', 'entireColumn', 'shiftRight', 'shiftDown'.
    /// </summary>
    public string? Shift { get; init; }
}
