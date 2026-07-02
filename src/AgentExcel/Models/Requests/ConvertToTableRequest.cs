namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for converting a range of cells into an Excel Table (ListObject).
/// </summary>
public record ConvertToTableRequest : WorksheetRequestBase
{
    /// <summary>
    /// The range address containing the table data (e.g., 'A1:D10').
    /// </summary>
    public required string Range { get; init; }

    /// <summary>
    /// Optional name for the new table. If null, a default name is generated.
    /// </summary>
    public string? Table { get; init; }

    /// <summary>
    /// True if the first row of the range contains column headers.
    /// </summary>
    public required bool HasHeaders { get; init; }
}
