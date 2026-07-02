namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for reading the content of an Excel Table (ListObject) by name.
/// </summary>
public record ReadTableRequest : WorksheetRequestBase
{
    /// <summary>
    /// The name of the Excel Table.
    /// </summary>
    public required string Table { get; init; }
}
