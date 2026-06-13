namespace AgentExcel.Models;

/// <summary>
/// Request for deleting a Power Query connection.
/// </summary>
public record DeleteQueryRequest : WorkbookRequestBase
{
    /// <summary>
    /// The name of the query to delete.
    /// </summary>
    public required string Name { get; init; }
}
