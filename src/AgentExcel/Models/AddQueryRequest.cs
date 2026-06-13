namespace AgentExcel.Models;

/// <summary>
/// Request for adding a Power Query connection.
/// </summary>
public record AddQueryRequest : WorkbookRequestBase
{
    /// <summary>
    /// The unique name of the query.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// The M formula of the query.
    /// </summary>
    public required string Formula { get; init; }

    /// <summary>
    /// Optional description of the query.
    /// </summary>
    public string? Description { get; init; }
}
