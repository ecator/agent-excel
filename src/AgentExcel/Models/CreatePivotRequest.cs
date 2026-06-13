namespace AgentExcel.Models;

/// <summary>
/// Request for creating a new PivotTable.
/// </summary>
public record CreatePivotRequest : WorkbookRequestBase
{
    /// <summary>
    /// The name of the worksheet containing the source data.
    /// </summary>
    public required string SourceSheet { get; init; }

    /// <summary>
    /// The range address of the source data (e.g., 'A1:D100').
    /// </summary>
    public required string SourceRange { get; init; }

    /// <summary>
    /// The name of the worksheet where the PivotTable will be placed.
    /// </summary>
    public required string TargetSheet { get; init; }

    /// <summary>
    /// The starting cell address for the PivotTable (e.g., 'A3').
    /// </summary>
    public required string TargetCell { get; init; }

    /// <summary>
    /// The name of the PivotTable to be created.
    /// </summary>
    public required string TableName { get; init; }
}
