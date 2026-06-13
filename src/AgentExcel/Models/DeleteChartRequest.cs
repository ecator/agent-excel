namespace AgentExcel.Models;

/// <summary>
/// Request for deleting an existing chart.
/// </summary>
public record DeleteChartRequest : WorksheetRequestBase
{
    /// <summary>
    /// The name of the chart to delete.
    /// </summary>
    public required string ChartName { get; init; }
}
