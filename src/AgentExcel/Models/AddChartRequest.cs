namespace AgentExcel.Models;

/// <summary>
/// Request for adding a new chart to a worksheet.
/// </summary>
public record AddChartRequest : WorksheetRequestBase
{
    /// <summary>
    /// The data source range address for the chart (e.g., 'A1:B10').
    /// </summary>
    public required string RangeAddress { get; init; }

    /// <summary>
    /// The type of the chart (e.g., 'xlColumnClustered', 'xlLine', 'xlPie').
    /// </summary>
    public required string ChartType { get; init; }

    /// <summary>
    /// The title of the chart.
    /// </summary>
    public required string Title { get; init; }
}
