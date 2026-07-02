namespace AgentExcel.Models.Requests;

/// <summary>
/// Request for adding a new chart to a worksheet.
/// </summary>
public record AddChartRequest : WorksheetRequestBase
{
    /// <summary>
    /// The data source range address for the chart (e.g., 'A1:B10').
    /// </summary>
    public required string Range { get; init; }

    /// <summary>
    /// The type of the chart. Supported values are: 'column', 'line', 'pie', 'bar'.
    /// </summary>
    public required string ChartType { get; init; }

    /// <summary>
    /// The title of the chart.
    /// </summary>
    public required string Title { get; init; }
}
