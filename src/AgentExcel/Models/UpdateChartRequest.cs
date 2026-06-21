namespace AgentExcel.Models;

/// <summary>
/// Request for updating an existing chart's properties.
/// </summary>
public record UpdateChartRequest : WorksheetRequestBase
{
    /// <summary>
    /// The name of the chart to update.
    /// </summary>
    public required string ChartName { get; init; }

    /// <summary>
    /// Optional new data source range address.
    /// </summary>
    public string? Range { get; init; }

    /// <summary>
    /// Optional new chart type. Supported values are: 'column', 'line', 'pie', 'bar'.
    /// </summary>
    public string? ChartType { get; init; }

    /// <summary>
    /// Optional new title.
    /// </summary>
    public string? Title { get; init; }
}
