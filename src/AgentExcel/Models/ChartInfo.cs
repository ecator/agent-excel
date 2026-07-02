namespace AgentExcel.Models;

/// <summary>
/// Information about a chart in a worksheet.
/// </summary>
/// <param name="Name">The unique name of the chart.</param>
/// <param name="Range">The data source range address of the chart.</param>
/// <param name="Type">The type of the chart.</param>
/// <param name="Title">The title text of the chart.</param>
public record ChartInfo(string Name, string Range, string Type, string Title);
