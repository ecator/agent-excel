using AgentExcel.Models;
using AgentExcel.Models.Requests;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;

namespace AgentExcel.Endpoints;

public static class ChartEndpoints
{
    public static void MapChartEndpoints(this IEndpointRouteBuilder app)
    {
        var charts = app.MapGroup("/charts").WithOpenApi();

        charts.MapPost("/list", (WorksheetRequest req, ChartService chartService) =>
            Results.Extensions.Yaml(chartService.ListCharts(req.Workbook, req.Sheet)))
            .WithTags("Charts")
            .WithSummary("List all charts in a sheet");

        charts.MapPost("/add", (AddChartRequest req, ChartService chartService) =>
            Results.Extensions.Yaml(chartService.AddChart(req.Workbook, req.Sheet, req.Range, req.ChartType, req.Title)))
            .WithTags("Charts")
            .WithSummary("Add a chart to a sheet");

        charts.MapPost("/update", (UpdateChartRequest req, ChartService chartService) =>
        {
            var updatedInfo = chartService.UpdateChart(req.Workbook, req.Sheet, req.ChartName, req.Range, req.ChartType, req.Title);
            var details = new List<string>();
            if (!string.IsNullOrEmpty(req.Range)) details.Add($"range updated to '{updatedInfo.Range}'");
            if (!string.IsNullOrEmpty(req.ChartType)) details.Add($"type updated to '{updatedInfo.Type}'");
            if (req.Title != null) details.Add($"title updated to '{updatedInfo.Title}'");

            string message = details.Count == 0
                ? $"Chart '{req.ChartName}' was not modified."
                : $"Chart '{req.ChartName}' updated:\n{string.Join("\n", details)}.";

            return Results.Text(message, "text/plain; charset=utf-8");
        })
        .WithTags("Charts")
        .WithSummary("Update chart properties (source data, type, title)");

        charts.MapPost("/delete", (DeleteChartRequest req, ChartService chartService) =>
        {
            chartService.DeleteChart(req.Workbook, req.Sheet, req.ChartName);
            string message = $"chart[{req.ChartName}] has deleted";
            return Results.Text(message, "text/plain; charset=utf-8");
        })
        .WithTags("Charts")
        .WithSummary("Delete a chart");
    }
}
