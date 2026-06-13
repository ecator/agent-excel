using AgentExcel.Models;
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
            Results.Extensions.Yaml(new { name = chartService.AddChart(req.Workbook, req.Sheet, req.RangeAddress, req.ChartType, req.Title) }))
            .WithTags("Charts")
            .WithSummary("Add a chart to a sheet");

        charts.MapPost("/update", (UpdateChartRequest req, ChartService chartService) =>
        {
            chartService.UpdateChart(req.Workbook, req.Sheet, req.ChartName, req.RangeAddress, req.ChartType, req.Title);
            return Results.Extensions.Yaml(new { status = "updated" });
        })
        .WithTags("Charts")
        .WithSummary("Update chart properties (source data, type, title)");

        charts.MapPost("/delete", (DeleteChartRequest req, ChartService chartService) =>
        {
            chartService.DeleteChart(req.Workbook, req.Sheet, req.ChartName);
            return Results.Extensions.Yaml(new { status = "deleted" });
        })
        .WithTags("Charts")
        .WithSummary("Delete a chart");
    }
}
