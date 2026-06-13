using AgentExcel.Models;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;

namespace AgentExcel.Endpoints;

public static class PivotTableEndpoints
{
    public static void MapPivotTableEndpoints(this IEndpointRouteBuilder app)
    {
        var pivot = app.MapGroup("/pivot").WithOpenApi();

        pivot.MapPost("/create", (CreatePivotRequest req, PivotTableService pivotTableService) =>
        {
            pivotTableService.CreatePivotTable(req.Workbook, req.SourceSheet, req.SourceRange, req.TargetSheet, req.TargetCell, req.TableName);
            return Results.Extensions.Yaml(new { status = "created" });
        })
        .WithTags("Pivot Tables")
        .WithSummary("Create a pivot table");
    }
}
