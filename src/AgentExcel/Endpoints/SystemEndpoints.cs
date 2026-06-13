using System.Diagnostics;

using AgentExcel.Models;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;
using Microsoft.Extensions.Hosting;

namespace AgentExcel.Endpoints;

public static class SystemEndpoints
{
    public static void MapSystemEndpoints(this IEndpointRouteBuilder app, int port, string listenHost)
    {
        app.MapGet("/status", (SystemService systemService) =>
        {
            string activeWorkbook = systemService.GetActiveWorkbookName();

            return Results.Extensions.Yaml(new
            {
                status = "running",
                pid = Process.GetCurrentProcess().Id,
                port = port,
                active_workbook = activeWorkbook,
                swagger_url = $"http://{listenHost}:{port}/swagger/index.html",
                openapi_url = $"http://{listenHost}:{port}/swagger/v1/swagger.json"
            });
        })
        .WithTags("System")
        .WithOpenApi(operation => { operation.Summary = "Get server and Excel status"; return operation; });

        app.MapPost("/exit", (IHostApplicationLifetime lifetime, SystemService systemService) =>
        {
            systemService.Shutdown();
            lifetime.StopApplication();
            return Results.Extensions.Yaml(new { status = "shutting_down" });
        })
        .WithTags("System")
        .WithOpenApi(operation => { operation.Summary = "Shutdown the daemon server"; return operation; });

        var system = app.MapGroup("/system").WithOpenApi();

        system.MapPost("/calculate", (SystemService systemService) =>
        {
            systemService.Calculate();
            return Results.Extensions.Yaml(new { status = "calculated" });
        })
        .WithTags("System")
        .WithSummary("Trigger a global calculation");

        var pq = app.MapGroup("/queries").WithOpenApi();

        pq.MapPost("/list", (WorkbookRequest req, SystemService systemService) =>
            Results.Extensions.Yaml(systemService.ListQueries(req.Workbook)))
            .WithTags("Queries")
            .WithSummary("List all Power Queries in a workbook");

        pq.MapPost("/add-update", (AddQueryRequest req, SystemService systemService) =>
        {
            systemService.AddOrUpdateQuery(req.Workbook, req.Name, req.Formula, req.Description);
            return Results.Extensions.Yaml(new { status = "success" });
        })
        .WithTags("Queries")
        .WithSummary("Add or update a Power Query (M Language)");

        pq.MapPost("/delete", (DeleteQueryRequest req, SystemService systemService) =>
        {
            systemService.DeleteQuery(req.Workbook, req.Name);
            return Results.Extensions.Yaml(new { status = "deleted" });
        })
        .WithTags("Queries")
        .WithSummary("Delete a Power Query");

        var conn = app.MapGroup("/connections").WithOpenApi();

        conn.MapPost("/refresh-all", (WorkbookRequest req, SystemService systemService) =>
        {
            systemService.RefreshAllDataConnections(req.Workbook);
            return Results.Extensions.Yaml(new { status = "refreshing" });
        })
        .WithTags("Connections")
        .WithSummary("Refresh all data connections, Power Queries, and Data Model");

        conn.MapPost("/refresh-model", (WorkbookRequest req, SystemService systemService) =>
        {
            systemService.RefreshModel(req.Workbook);
            return Results.Extensions.Yaml(new { status = "refreshing_model" });
        })
        .WithTags("Connections")
        .WithSummary("Refresh the Excel Data Model (Power Pivot)");
    }
}
