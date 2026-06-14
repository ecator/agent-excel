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
    }
}
