using System.Diagnostics;

using AgentExcel.Models;
using AgentExcel.Models.Requests;
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
            try
            {
                systemService.Shutdown();
            }
            finally
            {
                lifetime.StopApplication();
            }
            return Results.Text("stopped");
        })
        .WithTags("System")
        .WithOpenApi(operation => { operation.Summary = "Shutdown the daemon server"; return operation; });

        app.MapPost("/calculate", (SystemService systemService) =>
        {
            systemService.Calculate();
            return Results.Text("calculated");
        })
        .WithTags("System")
        .WithOpenApi(operation => { operation.Summary = "Trigger a global calculation"; return operation; });
    }
}
