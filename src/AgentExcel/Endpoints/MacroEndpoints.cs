using AgentExcel.Models;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;

namespace AgentExcel.Endpoints;

public static class MacroEndpoints
{
    public static void MapMacroEndpoints(this IEndpointRouteBuilder app)
    {
        var macros = app.MapGroup("/macros").WithOpenApi();

        macros.MapPost("/run", (RunMacroRequest req, MacroService macroService) =>
        {
            var result = macroService.RunMacro(req.Workbook, req.Macro, req.Args);
            return Results.Text($"result of {req.Macro}:\n{result}", "text/plain; charset=utf-8");
        })
            .WithTags("Macros")
            .WithSummary("Run a VBA macro");

        macros.MapPost("/list", (WorkbookRequest req, MacroService macroService) =>
            Results.Extensions.Yaml(macroService.ListMacros(req.Workbook)))
            .WithTags("Macros")
            .WithSummary("List all macros and their parameters in the workbook");
    }
}
