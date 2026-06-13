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
            Results.Extensions.Yaml(new { result = macroService.RunMacro(req.Workbook, req.MacroName, req.Args) }))
            .WithTags("Macros")
            .WithSummary("Run a VBA macro");
    }
}
