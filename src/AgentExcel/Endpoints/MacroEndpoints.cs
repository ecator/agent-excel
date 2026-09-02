using AgentExcel.Models;
using AgentExcel.Models.Requests;
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
            .WithSummary("List all macros and their parameters across all components (Module, Class, Form, Document) in the workbook");

        macros.MapPost("/export", (ExportModuleRequest req, MacroService macroService) =>
        {
            macroService.ExportModule(req.Workbook, req.Module, req.OutputFile);
            return Results.Text($"module[{req.Module}] has been exported to {req.OutputFile}", "text/plain; charset=utf-8");
        })
            .WithTags("Macros")
            .WithSummary("Export a VBA component (Module, Class, Form, or Document) to a file (encoded in system default ANSI)");

        macros.MapPost("/import", (ImportModuleRequest req, MacroService macroService) =>
        {
            var moduleName = macroService.ImportModule(req.Workbook, req.Path);
            return Results.Text($"module[{moduleName}] has been imported from {req.Path}", "text/plain; charset=utf-8");
        })
            .WithTags("Macros")
            .WithSummary("Import a VBA component file (Module, Class, Form) into the workbook");

        macros.MapPost("/delete", (DeleteModuleRequest req, MacroService macroService) =>
        {
            macroService.DeleteModule(req.Workbook, req.Module);
            return Results.Text($"module[{req.Module}] has been deleted", "text/plain; charset=utf-8");
        })
            .WithTags("Macros")
            .WithSummary("Delete a VBA component (Module, Class, Form) from the workbook");
    }
}
