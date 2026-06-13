using AgentExcel.Models;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;

namespace AgentExcel.Endpoints;

public static class WorkbookEndpoints
{
    public static void MapWorkbookEndpoints(this IEndpointRouteBuilder app)
    {
        var workbooks = app.MapGroup("/workbooks").WithOpenApi();

        workbooks.MapPost("/list", (WorkbookService workbookService) =>
        {
            var result = workbookService.ListWorkbooks().Select(i => new { i.Name, i.Path });
            return Results.Extensions.Yaml(result);
        })
            .WithTags("Workbooks")
            .WithSummary("List all open workbooks");


        workbooks.MapPost("/open", (OpenWorkbookRequest req, WorkbookService workbookService) =>
        {
            var result = workbookService.OpenWorkbook(req.Path);
            return Results.Extensions.Yaml(result);
        })
            .WithTags("Workbooks")
            .WithSummary("Open a workbook by path and return its info");

        workbooks.MapPost("/add", (WorkbookService workbookService) =>
        {
            var result = workbookService.AddWorkbook();
            return Results.Extensions.Yaml(new { result.Name, result.Sheets });
        })
            .WithTags("Workbooks")
            .WithSummary("Create a new workbook and return its info");

        workbooks.MapPost("/save", (WorkbookRequest req, WorkbookService workbookService) =>
        {
            workbookService.SaveWorkbook(req.Workbook);
            return Results.Ok($"saved");
        })
        .WithTags("Workbooks")
        .WithSummary("Save a workbook");

        workbooks.MapPost("/saveas", (SaveAsRequest req, WorkbookService workbookService) =>
        {
            workbookService.SaveAsWorkbook(req.Workbook, req.Path);
            return Results.Ok($"saved to `{req.Path}`");
        })
        .WithTags("Workbooks")
        .WithSummary("Save a workbook to a new path");

        workbooks.MapPost("/close", (CloseWorkbookRequest req, WorkbookService workbookService) =>
        {
            workbookService.CloseWorkbook(req.Workbook, req.SaveChanges);
            return Results.Ok("closed");
        })
        .WithTags("Workbooks")
        .WithSummary("Close a workbook");
    }
}
