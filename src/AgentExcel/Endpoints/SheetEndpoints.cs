using AgentExcel.Models;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;

namespace AgentExcel.Endpoints;

public static class SheetEndpoints
{
    public static void MapSheetEndpoints(this IEndpointRouteBuilder app)
    {
        var sheets = app.MapGroup("/sheets").WithOpenApi();

        sheets.MapPost("/list", (WorkbookRequest req, SheetService sheetService) =>
            Results.Extensions.Yaml(sheetService.ListSheets(req.Workbook)))
            .WithTags("Sheets")
            .WithSummary("List sheets in a workbook");

        sheets.MapPost("/add", (SheetRequest req, SheetService sheetService) =>
        {
            sheetService.AddSheet(req.Workbook, req.Name);
            return Results.Extensions.Yaml(new { status = "added" });
        })
        .WithTags("Sheets")
        .WithSummary("Add a new sheet");

        sheets.MapPost("/delete", (SheetRequest req, SheetService sheetService) =>
        {
            sheetService.DeleteSheet(req.Workbook, req.Name);
            return Results.Extensions.Yaml(new { status = "deleted" });
        })
        .WithTags("Sheets")
        .WithSummary("Delete a sheet");

        sheets.MapPost("/rename", (RenameSheetRequest req, SheetService sheetService) =>
        {
            sheetService.RenameSheet(req.Workbook, req.OldName, req.NewName);
            return Results.Extensions.Yaml(new { status = "renamed" });
        })
        .WithTags("Sheets")
        .WithSummary("Rename a sheet");
    }
}
