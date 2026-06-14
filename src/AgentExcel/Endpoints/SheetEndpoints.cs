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

        sheets.MapPost("/add", (WorksheetRequest req, SheetService sheetService) =>
        {
            sheetService.AddSheet(req.Workbook, req.Sheet);
            return Results.Text("added");
        })
        .WithTags("Sheets")
        .WithSummary("Add a new sheet");

        sheets.MapPost("/delete", (WorksheetRequest req, SheetService sheetService) =>
        {
            sheetService.DeleteSheet(req.Workbook, req.Sheet);
            return Results.Text("deleted");
        })
        .WithTags("Sheets")
        .WithSummary("Delete a sheet");

        sheets.MapPost("/rename", (RenameSheetRequest req, SheetService sheetService) =>
        {
            sheetService.RenameSheet(req.Workbook, req.OldName, req.NewName);
            return Results.Text("renamed");
        })
        .WithTags("Sheets")
        .WithSummary("Rename a sheet");

        sheets.MapPost("/copy", (CopySheetRequest req, SheetService sheetService) =>
        {
            sheetService.CopySheet(req.Workbook, req.OldName, req.NewName, req.TargetWorkbook, req.Position);
            return Results.Text("copied");
        })
        .WithTags("Sheets")
        .WithSummary("Copy a sheet to the target workbook. If target workbook is omitted, the source workbook is used. Position 0 for first, -1 for last.");

        sheets.MapPost("/move", (MoveSheetRequest req, SheetService sheetService) =>
        {
            sheetService.MoveSheet(req.Workbook, req.Sheet, req.TargetWorkbook, req.Position);
            return Results.Text("moved");
        })
        .WithTags("Sheets")
        .WithSummary("Move a sheet to the target workbook. If target workbook is omitted, the source workbook is used. Position 0 for first, -1 for last.");

        sheets.MapPost("/set-color", (SetSheetColorRequest req, SheetService sheetService) =>
        {
            sheetService.SetSheetColor(req.Workbook, req.Sheet, req.Color);
            return Results.Text($"color of {req.Sheet} is set to {req.Color}");
        })
        .WithTags("Sheets")
        .WithSummary("Set the tab color of a specified worksheet");

        sheets.MapPost("/set-visible", (SetSheetVisibilityRequest req, SheetService sheetService) =>
        {
            sheetService.SetSheetVisibility(req.Workbook, req.Sheet, req.Visibility);
            return Results.Text($"visibility of {req.Sheet} is set to {req.Visibility}");
        })
        .WithTags("Sheets")
        .WithSummary("Set the visibility state of a specified worksheet (Visible, Hidden, or VeryHidden)");

        sheets.MapPost("/get-active", (WorkbookRequest req, SheetService sheetService) =>
            Results.Extensions.Yaml(sheetService.GetActiveSheet(req.Workbook)))
            .WithTags("Sheets")
            .WithSummary("Get the active sheet in a workbook");

        sheets.MapPost("/set-active", (WorksheetRequest req, SheetService sheetService) =>
        {
            sheetService.SetActiveSheet(req.Workbook, req.Sheet);
            return Results.Text($"{req.Sheet} activated in {req.Workbook}");
        })
        .WithTags("Sheets")
        .WithSummary("Set the active sheet in a workbook");
    }
}
