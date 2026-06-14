using AgentExcel.Models;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;

namespace AgentExcel.Endpoints;

public static class DataEndpoints
{
    public static void MapDataEndpoints(this IEndpointRouteBuilder app)
    {
        var data = app.MapGroup("/data").WithOpenApi();

        data.MapPost("/read-range", (ReadRangeRequest req, DataService dataService) =>
        {
            var results = dataService.ReadRange(req.Workbook, req.Sheet, req.Range);
            return results.Count > 0 ? Results.Extensions.Yaml(results) : Results.Extensions.Yaml(null);
        })
        .WithTags("Data")
        .WithSummary("Read data from a range");


        data.MapPost("/list-tables", (ListTablesRequest req, DataService dataService) =>
        {
            var results = dataService.ListTables(req.Workbook, req.Sheet).Select(i => new { Sheet = i.Key, Tables = i.Value });
            return results.Count() > 0 ? Results.Extensions.Yaml(results) : Results.Extensions.Yaml(null);
        })
        .WithTags("Data")
        .WithSummary("List all Excel Tables in the workbook or a specific sheet");


        data.MapPost("/read-table", (ReadTableRequest req, DataService dataService) =>
        {
            var content = dataService.ReadTableAsMarkdown(req.Workbook, req.Sheet, req.Table);
            return Results.Text(content, "text/plain; charset=utf-8");
        })
        .WithTags("Data")
        .WithSummary("Read data from an Excel Table (ListObject) in Markdown format");

        data.MapPost("/rename-table", (RenameTableRequest req, DataService dataService) =>
        {
            dataService.RenameTable(req.Workbook, req.Sheet, req.Table, req.NewName);
            return Results.Text($"{req.Table} has renamed to {req.NewName}");
        })
        .WithTags("Data")
        .WithSummary("Rename an Excel Table (ListObject)");

        data.MapPost("/convert-to-table", (ConvertToTableRequest req, DataService dataService) =>
            Results.Extensions.Yaml(new { name = dataService.ConvertToTable(req.Workbook, req.Sheet, req.RangeAddress, req.TableName, req.HasHeaders) }))
            .WithTags("Data")
            .WithSummary("Convert a normal range to an Excel Table");

        data.MapPost("/convert-to-range", (ConvertToRangeRequest req, DataService dataService) =>
        {
            dataService.ConvertToRange(req.Workbook, req.Sheet, req.TableName);
            return Results.Extensions.Yaml(new { status = "converted" });
        })
        .WithTags("Data")
        .WithSummary("Convert an Excel Table back to a normal range");

        data.MapPost("/write-range", (RangeWriteRequest req, DataService dataService) =>
        {
            dataService.WriteRange(req.Workbook, req.Sheet, req.Address, req.Value);
            return Results.Extensions.Yaml(new { status = "success" });
        })
        .WithTags("Data")
        .WithSummary("Write data to a range");

        data.MapPost("/write-formula", (WriteFormulaRequest req, DataService dataService) =>
        {
            dataService.WriteFormula(req.Workbook, req.Sheet, req.Address, req.Formula);
            return Results.Extensions.Yaml(new { status = "success" });
        })
        .WithTags("Data")
        .WithSummary("Write a formula to a range");

        data.MapPost("/read-formula", (ReadRangeRequest req, DataService dataService) =>
            Results.Extensions.Yaml(new { formula = dataService.ReadFormula(req.Workbook, req.Sheet, req.Range ?? "") }))
            .WithTags("Data")
            .WithSummary("Read the formula from a range");

        data.MapPost("/find", (FindRequest req, DataService dataService) =>
            Results.Extensions.Yaml(new { results = dataService.Find(req.Workbook, req.Sheet, req.RangeAddress, req.What, req.MatchCase, req.WholeWord) }))
            .WithTags("Data")
            .WithSummary("Find all occurrences of a string");

        data.MapPost("/replace", (ReplaceRequest req, DataService dataService) =>
            Results.Extensions.Yaml(new { success = dataService.Replace(req.Workbook, req.Sheet, req.RangeAddress, req.What, req.Replacement, req.MatchCase, req.WholeWord) }))
            .WithTags("Data")
            .WithSummary("Replace all occurrences of a string");



        data.MapPost("/grep", (GrepRequest req, DataService dataService) =>
            Results.Extensions.Yaml(new { results = dataService.SearchInFolder(req.FolderPath, req.Pattern) }))
            .WithTags("Data")
            .WithSummary("Search for text in all Excel files within a folder (includes Cells, TextBoxes, and Shapes)");

        data.MapPost("/set-style", (SetStyleRequest req, DataService dataService) =>
        {
            dataService.SetStyle(req.Workbook, req.Sheet, req.Address, req.Style);
            return Results.Extensions.Yaml(new { status = "success" });
        })
        .WithTags("Data")
        .WithSummary("Set cell styles (Font, Color, Bold, Alignment, etc.)");

    }
}
