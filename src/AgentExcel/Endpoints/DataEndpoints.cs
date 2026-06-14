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

        data.MapPost("/write-range", (RangeWriteRequest req, DataService dataService) =>
        {
            dataService.WriteRange(req.Workbook, req.Sheet, req.Range, req.Value);
            return Results.Text($"range[{req.Range}] has be filled with data");
        })
        .WithTags("Data")
        .WithSummary("Write data to a range");

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
        {
            var tableName = dataService.ConvertToTable(req.Workbook, req.Sheet, req.Range, req.Table, req.HasHeaders);
            return Results.Text($"range[{req.Range}] has converted to table[{tableName}]");
        })
        .WithTags("Data")
        .WithSummary("Convert a normal range to an Excel Table");

        data.MapPost("/convert-to-range", (ConvertToRangeRequest req, DataService dataService) =>
        {
            var address = dataService.ConvertToRange(req.Workbook, req.Sheet, req.Table);
            return Results.Text($"table[{req.Table}] has converted to  range[{address}]");
        })
        .WithTags("Data")
        .WithSummary("Convert an Excel Table back to a normal range");


        data.MapPost("/read-formula", (ReadRangeRequest req, DataService dataService) =>
        {
            var results = dataService.ReadFormula(req.Workbook, req.Sheet, req.Range);
            return results.Count > 0 ? Results.Extensions.Yaml(results) : Results.Extensions.Yaml(null);
        })
        .WithTags("Data")
        .WithSummary("Read the formula from a range");

        data.MapPost("/write-formula", (WriteFormulaRequest req, DataService dataService) =>
        {
            dataService.WriteFormula(req.Workbook, req.Sheet, req.Range, req.Formula);
            return Results.Text($"range[{req.Range}] has be filled with formula");
        })
        .WithTags("Data")
        .WithSummary("Write formula to a range");

        data.MapPost("/find", (FindRequest req, DataService dataService) =>
        {
            var results = dataService.Find(req.Workbook, req.Sheet, req.Range, req.What, req.MatchCase, req.WholeWord);
            var content = FormatFindResults(results, "found");
            return Results.Text(content, "text/plain; charset=utf-8");
        })
        .WithTags("Data")
        .WithSummary("Find all occurrences of a string");

        data.MapPost("/replace", (ReplaceRequest req, DataService dataService) =>
        {
            var count = dataService.Replace(req.Workbook, req.Sheet, req.Range, req.What, req.Replacement, req.MatchCase, req.WholeWord);
            return Results.Text($"replaced {count} results");
        })
        .WithTags("Data")
        .WithSummary("Replace all occurrences of a string");

        data.MapPost("/get-style", (ReadRangeRequest req, DataService dataService) =>
        {
            var result = dataService.GetStyle(req.Workbook, req.Sheet, req.Range);
            return Results.Extensions.Yaml(result);
        })
        .WithTags("Data")
        .WithSummary("Get cell style. Note: only supports returning the style of the top-left (first) cell of the range.");

        data.MapPost("/set-style", (SetStyleRequest req, DataService dataService) =>
        {
            dataService.SetStyle(req.Workbook, req.Sheet, req.Range, req.Style);
            return Results.Text($"style of range[{req.Range}] has changed");
        })
        .WithTags("Data")
        .WithSummary("Set cell styles (Font, Color, Bold, Alignment, etc.)");

        data.MapPost("/clear", (ClearRequest req, DataService dataService) =>
        {
            var address = dataService.Clear(req.Workbook, req.Sheet, req.Range, req.Type);
            return Results.Text($"range[{address}] has been cleared");
        })
        .WithTags("Data")
        .WithSummary("Clear range (All, Formats, Contents, Comments, Hyperlinks)");

        data.MapPost("/get-selection", (WorksheetRequest req, DataService dataService) =>
        {
            var address = dataService.GetSelection(req.Workbook, req.Sheet);
            return Results.Text(address);
        })
        .WithTags("Data")
        .WithSummary("Get selection address of a specific sheet");

        data.MapPost("/set-selection", (SetSelectionRequest req, DataService dataService) =>
        {
            var address = dataService.SetSelection(req.Workbook, req.Sheet, req.Range);
            return Results.Text($"selection of sheet[{req.Sheet}] has been set to range[{address}]");
        })
        .WithTags("Data")
        .WithSummary("Set selection of a specific sheet and activate the range");

    }

    private static string FormatFindResults(List<FindResult> results, string action)
    {
        if (results.Count == 0)
        {
            return $"{action} 0 results";
        }

        var sb = new System.Text.StringBuilder();
        var groups = results.GroupBy(r => r.Sheet);
        int index = 0;
        foreach (var group in groups)
        {
            if (index > 0)
            {
                sb.AppendLine();
                sb.AppendLine();
            }

            var sheetName = group.Key;
            var count = group.Count();
            sb.AppendLine($"{action} {count} results in {sheetName}:");

            var headers = new[] { "address", "value" };
            var body = new string[count, 2];
            int rIdx = 0;
            foreach (var item in group)
            {
                body[rIdx, 0] = item.Address;
                body[rIdx, 1] = item.Value;
                rIdx++;
            }
            sb.Append(MarkdownTableHelper.ToMarkdownTable(headers, body));
            index++;
        }

        return sb.ToString();
    }
}
