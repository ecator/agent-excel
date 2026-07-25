using AgentExcel.Models;
using AgentExcel.Models.Requests;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;

namespace AgentExcel.Endpoints;

public static class RangeEndpoints
{
    public static void MapRangeEndpoints(this IEndpointRouteBuilder app)
    {
        var range = app.MapGroup("/range").WithOpenApi();

        range.MapPost("/read", (ReadRangeRequest req, RangeService rangeService) =>
        {
            var results = rangeService.ReadRange(req.Workbook, req.Sheet, req.Range);
            return results.Count > 0 ? Results.Extensions.Yaml(results) : Results.Extensions.Yaml(null);
        })
        .WithTags("Range")
        .WithSummary("Read data from a range");

        range.MapPost("/write", (RangeWriteRequest req, RangeService rangeService) =>
        {
            rangeService.WriteRange(req.Workbook, req.Sheet, req.Range, req.Value);
            return Results.Text($"range[{req.Range}] has be filled with data");
        })
        .WithTags("Range")
        .WithSummary("Write data to a range");

        range.MapPost("/list-tables", (ListTablesRequest req, RangeService rangeService) =>
        {
            var results = rangeService.ListTables(req.Workbook, req.Sheet).Select(i => new { Sheet = i.Key, Tables = i.Value });
            return results.Count() > 0 ? Results.Extensions.Yaml(results) : Results.Extensions.Yaml(null);
        })
        .WithTags("Range")
        .WithSummary("List all Excel Tables in the workbook or a specific sheet");


        range.MapPost("/read-table", (ReadTableRequest req, RangeService rangeService) =>
        {
            var content = rangeService.ReadTableAsMarkdown(req.Workbook, req.Sheet, req.Table);
            return Results.Text(content, "text/plain; charset=utf-8");
        })
        .WithTags("Range")
        .WithSummary("Read data from an Excel Table (ListObject) in Markdown format");

        range.MapPost("/rename-table", (RenameTableRequest req, RangeService rangeService) =>
        {
            rangeService.RenameTable(req.Workbook, req.Sheet, req.Table, req.NewName);
            return Results.Text($"{req.Table} has renamed to {req.NewName}");
        })
        .WithTags("Range")
        .WithSummary("Rename an Excel Table (ListObject)");

        range.MapPost("/convert-to-table", (ConvertToTableRequest req, RangeService rangeService) =>
        {
            var tableName = rangeService.ConvertToTable(req.Workbook, req.Sheet, req.Range, req.Table, req.HasHeaders);
            return Results.Text($"range[{req.Range}] has converted to table[{tableName}]");
        })
        .WithTags("Range")
        .WithSummary("Convert a normal range to an Excel Table");

        range.MapPost("/convert-to-range", (ConvertToRangeRequest req, RangeService rangeService) =>
        {
            var address = rangeService.ConvertToRange(req.Workbook, req.Sheet, req.Table);
            return Results.Text($"table[{req.Table}] has converted to  range[{address}]");
        })
        .WithTags("Range")
        .WithSummary("Convert an Excel Table back to a normal range");


        range.MapPost("/read-formula", (ReadRangeRequest req, RangeService rangeService) =>
        {
            var results = rangeService.ReadFormula(req.Workbook, req.Sheet, req.Range);
            return results.Count > 0 ? Results.Extensions.Yaml(results) : Results.Extensions.Yaml(null);
        })
        .WithTags("Range")
        .WithSummary("Read the formula from a range");

        range.MapPost("/write-formula", (WriteFormulaRequest req, RangeService rangeService) =>
        {
            rangeService.WriteFormula(req.Workbook, req.Sheet, req.Range, req.Formula);
            return Results.Text($"range[{req.Range}] has be filled with formula");
        })
        .WithTags("Range")
        .WithSummary("Write formula to a range");

        range.MapPost("/find", (FindRequest req, RangeService rangeService) =>
        {
            var results = rangeService.Find(req.Workbook, req.Sheet, req.Range, req.What, req.MatchCase, req.WholeWord);
            var content = FormatFindResults(results, "found");
            return Results.Text(content, "text/plain; charset=utf-8");
        })
        .WithTags("Range")
        .WithSummary("Find all occurrences of a string");

        range.MapPost("/replace", (ReplaceRequest req, RangeService rangeService) =>
        {
            var count = rangeService.Replace(req.Workbook, req.Sheet, req.Range, req.What, req.Replacement, req.MatchCase, req.WholeWord);
            return Results.Text($"replaced {count} results");
        })
        .WithTags("Range")
        .WithSummary("Replace all occurrences of a string");

        range.MapPost("/get-style", (ReadRangeRequest req, RangeService rangeService) =>
        {
            var result = rangeService.GetStyle(req.Workbook, req.Sheet, req.Range);
            return Results.Extensions.Yaml(result);
        })
        .WithTags("Range")
        .WithSummary("Get cell style. Note: only supports returning the style of the top-left (first) cell of the range.");

        range.MapPost("/set-style", (SetStyleRequest req, RangeService rangeService) =>
        {
            rangeService.SetStyle(req.Workbook, req.Sheet, req.Range, req.Style);
            return Results.Text($"style of range[{req.Range}] has changed");
        })
        .WithTags("Range")
        .WithSummary("Set cell styles (Font, Color, Bold, Alignment, etc.)");

        range.MapPost("/clear", (ClearRequest req, RangeService rangeService) =>
        {
            var address = rangeService.Clear(req.Workbook, req.Sheet, req.Range, req.Type);
            return Results.Text($"range[{address}] has been cleared");
        })
        .WithTags("Range")
        .WithSummary("Clear range (All, Formats, Contents, Comments, Hyperlinks)");

        range.MapPost("/delete", (DeleteRangeRequest req, RangeService rangeService) =>
        {
            var address = rangeService.Delete(req.Workbook, req.Sheet, req.Range, req.Shift);
            return Results.Text($"range[{address}] has been deleted");
        })
        .WithTags("Range")
        .WithSummary("Delete range with shift options (shift_left, shift_up, entire_row, entire_column)");

        range.MapPost("/get-comments", (GetCommentsRequest req, RangeService rangeService) =>
        {
            var results = rangeService.GetComments(req.Workbook, req.Sheet);
            return results.Count > 0 ? Results.Extensions.Yaml(results) : Results.Extensions.Yaml(null);
        })
        .WithTags("Range")
        .WithSummary("List all comments in a sheet");

        range.MapPost("/set-comment", (SetCommentRequest req, RangeService rangeService) =>
        {
            rangeService.SetComment(req.Workbook, req.Sheet, req.Range, req.Text, req.Visible);
            return Results.Text($"comment of range[{req.Range}] has been set");
        })
        .WithTags("Range")
        .WithSummary("Set comment for a specified range (overwrites existing comment)");


        range.MapPost("/get-selection", (WorksheetRequest req, RangeService rangeService) =>
        {
            var address = rangeService.GetSelection(req.Workbook, req.Sheet);
            return Results.Text(address);
        })
        .WithTags("Range")
        .WithSummary("Get selection address of a specific sheet");

        range.MapPost("/set-selection", (SetSelectionRequest req, RangeService rangeService) =>
        {
            var address = rangeService.SetSelection(req.Workbook, req.Sheet, req.Range);
            return Results.Text($"selection of sheet[{req.Sheet}] has been set to range[{address}]");
        })
        .WithTags("Range")
        .WithSummary("Set selection of a specific sheet and activate the range");

        range.MapPost("/set-list", (SetListValidationRequest req, RangeService rangeService) =>
        {
            rangeService.SetListValidation(req.Workbook, req.Sheet, req.Range, req.Formula);
            return Results.Text($"validation of range[{req.Range}] has changed to {req.Formula}");
        })
        .WithTags("Range")
        .WithSummary("Set a dropdown list validation");

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
