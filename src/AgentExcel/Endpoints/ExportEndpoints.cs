using AgentExcel.Models;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;

namespace AgentExcel.Endpoints;

public static class ExportEndpoints
{
    public static void MapExportEndpoints(this IEndpointRouteBuilder app)
    {
        var export = app.MapGroup("/export").WithOpenApi();

        export.MapPost("/range-image", (ExportImageRequest req, ExportService exportService) =>
        {
            exportService.ExportRangeAsImage(req.Workbook, req.Sheet, req.Range, req.OutputFile);
            return Results.Text($"range[{req.Range}] has been exported to {req.OutputFile}", "text/plain; charset=utf-8");
        })
        .WithTags("Export")
        .WithSummary("Export a range as a PNG image");

        export.MapPost("/workbook-pdf", (ExportPdfRequest req, ExportService exportService) =>
        {
            exportService.ExportAsPdf(req.Workbook, req.OutputFile);
            return Results.Text($"workbook[{req.Workbook}] has been exported to {req.OutputFile}", "text/plain; charset=utf-8");
        })
        .WithTags("Export")
        .WithSummary("Export the entire workbook as a PDF");
    }
}
