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
            exportService.ExportRangeAsImage(req.Workbook, req.Sheet, req.RangeAddress, req.OutputPath);
            return Results.Extensions.Yaml(new { status = "exported" });
        })
        .WithTags("Export")
        .WithSummary("Export a range as a PNG image");

        export.MapPost("/workbook-pdf", (ExportPdfRequest req, ExportService exportService) =>
        {
            exportService.ExportAsPdf(req.Workbook, req.OutputPath);
            return Results.Extensions.Yaml(new { status = "exported" });
        })
        .WithTags("Export")
        .WithSummary("Export the entire workbook as a PDF");
    }
}
