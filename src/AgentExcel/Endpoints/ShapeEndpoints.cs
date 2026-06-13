using AgentExcel.Models;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;

namespace AgentExcel.Endpoints;

public static class ShapeEndpoints
{
    public static void MapShapeEndpoints(this IEndpointRouteBuilder app)
    {
        var shapes = app.MapGroup("/shapes").WithOpenApi();

        shapes.MapPost("/list", (UsedRangeRequest req, ShapeService shapeService) =>
            Results.Extensions.Yaml(shapeService.ListShapes(req.Workbook, req.Sheet)))
            .WithTags("Shapes")
            .WithSummary("List all shapes and textboxes in a sheet");

        shapes.MapPost("/add", (AddShapeRequest req, ShapeService shapeService) =>
            Results.Extensions.Yaml(new { name = shapeService.AddShape(req.Workbook, req.Sheet, req.Type, req.Left, req.Top, req.Width, req.Height, req.Text) }))
            .WithTags("Shapes")
            .WithSummary("Add a shape or textbox");

        shapes.MapPost("/update", (UpdateShapeRequest req, ShapeService shapeService) =>
        {
            shapeService.UpdateShape(req.Workbook, req.Sheet, req.ShapeName, req.Left, req.Top, req.Width, req.Height, req.Text);
            return Results.Extensions.Yaml(new { status = "updated" });
        })
        .WithTags("Shapes")
        .WithSummary("Update shape properties (position, size, text)");

        shapes.MapPost("/delete", (DeleteShapeRequest req, ShapeService shapeService) =>
        {
            shapeService.DeleteShape(req.Workbook, req.Sheet, req.ShapeName);
            return Results.Extensions.Yaml(new { status = "deleted" });
        })
        .WithTags("Shapes")
        .WithSummary("Delete a shape");
    }
}
