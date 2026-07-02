using AgentExcel.Models;
using AgentExcel.Models.Requests;
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

        shapes.MapPost("/list", (WorksheetRequest req, ShapeService shapeService) =>
            Results.Extensions.Yaml(shapeService.ListShapes(req.Workbook, req.Sheet)))
            .WithTags("Shapes")
            .WithSummary("List all shapes and textboxes in a sheet");

        shapes.MapPost("/add", (AddShapeRequest req, ShapeService shapeService) =>
            Results.Extensions.Yaml(shapeService.AddShape(req.Workbook, req.Sheet, req.Type, req.Left, req.Top, req.Width, req.Height, req.Text)))
            .WithTags("Shapes")
            .WithSummary("Add a shape or textbox");

        shapes.MapPost("/update", (UpdateShapeRequest req, ShapeService shapeService) =>
        {
            var updatedInfo = shapeService.UpdateShape(req.Workbook, req.Sheet, req.ShapeName, req.Left, req.Top, req.Width, req.Height, req.Text);
            var details = new List<string>();
            if (req.Left.HasValue) details.Add($"left updated to {updatedInfo.Left}");
            if (req.Top.HasValue) details.Add($"top updated to {updatedInfo.Top}");
            if (req.Width.HasValue) details.Add($"width updated to {updatedInfo.Width}");
            if (req.Height.HasValue) details.Add($"height updated to {updatedInfo.Height}");
            if (req.Text != null) details.Add($"text updated to '{updatedInfo.Text}'");

            string message = details.Count == 0
                ? $"Shape '{req.ShapeName}' was not modified."
                : $"Shape '{req.ShapeName}' updated:\n{string.Join("\n", details)}.";

            return Results.Text(message, "text/plain; charset=utf-8");
        })
        .WithTags("Shapes")
        .WithSummary("Update shape properties (position, size, text)");

        shapes.MapPost("/delete", (DeleteShapeRequest req, ShapeService shapeService) =>
        {
            shapeService.DeleteShape(req.Workbook, req.Sheet, req.ShapeName);
            string message = $"shape[{req.ShapeName}] has deleted";
            return Results.Text(message, "text/plain; charset=utf-8");
        })
        .WithTags("Shapes")
        .WithSummary("Delete a shape");
    }
}
