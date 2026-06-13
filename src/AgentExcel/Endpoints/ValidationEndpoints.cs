using AgentExcel.Models;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;

namespace AgentExcel.Endpoints;

public static class ValidationEndpoints
{
    public static void MapValidationEndpoints(this IEndpointRouteBuilder app)
    {
        var validation = app.MapGroup("/validation").WithOpenApi();

        validation.MapPost("/set-list", (SetListValidationRequest req, ValidationService validationService) =>
        {
            validationService.SetListValidation(req.Workbook, req.Sheet, req.RangeAddress, req.Formula);
            return Results.Extensions.Yaml(new { status = "success" });
        })
        .WithTags("Validation")
        .WithSummary("Set a dropdown list validation");
    }
}
