using System.IO;
using System.Text;
using System.Threading.Tasks;

using AgentExcel.Middlewares;

using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging.Abstractions;

using NUnit.Framework;

namespace AgentExcel.Tests.Middlewares;

[TestFixture]
[Category("Unit")]
public class ExceptionHandlingMiddlewareTests : BaseTests
{
    [Test]
    public async Task InvokeAsync_WhenNoException_CallsNextMiddlewareAndPropagatesResponse()
    {
        // Arrange
        var context = new DefaultHttpContext();
        using var responseStream = new MemoryStream();
        context.Response.Body = responseStream;

        RequestDelegate next = async (ctx) =>
        {
            ctx.Response.StatusCode = 200;
            ctx.Response.ContentType = "text/plain; charset=utf-8";
            await ctx.Response.WriteAsync("Success");
        };

        var middleware = new ExceptionHandlingMiddleware(next, NullLogger<ExceptionHandlingMiddleware>.Instance);

        // Act
        await middleware.InvokeAsync(context);

        // Assert
        Assert.That(context.Response.StatusCode, Is.EqualTo(200));
        Assert.That(context.Response.ContentType, Is.EqualTo("text/plain; charset=utf-8"));

        responseStream.Position = 0;
        using var reader = new StreamReader(responseStream, Encoding.UTF8);
        string output = await reader.ReadToEndAsync();
        Assert.That(output, Is.EqualTo("Success"));
    }

    [Test]
    public async Task InvokeAsync_WhenExceptionThrown_Returns500AndErrorMessageAsPlainText()
    {
        // Arrange
        var context = new DefaultHttpContext();
        using var responseStream = new MemoryStream();
        context.Response.Body = responseStream;

        const string errorMessage = "Test exception message";
        RequestDelegate next = (ctx) =>
        {
            throw new Exception(errorMessage);
        };

        var middleware = new ExceptionHandlingMiddleware(next, NullLogger<ExceptionHandlingMiddleware>.Instance);

        // Act
        await middleware.InvokeAsync(context);

        // Assert
        Assert.That(context.Response.StatusCode, Is.EqualTo(500));
        Assert.That(context.Response.ContentType, Is.EqualTo("text/plain; charset=utf-8"));

        responseStream.Position = 0;
        using var reader = new StreamReader(responseStream, Encoding.UTF8);
        string output = await reader.ReadToEndAsync();
        Assert.That(output, Is.EqualTo(errorMessage));
    }

    [Test]
    public async Task InvokeAsync_WhenBadHttpRequestExceptionThrown_Returns400AndErrorMessageAsPlainText()
    {
        // Arrange
        var context = new DefaultHttpContext();
        using var responseStream = new MemoryStream();
        context.Response.Body = responseStream;

        const string errorMessage = "Failed to read parameter \"req\" from the request body as JSON.";
        RequestDelegate next = (ctx) =>
        {
            throw new BadHttpRequestException(errorMessage, StatusCodes.Status400BadRequest);
        };

        var middleware = new ExceptionHandlingMiddleware(next, NullLogger<ExceptionHandlingMiddleware>.Instance);

        // Act
        await middleware.InvokeAsync(context);

        // Assert
        Assert.That(context.Response.StatusCode, Is.EqualTo(400));
        Assert.That(context.Response.ContentType, Is.EqualTo("text/plain; charset=utf-8"));

        responseStream.Position = 0;
        using var reader = new StreamReader(responseStream, Encoding.UTF8);
        string output = await reader.ReadToEndAsync();
        Assert.That(output, Is.EqualTo(errorMessage));
    }

    [Test]
    public async Task InvokeAsync_WhenBadHttpRequestExceptionWithInnerExceptionThrown_Returns400AndErrorMessageWithDetails()
    {
        // Arrange
        var context = new DefaultHttpContext();
        using var responseStream = new MemoryStream();
        context.Response.Body = responseStream;

        const string errorMessage = "Failed to read parameter \"req\" from the request body as JSON.";
        const string detailMessage = "The JSON value could not be converted to System.String. Path: $.sheetName | LineNumber: 2 | BytePositionInLine: 24.";
        RequestDelegate next = (ctx) =>
        {
            var innerEx = new System.Text.Json.JsonException(detailMessage);
            throw new BadHttpRequestException(errorMessage, StatusCodes.Status400BadRequest, innerEx);
        };

        var middleware = new ExceptionHandlingMiddleware(next, NullLogger<ExceptionHandlingMiddleware>.Instance);

        // Act
        await middleware.InvokeAsync(context);

        // Assert
        Assert.That(context.Response.StatusCode, Is.EqualTo(400));
        Assert.That(context.Response.ContentType, Is.EqualTo("text/plain; charset=utf-8"));

        responseStream.Position = 0;
        using var reader = new StreamReader(responseStream, Encoding.UTF8);
        string output = await reader.ReadToEndAsync();
        Assert.That(output, Is.EqualTo($"{errorMessage} Details: {detailMessage}"));
    }
}
