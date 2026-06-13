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
}
