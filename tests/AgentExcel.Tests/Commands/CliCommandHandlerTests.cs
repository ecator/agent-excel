using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

using AgentExcel.Commands;

using NUnit.Framework;

namespace AgentExcel.Tests.Commands;

[TestFixture]
[Category("Unit")]
public class CliCommandHandlerTests : BaseTests
{
    private class FakeHttpMessageHandler : HttpMessageHandler
    {
        public Func<HttpRequestMessage, Task<HttpResponseMessage>>? HandlerFunc { get; set; }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            if (HandlerFunc == null)
            {
                return new HttpResponseMessage(HttpStatusCode.NotFound);
            }
            return await HandlerFunc(request);
        }
    }

    [Test]
    public async Task ExecuteApiRequest_WithNoEndpoint_PrintsUsage()
    {
        // Arrange
        var command = "get";
        string[] args = ["get"];
        using var outputWriter = new StringWriter();
        string exeName = Path.GetFileName(Process.GetCurrentProcess().MainModule?.FileName ?? "AgentExcel.exe");

        // Act
        await CliCommandHandler.ExecuteApiRequest(command, args, "127.0.0.1", null, null, outputWriter, 8080);

        // Assert
        string output = outputWriter.ToString();
        Assert.That(output, Does.Contain($"Usage: {exeName} get <endpoint>"));
    }

    [Test]
    public async Task ExecuteApiRequest_WhenDaemonNotRunning_PrintsError()
    {
        // Arrange
        var command = "get";
        string[] args = ["get", "/status"];
        using var outputWriter = new StringWriter();
        string exeName = Path.GetFileName(Process.GetCurrentProcess().MainModule?.FileName ?? "AgentExcel.exe");

        // Act
        await CliCommandHandler.ExecuteApiRequest(command, args, "127.0.0.1", null, null, outputWriter, null);

        // Assert
        string output = outputWriter.ToString();
        Assert.That(output, Does.Contain($"Error: Daemon server is not running. Please start the server first by running '{exeName} start'."));
    }

    [Test]
    public async Task ExecuteApiRequest_WithSuccessfulGet_SendsGetRequestAndPrintsResponse()
    {
        // Arrange
        var command = "get";
        string[] args = ["get", "status"];
        using var outputWriter = new StringWriter();
        var handler = new FakeHttpMessageHandler();
        using var client = new HttpClient(handler);

        HttpRequestMessage? interceptedRequest = null;
        handler.HandlerFunc = (req) =>
        {
            interceptedRequest = req;
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("status: running\npid: 1234")
            });
        };

        // Act
        await CliCommandHandler.ExecuteApiRequest(command, args, "127.0.0.1", client, null, outputWriter, 8080);

        // Assert
        string output = outputWriter.ToString();
        Assert.That(output, Does.Contain("status: running"));
        Assert.That(output, Does.Contain("pid: 1234"));
        Assert.That(interceptedRequest, Is.Not.Null);
        Assert.That(interceptedRequest!.Method, Is.EqualTo(HttpMethod.Get));
        Assert.That(interceptedRequest!.RequestUri!.ToString(), Is.EqualTo("http://127.0.0.1:8080/status"));
    }

    [Test]
    public async Task ExecuteApiRequest_WithSuccessfulPostAndStdin_SendsPostWithStdinBody()
    {
        // Arrange
        var command = "post";
        string[] args = ["post", "/workbooks/open", "--stdin"];
        using var outputWriter = new StringWriter();
        using var stdinReader = new StringReader("{\"path\":\"test.xlsx\"}");
        var handler = new FakeHttpMessageHandler();
        using var client = new HttpClient(handler);

        HttpRequestMessage? interceptedRequest = null;
        string? requestContent = null;
        handler.HandlerFunc = async (req) =>
        {
            interceptedRequest = req;
            if (req.Content != null)
            {
                requestContent = await req.Content.ReadAsStringAsync();
            }
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("workbook opened successfully")
            };
        };

        // Act
        await CliCommandHandler.ExecuteApiRequest(command, args, "127.0.0.1", client, stdinReader, outputWriter, 8080);

        // Assert
        string output = outputWriter.ToString();
        Assert.That(output, Does.Contain("workbook opened successfully"));
        Assert.That(interceptedRequest, Is.Not.Null);
        Assert.That(interceptedRequest!.Method, Is.EqualTo(HttpMethod.Post));
        Assert.That(interceptedRequest!.RequestUri!.ToString(), Is.EqualTo("http://127.0.0.1:8080/workbooks/open"));
        Assert.That(requestContent, Is.EqualTo("{\"path\":\"test.xlsx\"}"));
    }

    [Test]
    public async Task ExecuteApiRequest_WithSuccessfulPostAndBodyParam_SendsPostWithBodyParam()
    {
        // Arrange
        var command = "post";
        string[] args = ["post", "/workbooks/close", "{\"workbook\":\"wk1.xlsx\",\"saveChanges\":true}"];
        using var outputWriter = new StringWriter();
        var handler = new FakeHttpMessageHandler();
        using var client = new HttpClient(handler);

        HttpRequestMessage? interceptedRequest = null;
        string? requestContent = null;
        handler.HandlerFunc = async (req) =>
        {
            interceptedRequest = req;
            if (req.Content != null)
            {
                requestContent = await req.Content.ReadAsStringAsync();
            }
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("workbook closed")
            };
        };

        // Act
        await CliCommandHandler.ExecuteApiRequest(command, args, "127.0.0.1", client, null, outputWriter, 8080);

        // Assert
        string output = outputWriter.ToString();
        Assert.That(output, Does.Contain("workbook closed"));
        Assert.That(interceptedRequest, Is.Not.Null);
        Assert.That(interceptedRequest!.Method, Is.EqualTo(HttpMethod.Post));
        Assert.That(interceptedRequest!.RequestUri!.ToString(), Is.EqualTo("http://127.0.0.1:8080/workbooks/close"));
        Assert.That(requestContent, Is.EqualTo("{\"workbook\":\"wk1.xlsx\",\"saveChanges\":true}"));
    }

    [Test]
    public async Task ExecuteApiRequest_WithMultipleBodyParams_OnlyUsesThirdArgument()
    {
        // Arrange
        var command = "post";
        string[] args = ["post", "/workbooks/close", "{\"workbook\":\"wk1.xlsx\"}", "extra_argument"];
        using var outputWriter = new StringWriter();
        var handler = new FakeHttpMessageHandler();
        using var client = new HttpClient(handler);

        HttpRequestMessage? interceptedRequest = null;
        string? requestContent = null;
        handler.HandlerFunc = async (req) =>
        {
            interceptedRequest = req;
            if (req.Content != null)
            {
                requestContent = await req.Content.ReadAsStringAsync();
            }
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("workbook closed")
            };
        };

        // Act
        await CliCommandHandler.ExecuteApiRequest(command, args, "127.0.0.1", client, null, outputWriter, 8080);

        // Assert
        string output = outputWriter.ToString();
        Assert.That(output, Does.Contain("workbook closed"));
        Assert.That(interceptedRequest, Is.Not.Null);
        Assert.That(interceptedRequest!.Method, Is.EqualTo(HttpMethod.Post));
        Assert.That(interceptedRequest!.RequestUri!.ToString(), Is.EqualTo("http://127.0.0.1:8080/workbooks/close"));
        Assert.That(requestContent, Is.EqualTo("{\"workbook\":\"wk1.xlsx\"}"));
    }

    [Test]
    public async Task ExecuteApiRequest_WhenNot200_PrintsErrorCodeAndMessage()
    {
        // Arrange
        var command = "get";
        string[] args = ["get", "/invalid"];
        using var outputWriter = new StringWriter();
        var handler = new FakeHttpMessageHandler();
        using var client = new HttpClient(handler);

        handler.HandlerFunc = (req) =>
        {
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest)
            {
                Content = new StringContent("invalid request parameters")
            });
        };

        // Act
        await CliCommandHandler.ExecuteApiRequest(command, args, "127.0.0.1", client, null, outputWriter, 8080);

        // Assert
        string output = outputWriter.ToString();
        Assert.That(output, Does.Contain("Error Code: 400"));
        Assert.That(output, Does.Contain("invalid request parameters"));
    }

    [Test]
    public async Task ExecuteApiRequest_WhenHttpExceptionThrown_PrintsServerNotRunning()
    {
        // Arrange
        var command = "get";
        string[] args = ["get", "/status"];
        using var outputWriter = new StringWriter();
        var handler = new FakeHttpMessageHandler();
        using var client = new HttpClient(handler);
        string exeName = Path.GetFileName(Process.GetCurrentProcess().MainModule?.FileName ?? "AgentExcel.exe");

        handler.HandlerFunc = (req) =>
        {
            throw new HttpRequestException("Connection refused");
        };

        // Act
        await CliCommandHandler.ExecuteApiRequest(command, args, "127.0.0.1", client, null, outputWriter, 8080);

        // Assert
        string output = outputWriter.ToString();
        Assert.That(output, Does.Contain($"Error: Daemon server is not running. Please start the server first by running '{exeName} start'."));
    }
}
