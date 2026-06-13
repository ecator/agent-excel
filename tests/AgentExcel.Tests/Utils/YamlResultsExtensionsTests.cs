using System.IO;
using System.Text;
using System.Threading.Tasks;

using AgentExcel.Utils;

using Microsoft.AspNetCore.Http;

using NUnit.Framework;

namespace AgentExcel.Tests.Utils;

[TestFixture]
[Category("Unit")]
public class YamlResultsExtensionsTests : BaseTests
{
    [Test]
    public async Task Yaml_WithValidObject_SerializesToYamlAndSetsContentType()
    {
        // Arrange
        var testData = new TestPayload
        {
            Name = "AgentExcel",
            Version = "1.0",
            IsActive = true
        };

        var httpContext = new DefaultHttpContext();
        using var responseStream = new MemoryStream();
        httpContext.Response.Body = responseStream;

        // Act
        var result = Microsoft.AspNetCore.Http.Results.Extensions.Yaml(testData);
        await result.ExecuteAsync(httpContext);

        // Assert
        Assert.That(httpContext.Response.ContentType, Is.EqualTo("text/yaml; charset=utf-8"));
        Assert.That(httpContext.Response.StatusCode, Is.EqualTo(200));

        responseStream.Position = 0;
        using var reader = new StreamReader(responseStream, Encoding.UTF8);
        string yamlOutput = await reader.ReadToEndAsync();

        Assert.That(yamlOutput, Does.Contain("Name: AgentExcel"));
        Assert.That(yamlOutput, Does.Contain("Version: 1.0"));
        Assert.That(yamlOutput, Does.Contain("IsActive: true"));
    }

    private class TestPayload
    {
        public string Name { get; set; } = string.Empty;
        public string Version { get; set; } = string.Empty;
        public bool IsActive { get; set; }
    }
}
