using System;
using System.Linq;

using AgentExcel.Utils;

using NUnit.Framework;

namespace AgentExcel.Tests.Utils;

[TestFixture]
[Category("Unit")]
public class OpenApiParserTests : BaseTests
{
    private const string MockSwaggerJson = @"{
  ""openapi"": ""3.0.1"",
  ""paths"": {
    ""/data/read"": {
      ""get"": {
        ""tags"": [""Data""],
        ""summary"": ""Read range data""
      }
    },
    ""/data/write-range"": {
      ""post"": {
        ""tags"": [""Data""],
        ""summary"": ""Write data to range"",
        ""requestBody"": {
          ""content"": {
            ""application/json"": {
              ""schema"": {
                ""$ref"": ""#/components/schemas/WriteRequest""
              }
            }
          }
        }
      }
    },
    ""/charts/list"": {
      ""get"": {
        ""tags"": [""Charts""],
        ""summary"": ""List all charts""
      }
    }
  },
  ""components"": {
    ""schemas"": {
      ""WriteRequest"": {
        ""type"": ""object"",
        ""properties"": {
          ""range"": {
            ""type"": ""string"",
            ""description"": ""The range coordinate""
          }
        },
        ""required"": [ ""range"" ]
      }
    }
  }
}";

    [Test]
    public void Parse_ValidJson_ReturnsPopulatedCatalog()
    {
        // Act
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);

        // Assert
        Assert.That(catalog.Endpoints, Has.Count.EqualTo(3));
        Assert.That(catalog.Schemas, Has.Count.EqualTo(1));

        var readEp = catalog.Endpoints.FirstOrDefault(e => e.Path == "/data/read");
        Assert.That(readEp, Is.Not.Null);
        Assert.That(readEp!.Tag, Is.EqualTo("Data"));
        Assert.That(readEp.Method, Is.EqualTo("GET"));
        Assert.That(readEp.Summary, Is.EqualTo("Read range data"));
        Assert.That(readEp.RequestSchema, Is.Null);

        var writeEp = catalog.Endpoints.FirstOrDefault(e => e.Path == "/data/write-range");
        Assert.That(writeEp, Is.Not.Null);
        Assert.That(writeEp!.Tag, Is.EqualTo("Data"));
        Assert.That(writeEp.Method, Is.EqualTo("POST"));
        Assert.That(writeEp.RequestSchema, Is.Not.Null);
    }

    [Test]
    public void FormatCatalog_NoArguments_ShowsAllGroupsAndEndpointsWithoutDetails()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Contain("=== Charts ==="));
        Assert.That(output, Does.Contain("[GET] /charts/list - List all charts"));
        Assert.That(output, Does.Contain("=== Data ==="));
        Assert.That(output, Does.Contain("[GET] /data/read - Read range data"));
        Assert.That(output, Does.Contain("[POST] /data/write-range - Write data to range"));
        Assert.That(output, Does.Not.Contain("Body:"));
        Assert.That(output, Does.Not.Contain("\"range\":"));
    }

    [Test]
    public void FormatCatalog_OnlyGroupFilterNoGlob_ShowsGroupEndpointsWithoutDetails()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "data"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Not.Contain("=== Charts ==="));
        Assert.That(output, Does.Contain("=== Data ==="));
        Assert.That(output, Does.Contain("[GET] /data/read - Read range data"));
        Assert.That(output, Does.Contain("[POST] /data/write-range - Write data to range"));
        Assert.That(output, Does.Not.Contain("Body:"));
    }

    [Test]
    public void FormatCatalog_GroupAndEndpointFilterWithGlob_ShowsMatchingDetails()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "data", "*range*"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Not.Contain("=== Charts ==="));
        Assert.That(output, Does.Not.Contain("[GET] /data/read"));
        Assert.That(output, Does.Contain("=== Data ==="));
        Assert.That(output, Does.Contain("[POST] /data/write-range - Write data to range"));
        Assert.That(output, Does.Contain("  Body:"));
        Assert.That(output, Does.Contain("\"range\": \"string\" // (required) The range coordinate"));
    }

    [Test]
    public void FormatCatalog_GroupAndEndpointFilterNoGlob_ShowsMatchingDetails()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "data", "/data/write-range"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Not.Contain("=== Charts ==="));
        Assert.That(output, Does.Not.Contain("[GET] /data/read"));
        Assert.That(output, Does.Contain("=== Data ==="));
        Assert.That(output, Does.Contain("[POST] /data/write-range - Write data to range"));
        Assert.That(output, Does.Contain("  Body:"));
        Assert.That(output, Does.Contain("\"range\": \"string\" // (required) The range coordinate"));
    }

    [Test]
    public void FormatCatalog_GroupAndStarGlob_ShowsAllGroupEndpointsWithDetails()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "data", "*"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Not.Contain("=== Charts ==="));
        Assert.That(output, Does.Contain("=== Data ==="));
        Assert.That(output, Does.Contain("[GET] /data/read"));
        Assert.That(output, Does.Contain("[POST] /data/write-range"));
        Assert.That(output, Does.Contain("  Body:"));
    }

    [Test]
    public void FormatCatalog_StarStarGlob_ShowsAllGroupsAndEndpointsWithDetails()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "*", "*"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Contain("=== Charts ==="));
        Assert.That(output, Does.Contain("=== Data ==="));
        Assert.That(output, Does.Contain("[GET] /data/read"));
        Assert.That(output, Does.Contain("[POST] /data/write-range"));
        Assert.That(output, Does.Contain("  Body:"));
    }

    [Test]
    public void FormatCatalog_FilterCaseInsensitivity_MatchesCorrectly()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "DATA", "*RANGE*"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Contain("=== Data ==="));
        Assert.That(output, Does.Contain("[POST] /data/write-range - Write data to range"));
        Assert.That(output, Does.Contain("  Body:"));
    }
}
