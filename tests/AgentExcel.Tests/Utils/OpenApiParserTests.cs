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
    ""/range/read"": {
      ""get"": {
        ""tags"": [""Range""],
        ""summary"": ""Read range data""
      }
    },
    ""/range/write"": {
      ""post"": {
        ""tags"": [""Range""],
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

        var readEp = catalog.Endpoints.FirstOrDefault(e => e.Path == "/range/read");
        Assert.That(readEp, Is.Not.Null);
        Assert.That(readEp!.Tag, Is.EqualTo("Range"));
        Assert.That(readEp.Method, Is.EqualTo("GET"));
        Assert.That(readEp.Summary, Is.EqualTo("Read range data"));
        Assert.That(readEp.RequestSchema, Is.Null);

        var writeEp = catalog.Endpoints.FirstOrDefault(e => e.Path == "/range/write");
        Assert.That(writeEp, Is.Not.Null);
        Assert.That(writeEp!.Tag, Is.EqualTo("Range"));
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
        Assert.That(output, Does.Contain("=== Range ==="));
        Assert.That(output, Does.Contain("[GET] /range/read - Read range data"));
        Assert.That(output, Does.Contain("[POST] /range/write - Write data to range"));
        Assert.That(output, Does.Not.Contain("Body:"));
        Assert.That(output, Does.Not.Contain("\"range\":"));
    }

    [Test]
    public void FormatCatalog_OnlyGroupFilterNoGlob_ShowsGroupEndpointsWithoutDetails()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "range"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Not.Contain("=== Charts ==="));
        Assert.That(output, Does.Contain("=== Range ==="));
        Assert.That(output, Does.Contain("[GET] /range/read - Read range data"));
        Assert.That(output, Does.Contain("[POST] /range/write - Write data to range"));
        Assert.That(output, Does.Not.Contain("Body:"));
    }

    [Test]
    public void FormatCatalog_GroupAndEndpointFilterWithGlob_ShowsMatchingDetails()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "range", "*write*"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Not.Contain("=== Charts ==="));
        Assert.That(output, Does.Not.Contain("[GET] /range/read"));
        Assert.That(output, Does.Contain("=== Range ==="));
        Assert.That(output, Does.Contain("[POST] /range/write - Write data to range"));
        Assert.That(output, Does.Contain("  Body:"));
        Assert.That(output, Does.Contain("\"range\": \"string\" // (required) The range coordinate"));
    }

    [Test]
    public void FormatCatalog_GroupAndEndpointFilterNoGlob_ShowsMatchingDetails()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "range", "/range/write"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Not.Contain("=== Charts ==="));
        Assert.That(output, Does.Not.Contain("[GET] /range/read"));
        Assert.That(output, Does.Contain("=== Range ==="));
        Assert.That(output, Does.Contain("[POST] /range/write - Write data to range"));
        Assert.That(output, Does.Contain("  Body:"));
        Assert.That(output, Does.Contain("\"range\": \"string\" // (required) The range coordinate"));
    }

    [Test]
    public void FormatCatalog_GroupAndStarGlob_ShowsAllGroupEndpointsWithDetails()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "range", "*"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Not.Contain("=== Charts ==="));
        Assert.That(output, Does.Contain("=== Range ==="));
        Assert.That(output, Does.Contain("[GET] /range/read"));
        Assert.That(output, Does.Contain("[POST] /range/write"));
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
        Assert.That(output, Does.Contain("=== Range ==="));
        Assert.That(output, Does.Contain("[GET] /range/read"));
        Assert.That(output, Does.Contain("[POST] /range/write"));
        Assert.That(output, Does.Contain("  Body:"));
    }

    [Test]
    public void FormatCatalog_FilterCaseInsensitivity_MatchesCorrectly()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "RANGE", "*WRITE*"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Contain("=== Range ==="));
        Assert.That(output, Does.Contain("[POST] /range/write - Write data to range"));
        Assert.That(output, Does.Contain("  Body:"));
    }

    [Test]
    public void FormatCatalog_GroupFilterWithGlobButNoEndpointFilter_DoesNotShowDetails()
    {
        // Arrange
        using var catalog = OpenApiParser.Parse(MockSwaggerJson);
        string[] args = ["api", "ra*"];

        // Act
        string output = OpenApiParser.FormatCatalog(catalog, args);

        // Assert
        Assert.That(output, Does.Contain("=== Range ==="));
        Assert.That(output, Does.Not.Contain("  Body:"));
    }
}

