using AgentExcel.Utils;

using NUnit.Framework;

namespace AgentExcel.Tests.Utils;

[TestFixture]
[Category("Unit")]
public class MarkdownTableHelperTests : BaseTests
{
    [Test]
    public void ToMarkdownTable_WithValidInputs_FormatsCorrectly()
    {
        // Arrange
        var headers = new[] { "Col1", "Col2" };
        var body = new[,] { { "Row1|1", "Row1|2" }, { "Row2\\1", "Row2\n2" } };

        // Act
        var result = MarkdownTableHelper.ToMarkdownTable(headers, body);

        // Assert
        var expected = "| Col1 | Col2 |\r\n| --- | --- |\r\n| Row1\\|1 | Row1\\|2 |\r\n| Row2\\\\1 | Row2<br>2 |";
        Assert.That(result.Replace("\r\n", "\n"), Is.EqualTo(expected.Replace("\r\n", "\n")));
    }

    [Test]
    public void ToMarkdownTable_WithColumnMismatch_ThrowsArgumentException()
    {
        // Arrange
        var headers = new[] { "Col1", "Col2" };
        var body = new[,] { { "Row1" } }; // 1 column body vs 2 column headers

        // Act & Assert
        Assert.Throws<ArgumentException>(() => MarkdownTableHelper.ToMarkdownTable(headers, body));
    }
}
