using System.Text.Json;

using AgentExcel.Utils;

using NUnit.Framework;

namespace AgentExcel.Tests.Utils;

[TestFixture]
[Category("Unit")]
public class JsonHelperTests
{
    [Test]
    public void ConvertJsonValue_WithStringElement_ReturnsString()
    {
        // Arrange
        using var doc = JsonDocument.Parse("\"hello\"");
        var element = doc.RootElement;

        // Act
        var result = JsonHelper.ConvertJsonValue(element);

        // Assert
        Assert.That(result, Is.EqualTo("hello"));
    }

    [Test]
    public void ConvertJsonValue_WithIntegerElement_ReturnsInt()
    {
        // Arrange
        using var doc = JsonDocument.Parse("42");
        var element = doc.RootElement;

        // Act
        var result = JsonHelper.ConvertJsonValue(element);

        // Assert
        Assert.That(result, Is.EqualTo(42));
    }

    [Test]
    public void ConvertJsonValue_WithDoubleElement_ReturnsDouble()
    {
        // Arrange
        using var doc = JsonDocument.Parse("3.14159");
        var element = doc.RootElement;

        // Act
        var result = JsonHelper.ConvertJsonValue(element);

        // Assert
        Assert.That(result, Is.EqualTo(3.14159));
    }

    [Test]
    public void ConvertJsonValue_WithTrueElement_ReturnsTrue()
    {
        // Arrange
        using var doc = JsonDocument.Parse("true");
        var element = doc.RootElement;

        // Act
        var result = JsonHelper.ConvertJsonValue(element);

        // Assert
        Assert.That(result, Is.True);
    }

    [Test]
    public void ConvertJsonValue_WithFalseElement_ReturnsFalse()
    {
        // Arrange
        using var doc = JsonDocument.Parse("false");
        var element = doc.RootElement;

        // Act
        var result = JsonHelper.ConvertJsonValue(element);

        // Assert
        Assert.That(result, Is.False);
    }

    [Test]
    public void ConvertJsonValue_WithNullElement_ReturnsNull()
    {
        // Arrange
        using var doc = JsonDocument.Parse("null");
        var element = doc.RootElement;

        // Act
        var result = JsonHelper.ConvertJsonValue(element);

        // Assert
        Assert.That(result, Is.Null);
    }

    [Test]
    public void ConvertJsonValue_WithArrayElement_ReturnsObjectArray()
    {
        // Arrange
        using var doc = JsonDocument.Parse("[1, \"test\", true]");
        var element = doc.RootElement;

        // Act
        var result = JsonHelper.ConvertJsonValue(element);

        // Assert
        Assert.That(result, Is.InstanceOf<object?[]>());
        var arr = (object?[])result!;
        Assert.That(arr, Has.Length.EqualTo(3));
        Assert.That(arr[0], Is.EqualTo(1));
        Assert.That(arr[1], Is.EqualTo("test"));
        Assert.That(arr[2], Is.True);
    }

    [Test]
    public void ConvertJsonValue_WithNestedArrayElement_ReturnsNestedObjectArrays()
    {
        // Arrange
        using var doc = JsonDocument.Parse("[[1, 2], [\"a\", \"b\"]]");
        var element = doc.RootElement;

        // Act
        var result = JsonHelper.ConvertJsonValue(element);

        // Assert
        Assert.That(result, Is.InstanceOf<object?[]>());
        var arr = (object?[])result!;
        Assert.That(arr, Has.Length.EqualTo(2));

        var subArr1 = (object?[])arr[0]!;
        Assert.That(subArr1, Has.Length.EqualTo(2));
        Assert.That(subArr1[0], Is.EqualTo(1));
        Assert.That(subArr1[1], Is.EqualTo(2));

        var subArr2 = (object?[])arr[1]!;
        Assert.That(subArr2, Has.Length.EqualTo(2));
        Assert.That(subArr2[0], Is.EqualTo("a"));
        Assert.That(subArr2[1], Is.EqualTo("b"));
    }

    [Test]
    public void ConvertJsonValue_WithNonJsonElementValue_ReturnsOriginalValue()
    {
        // Arrange
        var original = new object();

        // Act
        var result = JsonHelper.ConvertJsonValue(original);

        // Assert
        Assert.That(result, Is.SameAs(original));
    }
}
