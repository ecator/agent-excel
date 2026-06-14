using AgentExcel.Utils;

using NUnit.Framework;

namespace AgentExcel.Tests.Utils;

[TestFixture]
[Category("Unit")]
public class ColorHelperTests : BaseTests
{
    [Test]
    public void HexToOleColor_With6DigitHex_ConvertsCorrectly()
    {
        // Arrange
        string redHex = "#FF0000";
        string greenHex = "00FF00";
        string blueHex = "#0000FF";

        // Act
        int redOle = ColorHelper.HexToOleColor(redHex);
        int greenOle = ColorHelper.HexToOleColor(greenHex);
        int blueOle = ColorHelper.HexToOleColor(blueHex);

        // Assert
        Assert.That(redOle, Is.EqualTo(255)); // 0x0000FF
        Assert.That(greenOle, Is.EqualTo(65280)); // 0x00FF00
        Assert.That(blueOle, Is.EqualTo(16711680)); // 0xFF0000
    }

    [Test]
    public void HexToOleColor_WithShortHex_PadsTrailingZerosAndConvertsCorrectly()
    {
        // Arrange
        string shortGreenHex = "#00ff"; // Should pad to "00ff00"
        string shortRedHex = "ff";     // Should pad to "ff0000"

        // Act
        int greenOle = ColorHelper.HexToOleColor(shortGreenHex);
        int redOle = ColorHelper.HexToOleColor(shortRedHex);

        // Assert
        Assert.That(greenOle, Is.EqualTo(65280)); // 0x00FF00
        Assert.That(redOle, Is.EqualTo(255));     // 0x0000FF
    }

    [Test]
    public void OleColorToHex_WithOleColors_ConvertsCorrectly()
    {
        // Arrange
        long redOle = 255;
        long greenOle = 65280;
        long blueOle = 16711680;

        // Act
        string redHex = ColorHelper.OleColorToHex(redOle);
        string greenHex = ColorHelper.OleColorToHex(greenOle);
        string blueHex = ColorHelper.OleColorToHex(blueOle);

        // Assert
        Assert.That(redHex, Is.EqualTo("#FF0000"));
        Assert.That(greenHex, Is.EqualTo("#00FF00"));
        Assert.That(blueHex, Is.EqualTo("#0000FF"));
    }
}
