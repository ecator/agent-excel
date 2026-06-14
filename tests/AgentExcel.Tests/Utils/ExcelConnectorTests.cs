using System.Runtime.InteropServices;

using AgentExcel.Utils;

using NUnit.Framework;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Tests.Utils;

[TestFixture]
[Category("COM")]
public class ExcelConnectorTests : BaseTests
{
    [Test]
    public void EnsureExcelApplication_WithCreateNewFalse_ReturnsExistingOrNull()
    {
        // Arrange & Act
        Excel.Application? app = ExcelConnector.EnsureExcelApplication(createNew: false);

        // Assert
        if (app != null)
        {
            Assert.That(Marshal.IsComObject(app), Is.True);
            ExcelConnector.SafeReleaseComObject(app);
            ExcelConnector.ForceGarbageCollection();
        }
        else
        {
            Assert.That(app, Is.Null);
        }
    }

    [Test]
    public void EnsureExcelApplication_WithCreateNewTrue_ReturnsInstance()
    {
        // Arrange & Act
        Excel.Application? app = null;
        try
        {
            app = ExcelConnector.EnsureExcelApplication(createNew: true);

            // Assert
            Assert.That(app, Is.Not.Null);
            Assert.That(Marshal.IsComObject(app), Is.True);
            Assert.That(app.Visible, Is.True);
        }
        finally
        {
            ExcelConnector.SafeReleaseComObject(app);
            ExcelConnector.ForceGarbageCollection();
        }
    }
}
