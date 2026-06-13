using System.Diagnostics;
using System.Runtime.InteropServices;

using AgentExcel.Providers;

using NUnit.Framework;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Tests.Providers;

[TestFixture]
[Category("COM")]
public class ExcelConnectionProviderTests : BaseTests
{
#pragma warning disable NUnit1032
    private static ExcelConnectionProvider? s_provider;
#pragma warning restore NUnit1032

    [OneTimeSetUp]
    public void OneTimeSetup()
    {
        s_provider = new ExcelConnectionProvider();
    }

    [OneTimeTearDown]
    public void OneTimeTeardown()
    {
        s_provider = null;
    }

    [Test]
    public void GetApp_WithCreateNewFalse_WhenNoExcelRunning_ReturnsNullOrActive()
    {
        // Arrange
        bool isExcelRunning = Process.GetProcessesByName("EXCEL").Length > 0;

        // Act & Assert
        if (!isExcelRunning)
        {
            Excel.Application? app = s_provider!.GetApp(createNew: false);
            Assert.That(app, Is.Null);
        }
        else
        {
            Excel.Application? app = null;
            try
            {
                app = s_provider!.GetApp(createNew: false);
                Assert.That(app, Is.Not.Null);
                Assert.That(Marshal.IsComObject(app), Is.True);
            }
            finally
            {
                // Note: provider.Dispose will release it, but we can verify it was obtained
            }
        }
    }

    [Test]
    public void GetApp_WithCreateNewTrue_AlwaysReturnsApplicationInstance()
    {
        // Arrange & Act
        Excel.Application? app = null;
        try
        {
            app = s_provider!.GetApp(createNew: true);

            // Assert
            Assert.That(app, Is.Not.Null);
            Assert.That(Marshal.IsComObject(app), Is.True);
            Assert.That(app.Visible, Is.True);
        }
        finally
        {
            if (app != null)
            {
                // If we created a new instance and it was not running before, we can close it
                // but let's let the provider dispose handles
            }
        }
    }
}
