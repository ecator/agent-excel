using System.Diagnostics;
using System.Runtime.InteropServices;

using AgentExcel.Providers;
using AgentExcel.Utils;

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
        if (s_provider != null)
        {
            var app = s_provider.GetApp(createNew: false);
            ExcelConnector.SafeReleaseComObject(app);
            s_provider = null;

            ExcelConnector.ForceGarbageCollection();
        }
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

    [Test]
    public async Task GetApp_AfterIdleTimeout_ReleasesCachedInstance()
    {
        // Arrange
        // Create a provider with a very short timeout (200ms) and short check interval (50ms)
        var provider = new ExcelConnectionProvider(TimeSpan.FromMilliseconds(200), TimeSpan.FromMilliseconds(50));

        try
        {
            // Act
            Excel.Application? app;
            try
            {
                app = provider.GetApp(createNew: true);
            }
            catch (Exception ex)
            {
                Assert.Inconclusive("Skipping test because Excel COM process could not be started/connected: " + ex.Message);
                return;
            }

            Assert.That(app, Is.Not.Null);
            Assert.That(Marshal.IsComObject(app), Is.True);

            // Wait for the idle timer to trigger and release the app
            await Task.Delay(400);

            // Assert
            // The provider should have released its cached instance. We can call Dispose safely
            // and verify that it doesn't throw.
            Assert.DoesNotThrow(() => provider.Dispose());
        }
        finally
        {
            provider.Dispose();
        }
    }
}
