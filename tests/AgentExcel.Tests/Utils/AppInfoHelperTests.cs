using System;

using AgentExcel.Utils;

using NUnit.Framework;

namespace AgentExcel.Tests.Utils;

[TestFixture]
[Category("Unit")]
public class AppInfoHelperTests : BaseTests
{
    [Test]
    public void GetExeName_ReturnsNonEmptyStringEndingWithExe()
    {
        // Act
        string exeName = AppInfoHelper.GetExeName();

        // Assert
        Assert.That(exeName, Is.Not.Null.And.Not.Empty);
        Assert.That(exeName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase), Is.True);
    }

    [Test]
    public void GetAppVersion_ReturnsNonEmptyString()
    {
        // Act
        string version = AppInfoHelper.GetAppVersion();

        // Assert
        Assert.That(version, Is.Not.Null.And.Not.Empty);
    }
}
