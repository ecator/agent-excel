using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace AgentExcel.Utils;

/// <summary>
/// Helper utilities for application information.
/// </summary>
internal static class AppInfoHelper
{
    /// <summary>
    /// Gets the file name of the current executable.
    /// </summary>
    public static string GetExeName()
    {
        return Path.GetFileName(Process.GetCurrentProcess().MainModule?.FileName ?? "AgentExcel.exe");
    }

    /// <summary>
    /// Gets the informational version of the application.
    /// </summary>
    public static string GetAppVersion()
    {
        var assembly = typeof(AppInfoHelper).Assembly;
        return assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion
            ?? assembly.GetName().Version?.ToString()
            ?? "unknown";
    }
}
