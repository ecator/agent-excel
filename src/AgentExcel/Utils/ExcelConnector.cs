using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using Microsoft.Win32;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Utils;

/// <summary>
/// Excel instance connector, responsible for probing and attaching to running Excel instances via COM interface.
/// </summary>
internal static class ExcelConnector
{
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject([In] ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

    /// <summary>
    /// Ensures that a running Excel instance is obtained (prefers attaching to an existing process; launches the excel.exe process and connects via COM if none exists and createNew is true).
    /// </summary>
    /// <param name="createNew">Whether to create/start a new Excel application process if none is running. Defaults to false.</param>
    /// <returns>Excel Application instance, or null if no running instance is found and createNew is false</returns>
    public static Excel.Application? EnsureExcelApplication(bool createNew = false)
    {
        Excel.Application? app = GetRunningExcelApplication();
        if (app == null && createNew)
        {
            string? excelPath = GetExcelPathFromRegistry();
            if (string.IsNullOrEmpty(excelPath))
            {
                excelPath = "excel.exe";
            }

            try
            {
                using var process = Process.Start(new ProcessStartInfo
                {
                    FileName = excelPath,
                    UseShellExecute = true
                });

                // Wait for Excel to start and register in ROT
                int retries = 50; // Max 5 seconds (50 * 100ms)
                for (int i = 0; i < retries; i++)
                {
                    System.Threading.Thread.Sleep(100);
                    app = GetRunningExcelApplication();
                    if (app != null)
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to start Excel process: {ex.Message}", ex);
            }

            if (app == null)
            {
                throw new Exception("Excel process started but failed to connect via COM.");
            }
        }

        if (app != null && createNew)
        {
            // Note: The newly created/started or attached instance must be set to Visible = true to comply with the "shared/user-aware mode".
            try
            {
                if (!app.Visible)
                {
                    app.Visible = true;
                }
            }
            catch (COMException)
            {
                // Ignore if setting visibility fails
            }
        }
        return app;
    }

    private static string? GetExcelPathFromRegistry()
    {
        string[] appPaths =
        [
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
            @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
        ];

        RegistryKey[] roots = [Registry.CurrentUser, Registry.LocalMachine];
        RegistryView[] views = [RegistryView.Registry64, RegistryView.Registry32];

        foreach (var root in roots)
        {
            foreach (var view in views)
            {
                try
                {
                    using var baseKey = RegistryKey.OpenBaseKey(
                        root == Registry.CurrentUser ? RegistryHive.CurrentUser : RegistryHive.LocalMachine,
                        view);

                    foreach (var path in appPaths)
                    {
                        using var key = baseKey.OpenSubKey(path);
                        var val = key?.GetValue(null)?.ToString();
                        if (!string.IsNullOrEmpty(val) && System.IO.File.Exists(val))
                        {
                            return val;
                        }
                    }
                }
                catch
                {
                    // Ignore and continue
                }
            }
        }

        string[] officeVersions = ["16.0", "15.0", "14.0"];
        foreach (var root in roots)
        {
            foreach (var view in views)
            {
                try
                {
                    using var baseKey = RegistryKey.OpenBaseKey(
                        root == Registry.CurrentUser ? RegistryHive.CurrentUser : RegistryHive.LocalMachine,
                        view);

                    foreach (var version in officeVersions)
                    {
                        string path = $@"SOFTWARE\Microsoft\Office\{version}\Excel\InstallRoot";
                        using var key = baseKey.OpenSubKey(path);
                        var dir = key?.GetValue("Path")?.ToString();
                        if (!string.IsNullOrEmpty(dir))
                        {
                            string fullPath = System.IO.Path.Combine(dir, "excel.exe");
                            if (System.IO.File.Exists(fullPath))
                            {
                                return fullPath;
                            }
                        }
                    }
                }
                catch
                {
                    // Ignore and continue
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Tries to get an existing Excel instance (does not throw exceptions, returns null if not found).
    /// </summary>
    public static Excel.Application? GetRunningExcelApplication()
    {
        try
        {
            Guid clsid;
            CLSIDFromProgID("Excel.Application", out clsid);
            GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
            if (obj != null) return obj as Excel.Application;
        }
        catch
        {
            // Ignore error; it is a common case when no running instance is found.
        }
        return null;
    }

    /// <summary>
    /// Safely releases a COM object if it is not null and is a valid COM object.
    /// </summary>
    public static void SafeReleaseComObject(object? obj)
    {
        if (obj != null && Marshal.IsComObject(obj))
        {
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch
            {
                // Ignore release errors
            }
        }
    }

    /// <summary>
    /// Explicitly invokes GC.Collect to clean up any remaining underlying COM components.
    /// </summary>
    public static void ForceGarbageCollection()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
