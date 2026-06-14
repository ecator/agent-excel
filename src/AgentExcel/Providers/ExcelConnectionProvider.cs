using System.Runtime.InteropServices;

using AgentExcel.Utils;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Providers;

/// <summary>
/// Provides and manages the connection to the Excel Application COM instance.
/// This class handles thread-safe initialization, caching of the connection,
/// and proper cleanup/disposal of Excel COM resources.
/// </summary>
public class ExcelConnectionProvider : IExcelConnectionProvider
{
    // Object used for synchronization locking during lazy initialization
    private readonly object _lock = new();

    // Cached reference to the active Excel application instance
    private Excel.Application? _app;

    // Flag to track whether the provider instance has been disposed
    private bool _disposed;

    /// <summary>
    /// Retrieves the active or a newly created Excel Application instance.
    /// Uses double-check locking to ensure thread safety during lazy initialization.
    /// </summary>
    /// <param name="createNew">
    /// If true, attempts to launch a new Excel process when no active instance is found.
    /// If false, only attaches to the existing instance and does not launch a new process.
    /// </param>
    /// <returns>The active or created <see cref="Excel.Application"/> instance, or null if none is available.</returns>
    /// <exception cref="Exception">Thrown if <paramref name="createNew"/> is true but Excel fails to start.</exception>
    public Excel.Application? GetApp(bool createNew = false)
    {
        if (_app != null)
        {
            if (createNew)
            {
                try
                {
                    if (!_app.Visible)
                    {
                        _app.Visible = true;
                    }
                }
                catch (COMException)
                {
                    // Ignore if setting visibility fails
                }
            }
            return _app;
        }

        lock (_lock)
        {
            if (_app != null)
            {
                if (createNew)
                {
                    try
                    {
                        if (!_app.Visible)
                        {
                            _app.Visible = true;
                        }
                    }
                    catch (COMException)
                    {
                        // Ignore if setting visibility fails
                    }
                }
                return _app;
            }

            // Attempt to connect or launch Excel using the utility helper
            _app = ExcelConnector.EnsureExcelApplication(createNew);

            if (_app == null && createNew)
            {
                throw new Exception("Failed to start Excel. Please verify that Microsoft Office 2016 or later is installed.");
            }
            return _app;
        }
    }

    /// <summary>
    /// Gracefully releases cached Excel COM objects and disposes the provider.
    /// If there are no open workbooks, it attempts to close the Excel application.
    /// Explicitly releases COM references to avoid leaving zombie Excel processes.
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        Excel.Application? app = _app;
        if (app != null)
        {
            // Retrieve the Workbooks collection to check if Excel can be closed safely
            Excel.Workbooks wbs = app.Workbooks;
            if (wbs.Count == 0)
            {
                try
                {
                    // Only quit if there are no workbooks open to avoid disrupting the user's active session
                    app.Quit();
                }
                catch
                {
                    // Ignore exceptions on quit
                }
            }
            // Explicitly release COM objects in reverse order of creation
            SafeReleaseComObject(wbs);
            SafeReleaseComObject(app);
            _app = null;
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }

    /// <inheritdoc />
    public void SafeReleaseComObject(object? obj)
    {
        ExcelConnector.SafeReleaseComObject(obj);
    }
}
