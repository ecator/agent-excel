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

    // Timer to periodically check for inactivity and release Excel COM resources
    private readonly Timer? _idleTimer;

    // Timestamp of the last activity
    private DateTime _lastActivityTime = DateTime.UtcNow;

    // Inactivity timeout threshold
    private readonly TimeSpan _idleTimeout;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelConnectionProvider"/> class.
    /// </summary>
    /// <param name="idleTimeout">The idle timeout after which Excel will be released. Defaults to 5 minutes.</param>
    /// <param name="checkInterval">The interval at which the idle timer checks for inactivity. Defaults to 1 minute.</param>
    public ExcelConnectionProvider(TimeSpan? idleTimeout = null, TimeSpan? checkInterval = null)
    {
        _idleTimeout = idleTimeout ?? TimeSpan.FromMinutes(5);
        var interval = checkInterval ?? TimeSpan.FromMinutes(1);
        _idleTimer = new Timer(CheckIdle, null, interval, interval);
    }

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
        NotifyActivity();

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

    /// <inheritdoc />
    public void NotifyActivity()
    {
        lock (_lock)
        {
            _lastActivityTime = DateTime.UtcNow;
        }
    }

    /// <summary>
    /// Gracefully releases cached Excel COM objects and disposes the provider.
    /// If there are no open workbooks, it attempts to close the Excel application.
    /// Explicitly releases COM references to avoid leaving zombie Excel processes.
    /// </summary>
    public void Dispose()
    {
        lock (_lock)
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
        }

        // Dispose the timer outside the lock to avoid deadlock if timer callback is waiting on the lock
        _idleTimer?.Dispose();

        lock (_lock)
        {
            ReleaseAppInstance(quitIfNoWorkbooks: true);
        }

        GC.SuppressFinalize(this);
    }

    private void ReleaseAppInstance(bool quitIfNoWorkbooks)
    {
        Excel.Application? app = _app;
        if (app != null)
        {
            Excel.Workbooks? wbs = null;
            try
            {
                // Retrieve the Workbooks collection to check if Excel can be closed safely
                wbs = app.Workbooks;
                if (quitIfNoWorkbooks && wbs.Count == 0)
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
            }
            catch
            {
                // Ignore exceptions during workbook count check or edit modes
            }
            finally
            {
                // Explicitly release COM objects in reverse order of creation
                if (wbs != null)
                {
                    SafeReleaseComObject(wbs);
                }
                SafeReleaseComObject(app);
                _app = null;
            }
        }
    }

    private void CheckIdle(object? state)
    {
        if (_disposed)
        {
            return;
        }

        lock (_lock)
        {
            if (_app == null)
            {
                return;
            }

            if (DateTime.UtcNow - _lastActivityTime >= _idleTimeout)
            {
                ReleaseAppInstance(quitIfNoWorkbooks: true);
            }
        }
    }

    /// <inheritdoc />
    public void SafeReleaseComObject(object? obj)
    {
        ExcelConnector.SafeReleaseComObject(obj);
    }
}
