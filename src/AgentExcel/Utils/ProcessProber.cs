using System.Diagnostics;
using System.Text.RegularExpressions;

namespace AgentExcel.Utils;

/// <summary>
/// Process and port prober.
/// </summary>
internal class ProcessProber
{
    /// <summary>
    /// Finds the port of the currently running server.
    /// </summary>
    /// <returns>The listening port, or null if not found.</returns>
    public static int? FindRunningServerPort()
    {
        string currentName = Process.GetCurrentProcess().ProcessName;
        int currentPid = Process.GetCurrentProcess().Id;

        // Find all background processes with the same name, excluding the current CLI process itself
        Process? backgroundProcess = Process.GetProcessesByName(currentName)
                                            .FirstOrDefault(p => p.Id != currentPid);

        if (backgroundProcess == null) return null;

        return GetListeningPortByPid(backgroundProcess.Id);
    }

    /// <summary>
    /// Gets the listening port by process PID.
    /// </summary>
    /// <param name="pid">The process PID.</param>
    /// <returns>The listening port, or null if not found.</returns>
    private static int? GetListeningPortByPid(int pid)
    {
        try
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "netstat",
                Arguments = "-ano",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using Process? p = Process.Start(psi);
            if (p is null) return null;
            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit();

            // Extract the local TCP port belonging to the target PID in LISTENING state
            string pattern = $@"\s+TCP\s+(?:127\.0\.0\.1|0\.0\.0\.0):(\d+)\s+.*?\s+LISTENING\s+{pid}\b";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

            Match match = regex.Match(output);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int port))
            {
                return port;
            }
        }
        catch
        {
            // Ignore permission exceptions during system probing
        }
        return null;
    }
}
