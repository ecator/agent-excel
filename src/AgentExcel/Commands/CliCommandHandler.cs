using System.Diagnostics;
using System.Net;
using System.Net.Sockets;

using AgentExcel.Utils;

namespace AgentExcel.Commands;

public static class CliCommandHandler
{
    public static void StartBackgroundServer(string listenHost)
    {
        int? existingPort = ProcessProber.FindRunningServerPort(); if (existingPort != null)
        {
            Console.WriteLine($"Server already running at http://{listenHost}:{existingPort}");
            return;
        }

        int finalPort = GetAvailablePort();
        string exePath = Process.GetCurrentProcess().MainModule?.FileName
            ?? throw new InvalidOperationException("Failed to retrieve the current executable path.");

        ProcessStartInfo startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = $"--run-server {finalPort}",
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        Process.Start(startInfo);
        Console.WriteLine($"Server started at http://{listenHost}:{finalPort}");
        Console.WriteLine($"Swagger UI available at http://{listenHost}:{finalPort}/swagger/index.html");
        Console.WriteLine($"OpenAPI Specification available at http://{listenHost}:{finalPort}/swagger/v1/swagger.json");
    }

    public static async Task StopBackgroundServer(string listenHost)
    {
        int? port = ProcessProber.FindRunningServerPort();
        if (port == null)
        {
            Console.WriteLine("Server is not running.");
            return;
        }

        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
            await client.PostAsync($"http://{listenHost}:{port}/exit", null);
            Console.WriteLine("Server stopped.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to stop the server: {ex.Message}");
        }
    }

    public static async Task ShowStatus(string listenHost)
    {
        int? port = ProcessProber.FindRunningServerPort();
        if (port == null)
        {
            Console.WriteLine("Status: Not running");
            return;
        }

        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(2) };
            var response = await client.GetAsync($"http://{listenHost}:{port}/status");
            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync();
                Console.WriteLine(json);
            }
            else
            {
                Console.WriteLine($"Server returned error code: {response.StatusCode}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error querying server status: {ex.Message}");
        }
    }

    public static void OpenSwaggerPage(string listenHost)
    {
        int? port = ProcessProber.FindRunningServerPort();
        if (port == null)
        {
            Console.WriteLine("Server is not running. Starting it now...");
            StartBackgroundServer(listenHost);
            for (int i = 0; i < 10; i++)
            {
                Thread.Sleep(500);
                port = ProcessProber.FindRunningServerPort();
                if (port != null) break;
            }
        }

        if (port != null)
        {
            string url = $"http://{listenHost}:{port}/swagger/index.html";
            Console.WriteLine($"Opening Swagger UI: {url}");
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to open browser: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("Failed to start server or detect the listening port.");
        }
    }

    private static int GetAvailablePort()
    {
        TcpListener listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        int port = ((IPEndPoint)listener.LocalEndpoint).Port;
        listener.Stop();
        return port;
    }
}
