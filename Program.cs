using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.Json;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel
{
    /// <summary>
    /// 应用程序主入口点
    /// </summary>
    class Program
    {
        // 全局互斥锁，确保后台守护进程(Daemon)只有一个实例
        const string MUTEXT_NAME = "Global\\AgentExcelServerMutex";

        // 监听地址，强制绑定到localhost，避免外部访问风险
        const string LISTEN_HOST = "127.0.0.1";

        static async Task Main(string[] args)
        {
            if (args.Length > 0)
            {
                string command = args[0].ToLower();

                if (command == "start")
                {
                    StartBackgroundServer();
                    return;
                }
                else if (command == "stop")
                {
                    await StopBackgroundServer();
                    return;
                }
                else if (command == "status")
                {
                    await ShowStatus();
                    return;
                }
                else if (command == "--run-server")
                {
                    // 内部隐藏命令：真正的 Web 服务启动入口
                    int port = int.Parse(args[1]);
                    RunServer(port);
                    return;
                }
            }
            var exeName = Path.GetFileName(Process.GetCurrentProcess().MainModule.FileName);
            Console.WriteLine("help:");
            Console.WriteLine($"  {exeName} start   (启动后台服务)");
            Console.WriteLine($"  {exeName} stop    (停止后台服务)");
            Console.WriteLine($"  {exeName} status  (查看服务与Excel状态)");
        }

        /// <summary>
        /// 启动后台服务 (Daemon)
        /// </summary>
        static void StartBackgroundServer()
        {
            Console.WriteLine("准备启动 AgentExcel 后台服务...");

            // 1. 检查是否已经有服务在运行
            int? existingPort = ProcessProber.FindRunningServerPort();
            if (existingPort != null)
            {
                Console.WriteLine($"服务已经在后台运行中 (监听地址: http://{LISTEN_HOST}:{existingPort})，无需重复启动。");
                return;
            }

            // 2. 获取一个系统级绝对空闲的端口
            int finalPort = GetAvailablePort();

            string exePath = Process.GetCurrentProcess().MainModule.FileName;

            // 3. 启动后台子进程
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"--run-server {finalPort}",
                UseShellExecute = false,
                CreateNoWindow = true,      // 不创建控制台黑框
                WindowStyle = ProcessWindowStyle.Hidden
            };

            Process.Start(startInfo);
            Console.WriteLine($"服务已成功在后台启动！");
            Console.WriteLine($"监听地址: http://{LISTEN_HOST}:{finalPort}");
        }

        /// <summary>
        /// 停止后台服务
        /// </summary>
        static async Task StopBackgroundServer()
        {
            Console.WriteLine("准备停止 AgentExcel 服务...");

            int? port = ProcessProber.FindRunningServerPort();
            if (port == null)
            {
                Console.WriteLine("未探测到后台运行的服务，进程可能已关闭。");
                return;
            }

            string serverUrl = $"http://{LISTEN_HOST}:{port}";
            Console.WriteLine($"正在向后台服务 ({serverUrl}) 发送退出指令...");

            try
            {
                using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
                var response = await client.PostAsync($"{serverUrl}/exit", null);

                if (response.IsSuccessStatusCode)
                {
                    Console.WriteLine("服务已安全关闭。");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"无法优雅关闭服务 (它可能陷入了死循环): {ex.Message}");
                Environment.Exit(1);
            }
        }

        /// <summary>
        /// 查看系统状态
        /// </summary>
        static async Task ShowStatus()
        {
            Console.WriteLine("正在探测 AgentExcel 后台服务...");

            int? port = ProcessProber.FindRunningServerPort();

            if (port == null)
            {
                Console.WriteLine("状态: 未运行 (未检测到后台进程)");
                return;
            }

            Console.WriteLine($"发现后台进程！(监听地址: http://{LISTEN_HOST}:{port})");
            Console.WriteLine("正在执行深度健康检查...");

            try
            {
                using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(2) };
                var response = await client.GetAsync($"http://{LISTEN_HOST}:{port}/status");

                if (response.IsSuccessStatusCode)
                {
                    string jsonString = await response.Content.ReadAsStringAsync();
                    using JsonDocument doc = JsonDocument.Parse(jsonString);
                    JsonElement root = doc.RootElement;

                    Console.WriteLine("\n--- 服务状态报告 ---");
                    Console.WriteLine($"[服务状态] : {root.GetProperty("status").GetString()}");
                    Console.WriteLine($"[进程 PID] : {root.GetProperty("pid").GetInt32()}");
                    Console.WriteLine($"[监听地址] : http://{LISTEN_HOST}:{root.GetProperty("port").GetInt32()}");

                    bool hasExcel = root.GetProperty("excel_attached").GetBoolean();
                    Console.WriteLine($"[COM 连接] : {(hasExcel ? "已接管 Excel" : "未找到运行的 Excel")}");
                    Console.WriteLine($"[活动表格] : {root.GetProperty("active_workbook").GetString()}");
                    Console.WriteLine("------------------------\n");
                }
                else
                {
                    Console.WriteLine($"服务进程存在，但健康检查返回了异常 HTTP 状态码: {response.StatusCode}");
                    Environment.Exit(9);
                }
            }
            catch (TaskCanceledException)
            {
                Console.WriteLine("健康检查超时！服务可能被 Excel 的模态弹窗(如另存为、报错提示)阻塞了。");
                Environment.Exit(2);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"无法与服务通信: {ex.Message}");
                Environment.Exit(1);
            }
        }

        /// <summary>
        /// 运行长驻内存的 Web 服务
        /// </summary>
        /// <param name="port">监听端口</param>
        static void RunServer(int port)
        {
            // 后台进程尝试获取互斥锁，作为单例的双重保险
            using var mutex = new Mutex(true, MUTEXT_NAME, out bool createdNew);
            if (!createdNew)
            {
                return; // 如果锁已被占用，说明发生了极小概率的并发启动，静默退出
            }

            var builder = WebApplication.CreateBuilder();
            string serverUrl = $"http://localhost:{port}";
            builder.WebHost.UseUrls(serverUrl);
            var app = builder.Build();

            Excel.Application excelApp = null;
            try
            {
                // 启动时尝试接管已经激活的 Excel
                excelApp = ExcelConnector.GetRunningExcelApplication();
            }
            catch
            {

            }

            // --- 路由 1：获取详细状态 ---
            app.MapGet("/status", () =>
            {
                bool isExcelAttached = excelApp != null;
                string activeWorkbookName = "无 (可能未打开任何文件)";

                if (isExcelAttached)
                {
                    Excel.Workbooks workbooks = null;
                    Excel.Workbook activeWorkbook = null;
                    try
                    {
                        workbooks = excelApp.Workbooks;
                        if (workbooks.Count > 0)
                        {
                            activeWorkbook = excelApp.ActiveWorkbook;
                            activeWorkbookName = activeWorkbook.Name;
                        }
                    }
                    catch
                    {
                        activeWorkbookName = "获取失败 (当前单元格可能处于编辑模式，请按回车)";
                    }
                    finally
                    {
                        // 显式释放 COM 对象，避免内存泄漏和 Excel 进程残留
                        if (activeWorkbook != null) Marshal.ReleaseComObject(activeWorkbook);
                        if (workbooks != null) Marshal.ReleaseComObject(workbooks);
                    }
                }

                return Results.Ok(new
                {
                    status = "running",
                    pid = Process.GetCurrentProcess().Id,
                    port = port,
                    excel_attached = isExcelAttached,
                    active_workbook = activeWorkbookName
                });
            });

            // --- 路由 2 & 3：读写接口演示 ---
            app.MapGet("/read", (string range) =>
            {
                // 你的数据读取逻辑
                return Results.Ok(new { range = range, value = "示例数据" });
            });
            app.MapPost("/write", (WriteRequest req) =>
            {
                // 你的数据写入逻辑
                return Results.Ok(new { status = "success" });
            });

            // --- 路由 4：退出并释放资源 ---
            app.MapPost("/exit", (IHostApplicationLifetime lifetime) =>
            {
                if (excelApp != null)
                {
                    Marshal.ReleaseComObject(excelApp);
                }
                lifetime.StopApplication();
                return Results.Ok(new { status = "shutting_down" });
            });

            // 阻塞并运行服务
            app.Run();
        }

        /// <summary>
        /// 获取系统的绝对空闲端口
        /// </summary>
        /// <returns>可用端口号</returns>
        static int GetAvailablePort()
        {
            TcpListener listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            int port = ((IPEndPoint)listener.LocalEndpoint).Port;
            listener.Stop();
            return port;
        }
    }

    // 请求模型记录
    record WriteRequest(string Range, object Value);
}
