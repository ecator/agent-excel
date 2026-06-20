using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;

using AgentExcel.Commands;
using AgentExcel.Endpoints;
using AgentExcel.Middlewares;
using AgentExcel.Providers;
using AgentExcel.Services;
using AgentExcel.Utils;

using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.OpenApi.Models;

namespace AgentExcel;

class Program
{
    const string MutexName = "Global\\AgentExcelServerMutex";
    const string ListenHost = "127.0.0.1";

    static async Task Main(string[] args)
    {
        if (args.Length > 0)
        {
            string command = args[0].ToLower();

            if (command == "start")
            {
                CliCommandHandler.StartBackgroundServer(ListenHost);
                return;
            }
            else if (command == "stop")
            {
                await CliCommandHandler.StopBackgroundServer(ListenHost);
                return;
            }
            else if (command == "status")
            {
                await CliCommandHandler.ShowStatus(ListenHost);
                return;
            }
            else if (command == "swagger" || command == "ui")
            {
                CliCommandHandler.OpenSwaggerPage(ListenHost);
                return;
            }
            else if (command == "version")
            {
                Console.WriteLine(AppInfoHelper.GetAppVersion());
                return;
            }
            else if (command == "get" || command == "post")
            {
                await CliCommandHandler.ExecuteApiRequest(command, args, ListenHost);
                return;
            }
            else if (command == "api")
            {
                await CliCommandHandler.ShowApiCatalog(ListenHost, args);
                return;
            }
            else if (command == "--run-server")
            {
                int port = int.Parse(args[1]);
                RunServer(port);
                return;
            }
        }

        var exeName = AppInfoHelper.GetExeName();
        Console.WriteLine("AgentExcel - AI Power for Excel");
        Console.WriteLine("Usage:");
        Console.WriteLine($"  {exeName} start                     (Start daemon)");
        Console.WriteLine($"  {exeName} stop                      (Stop daemon)");
        Console.WriteLine($"  {exeName} status                    (Show status)");
        Console.WriteLine($"  {exeName} swagger                   (Open Swagger API testing page)");
        Console.WriteLine($"  {exeName} api [group] [endpoint]    (Show API catalog / OpenAPI info with filtering)");
        Console.WriteLine($"  {exeName} version                   (Show version)");
        Console.WriteLine($"  {exeName} get <endpoint> [--stdin] [body_param]");
        Console.WriteLine($"  {exeName} post <endpoint> [--stdin] [body_param]");
    }

    static void RunServer(int port)
    {
        using var mutex = new Mutex(true, MutexName, out bool createdNew);
        if (!createdNew) return;

        var builder = WebApplication.CreateBuilder();

        builder.Services.AddEndpointsApiExplorer();
        builder.Services.AddSwaggerGen(c =>
        {
            c.SwaggerDoc("v1", new OpenApiInfo { Title = "AgentExcel API", Version = "v1" });
            c.SchemaFilter<RequiredSchemaFilter>();
            var xmlFile = $"{System.Reflection.Assembly.GetExecutingAssembly().GetName().Name}.xml";
            var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
            if (File.Exists(xmlPath))
            {
                c.IncludeXmlComments(xmlPath);
            }
        });

        // Register core COM manager and domain-specific services
        builder.Services.AddSingleton<IExcelConnectionProvider, ExcelConnectionProvider>();
        builder.Services.AddSingleton<WorkbookService>();
        builder.Services.AddSingleton<SheetService>();
        builder.Services.AddSingleton<RangeService>();
        builder.Services.AddSingleton<ShapeService>();
        builder.Services.AddSingleton<ChartService>();
        builder.Services.AddSingleton<ExportService>();
        builder.Services.AddSingleton<PivotTableService>();
        builder.Services.AddSingleton<MacroService>();
        builder.Services.AddSingleton<ValidationService>();
        builder.Services.AddSingleton<SystemService>();

        var app = builder.Build();

        app.UseMiddleware<ExceptionHandlingMiddleware>();

        app.UseSwagger();
        app.UseSwaggerUI();

        // Register and map all application endpoints via extension method
        app.MapSystemEndpoints(port, ListenHost);
        app.MapWorkbookEndpoints();
        app.MapSheetEndpoints();
        app.MapRangeEndpoints();
        app.MapShapeEndpoints();
        app.MapChartEndpoints();
        app.MapExportEndpoints();
        app.MapPivotTableEndpoints();
        app.MapMacroEndpoints();
        app.MapValidationEndpoints();

        // Retrieve connection provider and register system shutdown handler
        var connectionProvider = app.Services.GetRequiredService<IExcelConnectionProvider>();
        RegisterConsoleCtrlHandler(connectionProvider);

        app.Run($"http://{ListenHost}:{port}");
    }

    private static ConsoleCtrlDelegate? s_consoleCtrlHandler;

    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetConsoleCtrlHandler(ConsoleCtrlDelegate handler, [MarshalAs(UnmanagedType.Bool)] bool add);

    private delegate bool ConsoleCtrlDelegate(CtrlType sig);

    private enum CtrlType
    {
        CTRL_C_EVENT = 0,
        CTRL_BREAK_EVENT = 1,
        CTRL_CLOSE_EVENT = 2,
        CTRL_LOGOFF_EVENT = 5,
        CTRL_SHUTDOWN_EVENT = 6
    }

    private static void RegisterConsoleCtrlHandler(IExcelConnectionProvider provider)
    {
        s_consoleCtrlHandler = (sig) =>
        {
            switch (sig)
            {
                case CtrlType.CTRL_C_EVENT:
                case CtrlType.CTRL_CLOSE_EVENT:
                case CtrlType.CTRL_LOGOFF_EVENT:
                case CtrlType.CTRL_SHUTDOWN_EVENT:
                    try
                    {
                        provider.Dispose();
                    }
                    catch
                    {
                        // Ignore exceptions during cleanup
                    }
                    break;
            }
            return false;
        };

        SetConsoleCtrlHandler(s_consoleCtrlHandler, true);
    }
}
