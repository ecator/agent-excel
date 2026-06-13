using System.Diagnostics;

using AgentExcel.Commands;
using AgentExcel.Endpoints;
using AgentExcel.Middlewares;
using AgentExcel.Providers;
using AgentExcel.Services;

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
            else if (command == "--run-server")
            {
                int port = int.Parse(args[1]);
                RunServer(port);
                return;
            }
        }

        var exeName = Path.GetFileName(Process.GetCurrentProcess().MainModule?.FileName ?? "AgentExcel.exe");
        Console.WriteLine("AgentExcel - AI Power for Excel");
        Console.WriteLine("Usage:");
        Console.WriteLine($"  {exeName} start   (Start daemon)");
        Console.WriteLine($"  {exeName} stop    (Stop daemon)");
        Console.WriteLine($"  {exeName} status  (Show status)");
        Console.WriteLine($"  {exeName} swagger (Open Swagger API testing page)");
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
        builder.Services.AddSingleton<DataService>();
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
        app.MapDataEndpoints();
        app.MapShapeEndpoints();
        app.MapChartEndpoints();
        app.MapExportEndpoints();
        app.MapPivotTableEndpoints();
        app.MapMacroEndpoints();
        app.MapValidationEndpoints();

        app.Run($"http://{ListenHost}:{port}");
    }
}
