using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using AgentExcel;
using Microsoft.OpenApi.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel
{
    class Program
    {
        const string MUTEXT_NAME = "Global\\AgentExcelServerMutex";
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
                else if (command == "swagger" || command == "ui")
                {
                    OpenSwaggerPage();
                    return;
                }
                else if (command == "--run-server")
                {
                    int port = int.Parse(args[1]);
                    RunServer(port);
                    return;
                }
            }
            var exeName = Path.GetFileName(Process.GetCurrentProcess().MainModule.FileName);
            Console.WriteLine("AgentExcel - AI Power for Excel");
            Console.WriteLine($"Usage:");
            Console.WriteLine($"  {exeName} start   (Start daemon)");
            Console.WriteLine($"  {exeName} stop    (Stop daemon)");
            Console.WriteLine($"  {exeName} status  (Show status)");
            Console.WriteLine($"  {exeName} swagger (Open Swagger API testing page)");
        }

        static void StartBackgroundServer()
        {
            int? existingPort = ProcessProber.FindRunningServerPort();
            if (existingPort != null)
            {
                Console.WriteLine($"Server already running at http://{LISTEN_HOST}:{existingPort}");
                return;
            }

            int finalPort = GetAvailablePort();
            string exePath = Process.GetCurrentProcess().MainModule.FileName;

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"--run-server {finalPort}",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            Process.Start(startInfo);
            Console.WriteLine($"Server started at http://{LISTEN_HOST}:{finalPort}");
            Console.WriteLine($"Swagger UI available at http://{LISTEN_HOST}:{finalPort}/swagger/index.html");
            Console.WriteLine($"OpenAPI Specification available at http://{LISTEN_HOST}:{finalPort}/swagger/v1/swagger.json");
        }

        static async Task StopBackgroundServer()
        {
            int? port = ProcessProber.FindRunningServerPort();
            if (port == null) return;

            try
            {
                using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
                await client.PostAsync($"http://{LISTEN_HOST}:{port}/exit", null);
                Console.WriteLine("Server stopped.");
            }
            catch { }
        }

        static async Task ShowStatus()
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
                var response = await client.GetAsync($"http://{LISTEN_HOST}:{port}/status");
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    Console.WriteLine(json);
                }
            }
            catch (Exception ex) { Console.WriteLine($"Error: {ex.Message}"); }
        }

        static void OpenSwaggerPage()
        {
            int? port = ProcessProber.FindRunningServerPort();
            if (port == null)
            {
                Console.WriteLine("Server is not running. Starting it now...");
                StartBackgroundServer();
                for (int i = 0; i < 10; i++)
                {
                    Thread.Sleep(500);
                    port = ProcessProber.FindRunningServerPort();
                    if (port != null) break;
                }
            }

            if (port != null)
            {
                string url = $"http://{LISTEN_HOST}:{port}/swagger/index.html";
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

        static void RunServer(int port)
        {
            using var mutex = new Mutex(true, MUTEXT_NAME, out bool createdNew);
            if (!createdNew) return;

            var builder = WebApplication.CreateBuilder();
            
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new OpenApiInfo { Title = "AgentExcel API", Version = "v1" });
            });

            var app = builder.Build();

            app.UseSwagger();
            app.UseSwaggerUI();

            Excel.Application excelApp = ExcelConnector.EnsureExcelApplication();
            var excelService = new ExcelService(excelApp);

            // --- Status & Lifecycle ---

            app.MapGet("/status", () =>
            {
                string activeWorkbook = "None";
                try { activeWorkbook = excelApp.ActiveWorkbook?.Name ?? "None"; } catch { activeWorkbook = "Busy"; }
                
                return Results.Ok(new
                {
                    status = "running",
                    pid = Process.GetCurrentProcess().Id,
                    port = port,
                    active_workbook = activeWorkbook,
                    swagger_url = $"http://{LISTEN_HOST}:{port}/swagger/index.html",
                    openapi_url = $"http://{LISTEN_HOST}:{port}/swagger/v1/swagger.json"
                });
            })
            .WithOpenApi(operation => { operation.Summary = "Get server and Excel status"; return operation; });

            app.MapPost("/exit", (IHostApplicationLifetime lifetime) =>
            {
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                lifetime.StopApplication();
                return Results.Ok(new { status = "shutting_down" });
            })
            .WithOpenApi(operation => { operation.Summary = "Shutdown the daemon server"; return operation; });

            // --- Workbook Operations ---

            var workbooks = app.MapGroup("/workbooks").WithOpenApi();

            workbooks.MapGet("/", () => Results.Ok(excelService.ListWorkbooks()))
                .WithSummary("List all open workbooks");

            workbooks.MapPost("/open", (OpenWorkbookRequest req) => Results.Ok(new { name = excelService.OpenWorkbook(req.Path) }))
                .WithSummary("Open a workbook by path");

            workbooks.MapPost("/add", () => Results.Ok(new { name = excelService.AddWorkbook() }))
                .WithSummary("Create a new workbook");

            workbooks.MapPost("/save", (WorkbookRequest req) => { excelService.SaveWorkbook(req.Name); return Results.Ok(new { status = "saved" }); })
                .WithSummary("Save a workbook");

            workbooks.MapPost("/saveas", (SaveAsRequest req) => { excelService.SaveAsWorkbook(req.Name, req.Path); return Results.Ok(new { status = "saved_as" }); })
                .WithSummary("Save a workbook to a new path");

            workbooks.MapPost("/close", (CloseWorkbookRequest req) => { excelService.CloseWorkbook(req.Name, req.SaveChanges); return Results.Ok(new { status = "closed" }); })
                .WithSummary("Close a workbook");

            // --- Sheet Operations ---

            var sheets = app.MapGroup("/sheets").WithOpenApi();

            sheets.MapPost("/list", (WorkbookRequest req) => Results.Ok(excelService.ListSheets(req.Name)))
                .WithSummary("List sheets in a workbook");

            sheets.MapPost("/add", (SheetRequest req) => { excelService.AddSheet(req.Workbook, req.Name); return Results.Ok(new { status = "added" }); })
                .WithSummary("Add a new sheet");

            sheets.MapPost("/delete", (SheetRequest req) => { excelService.DeleteSheet(req.Workbook, req.Name); return Results.Ok(new { status = "deleted" }); })
                .WithSummary("Delete a sheet");

            sheets.MapPost("/rename", (RenameSheetRequest req) => { excelService.RenameSheet(req.Workbook, req.OldName, req.NewName); return Results.Ok(new { status = "renamed" }); })
                .WithSummary("Rename a sheet");

            // --- Data Operations ---

            var data = app.MapGroup("/data").WithOpenApi();

            data.MapPost("/read-range", (ReadRangeRequest req) => Results.Ok(new { values = excelService.ReadRange(req.Workbook, req.Sheet, req.Address) }))
                .WithSummary("Read data from a range");

            data.MapPost("/read-table", (ReadTableRequest req) => Results.Ok(new { values = excelService.ReadTable(req.Workbook, req.Sheet, req.Name) }))
                .WithSummary("Read data from an Excel Table (ListObject)");

            data.MapPost("/convert-to-table", (ConvertToTableRequest req) => Results.Ok(new { name = excelService.ConvertToTable(req.Workbook, req.Sheet, req.RangeAddress, req.TableName, req.HasHeaders) }))
                .WithSummary("Convert a normal range to an Excel Table");

            data.MapPost("/convert-to-range", (ConvertToRangeRequest req) => { excelService.ConvertToRange(req.Workbook, req.Sheet, req.TableName); return Results.Ok(new { status = "converted" }); })
                .WithSummary("Convert an Excel Table back to a normal range");

            data.MapPost("/write-range", (RangeWriteRequest req) => { excelService.WriteRange(req.Workbook, req.Sheet, req.Address, req.Value); return Results.Ok(new { status = "success" }); })
                .WithSummary("Write data to a range");

            data.MapPost("/write-formula", (WriteFormulaRequest req) => { excelService.WriteFormula(req.Workbook, req.Sheet, req.Address, req.Formula); return Results.Ok(new { status = "success" }); })
                .WithSummary("Write a formula to a range");

            data.MapPost("/read-formula", (ReadRangeRequest req) => Results.Ok(new { formula = excelService.ReadFormula(req.Workbook, req.Sheet, req.Address) }))
                .WithSummary("Read the formula from a range");

            data.MapPost("/find", (FindRequest req) => Results.Ok(new { results = excelService.Find(req.Workbook, req.Sheet, req.RangeAddress, req.What, req.MatchCase, req.WholeWord) }))
                .WithSummary("Find all occurrences of a string");

            data.MapPost("/replace", (ReplaceRequest req) => Results.Ok(new { success = excelService.Replace(req.Workbook, req.Sheet, req.RangeAddress, req.What, req.Replacement, req.MatchCase, req.WholeWord) }))
                .WithSummary("Replace all occurrences of a string");

            data.MapPost("/used-range", (UsedRangeRequest req) => Results.Ok(new { address = excelService.GetUsedRangeAddress(req.Workbook, req.Sheet) }))
                .WithSummary("Get the used range address");

            data.MapPost("/grep", (GrepRequest req) => Results.Ok(new { results = excelService.SearchInFolder(req.FolderPath, req.Pattern) }))
                .WithSummary("Search for text in all Excel files within a folder (includes Cells, TextBoxes, and Shapes)");

            data.MapPost("/set-style", (SetStyleRequest req) => { excelService.SetStyle(req.Workbook, req.Sheet, req.Address, req.Style); return Results.Ok(new { status = "success" }); })
                .WithSummary("Set cell styles (Font, Color, Bold, Alignment, etc.)");

            // --- Shape Operations ---

            var shapes = app.MapGroup("/shapes").WithOpenApi();

            shapes.MapPost("/list", (UsedRangeRequest req) => Results.Ok(excelService.ListShapes(req.Workbook, req.Sheet)))
                .WithSummary("List all shapes and textboxes in a sheet");

            shapes.MapPost("/add", (AddShapeRequest req) => Results.Ok(new { name = excelService.AddShape(req.Workbook, req.Sheet, req.Type, req.Left, req.Top, req.Width, req.Height, req.Text) }))
                .WithSummary("Add a shape or textbox");

            shapes.MapPost("/update", (UpdateShapeRequest req) => { excelService.UpdateShape(req.Workbook, req.Sheet, req.ShapeName, req.Left, req.Top, req.Width, req.Height, req.Text); return Results.Ok(new { status = "updated" }); })
                .WithSummary("Update shape properties (position, size, text)");

            shapes.MapPost("/delete", (DeleteShapeRequest req) => { excelService.DeleteShape(req.Workbook, req.Sheet, req.ShapeName); return Results.Ok(new { status = "deleted" }); })
                .WithSummary("Delete a shape");

            // --- Charts Operations ---

            var charts = app.MapGroup("/charts").WithOpenApi();

            charts.MapPost("/list", (WorkbookRequest req) => Results.Ok(excelService.ListCharts(req.Name, null)))
                .WithSummary("List all charts in a sheet");

            charts.MapPost("/add", (AddChartRequest req) => Results.Ok(new { name = excelService.AddChart(req.Workbook, req.Sheet, req.RangeAddress, req.ChartType, req.Title) }))
                .WithSummary("Add a chart to a sheet");

            charts.MapPost("/update", (UpdateChartRequest req) => { excelService.UpdateChart(req.Workbook, req.Sheet, req.ChartName, req.RangeAddress, req.ChartType, req.Title); return Results.Ok(new { status = "updated" }); })
                .WithSummary("Update chart properties (source data, type, title)");

            charts.MapPost("/delete", (DeleteChartRequest req) => { excelService.DeleteChart(req.Workbook, req.Sheet, req.ChartName); return Results.Ok(new { status = "deleted" }); })
                .WithSummary("Delete a chart");

            // --- Export Operations ---

            var export = app.MapGroup("/export").WithOpenApi();

            export.MapPost("/range-image", (ExportImageRequest req) => { excelService.ExportRangeAsImage(req.Workbook, req.Sheet, req.RangeAddress, req.OutputPath); return Results.Ok(new { status = "exported" }); })
                .WithSummary("Export a range as a PNG image");

            export.MapPost("/workbook-pdf", (ExportPdfRequest req) => { excelService.ExportAsPdf(req.Workbook, req.OutputPath); return Results.Ok(new { status = "exported" }); })
                .WithSummary("Export the entire workbook as a PDF");

            // --- Pivot Table Operations ---

            var pivot = app.MapGroup("/pivot").WithOpenApi();

            pivot.MapPost("/create", (CreatePivotRequest req) => { excelService.CreatePivotTable(req.Workbook, req.SourceSheet, req.SourceRange, req.TargetSheet, req.TargetCell, req.TableName); return Results.Ok(new { status = "created" }); })
                .WithSummary("Create a pivot table");

            // --- Macro Operations ---

            var macros = app.MapGroup("/macros").WithOpenApi();

            macros.MapPost("/run", (RunMacroRequest req) => Results.Ok(new { result = excelService.RunMacro(req.MacroName, req.Args) }))
                .WithSummary("Run a VBA macro");

            // --- Data Validation ---

            var validation = app.MapGroup("/validation").WithOpenApi();

            validation.MapPost("/set-list", (SetListValidationRequest req) => { excelService.SetListValidation(req.Workbook, req.Sheet, req.RangeAddress, req.Formula); return Results.Ok(new { status = "success" }); })
                .WithSummary("Set a dropdown list validation");

            // --- System Operations ---

            var system = app.MapGroup("/system").WithOpenApi();

            system.MapPost("/calculate", () => { excelService.Calculate(); return Results.Ok(new { status = "calculated" }); })
                .WithSummary("Trigger a global calculation");

            // --- Power Query & Connections ---

            var pq = app.MapGroup("/queries").WithOpenApi();

            pq.MapPost("/list", (WorkbookRequest req) => Results.Ok(excelService.ListQueries(req.Name)))
                .WithSummary("List all Power Queries in a workbook");

            pq.MapPost("/add-update", (AddQueryRequest req) => { excelService.AddOrUpdateQuery(req.Workbook, req.Name, req.Formula, req.Description); return Results.Ok(new { status = "success" }); })
                .WithSummary("Add or update a Power Query (M Language)");

            pq.MapPost("/delete", (DeleteQueryRequest req) => { excelService.DeleteQuery(req.Workbook, req.Name); return Results.Ok(new { status = "deleted" }); })
                .WithSummary("Delete a Power Query");

            var conn = app.MapGroup("/connections").WithOpenApi();

            conn.MapPost("/refresh-all", (WorkbookRequest req) => { excelService.RefreshAllDataConnections(req.Name); return Results.Ok(new { status = "refreshing" }); })
                .WithSummary("Refresh all data connections, Power Queries, and Data Model");

            conn.MapPost("/refresh-model", (WorkbookRequest req) => { excelService.RefreshModel(req.Name); return Results.Ok(new { status = "refreshing_model" }); })
                .WithSummary("Refresh the Excel Data Model (Power Pivot)");

            app.Run($"http://{LISTEN_HOST}:{port}");
        }

        static int GetAvailablePort()
        {
            TcpListener listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            int port = ((IPEndPoint)listener.LocalEndpoint).Port;
            listener.Stop();
            return port;
        }
    }

    // --- Request Models ---

    public record WorkbookRequest(string? Name);
    public record OpenWorkbookRequest(string Path);
    public record SaveAsRequest(string? Name, string Path);
    public record CloseWorkbookRequest(string? Name, bool SaveChanges);
    
    public record SheetRequest(string? Workbook, string Name);
    public record RenameSheetRequest(string? Workbook, string OldName, string NewName);
    
    public record ReadRangeRequest(string? Workbook, string? Sheet, string Address);
    public record ReadTableRequest(string? Workbook, string? Sheet, string Name);
    public record UsedRangeRequest(string? Workbook, string? Sheet);
    public record RangeWriteRequest(string? Workbook, string? Sheet, string Address, object Value);
    public record GrepRequest(string FolderPath, string Pattern);
    public record SetStyleRequest(string? Workbook, string? Sheet, string Address, ExcelService.CellStyle Style);
    public record AddShapeRequest(string? Workbook, string? Sheet, string Type, float Left, float Top, float Width, float Height, string? Text);
    public record UpdateShapeRequest(string? Workbook, string? Sheet, string ShapeName, float? Left, float? Top, float? Width, float? Height, string? Text);
    public record DeleteShapeRequest(string? Workbook, string? Sheet, string ShapeName);

    public record AddChartRequest(string? Workbook, string? Sheet, string RangeAddress, string ChartType, string Title);
    public record UpdateChartRequest(string? Workbook, string? Sheet, string ChartName, string? RangeAddress, string? ChartType, string? Title);
    public record DeleteChartRequest(string? Workbook, string? Sheet, string ChartName);
    public record ExportImageRequest(string? Workbook, string? Sheet, string RangeAddress, string OutputPath);
    public record ExportPdfRequest(string? Workbook, string OutputPath);
    public record CreatePivotRequest(string? Workbook, string SourceSheet, string SourceRange, string TargetSheet, string TargetCell, string TableName);
    public record RunMacroRequest(string MacroName, object[]? Args);
    public record SetListValidationRequest(string? Workbook, string? Sheet, string RangeAddress, string Formula);

    public record AddQueryRequest(string? Workbook, string Name, string Formula, string? Description);
    public record DeleteQueryRequest(string? Workbook, string Name);
    public record WriteFormulaRequest(string? Workbook, string? Sheet, string Address, string Formula);
    public record ConvertToTableRequest(string? Workbook, string? Sheet, string RangeAddress, string? TableName, bool HasHeaders);
    public record ConvertToRangeRequest(string? Workbook, string? Sheet, string TableName);

    public record FindRequest(string? Workbook, string? Sheet, string? RangeAddress, string What, bool MatchCase, bool WholeWord);
    public record ReplaceRequest(string? Workbook, string? Sheet, string? RangeAddress, string What, string Replacement, bool MatchCase, bool WholeWord);
}
