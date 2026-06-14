using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

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

    public static async Task ExecuteApiRequest(
        string command,
        string[] args,
        string listenHost,
        HttpClient? client = null,
        System.IO.TextReader? stdinReader = null,
        System.IO.TextWriter? outputWriter = null,
        int? overridePort = null)
    {
        var exeName = Path.GetFileName(Process.GetCurrentProcess().MainModule?.FileName ?? "AgentExcel.exe");
        var outWriter = outputWriter ?? Console.Out;
        if (args.Length < 2)
        {
            outWriter.WriteLine($"Usage: {exeName} {command} <endpoint> [--stdin] [body_param]");
            return;
        }

        string endpoint = args[1];
        if (!endpoint.StartsWith("/"))
        {
            endpoint = "/" + endpoint;
        }
        if (endpoint.EndsWith("/")){
            endpoint = endpoint.TrimEnd('/');
        }

        bool useStdin = args.Skip(2).Any(arg => string.Equals(arg, "--stdin", StringComparison.OrdinalIgnoreCase));
        string? requestBody = null;

        if (useStdin)
        {
            var reader = stdinReader ?? Console.In;
            requestBody = await reader.ReadToEndAsync();
        }
        else if (args.Length > 2)
        {
            requestBody = args[2];
        }

        int? port = overridePort ?? ProcessProber.FindRunningServerPort();
        if (port == null)
        {
            outWriter.WriteLine($"Error: Daemon server is not running. Please start the server first by running '{exeName} start'.");
            return;
        }

        string url = $"http://{listenHost}:{port}{endpoint}";
        var method = string.Equals(command, "post", StringComparison.OrdinalIgnoreCase) ? HttpMethod.Post : HttpMethod.Get;

        using var request = new HttpRequestMessage(method, url);
        if (requestBody != null)
        {
            request.Content = new StringContent(requestBody, System.Text.Encoding.UTF8, "application/json");
        }

        HttpClient? localClient = null;
        try
        {
            if (client == null)
            {
                localClient = new HttpClient { Timeout = TimeSpan.FromSeconds(300) };
            }
            HttpClient activeClient = client ?? localClient!;

            using var response = await activeClient.SendAsync(request);
            string content = await response.Content.ReadAsStringAsync();

            if ((int)response.StatusCode == 200)
            {
                outWriter.WriteLine(content);
            }
            else
            {
                outWriter.WriteLine($"Error Code: {(int)response.StatusCode}");
                outWriter.WriteLine(content);
            }
        }
        catch (HttpRequestException)
        {
            outWriter.WriteLine($"Error: Daemon server is not running. Please start the server first by running '{exeName} start'.");
        }
        catch (Exception ex)
        {
            outWriter.WriteLine($"Error sending request: {ex.Message}");
        }
        finally
        {
            localClient?.Dispose();
        }
    }

    public static async Task ShowApiCatalog(
        string listenHost,
        HttpClient? client = null,
        System.IO.TextWriter? outputWriter = null,
        int? overridePort = null)
    {
        var exeName = Path.GetFileName(Process.GetCurrentProcess().MainModule?.FileName ?? "AgentExcel.exe");
        var outWriter = outputWriter ?? Console.Out;

        int? port = overridePort ?? ProcessProber.FindRunningServerPort();
        if (port == null)
        {
            outWriter.WriteLine($"Error: Daemon server is not running. Please start the server first by running '{exeName} start'.");
            return;
        }

        string url = $"http://{listenHost}:{port}/swagger/v1/swagger.json";
        HttpClient? localClient = null;
        try
        {
            HttpClient activeClient = client ?? (localClient = new HttpClient { Timeout = TimeSpan.FromSeconds(10) });
            string jsonString = await activeClient.GetStringAsync(url);
            string formatted = FormatApiCatalog(jsonString);
            outWriter.WriteLine(formatted);
        }
        catch (HttpRequestException)
        {
            outWriter.WriteLine($"Error: Daemon server is not running. Please start the server first by running '{exeName} start'.");
        }
        catch (Exception ex)
        {
            outWriter.WriteLine($"Error fetching or parsing API catalog: {ex.Message}");
        }
        finally
        {
            localClient?.Dispose();
        }
    }

    private static string FormatApiCatalog(string jsonString)
    {
        using var doc = JsonDocument.Parse(jsonString);
        var root = doc.RootElement;

        var schemas = new Dictionary<string, JsonElement>();
        if (root.TryGetProperty("components", out var components) &&
            components.TryGetProperty("schemas", out var schemasElem))
        {
            foreach (var prop in schemasElem.EnumerateObject())
            {
                schemas[prop.Name] = prop.Value;
            }
        }

        var endpoints = new List<ApiEndpointInfo>();

        if (root.TryGetProperty("paths", out var paths))
        {
            foreach (var pathProp in paths.EnumerateObject())
            {
                string path = pathProp.Name;
                foreach (var methodProp in pathProp.Value.EnumerateObject())
                {
                    string method = methodProp.Name.ToUpper();
                    var op = methodProp.Value;

                    string summary = "";
                    if (op.TryGetProperty("summary", out var sumProp))
                    {
                        summary = CleanDescription(sumProp.GetString());
                    }
                    else if (op.TryGetProperty("description", out var descProp))
                    {
                        summary = CleanDescription(descProp.GetString());
                    }

                    string tag = "General";
                    if (op.TryGetProperty("tags", out var tagsProp) && tagsProp.GetArrayLength() > 0)
                    {
                        tag = tagsProp[0].GetString() ?? "General";
                    }

                    JsonElement? requestSchema = null;
                    if (op.TryGetProperty("requestBody", out var reqBody))
                    {
                        if (reqBody.TryGetProperty("content", out var content) &&
                            content.TryGetProperty("application/json", out var appJson) &&
                            appJson.TryGetProperty("schema", out var schemaRef))
                        {
                            requestSchema = schemaRef;
                        }
                    }

                    endpoints.Add(new ApiEndpointInfo(tag, method, path, summary, requestSchema));
                }
            }
        }

        var sb = new StringBuilder();
        var grouped = endpoints.GroupBy(e => e.Tag).OrderBy(g => g.Key);

        foreach (var group in grouped)
        {
            sb.AppendLine($"=== {group.Key} ===");
            foreach (var ep in group.OrderBy(e => e.Path))
            {
                sb.AppendLine($"[{ep.Method}] {ep.Path} - {ep.Summary}");
                if (ep.RequestSchema != null)
                {
                    sb.AppendLine("  Body:");
                    var bodyJson = FormatSchemaJson(ep.RequestSchema.Value, schemas, 2);
                    sb.AppendLine($"  {bodyJson}");
                }
                sb.AppendLine();
            }
        }

        return sb.ToString().TrimEnd();
    }

    private static string FormatSchemaJson(JsonElement schema, Dictionary<string, JsonElement> schemas, int indent)
    {
        if (schema.ValueKind == JsonValueKind.Object && schema.TryGetProperty("$ref", out var refProp))
        {
            string refVal = refProp.GetString() ?? "";
            const string prefix = "#/components/schemas/";
            if (refVal.StartsWith(prefix))
            {
                string schemaName = refVal.Substring(prefix.Length);
                if (schemas.TryGetValue(schemaName, out var resolvedSchema))
                {
                    return FormatSchemaJson(resolvedSchema, schemas, indent);
                }
            }
            return "null";
        }

        string type = "any";
        if (schema.TryGetProperty("type", out var typeProp))
        {
            type = typeProp.GetString() ?? "any";
        }
        else if (schema.TryGetProperty("properties", out _))
        {
            type = "object";
        }

        if (type == "object")
        {
            if (schema.TryGetProperty("properties", out var props) && props.ValueKind == JsonValueKind.Object)
            {
                var sb = new StringBuilder();
                sb.AppendLine("{");
                var properties = props.EnumerateObject().ToList();
                HashSet<string> requiredProps = new HashSet<string>();
                if (schema.TryGetProperty("required", out var reqArray))
                {
                    foreach (var r in reqArray.EnumerateArray())
                    {
                        var rStr = r.GetString();
                        if (rStr != null) requiredProps.Add(rStr);
                    }
                }

                var lines = new List<PropertyLineInfo>();
                for (int i = 0; i < properties.Count; i++)
                {
                    var prop = properties[i];
                    string propName = prop.Name;
                    var propDetails = prop.Value;
                    bool isLast = i == properties.Count - 1;

                    var (valStr, comment) = GetPropValueAndComment(propDetails, schemas, indent + 2, requiredProps.Contains(propName));
                    lines.Add(new PropertyLineInfo(propName, valStr, comment, isLast));
                }

                int maxLhsWidth = 0;
                foreach (var line in lines)
                {
                    if (!line.ValStr.Contains('\n') && !line.ValStr.Contains('\r'))
                    {
                        string comma = line.IsLast ? "" : ",";
                        string lhs = $"{new string(' ', indent + 2)}\"{line.Name}\": {line.ValStr}{comma}";
                        if (lhs.Length > maxLhsWidth)
                        {
                            maxLhsWidth = lhs.Length;
                        }
                    }
                }

                for (int i = 0; i < lines.Count; i++)
                {
                    var line = lines[i];
                    string comma = line.IsLast ? "" : ",";

                    if (line.ValStr.Contains('\n') || line.ValStr.Contains('\r'))
                    {
                        var valLines = line.ValStr.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                        string commentStr = string.IsNullOrEmpty(line.Comment) ? "" : $" // {line.Comment}";
                        sb.AppendLine($"{new string(' ', indent + 2)}\"{line.Name}\": {valLines[0]}{commentStr}");
                        for (int j = 1; j < valLines.Length; j++)
                        {
                            if (j == valLines.Length - 1)
                            {
                                sb.AppendLine($"{valLines[j]}{comma}");
                            }
                            else
                            {
                                sb.AppendLine(valLines[j]);
                            }
                        }
                    }
                    else
                    {
                        string lhs = $"{new string(' ', indent + 2)}\"{line.Name}\": {line.ValStr}{comma}";
                        string commentStr = string.IsNullOrEmpty(line.Comment) ? "" : $" // {line.Comment}";
                        if (!string.IsNullOrEmpty(commentStr))
                        {
                            sb.AppendLine(lhs.PadRight(maxLhsWidth) + commentStr);
                        }
                        else
                        {
                            sb.AppendLine(lhs);
                        }
                    }
                }

                sb.Append(new string(' ', indent));
                sb.Append("}");
                return sb.ToString();
            }
            else
            {
                return "{}";
            }
        }
        else if (type == "array")
        {
            if (schema.TryGetProperty("items", out var items))
            {
                var (itemVal, _) = GetPropValueAndComment(items, schemas, indent, false);
                return $"[{itemVal}]";
            }
            return "[]";
        }
        else if (type == "string")
        {
            return "\"string\"";
        }
        else if (type == "integer" || type == "number")
        {
            return "number";
        }
        else if (type == "boolean")
        {
            return "boolean";
        }
        else
        {
            return "any";
        }
    }

    private static (string valStr, string comment) GetPropValueAndComment(JsonElement propDetails, Dictionary<string, JsonElement> schemas, int indent, bool isRequired)
    {
        if (propDetails.TryGetProperty("$ref", out var refProp))
        {
            string refVal = refProp.GetString() ?? "";
            const string prefix = "#/components/schemas/";
            if (refVal.StartsWith(prefix))
            {
                string schemaName = refVal.Substring(prefix.Length);
                if (schemas.TryGetValue(schemaName, out var resolvedSchema))
                {
                    string valStr = FormatSchemaJson(resolvedSchema, schemas, indent);
                    string desc = "";
                    if (resolvedSchema.TryGetProperty("description", out var descProp))
                    {
                        desc = CleanDescription(descProp.GetString());
                    }
                    string status = isRequired ? "required" : "optional";
                    string comment = string.IsNullOrWhiteSpace(desc) ? $"({status})" : $"({status}) {desc.Trim()}";
                    return (valStr, comment);
                }
            }
        }

        string type = "any";
        if (propDetails.TryGetProperty("type", out var typeProp))
        {
            type = typeProp.GetString() ?? "any";
        }
        else if (propDetails.TryGetProperty("properties", out _))
        {
            type = "object";
        }

        string propDesc = "";
        if (propDetails.TryGetProperty("description", out var descVal))
        {
            propDesc = CleanDescription(descVal.GetString());
        }

        string reqStatus = isRequired ? "required" : "optional";
        string propComment = string.IsNullOrWhiteSpace(propDesc) ? $"({reqStatus})" : $"({reqStatus}) {propDesc.Trim()}";

        if (type == "object")
        {
            string valStr = FormatSchemaJson(propDetails, schemas, indent);
            return (valStr, propComment);
        }
        else if (type == "array")
        {
            if (propDetails.TryGetProperty("items", out var items))
            {
                var (itemVal, _) = GetPropValueAndComment(items, schemas, indent, false);
                return ($"[{itemVal}]", propComment);
            }
            return ("[]", propComment);
        }
        else if (type == "string")
        {
            return ("\"string\"", propComment);
        }
        else if (type == "integer" || type == "number")
        {
            return ("number", propComment);
        }
        else if (type == "boolean")
        {
            return ("boolean", propComment);
        }
        else
        {
            return ("any", propComment);
        }
    }

    private class ApiEndpointInfo(string tag, string method, string path, string summary, JsonElement? requestSchema)
    {
        public string Tag { get; } = tag;
        public string Method { get; } = method;
        public string Path { get; } = path;
        public string Summary { get; } = summary;
        public JsonElement? RequestSchema { get; } = requestSchema;
    }

    private class PropertyLineInfo(string name, string valStr, string comment, bool isLast)
    {
        public string Name { get; } = name;
        public string ValStr { get; } = valStr;
        public string Comment { get; } = comment;
        public bool IsLast { get; } = isLast;
    }

    private static string CleanDescription(string? desc)
    {
        if (string.IsNullOrEmpty(desc)) return "";
        string cleaned = desc.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        return Regex.Replace(cleaned, @"\s+", " ").Trim();
    }
}
