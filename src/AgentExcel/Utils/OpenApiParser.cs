using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace AgentExcel.Utils;

/// <summary>
/// A representation of the parsed OpenAPI catalog, holding references to all endpoints and schemas.
/// </summary>
internal sealed class ApiCatalog : IDisposable
{
    private readonly JsonDocument _document;

    public IReadOnlyList<ApiEndpoint> Endpoints { get; }
    public IReadOnlyDictionary<string, JsonElement> Schemas { get; }

    public ApiCatalog(JsonDocument document, List<ApiEndpoint> endpoints, Dictionary<string, JsonElement> schemas)
    {
        _document = document;
        Endpoints = endpoints;
        Schemas = schemas;
    }

    public void Dispose()
    {
        _document.Dispose();
    }
}

/// <summary>
/// Represents an OpenAPI endpoint path and method metadata.
/// </summary>
internal sealed class ApiEndpoint
{
    public string Tag { get; }
    public string Method { get; }
    public string Path { get; }
    public string Summary { get; }
    public JsonElement? RequestSchema { get; }

    public ApiEndpoint(string tag, string method, string path, string summary, JsonElement? requestSchema)
    {
        Tag = tag;
        Method = method;
        Path = path;
        Summary = summary;
        RequestSchema = requestSchema;
    }
}

/// <summary>
/// Helper utilities to parse, filter, and format OpenAPI schemas.
/// </summary>
internal static class OpenApiParser
{
    /// <summary>
    /// Parses an OpenAPI JSON string into a structured catalog of endpoints and schemas.
    /// </summary>
    public static ApiCatalog Parse(string jsonString)
    {
        var doc = JsonDocument.Parse(jsonString);
        var root = doc.RootElement;

        var schemas = new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);
        if (root.TryGetProperty("components", out var components) &&
            components.TryGetProperty("schemas", out var schemasElem))
        {
            foreach (var prop in schemasElem.EnumerateObject())
            {
                schemas[prop.Name] = prop.Value;
            }
        }

        var endpoints = new List<ApiEndpoint>();
        if (root.TryGetProperty("paths", out var paths))
        {
            foreach (var pathProp in paths.EnumerateObject())
            {
                string path = pathProp.Name;
                foreach (var methodProp in pathProp.Value.EnumerateObject())
                {
                    string method = methodProp.Name.ToUpperInvariant();
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

                    endpoints.Add(new ApiEndpoint(tag, method, path, summary, requestSchema));
                }
            }
        }

        return new ApiCatalog(doc, endpoints, schemas);
    }

    /// <summary>
    /// Filters and formats the catalog according to provided command-line arguments.
    /// </summary>
    public static string FormatCatalog(ApiCatalog catalog, string[]? args)
    {
        string? groupFilter = null;
        string? endpointFilter = null;

        if (args is not null)
        {
            if (args.Length > 1)
            {
                groupFilter = args[1];
            }
            if (args.Length > 2)
            {
                endpointFilter = args[2];
            }
        }

        // Determine if we should show details:
        // 1. If endpoint filter is specified (args.Length > 2)
        // 2. Or if group filter contains a glob '*'
        bool showDetails = (args is not null && args.Length > 2) ||
                           (groupFilter is not null && groupFilter.Contains('*'));

        var filteredEndpoints = catalog.Endpoints.Where(ep =>
        {
            if (groupFilter is not null && !MatchGroup(ep.Tag, groupFilter))
            {
                return false;
            }
            if (endpointFilter is not null && !MatchEndpoint(ep.Path, endpointFilter))
            {
                return false;
            }
            return true;
        }).ToList();

        var sb = new StringBuilder();
        var grouped = filteredEndpoints.GroupBy(e => e.Tag).OrderBy(g => g.Key);

        foreach (var group in grouped)
        {
            sb.AppendLine($"=== {group.Key} ===");
            foreach (var ep in group.OrderBy(e => e.Path))
            {
                sb.AppendLine($"[{ep.Method}] {ep.Path} - {ep.Summary}");
                if (showDetails && ep.RequestSchema is not null)
                {
                    sb.AppendLine("  Body:");
                    var bodyJson = FormatSchemaJson(ep.RequestSchema.Value, catalog.Schemas, 2);
                    sb.AppendLine($"  {bodyJson}");
                }
                sb.AppendLine();
            }
        }

        return sb.ToString().TrimEnd();
    }

    public static string FormatSchemaJson(JsonElement schema, IReadOnlyDictionary<string, JsonElement> schemas, int indent)
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
                var requiredProps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (schema.TryGetProperty("required", out var reqArray))
                {
                    foreach (var r in reqArray.EnumerateArray())
                    {
                        var rStr = r.GetString();
                        if (rStr is not null) requiredProps.Add(rStr);
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

    private static (string valStr, string comment) GetPropValueAndComment(
        JsonElement propDetails,
        IReadOnlyDictionary<string, JsonElement> schemas,
        int indent,
        bool isRequired)
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

    private static bool MatchGroup(string tag, string filter)
    {
        if (string.IsNullOrEmpty(filter)) return true;

        if (filter.Contains('*'))
        {
            return MatchesGlob(tag, filter);
        }

        return string.Equals(tag, filter, StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchEndpoint(string path, string filter)
    {
        if (string.IsNullOrEmpty(filter)) return true;

        if (filter.Contains('*'))
        {
            return MatchesGlob(path, filter);
        }

        // Exact match or suffix match with leading slash normalization
        string normalizedPath = path.StartsWith("/") ? path : "/" + path;
        string normalizedFilter = filter.StartsWith("/") ? filter : "/" + filter;

        if (string.Equals(normalizedPath, normalizedFilter, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        // Also support matching endpoint name without slash (e.g. "write" matching "/data/write" or "/write")
        if (!filter.StartsWith("/"))
        {
            if (path.EndsWith("/" + filter, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool MatchesGlob(string input, string pattern)
    {
        string regexPattern = "^" + Regex.Escape(pattern).Replace("\\*", ".*") + "$";
        return Regex.IsMatch(input, regexPattern, RegexOptions.IgnoreCase);
    }

    public static string CleanDescription(string? desc)
    {
        if (string.IsNullOrEmpty(desc)) return "";
        string cleaned = desc.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        return Regex.Replace(cleaned, @"\s+", " ").Trim();
    }

    private sealed class PropertyLineInfo
    {
        public string Name { get; }
        public string ValStr { get; }
        public string Comment { get; }
        public bool IsLast { get; }

        public PropertyLineInfo(string name, string valStr, string comment, bool isLast)
        {
            Name = name;
            ValStr = valStr;
            Comment = comment;
            IsLast = isLast;
        }
    }
}
