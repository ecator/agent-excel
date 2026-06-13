using Microsoft.AspNetCore.Http;

using YamlDotNet.Serialization;

namespace AgentExcel.Utils;

/// <summary>
/// Provides extension methods for returning YAML responses from Minimal APIs.
/// </summary>
public static class YamlResultsExtensions
{
    private static readonly ISerializer Serializer = new SerializerBuilder()
        .Build();

    /// <summary>
    /// Returns a YAML-formatted response.
    /// </summary>
    /// <param name="_">The IResultExtensions instance.</param>
    /// <param name="value">The value to serialize as YAML.</param>
    /// <param name="statusCode">The HTTP status code. Defaults to 200 OK.</param>
    /// <returns>A YamlResult instance.</returns>
    public static IResult Yaml(this IResultExtensions _, object? value, int statusCode = 200)
    {
        return new YamlResult(value, statusCode);
    }

    private class YamlResult : IResult
    {
        private readonly object? _value;
        private readonly int _statusCode;

        public YamlResult(object? value, int statusCode)
        {
            _value = value;
            _statusCode = statusCode;
        }

        public async Task ExecuteAsync(HttpContext httpContext)
        {
            httpContext.Response.StatusCode = _statusCode;
            httpContext.Response.ContentType = "text/yaml; charset=utf-8";

            if (_value is not null)
            {
                string yaml = Serializer.Serialize(_value);
                await httpContext.Response.WriteAsync(yaml);
            }
        }
    }
}
