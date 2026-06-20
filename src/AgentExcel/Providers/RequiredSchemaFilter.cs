using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.Json;

using Microsoft.OpenApi.Models;

using Swashbuckle.AspNetCore.SwaggerGen;

namespace AgentExcel.Providers;

/// <summary>
/// Schema filter to automatically mark C# 11 required properties as required in Swagger schema.
/// </summary>
public class RequiredSchemaFilter : ISchemaFilter
{
    public void Apply(OpenApiSchema schema, SchemaFilterContext context)
    {
        if (schema.Properties == null) return;

        foreach (var property in context.Type.GetProperties())
        {
            // Check for C# 11 'required' modifier
            var isRequired = Attribute.IsDefined(property, typeof(RequiredMemberAttribute));
            if (isRequired)
            {
                string jsonName = JsonNamingPolicy.CamelCase.ConvertName(property.Name);
                if (schema.Properties.ContainsKey(jsonName))
                {
                    schema.Required.Add(jsonName);
                }
            }
        }
    }
}
