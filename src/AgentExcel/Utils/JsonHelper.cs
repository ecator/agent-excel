using System;
using System.Collections.Generic;
using System.Text.Json;

namespace AgentExcel.Utils;

public static class JsonHelper
{
    public static object? ConvertJsonElement(JsonElement element)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.String:
                return element.GetString();
            case JsonValueKind.Number:
                if (element.TryGetInt32(out int i)) return i;
                if (element.TryGetInt64(out long l)) return l;
                if (element.TryGetDouble(out double d)) return d;
                return element.GetRawText();
            case JsonValueKind.True:
                return true;
            case JsonValueKind.False:
                return false;
            case JsonValueKind.Null:
                return null;
            case JsonValueKind.Array:
                var list = new List<object?>();
                foreach (var item in element.EnumerateArray())
                {
                    list.Add(ConvertJsonValue(item));
                }
                return list.ToArray();
            default:
                return element.GetRawText();
        }
    }

    public static object? ConvertJsonValue(object? value)
    {
        if (value is JsonElement el)
        {
            return ConvertJsonElement(el);
        }
        return value;
    }
}
