using System.Text.Json;

namespace Office365MCP.Tools.Search;

public static class JsonPath
{
    // Returns null when any step is missing or not an object, which Graph does for absent senders.
    public static string? GetString(JsonElement element, params string[] path)
    {
        foreach (var name in path)
        {
            if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out element))
                return null;
        }
        return element.ValueKind == JsonValueKind.String ? element.GetString() : null;
    }
}
