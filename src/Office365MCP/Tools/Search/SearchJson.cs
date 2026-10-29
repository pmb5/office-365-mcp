using System.Text.Encodings.Web;
using System.Text.Json;

namespace Office365MCP.Tools.Search;

public static class SearchJson
{
    // Responses are read by an agent, not rendered as HTML, so the default escaping of quotes and
    // angle brackets (", <) only makes hints and previews harder to read.
    public static readonly JsonSerializerOptions Options = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };
}
