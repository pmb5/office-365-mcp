using System.Text;

namespace Office365MCP.Tools.Search;

// nextPage tokens are opaque to the agent so it passes them back unchanged rather than editing them.
public static class SearchCursor
{
    public static string Encode(string value) =>
        Convert.ToBase64String(Encoding.UTF8.GetBytes(value)).TrimEnd('=').Replace('+', '-').Replace('/', '_');

    public static string Decode(string token)
    {
        try
        {
            var base64 = token.Trim().Replace('-', '+').Replace('_', '/');
            base64 = base64.PadRight(base64.Length + (4 - base64.Length % 4) % 4, '=');
            return Encoding.UTF8.GetString(Convert.FromBase64String(base64));
        }
        catch (FormatException)
        {
            throw new ArgumentException("nextPage is not a token from a previous search response.");
        }
    }
}
