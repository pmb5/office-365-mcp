namespace Office365MCP.Auth;

public interface IScopeProvider
{
    string[] GetDefaultScopes();
    string[] GetCurrentScopes();
}
