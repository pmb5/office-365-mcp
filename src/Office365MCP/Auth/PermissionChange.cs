namespace Office365MCP.Auth;

public class PermissionChange
{
    public string Permission { get; set; } = string.Empty;
    public DateTime Timestamp { get; set; }
    public string Reason { get; set; } = string.Empty;
    public string AddedBy { get; set; } = "User";
    public bool IsReverted { get; set; } = false;
}

public class PermissionChangeHistory
{
    public List<PermissionChange> Changes { get; set; } = new();
    public DateTime LastModified { get; set; }
}
