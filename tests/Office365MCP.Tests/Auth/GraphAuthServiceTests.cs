using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using NSubstitute;
using Office365MCP.Auth;
using Xunit;

namespace Office365MCP.Tests.Auth;

public class GraphAuthServiceTests
{
    private IPermissionHelper CreateMockPermissionHelper()
    {
        var permissionHelper = Substitute.For<IPermissionHelper>();
        permissionHelper.GetDefaultScopes().Returns(new[] { "User.Read", "Mail.Read", "Calendars.Read", "Chat.Read", "Notes.Read" });
        permissionHelper.GetCurrentScopes().Returns(new[] { "User.Read", "Mail.Read", "Calendars.Read", "Chat.Read", "Notes.Read" });
        return permissionHelper;
    }

    [Fact]
    public void GraphAuthService_Implements_IGraphAuthService()
    {
        var inMemorySettings = new Dictionary<string, string>
        {
            ["Graph:ClientId"] = "test-client-id",
            ["Graph:TenantId"] = "organizations"
        };

        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(inMemorySettings!)
            .Build();

        var permissionHelper = CreateMockPermissionHelper();
        var logger = Substitute.For<ILogger<GraphAuthService>>();

        var service = new GraphAuthService(configuration, permissionHelper, logger);

        Assert.IsAssignableFrom<IGraphAuthService>(service);
    }

    [Fact]
    public void Constructor_WithMissingClientId_ThrowsException()
    {
        var inMemorySettings = new Dictionary<string, string>
        {
            ["Graph:TenantId"] = "organizations"
        };

        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(inMemorySettings!)
            .Build();

        var permissionHelper = CreateMockPermissionHelper();
        var logger = Substitute.For<ILogger<GraphAuthService>>();

        Assert.Throws<InvalidOperationException>(() =>
            new GraphAuthService(configuration, permissionHelper, logger));
    }

    [Fact]
    public void Constructor_WithMissingTenantId_ThrowsException()
    {
        var inMemorySettings = new Dictionary<string, string>
        {
            ["Graph:ClientId"] = "test-client-id"
        };

        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(inMemorySettings!)
            .Build();

        var permissionHelper = CreateMockPermissionHelper();
        var logger = Substitute.For<ILogger<GraphAuthService>>();

        Assert.Throws<InvalidOperationException>(() =>
            new GraphAuthService(configuration, permissionHelper, logger));
    }

    [Fact]
    public void Constructor_CreatesInstanceSuccessfully()
    {
        var inMemorySettings = new Dictionary<string, string>
        {
            ["Graph:ClientId"] = "test-client-id",
            ["Graph:TenantId"] = "organizations"
        };

        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(inMemorySettings!)
            .Build();

        var permissionHelper = CreateMockPermissionHelper();
        var logger = Substitute.For<ILogger<GraphAuthService>>();

        var service = new GraphAuthService(configuration, permissionHelper, logger);
        Assert.NotNull(service);
    }
}
