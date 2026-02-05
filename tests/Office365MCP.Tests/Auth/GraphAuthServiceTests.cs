using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using NSubstitute;
using Office365MCP.Auth;
using Xunit;

namespace Office365MCP.Tests.Auth;

public class GraphAuthServiceTests
{
    [Fact]
    public void GraphAuthService_Implements_IGraphAuthService()
    {
        var inMemorySettings = new Dictionary<string, string>
        {
            ["Graph:ClientId"] = "test-client-id",
            ["Graph:TenantId"] = "organizations",
            ["Graph:Scopes:0"] = "User.Read",
            ["Graph:Scopes:1"] = "Mail.Read"
        };

        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(inMemorySettings!)
            .Build();

        var logger = Substitute.For<ILogger<GraphAuthService>>();

        var service = new GraphAuthService(configuration, logger);

        Assert.IsAssignableFrom<IGraphAuthService>(service);
    }

    [Fact]
    public void Constructor_WithMissingClientId_ThrowsException()
    {
        var inMemorySettings = new Dictionary<string, string>
        {
            ["Graph:TenantId"] = "organizations",
            ["Graph:Scopes:0"] = "User.Read"
        };

        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(inMemorySettings!)
            .Build();

        var logger = Substitute.For<ILogger<GraphAuthService>>();

        Assert.Throws<InvalidOperationException>(() =>
            new GraphAuthService(configuration, logger));
    }

    [Fact]
    public void Constructor_WithMissingTenantId_ThrowsException()
    {
        var inMemorySettings = new Dictionary<string, string>
        {
            ["Graph:ClientId"] = "test-client-id",
            ["Graph:Scopes:0"] = "User.Read"
        };

        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(inMemorySettings!)
            .Build();

        var logger = Substitute.For<ILogger<GraphAuthService>>();

        Assert.Throws<InvalidOperationException>(() =>
            new GraphAuthService(configuration, logger));
    }

    [Fact]
    public void Constructor_WithMissingScopes_ThrowsException()
    {
        var inMemorySettings = new Dictionary<string, string>
        {
            ["Graph:ClientId"] = "test-client-id",
            ["Graph:TenantId"] = "organizations"
        };

        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(inMemorySettings!)
            .Build();

        var logger = Substitute.For<ILogger<GraphAuthService>>();

        Assert.Throws<InvalidOperationException>(() =>
            new GraphAuthService(configuration, logger));
    }
}
