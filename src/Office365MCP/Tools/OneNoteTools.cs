using System.ComponentModel;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using Office365MCP.Graph;

namespace Office365MCP.Tools;

[McpServerToolType]
public class OneNoteTools
{
    private readonly IGraphClientService _graphClient;
    private readonly ILogger<OneNoteTools> _logger;

    public OneNoteTools(IGraphClientService graphClient, ILogger<OneNoteTools> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    [McpServerTool, Description("List all OneNote notebooks")]
    public async Task<string> ListNotebooks()
    {
        try
        {
            var response = await _graphClient.GetAsync("me/onenote/notebooks");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list notebooks");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("List sections in a specific notebook")]
    public async Task<string> ListSections(
        [Description("The ID of the notebook")] string notebookId)
    {
        try
        {
            var response = await _graphClient.GetAsync($"me/onenote/notebooks/{notebookId}/sections");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list sections for notebook {NotebookId}", notebookId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("List pages in a specific section")]
    public async Task<string> ListPages(
        [Description("The ID of the section")] string sectionId)
    {
        try
        {
            var response = await _graphClient.GetAsync($"me/onenote/sections/{sectionId}/pages");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list pages for section {SectionId}", sectionId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Get the HTML content of a specific OneNote page")]
    public async Task<string> GetPageContent(
        [Description("The ID of the page")] string pageId)
    {
        try
        {
            var response = await _graphClient.GetAsync($"me/onenote/pages/{pageId}/content");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to get content for page {PageId}", pageId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Create a new OneNote page")]
    public async Task<string> CreatePage(
        [Description("The ID of the section where the page will be created")] string sectionId,
        [Description("Page title")] string title,
        [Description("Page content in HTML format")] string htmlContent)
    {
        try
        {
            var html = $@"<!DOCTYPE html>
<html>
  <head>
    <title>{title}</title>
  </head>
  <body>
    {htmlContent}
  </body>
</html>";

            var response = await _graphClient.PostAsync($"me/onenote/sections/{sectionId}/pages", html);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to create page in section {SectionId}", sectionId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }
}
