using System.ComponentModel;
using System.Diagnostics;
using System.Text.Json;
using ModelContextProtocol.Server;
using WpfMcp.Server.Models;
using WpfMcp.Server.Services;

namespace WpfMcp.Server.Tools;

/// <summary>
/// MCP tools for reading captured console output from the WPF application.
/// </summary>
[McpServerToolType]
public sealed class WpfConsoleTools
{
    private readonly IApplicationManager _applicationManager;

    public WpfConsoleTools(IApplicationManager applicationManager)
    {
        _applicationManager = applicationManager;
    }

    [McpServerTool(Name = "wpf_console_messages"), Description("Returns all console messages")]
    public string GetConsoleMessages(
        [Description("Level of the console messages to return. Each level includes the messages of more severe levels. Defaults to \"info\".")] string level = "info",
        [Description("Filename to save the console messages to. If not provided, messages are returned as text.")] string? filename = null)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            // Map level to filter: "error" = stderr only, "warning"/"info"/"debug" = all, "stdout"/"stderr" = exact match
            string? filterLevel = level.ToLowerInvariant() switch
            {
                "error" => "stderr",
                "warning" or "info" or "debug" => null, // all messages
                "stdout" => "stdout",
                "stderr" => "stderr",
                _ => null
            };

            var messages = _applicationManager.GetConsoleOutput(filterLevel);

            var formatted = messages.Select(m => $"[{m.Timestamp:HH:mm:ss.fff}] [{m.Level}] {m.Text}").ToList();

            if (!string.IsNullOrEmpty(filename))
            {
                File.WriteAllLines(filename, formatted);
                return JsonSerializer.Serialize(ToolResponse<object>.Ok(new
                {
                    saved_to = filename,
                    message_count = formatted.Count
                }, new ResponseMetadata
                {
                    ExecutionTimeMs = stopwatch.ElapsedMilliseconds
                }));
            }

            return JsonSerializer.Serialize(ToolResponse<object>.Ok(new
            {
                messages = formatted,
                message_count = formatted.Count
            }, new ResponseMetadata
            {
                ExecutionTimeMs = stopwatch.ElapsedMilliseconds
            }));
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.AppNotResponding,
                $"Failed to retrieve console messages: {ex.Message}",
                "The application may not be attached or responding"));
        }
    }
}
