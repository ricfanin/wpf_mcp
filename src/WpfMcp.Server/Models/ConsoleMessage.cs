namespace WpfMcp.Server.Models;

/// <summary>
/// Represents a captured console output line from the WPF application.
/// </summary>
public sealed record ConsoleMessage(string Level, string Text, DateTime Timestamp);
