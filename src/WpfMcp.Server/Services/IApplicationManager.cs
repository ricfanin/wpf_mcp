using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using WpfMcp.Server.Models;

namespace WpfMcp.Server.Services;

/// <summary>
/// Manages the lifecycle of WPF applications for automation.
/// </summary>
public interface IApplicationManager
{
    /// <summary>
    /// The UIA3 automation instance used for element operations.
    /// </summary>
    UIA3Automation Automation { get; }

    /// <summary>
    /// Whether an application is currently attached.
    /// </summary>
    bool IsAttached { get; }

    /// <summary>
    /// The main window of the attached application.
    /// </summary>
    Window? MainWindow { get; }

    /// <summary>
    /// Launches a WPF application and waits for its main window.
    /// </summary>
    /// <param name="path">Path to the executable.</param>
    /// <param name="arguments">Command line arguments.</param>
    /// <param name="timeoutMs">Timeout in milliseconds.</param>
    /// <returns>The main window of the launched application.</returns>
    Task<Window> LaunchApplicationAsync(string path, string[]? arguments = null, int timeoutMs = 30000);

    /// <summary>
    /// Attaches to a running application by process name.
    /// </summary>
    /// <param name="processName">Name of the process (without .exe).</param>
    /// <returns>The main window of the attached application.</returns>
    Task<Window> AttachByNameAsync(string processName);

    /// <summary>
    /// Attaches to a running application by process ID.
    /// </summary>
    /// <param name="processId">The process ID.</param>
    /// <returns>The main window of the attached application.</returns>
    Task<Window> AttachByIdAsync(int processId);

    /// <summary>
    /// Closes the attached application.
    /// </summary>
    /// <param name="force">Whether to force kill if graceful close fails.</param>
    Task CloseApplicationAsync(bool force = false);

    /// <summary>
    /// Gets all windows belonging to the attached application.
    /// </summary>
    IReadOnlyList<Window> GetAllWindows();

    /// <summary>
    /// Checks if the attached application has crashed.
    /// </summary>
    bool HasApplicationCrashed();

    /// <summary>
    /// Gets captured console output from the launched application.
    /// </summary>
    IReadOnlyList<ConsoleMessage> GetConsoleOutput(string? level = null, int? limit = null);
}
