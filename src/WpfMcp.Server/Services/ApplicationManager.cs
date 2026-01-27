using System.Collections.Concurrent;
using System.Diagnostics;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using WpfMcp.Server.Models;

namespace WpfMcp.Server.Services;

/// <summary>
/// Manages WPF application lifecycle for automation.
/// </summary>
public sealed class ApplicationManager : IApplicationManager, IDisposable
{
    private readonly UIA3Automation _automation;
    private Application? _application;
    private Window? _mainWindow;
    private bool _disposed;
    private readonly ConcurrentQueue<ConsoleMessage> _consoleBuffer = new();
    private const int MaxBufferSize = 1000;

    public ApplicationManager()
    {
        _automation = new UIA3Automation();
    }

    public UIA3Automation Automation => _automation;

    public bool IsAttached => _application != null && !HasApplicationCrashed();

    public Window? MainWindow => _mainWindow;

    public async Task<Window> LaunchApplicationAsync(string path, string[]? arguments = null, int timeoutMs = 30000)
    {
        ThrowIfDisposed();

        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"Application not found: {path}", path);
        }

        // Close any existing application
        if (_application != null)
        {
            await CloseApplicationAsync(force: true);
        }

        var processStartInfo = new ProcessStartInfo(path)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = false
        };

        if (arguments != null)
        {
            processStartInfo.Arguments = string.Join(" ", arguments);
        }

        var process = Process.Start(processStartInfo)
            ?? throw new InvalidOperationException($"Failed to start process: {path}");

        process.OutputDataReceived += (_, e) => { if (e.Data != null) EnqueueMessage("stdout", e.Data); };
        process.ErrorDataReceived += (_, e) => { if (e.Data != null) EnqueueMessage("stderr", e.Data); };
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        _application = Application.Attach(process);

        // Wait for main window
        using var cts = new CancellationTokenSource(timeoutMs);
        _mainWindow = await WaitForMainWindowAsync(_application, cts.Token);

        return _mainWindow;
    }

    public Task<Window> AttachByNameAsync(string processName)
    {
        ThrowIfDisposed();

        var processes = Process.GetProcessesByName(processName);
        if (processes.Length == 0)
        {
            throw new InvalidOperationException($"No process found with name: {processName}");
        }

        // Use the first matching process
        return AttachByIdAsync(processes[0].Id);
    }

    public async Task<Window> AttachByIdAsync(int processId)
    {
        ThrowIfDisposed();

        var process = Process.GetProcessById(processId);
        _application = Application.Attach(process);

        using var cts = new CancellationTokenSource(30000);
        _mainWindow = await WaitForMainWindowAsync(_application, cts.Token);

        return _mainWindow;
    }

    public async Task CloseApplicationAsync(bool force = false)
    {
        ThrowIfDisposed();

        if (_application == null) return;

        try
        {
            if (!force)
            {
                // Try graceful close first
                _application.Close();

                // Wait briefly for graceful close
                await Task.Delay(1000);
            }

            if (force || !_application.HasExited)
            {
                _application.Kill();
            }
        }
        finally
        {
            _application = null;
            _mainWindow = null;
            _consoleBuffer.Clear();
        }
    }

    public IReadOnlyList<Window> GetAllWindows()
    {
        ThrowIfDisposed();

        if (_application == null)
        {
            return Array.Empty<Window>();
        }

        return _application.GetAllTopLevelWindows(_automation);
    }

    public bool HasApplicationCrashed()
    {
        if (_application == null) return false;

        try
        {
            return _application.HasExited;
        }
        catch
        {
            return true;
        }
    }

    public IReadOnlyList<ConsoleMessage> GetConsoleOutput(string? level = null, int? limit = null)
    {
        var messages = _consoleBuffer.ToArray().AsEnumerable();

        if (!string.IsNullOrEmpty(level))
        {
            messages = messages.Where(m => string.Equals(m.Level, level, StringComparison.OrdinalIgnoreCase));
        }

        if (limit.HasValue && limit.Value > 0)
        {
            messages = messages.TakeLast(limit.Value);
        }

        return messages.ToList();
    }

    private void EnqueueMessage(string level, string text)
    {
        _consoleBuffer.Enqueue(new ConsoleMessage(level, text, DateTime.UtcNow));

        while (_consoleBuffer.Count > MaxBufferSize)
        {
            _consoleBuffer.TryDequeue(out _);
        }
    }

    private async Task<Window> WaitForMainWindowAsync(Application application, CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                var mainWindow = application.GetMainWindow(_automation);
                if (mainWindow != null)
                {
                    return mainWindow;
                }
            }
            catch
            {
                // Window not ready yet
            }

            await Task.Delay(100, cancellationToken);
        }

        throw new TimeoutException("Timed out waiting for main window");
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }

    public void Dispose()
    {
        if (_disposed) return;

        _application?.Dispose();
        _automation.Dispose();
        _disposed = true;
    }
}
