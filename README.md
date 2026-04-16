# WPF-MCP Server

[![CI](https://github.com/ricfanin/wpf_mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/ricfanin/wpf_mcp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-8.0-512BD4)](https://dotnet.microsoft.com/)

A Model Context Protocol (MCP) server that enables AI agents to programmatically interact with Windows Presentation Foundation (WPF) applications. Think of it as **Playwright for WPF** - allowing Claude and other AI assistants to navigate, inspect, and interact with desktop applications.

## Features

- **Application Management** - Launch, attach to, and close WPF applications
- **Element Discovery** - Get accessibility tree snapshots in LLM-friendly YAML format
- **UI Interaction** - Click, type, toggle, select, scroll, and more
- **Window Management** - Switch between windows, minimize, maximize, restore
- **Screenshots** - Capture application or element screenshots as base64
- **Wait Conditions** - Wait for elements to become visible, enabled, or focused

## Quick Start

### Prerequisites

- Windows 10/11
- .NET 8.0 SDK
- A WPF application to automate

### Installation

#### Option 1: Install as Global Tool (Recommended)

```bash
# Install from NuGet (when published)
dotnet tool install --global WpfMcp.Server

# Or install from local build
.\scripts\install-local.ps1
```

#### Option 2: Clone and Run

```bash
# Clone the repository
git clone https://github.com/ricfanin/wpf_mcp.git
cd wpf_mcp

# Build the project
dotnet build

# Run the server
dotnet run --project src/WpfMcp.Server
```

### Configure MCP Client

#### If installed as Global Tool:

Add to your Claude Desktop/Claude Code configuration:

```json
{
  "mcpServers": {
    "wpf-mcp": {
      "command": "wpf-mcp",
      "args": []
    }
  }
}
```

#### If running from source:

For Claude Code, copy `.mcp.json.example` to `.mcp.json`:

```bash
cp .mcp.json.example .mcp.json
```

For Claude Desktop (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "wpf-mcp": {
      "command": "dotnet",
      "args": ["run", "--project", "C:/path/to/WPF-mcp/src/WpfMcp.Server"]
    }
  }
}
```

### Basic Usage

Once connected, you can ask Claude to interact with WPF applications:

```
"Launch the calculator app at C:\Windows\System32\calc.exe"
"Take a snapshot of the current UI"
"Click the '7' button"
"Type '123' in the input field"
```

## Available Tools

| Tool | Description |
|------|-------------|
| `wpf_launch_application` | Launch a WPF executable |
| `wpf_attach_application` | Attach to a running process |
| `wpf_close_application` | Close the attached application |
| `wpf_snapshot` | Get accessibility tree snapshot |
| `wpf_find_element` | Find elements by criteria |
| `wpf_get_element_properties` | Get detailed element properties |
| `wpf_click` | Click an element |
| `wpf_type` | Type text into an element |
| `wpf_set_value` | Set element value directly |
| `wpf_toggle` | Toggle checkbox/toggle button |
| `wpf_select` | Select item in dropdown/list |
| `wpf_expand_collapse` | Expand or collapse tree nodes |
| `wpf_press_key` | Send keyboard input |
| `wpf_scroll` | Scroll within a container |
| `wpf_scroll_into_view` | Scroll element into view |
| `wpf_focus` | Set keyboard focus |
| `wpf_list_windows` | List all application windows |
| `wpf_switch_window` | Switch to a different window |
| `wpf_window_action` | Minimize/maximize/restore/close |
| `wpf_take_screenshot` | Capture screenshot as base64 |
| `wpf_wait_for` | Wait for element condition |

## Documentation

- [Architecture Guide](docs/ARCHITECTURE.md) - Technical architecture and design decisions
- [Tools Reference](docs/TOOLS_REFERENCE.md) - Complete tool documentation with examples
- [Integration Guide](docs/INTEGRATION_GUIDE.md) - How to integrate with Claude and MCP clients

## Example Workflow

```
1. Launch application
   wpf_launch_application(path="C:\MyApp\App.exe")

2. Take a snapshot to see the UI
   wpf_snapshot()

   Output (YAML):
   - window "My Application" [ref=e1]
     - button "Login" [ref=e2] [enabled]
     - textbox "Username" [ref=e3] [value=""]
     - textbox "Password" [ref=e4] [value=""]

3. Type credentials
   wpf_type(element="Username field", ref="e3", text="user@example.com")
   wpf_type(element="Password field", ref="e4", text="secret123")

4. Click login
   wpf_click(element="Login button", ref="e2")

5. Wait for next screen
   wpf_wait_for(condition="exists", timeout_ms=5000)

6. Take new snapshot
   wpf_snapshot()
```

## Technical Stack

| Component | Technology |
|-----------|------------|
| Runtime | .NET 8.0 (Windows) |
| MCP SDK | [ModelContextProtocol](https://github.com/modelcontextprotocol/csharp-sdk) |
| UI Automation | [FlaUI](https://github.com/FlaUI/FlaUI) (UIA3) |
| Transport | stdio (JSON-RPC 2.0) |

## Project Structure

```
WPF-mcp/
├── src/WpfMcp.Server/
│   ├── Program.cs              # Entry point
│   ├── Models/                 # Data structures
│   │   ├── ToolResponse.cs     # Standard response schema
│   │   ├── ErrorCodes.cs       # Error constants
│   │   ├── SnapshotElement.cs  # Tree node model
│   │   └── ElementReference.cs # Element ref metadata
│   ├── Services/               # Core business logic
│   │   ├── ApplicationManager.cs
│   │   └── ElementReferenceManager.cs
│   └── Tools/                  # MCP tool implementations
│       ├── WpfApplicationTools.cs
│       ├── WpfSnapshotTools.cs
│       ├── WpfInteractionTools.cs
│       ├── WpfNavigationTools.cs
│       ├── WpfWindowTools.cs
│       └── WpfUtilityTools.cs
├── tests/WpfMcp.Server.Tests/  # Unit tests
└── docs/                       # Documentation
```

## Requirements for Target Applications

For best results, WPF applications should:

- Set `AutomationProperties.AutomationId` on interactive elements
- Use standard WPF controls (or custom controls with `AutomationPeer`)
- Avoid full-window overlays that block the automation tree
- Have meaningful `AutomationProperties.Name` values

## Building

```bash
# Debug build
dotnet build
# or
.\scripts\build.ps1

# Release build
dotnet build -c Release
# or
.\scripts\build.ps1 -Release

# Run tests
dotnet test

# Development mode (auto-rebuild on changes)
.\scripts\dev.ps1

# Create NuGet package
.\scripts\pack.ps1

# Install locally as global tool
.\scripts\install-local.ps1
```

## License

MIT License - See [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please read the contributing guidelines before submitting PRs.
