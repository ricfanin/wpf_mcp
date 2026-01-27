using System.ComponentModel;
using System.Diagnostics;
using System.Text.Json;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using ModelContextProtocol.Server;
using WpfMcp.Server.Models;
using WpfMcp.Server.Services;

namespace WpfMcp.Server.Tools;

/// <summary>
/// MCP tools for interacting with UI elements.
/// </summary>
[McpServerToolType]
public sealed class WpfInteractionTools
{
    private readonly IApplicationManager _applicationManager;
    private readonly IElementReferenceManager _elementReferenceManager;

    private const int MaxTextLength = 10000;

    public WpfInteractionTools(IApplicationManager applicationManager, IElementReferenceManager elementReferenceManager)
    {
        _applicationManager = applicationManager;
        _elementReferenceManager = elementReferenceManager;
    }

    [McpServerTool(Name = "wpf_click"), Description("Clicks an element using InvokePattern or mouse simulation")]
    public string Click(
        [Description("Human-readable element description for permission")] string element,
        [Description("Element reference from snapshot")] string @ref,
        [Description("Type of click to perform: single, double, or right")] string click_type = "single")
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            var validationResult = ValidateElementAccess(@ref, out var automationElement);
            if (validationResult != null) return validationResult;

            // Check if element is enabled
            if (!automationElement!.Properties.IsEnabled.ValueOrDefault)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.ElementNotEnabled,
                    $"Element '{element}' is disabled",
                    "Wait for the element to be enabled before clicking"));
            }

            // Check if element is visible
            if (automationElement.Properties.IsOffscreen.ValueOrDefault)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.ElementNotVisible,
                    $"Element '{element}' is not visible on screen",
                    "Call wpf_scroll_into_view to make the element visible first"));
            }

            // Try InvokePattern first for single clicks
            if (click_type == "single" && automationElement.Patterns.Invoke.IsSupported)
            {
                automationElement.Patterns.Invoke.Pattern.Invoke();
            }
            else
            {
                // Fall back to mouse simulation
                var point = automationElement.GetClickablePoint();

                switch (click_type.ToLowerInvariant())
                {
                    case "single":
                        Mouse.Click(point);
                        break;
                    case "double":
                        Mouse.DoubleClick(point);
                        break;
                    case "right":
                        Mouse.RightClick(point);
                        break;
                    default:
                        return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                            ErrorCodes.InvalidParameter,
                            $"Invalid click_type: {click_type}",
                            "Use 'single', 'double', or 'right'"));
                }
            }

            return JsonSerializer.Serialize(ToolResponse<object>.Ok(new
            {
                clicked = true,
                element_description = element,
                click_type
            }, new ResponseMetadata
            {
                ExecutionTimeMs = stopwatch.ElapsedMilliseconds
            }));
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.ElementDisappeared,
                $"Failed to click element: {ex.Message}",
                "The element may have been removed. Call wpf_snapshot to refresh"));
        }
    }

    [McpServerTool(Name = "wpf_type"), Description("Types text into a text input element")]
    public string Type(
        [Description("Human-readable element description")] string element,
        [Description("Element reference from snapshot")] string @ref,
        [Description("Text to type")] string text,
        [Description("Clear existing text before typing")] bool clear_first = true)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            // Validate text length
            if (text.Length > MaxTextLength)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.ValueTooLong,
                    $"Text exceeds maximum length of {MaxTextLength} characters",
                    "Reduce the text length or split into multiple type operations"));
            }

            var validationResult = ValidateElementAccess(@ref, out var automationElement);
            if (validationResult != null) return validationResult;

            // Check if element is enabled
            if (!automationElement!.Properties.IsEnabled.ValueOrDefault)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.ElementNotEnabled,
                    $"Element '{element}' is disabled",
                    "Wait for the element to be enabled before typing"));
            }

            // Check if element is read-only
            if (automationElement.Patterns.Value.IsSupported &&
                automationElement.Patterns.Value.Pattern.IsReadOnly.Value)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.ElementReadOnly,
                    $"Element '{element}' is read-only",
                    "This element cannot be modified"));
            }

            // Focus the element
            automationElement.Focus();
            Wait.UntilInputIsProcessed();

            // Clear existing text if requested
            if (clear_first)
            {
                if (automationElement.Patterns.Value.IsSupported)
                {
                    automationElement.Patterns.Value.Pattern.SetValue(string.Empty);
                }
                else
                {
                    // Select all and delete
                    Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
                    Keyboard.Type(VirtualKeyShort.DELETE);
                    Wait.UntilInputIsProcessed();
                }
            }

            // Type the text
            if (string.IsNullOrEmpty(text))
            {
                // Nothing to type, just cleared
            }
            else
            {
                Keyboard.Type(text);
                Wait.UntilInputIsProcessed();
            }

            return JsonSerializer.Serialize(ToolResponse<object>.Ok(new
            {
                typed = true,
                element_description = element,
                text_length = text.Length,
                cleared_first = clear_first
            }, new ResponseMetadata
            {
                ExecutionTimeMs = stopwatch.ElapsedMilliseconds
            }));
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.ElementDisappeared,
                $"Failed to type into element: {ex.Message}",
                "The element may have been removed. Call wpf_snapshot to refresh"));
        }
    }

    [McpServerTool(Name = "wpf_set_value"), Description("Sets value directly using ValuePattern")]
    public string SetValue(
        [Description("Human-readable element description")] string element,
        [Description("Element reference from snapshot")] string @ref,
        [Description("Value to set")] string value)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            // Validate value length
            if (value.Length > MaxTextLength)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.ValueTooLong,
                    $"Value exceeds maximum length of {MaxTextLength} characters",
                    "Reduce the value length"));
            }

            var validationResult = ValidateElementAccess(@ref, out var automationElement);
            if (validationResult != null) return validationResult;

            // Check if ValuePattern is supported
            if (!automationElement!.Patterns.Value.IsSupported)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.PatternNotSupported,
                    $"Element '{element}' does not support ValuePattern",
                    "Use wpf_type instead for keyboard input"));
            }

            // Check if element is read-only
            if (automationElement.Patterns.Value.Pattern.IsReadOnly.Value)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.ElementReadOnly,
                    $"Element '{element}' is read-only",
                    "This element cannot be modified"));
            }

            automationElement.Patterns.Value.Pattern.SetValue(value);

            return JsonSerializer.Serialize(ToolResponse<object>.Ok(new
            {
                value_set = true,
                element_description = element
            }, new ResponseMetadata
            {
                ExecutionTimeMs = stopwatch.ElapsedMilliseconds
            }));
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.ElementDisappeared,
                $"Failed to set value: {ex.Message}",
                "The element may have been removed. Call wpf_snapshot to refresh"));
        }
    }

    [McpServerTool(Name = "wpf_toggle"), Description("Toggles a checkbox or toggle button")]
    public string Toggle(
        [Description("Human-readable element description")] string element,
        [Description("Element reference from snapshot")] string @ref,
        [Description("Desired state: on, off, or toggle")] string target_state = "toggle")
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            var validationResult = ValidateElementAccess(@ref, out var automationElement);
            if (validationResult != null) return validationResult;

            // Check if TogglePattern is supported
            if (!automationElement!.Patterns.Toggle.IsSupported)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.PatternNotSupported,
                    $"Element '{element}' does not support TogglePattern",
                    "This element cannot be toggled"));
            }

            // Check if element is enabled
            if (!automationElement.Properties.IsEnabled.ValueOrDefault)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.ElementNotEnabled,
                    $"Element '{element}' is disabled",
                    "Wait for the element to be enabled before toggling"));
            }

            var currentState = automationElement.Patterns.Toggle.Pattern.ToggleState.Value;
            bool shouldToggle = target_state.ToLowerInvariant() switch
            {
                "on" => currentState != ToggleState.On,
                "off" => currentState != ToggleState.Off,
                "toggle" => true,
                _ => throw new ArgumentException($"Invalid target_state: {target_state}")
            };

            if (shouldToggle)
            {
                automationElement.Patterns.Toggle.Pattern.Toggle();
            }

            var newState = automationElement.Patterns.Toggle.Pattern.ToggleState.Value;

            return JsonSerializer.Serialize(ToolResponse<object>.Ok(new
            {
                toggled = shouldToggle,
                element_description = element,
                previous_state = currentState.ToString().ToLowerInvariant(),
                new_state = newState.ToString().ToLowerInvariant()
            }, new ResponseMetadata
            {
                ExecutionTimeMs = stopwatch.ElapsedMilliseconds
            }));
        }
        catch (ArgumentException ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.InvalidParameter,
                ex.Message,
                "Use 'on', 'off', or 'toggle' for target_state"));
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.ElementDisappeared,
                $"Failed to toggle element: {ex.Message}",
                "The element may have been removed. Call wpf_snapshot to refresh"));
        }
    }

    [McpServerTool(Name = "wpf_select"), Description("Selects an item in a selection control")]
    public string Select(
        [Description("Human-readable element description")] string element,
        [Description("Element reference for the container")] string @ref,
        [Description("Item text or reference to select")] string? item = null,
        [Description("Direct reference to item element")] string? item_ref = null)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            var validationResult = ValidateElementAccess(@ref, out var automationElement);
            if (validationResult != null) return validationResult;

            if (string.IsNullOrEmpty(item) && string.IsNullOrEmpty(item_ref))
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.InvalidParameter,
                    "Either 'item' or 'item_ref' must be provided",
                    "Specify the item to select by text or reference"));
            }

            AutomationElement? itemElement = null;

            // Try to find item by reference first
            if (!string.IsNullOrEmpty(item_ref))
            {
                itemElement = _elementReferenceManager.GetElement(item_ref);
                if (itemElement == null)
                {
                    return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                        ErrorCodes.ElementNotFound,
                        $"Item with ref '{item_ref}' not found",
                        "Call wpf_snapshot to refresh element references"));
                }
            }
            else if (!string.IsNullOrEmpty(item))
            {
                // Find item by text
                var children = automationElement!.FindAllChildren();
                itemElement = children.FirstOrDefault(c =>
                {
                    var name = c.Properties.Name.ValueOrDefault;
                    return name != null && name.Contains(item, StringComparison.OrdinalIgnoreCase);
                });

                if (itemElement == null)
                {
                    var availableItems = children
                        .Select(c => c.Properties.Name.ValueOrDefault)
                        .Where(n => !string.IsNullOrEmpty(n))
                        .Take(10)
                        .ToList();

                    return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                        ErrorCodes.ItemNotFound,
                        $"Item '{item}' not found in the selection control",
                        $"Available items: {string.Join(", ", availableItems)}"));
                }
            }

            // Try to select the item
            if (itemElement!.Patterns.SelectionItem.IsSupported)
            {
                itemElement.Patterns.SelectionItem.Pattern.Select();
            }
            else
            {
                // Fall back to clicking the item
                itemElement.Click();
            }

            return JsonSerializer.Serialize(ToolResponse<object>.Ok(new
            {
                selected = true,
                element_description = element,
                selected_item = itemElement.Properties.Name.ValueOrDefault
            }, new ResponseMetadata
            {
                ExecutionTimeMs = stopwatch.ElapsedMilliseconds
            }));
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.ElementDisappeared,
                $"Failed to select item: {ex.Message}",
                "The element may have been removed. Call wpf_snapshot to refresh"));
        }
    }

    [McpServerTool(Name = "wpf_expand_collapse"), Description("Expands or collapses an element")]
    public string ExpandCollapse(
        [Description("Human-readable element description")] string element,
        [Description("Element reference from snapshot")] string @ref,
        [Description("Action to perform: expand, collapse, or toggle")] string action = "toggle")
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            var validationResult = ValidateElementAccess(@ref, out var automationElement);
            if (validationResult != null) return validationResult;

            // Check if ExpandCollapsePattern is supported
            if (!automationElement!.Patterns.ExpandCollapse.IsSupported)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.PatternNotSupported,
                    $"Element '{element}' does not support ExpandCollapsePattern",
                    "This element cannot be expanded or collapsed"));
            }

            var currentState = automationElement.Patterns.ExpandCollapse.Pattern.ExpandCollapseState.Value;

            switch (action.ToLowerInvariant())
            {
                case "expand":
                    automationElement.Patterns.ExpandCollapse.Pattern.Expand();
                    break;
                case "collapse":
                    automationElement.Patterns.ExpandCollapse.Pattern.Collapse();
                    break;
                case "toggle":
                    if (currentState == ExpandCollapseState.Expanded)
                        automationElement.Patterns.ExpandCollapse.Pattern.Collapse();
                    else
                        automationElement.Patterns.ExpandCollapse.Pattern.Expand();
                    break;
                default:
                    return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                        ErrorCodes.InvalidParameter,
                        $"Invalid action: {action}",
                        "Use 'expand', 'collapse', or 'toggle'"));
            }

            var newState = automationElement.Patterns.ExpandCollapse.Pattern.ExpandCollapseState.Value;

            return JsonSerializer.Serialize(ToolResponse<object>.Ok(new
            {
                action_performed = action,
                element_description = element,
                previous_state = currentState.ToString().ToLowerInvariant(),
                new_state = newState.ToString().ToLowerInvariant()
            }, new ResponseMetadata
            {
                ExecutionTimeMs = stopwatch.ElapsedMilliseconds
            }));
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.ElementDisappeared,
                $"Failed to expand/collapse element: {ex.Message}",
                "The element may have been removed. Call wpf_snapshot to refresh"));
        }
    }

    [McpServerTool(Name = "wpf_press_key"), Description("Sends keyboard key press to focused element")]
    public string PressKey(
        [Description("Key to press (e.g., 'Enter', 'Tab', 'Escape', 'F1', 'Ctrl+S')")] string key,
        [Description("Optional element reference to focus first")] string? @ref = null)
    {
        var stopwatch = Stopwatch.StartNew();

        try
        {
            if (!_applicationManager.IsAttached)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.AppNotAttached,
                    "No application is currently attached",
                    "Call wpf_launch_application or wpf_attach_application first"));
            }

            // Focus element if specified
            if (!string.IsNullOrEmpty(@ref))
            {
                var element = _elementReferenceManager.GetElement(@ref);
                if (element == null)
                {
                    return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                        ErrorCodes.ElementNotFound,
                        $"Element with ref '{@ref}' not found",
                        "Call wpf_snapshot to refresh element references"));
                }
                element.Focus();
                Wait.UntilInputIsProcessed();
            }

            // Parse and send the key
            var keys = ParseKeySequence(key);
            if (keys == null)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.InvalidParameter,
                    $"Invalid key specification: {key}",
                    "Use key names like 'Enter', 'Tab', 'Escape', 'F1', or combinations like 'Ctrl+S'"));
            }

            if (keys.Length == 1)
            {
                Keyboard.Type(keys[0]);
            }
            else
            {
                Keyboard.TypeSimultaneously(keys);
            }
            Wait.UntilInputIsProcessed();

            return JsonSerializer.Serialize(ToolResponse<object>.Ok(new
            {
                key_pressed = key,
                focused_ref = @ref
            }, new ResponseMetadata
            {
                ExecutionTimeMs = stopwatch.ElapsedMilliseconds
            }));
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.AppNotResponding,
                $"Failed to press key: {ex.Message}",
                "The application may not be responding"));
        }
    }

    private string? ValidateElementAccess(string refId, out AutomationElement? element)
    {
        element = null;

        if (!_applicationManager.IsAttached)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.AppNotAttached,
                "No application is currently attached",
                "Call wpf_launch_application or wpf_attach_application first"));
        }

        if (_applicationManager.HasApplicationCrashed())
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.AppCrashed,
                "The attached application has terminated",
                "Call wpf_launch_application or wpf_attach_application to connect to a new application"));
        }

        element = _elementReferenceManager.GetElement(refId);
        if (element == null)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.ElementNotFound,
                $"Element with ref '{refId}' not found in current snapshot",
                "Call wpf_snapshot to refresh element references"));
        }

        return null;
    }

    private static VirtualKeyShort[]? ParseKeySequence(string key)
    {
        var parts = key.Split('+').Select(p => p.Trim().ToUpperInvariant()).ToArray();
        var keys = new List<VirtualKeyShort>();

        foreach (var part in parts)
        {
            var vk = part switch
            {
                "CTRL" or "CONTROL" => VirtualKeyShort.CONTROL,
                "ALT" => VirtualKeyShort.ALT,
                "SHIFT" => VirtualKeyShort.SHIFT,
                "WIN" or "WINDOWS" => VirtualKeyShort.LWIN,
                "ENTER" or "RETURN" => VirtualKeyShort.RETURN,
                "TAB" => VirtualKeyShort.TAB,
                "ESCAPE" or "ESC" => VirtualKeyShort.ESCAPE,
                "SPACE" => VirtualKeyShort.SPACE,
                "BACKSPACE" or "BACK" => VirtualKeyShort.BACK,
                "DELETE" or "DEL" => VirtualKeyShort.DELETE,
                "INSERT" or "INS" => VirtualKeyShort.INSERT,
                "HOME" => VirtualKeyShort.HOME,
                "END" => VirtualKeyShort.END,
                "PAGEUP" or "PGUP" => VirtualKeyShort.PRIOR,
                "PAGEDOWN" or "PGDN" => VirtualKeyShort.NEXT,
                "UP" => VirtualKeyShort.UP,
                "DOWN" => VirtualKeyShort.DOWN,
                "LEFT" => VirtualKeyShort.LEFT,
                "RIGHT" => VirtualKeyShort.RIGHT,
                "F1" => VirtualKeyShort.F1,
                "F2" => VirtualKeyShort.F2,
                "F3" => VirtualKeyShort.F3,
                "F4" => VirtualKeyShort.F4,
                "F5" => VirtualKeyShort.F5,
                "F6" => VirtualKeyShort.F6,
                "F7" => VirtualKeyShort.F7,
                "F8" => VirtualKeyShort.F8,
                "F9" => VirtualKeyShort.F9,
                "F10" => VirtualKeyShort.F10,
                "F11" => VirtualKeyShort.F11,
                "F12" => VirtualKeyShort.F12,
                "A" => VirtualKeyShort.KEY_A,
                "B" => VirtualKeyShort.KEY_B,
                "C" => VirtualKeyShort.KEY_C,
                "D" => VirtualKeyShort.KEY_D,
                "E" => VirtualKeyShort.KEY_E,
                "F" => VirtualKeyShort.KEY_F,
                "G" => VirtualKeyShort.KEY_G,
                "H" => VirtualKeyShort.KEY_H,
                "I" => VirtualKeyShort.KEY_I,
                "J" => VirtualKeyShort.KEY_J,
                "K" => VirtualKeyShort.KEY_K,
                "L" => VirtualKeyShort.KEY_L,
                "M" => VirtualKeyShort.KEY_M,
                "N" => VirtualKeyShort.KEY_N,
                "O" => VirtualKeyShort.KEY_O,
                "P" => VirtualKeyShort.KEY_P,
                "Q" => VirtualKeyShort.KEY_Q,
                "R" => VirtualKeyShort.KEY_R,
                "S" => VirtualKeyShort.KEY_S,
                "T" => VirtualKeyShort.KEY_T,
                "U" => VirtualKeyShort.KEY_U,
                "V" => VirtualKeyShort.KEY_V,
                "W" => VirtualKeyShort.KEY_W,
                "X" => VirtualKeyShort.KEY_X,
                "Y" => VirtualKeyShort.KEY_Y,
                "Z" => VirtualKeyShort.KEY_Z,
                "0" => VirtualKeyShort.KEY_0,
                "1" => VirtualKeyShort.KEY_1,
                "2" => VirtualKeyShort.KEY_2,
                "3" => VirtualKeyShort.KEY_3,
                "4" => VirtualKeyShort.KEY_4,
                "5" => VirtualKeyShort.KEY_5,
                "6" => VirtualKeyShort.KEY_6,
                "7" => VirtualKeyShort.KEY_7,
                "8" => VirtualKeyShort.KEY_8,
                "9" => VirtualKeyShort.KEY_9,
                _ => (VirtualKeyShort?)null
            };

            if (vk == null) return null;
            keys.Add(vk.Value);
        }

        return keys.ToArray();
    }
}
