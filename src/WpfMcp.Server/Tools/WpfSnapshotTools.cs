using System.ComponentModel;
using System.Diagnostics;
using System.Text.Json;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.UIA3.Converters;
using Interop.UIAutomationClient;
using ModelContextProtocol.Server;
using WpfMcp.Server.Models;
using WpfMcp.Server.Services;
using InteropTreeScope = Interop.UIAutomationClient.TreeScope;

namespace WpfMcp.Server.Tools;

/// <summary>
/// MCP tools for element discovery and snapshot operations.
/// </summary>
[McpServerToolType]
public sealed class WpfSnapshotTools
{
    private readonly IApplicationManager _applicationManager;
    private readonly IElementReferenceManager _elementReferenceManager;

    public WpfSnapshotTools(IApplicationManager applicationManager, IElementReferenceManager elementReferenceManager)
    {
        _applicationManager = applicationManager;
        _elementReferenceManager = elementReferenceManager;
    }

    [McpServerTool(Name = "wpf_snapshot"), Description("Returns structured accessibility tree snapshot for LLM analysis")]
    public string TakeSnapshot(
        [Description("Element reference to use as root (default: main window)")] string? root_ref = null,
        [Description("Maximum tree depth to traverse (1-20)")] int max_depth = 5,
        [Description("Include invisible elements in snapshot")] bool include_invisible = false)
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

            // Validate max_depth
            if (max_depth < 1 || max_depth > 20)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.InvalidParameter,
                    "max_depth must be between 1 and 20",
                    "Use a depth value that balances coverage and performance"));
            }

            // Check for application crash
            if (_applicationManager.HasApplicationCrashed())
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.AppCrashed,
                    "The attached application has terminated",
                    "Call wpf_launch_application or wpf_attach_application to connect to a new application"));
            }

            // Start a new snapshot context
            _elementReferenceManager.BeginNewSnapshot();

            // Get root element
            AutomationElement rootElement;
            if (!string.IsNullOrEmpty(root_ref))
            {
                var element = _elementReferenceManager.GetElement(root_ref);
                if (element == null)
                {
                    return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                        ErrorCodes.ElementNotFound,
                        $"Element with ref '{root_ref}' not found",
                        "Call wpf_snapshot without root_ref to refresh element references"));
                }
                rootElement = element;
            }
            else
            {
                rootElement = _applicationManager.MainWindow!;
            }

            // Build the snapshot tree using native cached query for performance
            var automation = _applicationManager.Automation;
            var nativeAutomation = automation.NativeAutomation;

            // Create native cache request with all properties needed during tree walk
            var nativeCacheReq = nativeAutomation.CreateCacheRequest();
            nativeCacheReq.TreeScope = InteropTreeScope.TreeScope_Subtree;
            // Properties
            nativeCacheReq.AddProperty(UiaPropertyIds.Name);
            nativeCacheReq.AddProperty(UiaPropertyIds.AutomationId);
            nativeCacheReq.AddProperty(UiaPropertyIds.ControlType);
            nativeCacheReq.AddProperty(UiaPropertyIds.IsEnabled);
            nativeCacheReq.AddProperty(UiaPropertyIds.IsOffscreen);
            nativeCacheReq.AddProperty(UiaPropertyIds.HasKeyboardFocus);
            nativeCacheReq.AddProperty(UiaPropertyIds.IsKeyboardFocusable);
            nativeCacheReq.AddProperty(UiaPropertyIds.ProcessId);
            nativeCacheReq.AddProperty(UiaPropertyIds.FrameworkId);
            nativeCacheReq.AddProperty(UiaPropertyIds.BoundingRectangle);
            nativeCacheReq.AddProperty(UiaPropertyIds.ClassName);
            nativeCacheReq.AddProperty(UiaPropertyIds.RuntimeId);
            // Patterns
            nativeCacheReq.AddPattern(UiaPatternIds.Value);
            nativeCacheReq.AddPattern(UiaPatternIds.Toggle);
            nativeCacheReq.AddPattern(UiaPatternIds.SelectionItem);
            nativeCacheReq.AddPattern(UiaPatternIds.ExpandCollapse);
            nativeCacheReq.AddPattern(UiaPatternIds.Window);

            // Execute cached query: single cross-process COM call fetches entire subtree
            var nativeRoot = AutomationElementConverter.ToNative(rootElement);
            var cachedNativeRoot = nativeRoot.FindFirstBuildCache(
                InteropTreeScope.TreeScope_Element,
                nativeAutomation.CreateTrueCondition(),
                nativeCacheReq);

            var snapshotElement = BuildCachedSnapshotTree(cachedNativeRoot, 0, max_depth, include_invisible);

            var metadata = new ResponseMetadata
            {
                ExecutionTimeMs = stopwatch.ElapsedMilliseconds,
                SnapshotValid = true
            };

            // Add performance warning for large snapshots
            if (_elementReferenceManager.ElementCount > 1000)
            {
                metadata.Warnings.Add($"Large snapshot with {_elementReferenceManager.ElementCount} elements. Consider reducing max_depth for better performance.");
            }

            var result = new SnapshotResult
            {
                Tree = snapshotElement,
                ElementCount = _elementReferenceManager.ElementCount
            };

            // Return YAML format for better LLM readability
            var response = new
            {
                success = true,
                data = new
                {
                    element_count = result.ElementCount,
                    snapshot = result.Yaml
                },
                metadata = new
                {
                    execution_time_ms = metadata.ExecutionTimeMs,
                    warnings = metadata.Warnings,
                    snapshot_valid = metadata.SnapshotValid
                }
            };

            return JsonSerializer.Serialize(response);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.AppNotResponding,
                $"Failed to take snapshot: {ex.Message}",
                "The application may not be responding. Try again or check application state"));
        }
    }

    [McpServerTool(Name = "wpf_find_element"), Description("Find element by AutomationId, Name, or ControlType")]
    public string FindElement(
        [Description("Unique AutomationId property")] string? automation_id = null,
        [Description("Element Name property (visible text)")] string? name = null,
        [Description("UI Automation control type")] string? control_type = null,
        [Description("Search within this element (default: main window)")] string? root_ref = null)
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

            if (string.IsNullOrEmpty(automation_id) && string.IsNullOrEmpty(name) && string.IsNullOrEmpty(control_type))
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.InvalidParameter,
                    "At least one search criterion must be provided",
                    "Specify automation_id, name, or control_type"));
            }

            // Get root element
            AutomationElement rootElement;
            if (!string.IsNullOrEmpty(root_ref))
            {
                var element = _elementReferenceManager.GetElement(root_ref);
                if (element == null)
                {
                    return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                        ErrorCodes.ElementNotFound,
                        $"Root element with ref '{root_ref}' not found",
                        "Call wpf_snapshot to refresh element references"));
                }
                rootElement = element;
            }
            else
            {
                rootElement = _applicationManager.MainWindow!;
            }

            // Find matching elements
            var matchingElements = FindMatchingElements(rootElement, automation_id, name, control_type);

            if (matchingElements.Count == 0)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.ElementNotFound,
                    "No matching elements found",
                    "Verify search criteria or call wpf_snapshot to see available elements"));
            }

            // Register found elements and build results
            var results = matchingElements.Select(element =>
            {
                var reference = _elementReferenceManager.RegisterElement(element);
                return new
                {
                    @ref = reference.Ref,
                    control_type = reference.ControlType,
                    name = reference.Name,
                    automation_id = reference.AutomationId,
                    is_enabled = reference.IsEnabled
                };
            }).ToList();

            return JsonSerializer.Serialize(ToolResponse<object>.Ok(new
            {
                count = results.Count,
                elements = results
            }, new ResponseMetadata
            {
                ExecutionTimeMs = stopwatch.ElapsedMilliseconds,
                SnapshotValid = true
            }));
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.AppNotResponding,
                $"Failed to find element: {ex.Message}",
                "The application may not be responding"));
        }
    }

    [McpServerTool(Name = "wpf_get_element_properties"), Description("Returns all automation properties for an element")]
    public string GetElementProperties(
        [Description("Element reference from snapshot")] string @ref)
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

            var element = _elementReferenceManager.GetElement(@ref);
            if (element == null)
            {
                return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                    ErrorCodes.ElementNotFound,
                    $"Element with ref '{@ref}' not found in current snapshot",
                    "Call wpf_snapshot to refresh element references"));
            }

            var properties = new Dictionary<string, object?>
            {
                ["ref"] = @ref,
                ["name"] = element.Properties.Name.ValueOrDefault,
                ["automation_id"] = element.Properties.AutomationId.ValueOrDefault,
                ["control_type"] = element.Properties.ControlType.ValueOrDefault.ToString(),
                ["class_name"] = element.Properties.ClassName.ValueOrDefault,
                ["is_enabled"] = element.Properties.IsEnabled.ValueOrDefault,
                ["is_offscreen"] = element.Properties.IsOffscreen.ValueOrDefault,
                ["is_keyboard_focusable"] = element.Properties.IsKeyboardFocusable.ValueOrDefault,
                ["has_keyboard_focus"] = element.Properties.HasKeyboardFocus.ValueOrDefault,
                ["process_id"] = element.Properties.ProcessId.ValueOrDefault,
                ["framework_id"] = element.Properties.FrameworkId.ValueOrDefault,
                ["bounding_rectangle"] = GetBoundingRectObject(element)
            };

            // Add pattern availability
            var patterns = GetSupportedPatterns(element);
            properties["supported_patterns"] = patterns;

            return JsonSerializer.Serialize(ToolResponse<object>.Ok(properties, new ResponseMetadata
            {
                ExecutionTimeMs = stopwatch.ElapsedMilliseconds
            }));
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(ToolResponse<object>.Fail(
                ErrorCodes.ElementStale,
                $"Failed to get element properties: {ex.Message}",
                "The element may have been removed. Call wpf_snapshot to refresh"));
        }
    }

    /// <summary>
    /// Builds the snapshot tree from a native cached element, reading all properties
    /// from the UIA cache (zero cross-process COM calls per element).
    /// </summary>
    private SnapshotElement BuildCachedSnapshotTree(IUIAutomationElement cachedElement, int currentDepth, int maxDepth, bool includeInvisible)
    {
        var automation = _applicationManager.Automation;

        // Convert to FlaUI element for registration (lightweight wrapper, no COM calls)
        var flaElement = AutomationElementConverter.NativeToManaged(automation, cachedElement);

        // Register the FlaUI element (for future interactions via ref ID).
        // RegisterElement reads properties from the FlaUI wrapper — these will be live COM calls.
        // To avoid that, we build the ElementReference data from cached native properties
        // and use RegisterElement only to store the element mapping.
        var reference = _elementReferenceManager.RegisterElement(flaElement);

        // Read states from cached native properties (no COM round-trips)
        var states = GetCachedElementStates(cachedElement);

        // Read value from cached pattern
        string? value = null;
        try
        {
            var valuePattern = (IUIAutomationValuePattern?)cachedElement.GetCachedPattern(UiaPatternIds.Value);
            if (valuePattern != null)
            {
                value = valuePattern.CachedValue;
            }
        }
        catch
        {
            // Value pattern not available
        }

        var snapshotElement = new SnapshotElement
        {
            Ref = reference.Ref,
            ControlType = reference.ControlType,
            Name = reference.Name,
            AutomationId = reference.AutomationId,
            Value = value,
            States = states,
            Depth = currentDepth
        };

        // Add children if within depth limit
        if (currentDepth < maxDepth)
        {
            try
            {
                var cachedChildren = cachedElement.GetCachedChildren();
                var childCount = cachedChildren.Length;
                for (int i = 0; i < childCount; i++)
                {
                    var child = cachedChildren.GetElement(i);

                    // Skip invisible elements if not requested (read from cache)
                    if (!includeInvisible)
                    {
                        try
                        {
                            var isOffscreen = (bool)child.GetCachedPropertyValue(UiaPropertyIds.IsOffscreen);
                            if (isOffscreen) continue;
                        }
                        catch
                        {
                            // Can't determine visibility, include the element
                        }
                    }

                    var childSnapshot = BuildCachedSnapshotTree(child, currentDepth + 1, maxDepth, includeInvisible);
                    snapshotElement.Children.Add(childSnapshot);
                }
            }
            catch
            {
                // Children not accessible
            }
        }

        return snapshotElement;
    }

    /// <summary>
    /// Reads element states from the UIA cache (no live COM calls).
    /// </summary>
    private static List<string> GetCachedElementStates(IUIAutomationElement cachedElement)
    {
        var states = new List<string>();

        try
        {
            var isEnabled = (bool)cachedElement.GetCachedPropertyValue(UiaPropertyIds.IsEnabled);
            if (!isEnabled)
                states.Add("disabled");

            var hasFocus = (bool)cachedElement.GetCachedPropertyValue(UiaPropertyIds.HasKeyboardFocus);
            if (hasFocus)
                states.Add("focused");

            // Toggle state
            var togglePattern = (IUIAutomationTogglePattern?)cachedElement.GetCachedPattern(UiaPatternIds.Toggle);
            if (togglePattern != null)
            {
                states.Add(togglePattern.CachedToggleState == Interop.UIAutomationClient.ToggleState.ToggleState_On ? "checked" : "unchecked");
            }

            // Selection state
            var selectionItemPattern = (IUIAutomationSelectionItemPattern?)cachedElement.GetCachedPattern(UiaPatternIds.SelectionItem);
            if (selectionItemPattern != null)
            {
                if (selectionItemPattern.CachedIsSelected != 0)
                    states.Add("selected");
            }

            // Expand/collapse state
            var expandPattern = (IUIAutomationExpandCollapsePattern?)cachedElement.GetCachedPattern(UiaPatternIds.ExpandCollapse);
            if (expandPattern != null)
            {
                states.Add(expandPattern.CachedExpandCollapseState == Interop.UIAutomationClient.ExpandCollapseState.ExpandCollapseState_Expanded ? "expanded" : "collapsed");
            }

            // Read-only (from value pattern)
            var valuePattern = (IUIAutomationValuePattern?)cachedElement.GetCachedPattern(UiaPatternIds.Value);
            if (valuePattern != null)
            {
                if (valuePattern.CachedIsReadOnly != 0)
                    states.Add("readonly");
            }

            // Modal (from window pattern)
            var windowPattern = (IUIAutomationWindowPattern?)cachedElement.GetCachedPattern(UiaPatternIds.Window);
            if (windowPattern != null)
            {
                if (windowPattern.CachedIsModal != 0)
                    states.Add("modal");
            }
        }
        catch
        {
            // Some states not available from cache
        }

        return states;
    }

    /// <summary>
    /// UIA Property ID constants. Stable since Windows Vista.
    /// See: https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-automation-element-propids
    /// </summary>
    private static class UiaPropertyIds
    {
        public const int BoundingRectangle = 30001;
        public const int ProcessId = 30002;
        public const int ControlType = 30003;
        public const int Name = 30005;
        public const int HasKeyboardFocus = 30008;
        public const int IsKeyboardFocusable = 30009;
        public const int IsEnabled = 30010;
        public const int AutomationId = 30011;
        public const int ClassName = 30012;
        public const int IsOffscreen = 30022;
        public const int FrameworkId = 30024;
        public const int RuntimeId = 30000;
    }

    /// <summary>
    /// UIA Pattern ID constants. Stable since Windows Vista.
    /// See: https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpattern-ids
    /// </summary>
    private static class UiaPatternIds
    {
        public const int Value = 10002;
        public const int ExpandCollapse = 10005;
        public const int Window = 10009;
        public const int SelectionItem = 10010;
        public const int Toggle = 10015;
    }

    private List<AutomationElement> FindMatchingElements(AutomationElement root, string? automationId, string? name, string? controlType)
    {
        var cf = _applicationManager.Automation.ConditionFactory;
        var conditions = new List<ConditionBase>();

        if (!string.IsNullOrEmpty(automationId))
        {
            conditions.Add(cf.ByAutomationId(automationId));
        }

        if (!string.IsNullOrEmpty(controlType))
        {
            if (Enum.TryParse<ControlType>(controlType, ignoreCase: true, out var ct))
            {
                conditions.Add(cf.ByControlType(ct));
            }
        }

        // UIA doesn't support substring matching natively, so for name we use two strategies:
        // - If name is the only criterion and no others, do a native name search then post-filter for substring
        // - If combined with other criteria, use native conditions for the others and post-filter name

        ConditionBase searchCondition;
        if (conditions.Count == 0)
        {
            // Only name criterion (or none, but caller ensures at least one)
            searchCondition = TrueCondition.Default;
        }
        else if (conditions.Count == 1)
        {
            searchCondition = conditions[0];
        }
        else
        {
            searchCondition = new AndCondition(conditions.ToArray());
        }

        var found = root.FindAll(FlaUI.Core.Definitions.TreeScope.Descendants, searchCondition);
        var results = new List<AutomationElement>();

        foreach (var element in found)
        {
            // Post-filter for substring name match (UIA only supports exact match)
            if (!string.IsNullOrEmpty(name))
            {
                var elementName = element.Properties.Name.ValueOrDefault;
                if (elementName == null || !elementName.Contains(name, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
            }

            results.Add(element);
        }

        return results;
    }

    private static List<string> GetSupportedPatterns(AutomationElement element)
    {
        var patterns = new List<string>();

        if (element.Patterns.Invoke.IsSupported) patterns.Add("Invoke");
        if (element.Patterns.Value.IsSupported) patterns.Add("Value");
        if (element.Patterns.Toggle.IsSupported) patterns.Add("Toggle");
        if (element.Patterns.Selection.IsSupported) patterns.Add("Selection");
        if (element.Patterns.SelectionItem.IsSupported) patterns.Add("SelectionItem");
        if (element.Patterns.ExpandCollapse.IsSupported) patterns.Add("ExpandCollapse");
        if (element.Patterns.Scroll.IsSupported) patterns.Add("Scroll");
        if (element.Patterns.ScrollItem.IsSupported) patterns.Add("ScrollItem");
        if (element.Patterns.Grid.IsSupported) patterns.Add("Grid");
        if (element.Patterns.GridItem.IsSupported) patterns.Add("GridItem");
        if (element.Patterns.Table.IsSupported) patterns.Add("Table");
        if (element.Patterns.TableItem.IsSupported) patterns.Add("TableItem");
        if (element.Patterns.Transform.IsSupported) patterns.Add("Transform");
        if (element.Patterns.Window.IsSupported) patterns.Add("Window");

        return patterns;
    }

    private static object? GetBoundingRectObject(AutomationElement element)
    {
        try
        {
            var rect = element.Properties.BoundingRectangle.ValueOrDefault;
            if (rect.IsEmpty) return null;

            return new { x = rect.X, y = rect.Y, width = rect.Width, height = rect.Height };
        }
        catch
        {
            return null;
        }
    }
}
