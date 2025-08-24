using System;
using System.Windows.Automation;

namespace BrowserTabs
{
    /// <summary>
    /// Represents a single browser tab in a Chromium-based browser.
    /// </summary>
    public class BrowserTab
    {
        /// <summary>
        /// Unique identifier for the tab.
        /// </summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>
        /// The title of the tab.
        /// </summary>
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// The UI Automation element representing the tab.
        /// </summary>
        public AutomationElement? AutomationElement { get; set; }

        /// <summary>
        /// The window handle for restoring the window before activating the tab.
        /// </summary>
        public IntPtr Hwnd { get; set; }

        /// <summary>
        /// Is the tab part of a minimized window.
        /// </summary>
        public bool IsMinimized { get; set; }

        /// <summary>
        /// The name of the browser process that owns this tab.
        /// </summary>
        public string BrowserName { get; set; } = string.Empty;

        /// <summary>
        /// Activates this browser tab.
        /// </summary>
        /// <returns>True if successful, false otherwise.</returns>
        public bool ActivateTab()
        {
            try
            {
                if (AutomationElement is null)
                {
                    Console.Error.WriteLine("Cannot activate tab because AutomationElement is null");
                    return false;
                }

                if (IsMinimized)
                    NativeMethods.ShowWindow(Hwnd, NativeMethods.SW_RESTORE);

                // Chromium
                if (AutomationElement.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var pattern) && pattern is SelectionItemPattern selectionPattern)
                {
                    selectionPattern.Select();
                    return true;
                }
                
                // Firefox
                if (AutomationElement.TryGetCurrentPattern(InvokePattern.Pattern, out var invokePattern) && invokePattern is InvokePattern invoke)
                {
                    invoke.Invoke();
                    return true;
                }
            }
            catch (ElementNotAvailableException ex)
            {
                Console.Error.WriteLine($"Element not available: {ex}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error activating tab: {ex}");
            }

            return false;
        }

        /// <summary>
        /// Closes this browser tab.
        /// </summary>
        /// <returns>True if successful, false otherwise.</returns>
        public bool CloseTab()
        {
            if (AutomationElement is null)
            {
                Console.Error.WriteLine("Cannot close tab because AutomationElement is null");
                return false;
            }

            try
            {
                // Special handling for closing Microsoft Edge tab- needs to be activated first
                if (BrowserName.Equals("msedge", StringComparison.OrdinalIgnoreCase))
                    ActivateTab();

                var closeButtonCondition = new OrCondition(
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                        new PropertyCondition(AutomationElement.NameProperty, "Close")
                    ),
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                        new PropertyCondition(AutomationElement.NameProperty, "Close tab", PropertyConditionFlags.IgnoreCase)
                    )
                );

                var closeButtons = AutomationElement.FindAll(TreeScope.Children, closeButtonCondition);
                if (closeButtons.Count > 0)
                {
                    var closeButton = closeButtons[0];
                    var invokePattern = closeButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    if (invokePattern != null)
                    {
                        invokePattern.Invoke();
                        return true;
                    }
                }
            }
            catch (ElementNotAvailableException ex)
            {
                Console.Error.WriteLine($"Element not available: {ex}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error closing tab: {ex}");
            }

            return false;
        }
    }
}
