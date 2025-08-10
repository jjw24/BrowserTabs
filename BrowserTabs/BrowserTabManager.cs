using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Automation;
using System.Linq;

namespace BrowserTabs
{
    /// <summary>
    /// Retrieve, activate, and close tabs in Chromium-based browsers.
    /// </summary>
    public class BrowserTabManager
    {
        private static readonly string[] ChromiumProcessNames = new[]
        {
            "msedge", "chrome", "brave", "vivaldi", "opera", "chromium"
        };

        // List of known tab class names for Chromium browsers
        private static readonly string[] ChromiumTabClassNames = new[]
        {
            "EdgeTab", // Microsoft Edge
            "Tab"      // Chrome, Brave, Vivaldi, Opera, Chromium
        };

        /// <summary>
        /// Retrieves all open tabs from all Chromium-based browser windows.
        /// </summary>
        /// <returns>List of BrowserTab objects representing each open tab.</returns>
        public static List<BrowserTab> GetChromiumTabs()
        {
            var tabs = new List<BrowserTab>();

            try
            {
                var browserWindows = GetAllChromiumWindows();

                foreach (var (hwnd, processId) in browserWindows)
                {
                    try
                    {
                        var process = Process.GetProcessById(processId);
                        var mainWindow = AutomationElement.FromHandle(hwnd);
                        if (mainWindow != null)
                        {
                            var windowTabs = GetTabsFromWindow(mainWindow, process);
                            tabs.AddRange(windowTabs);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Error processing Chromium window: {ex}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error getting Chromium tabs: {ex}");
            }

            return tabs;
        }

        /// <summary>
        /// Retrieves all tabs from a specific browser window.
        /// </summary>
        /// <param name="mainWindow">AutomationElement representing the browser window.</param>
        /// <param name="process">Process object for the browser.</param>
        /// <returns>List of BrowserTab objects found in the window.</returns>
        private static List<BrowserTab> GetTabsFromWindow(AutomationElement mainWindow, Process process)
        {
            var tabs = new List<BrowserTab>();

            try
            {
                var tabCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);
                var tabElements = mainWindow.FindAll(TreeScope.Descendants, tabCondition);

                if (tabElements.Count > 0)
                {
                    for (int i = 0; i < tabElements.Count; i++)
                    {
                        try
                        {
                            var tabElement = tabElements[i];

                            // Only include tabs with known Chromium tab class names
                            if (!ChromiumTabClassNames.Contains(tabElement.Current.ClassName, StringComparer.Ordinal))
                                continue;

                            var tab = CreateTabFromElement(tabElement, process, i);
                            if (tab != null)
                            {
                                tabs.Add(tab);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine($"Error processing tab element: {ex}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error getting tabs from window: {ex}");
            }

            return tabs;
        }

        /// <summary>
        /// Creates a BrowserTab object from an AutomationElement representing a tab.
        /// </summary>
        /// <param name="tabElement">AutomationElement for the tab.</param>
        /// <param name="process">Process object for the browser.</param>
        /// <param name="index">Tab index.</param>
        /// <returns>BrowserTab object or null if invalid.</returns>
        private static BrowserTab? CreateTabFromElement(AutomationElement tabElement, Process process, int index)
        {
            try
            {
                var name = tabElement.Current.Name;
                if (string.IsNullOrEmpty(name) || name == "New Tab" || name.Contains("about:blank"))
                    return null;

                var isSelected = false;
                try
                {
                    var selectionPattern = tabElement.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                    isSelected = selectionPattern?.Current.IsSelected ?? false;
                }
                catch
                {
                }

                return new BrowserTab
                {
                    Id = $"{process.Id}_{index}",
                    Title = name,
                    Url = ExtractUrlFromTitle(name),
                    IsActive = isSelected,
                    TabIndex = index,
                    AutomationElement = tabElement,
                    BrowserName = process.ProcessName
                };
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error creating tab from element: {ex}");
                return null;
            }
        }

        /// <summary>
        /// Activates the specified browser tab.
        /// </summary>
        /// <param name="tab">BrowserTab to activate.</param>
        /// <returns>True if successful, false otherwise.</returns>
        public static bool ActivateTab(BrowserTab tab)
        {
            try
            {
                if (tab.AutomationElement is null)
                {
                    Console.Error.WriteLine("Cannot activate tab because AutomationElement is null");
                    return false;
                }

                var selectionPattern = tab.AutomationElement.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                if (selectionPattern != null)
                {
                    selectionPattern.Select();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error activating tab: {ex}");
            }

            return false;
        }

        /// <summary>
        /// Closes the specified browser tab.
        /// </summary>
        /// <param name="tab">BrowserTab to close.</param>
        /// <returns>True if successful, false otherwise.</returns>
        public static bool CloseTab(BrowserTab tab)
        {
            try
            {
                if (tab.AutomationElement != null)
                {
                    ActivateTab(tab);

                    var closeButtonCondition = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                        new PropertyCondition(AutomationElement.NameProperty, "Close")
                    );

                    var closeButtons = tab.AutomationElement.FindAll(TreeScope.Descendants, closeButtonCondition);
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
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error closing tab: {ex}");
            }

            return false;
        }

        /// <summary>
        /// Extracts the URL or page title from a browser tab's title string.
        /// </summary>
        /// <param name="title">Tab title string.</param>
        /// <returns>Extracted URL or page title.</returns>
        private static string ExtractUrlFromTitle(string title)
        {
            try
            {
                string[] suffixes = new[]
                {
                    " - Microsoft Edge",
                    " - Google Chrome",
                    " - Brave",
                    " - Vivaldi",
                    " - Opera",
                    " - Chromium"
                };

                foreach (var suffix in suffixes)
                {
                    if (title.Contains(suffix))
                        return title.Replace(suffix, "");
                }

                return title;
            }
            catch
            {
                return title;
            }
        }

        /// <summary>
        /// Finds all top-level Chromium browser windows on the system.
        /// </summary>
        /// <returns>List of tuples containing window handle and process ID.</returns>
        private static List<(IntPtr hwnd, int processId)> GetAllChromiumWindows()
        {
            var browserWindows = new List<(IntPtr, int)>();
            NativeMethods.EnumWindows((hwnd, lParam) =>
            {
                int length = NativeMethods.GetWindowTextLength(hwnd);
                if (length == 0) return true;

                var builder = new StringBuilder(length + 1);
                NativeMethods.GetWindowText(hwnd, builder, builder.Capacity);
                string title = builder.ToString();

                uint pid;
                NativeMethods.GetWindowThreadProcessId(hwnd, out pid);

                try
                {
                    var process = Process.GetProcessById((int)pid);
                    foreach (var processName in ChromiumProcessNames)
                    {
                        if (process.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase))
                        {
                            browserWindows.Add((hwnd, (int)pid));
                            break;
                        }
                    }
                }
                catch { }

                return true;
            }, IntPtr.Zero);

            return browserWindows;
        }
    }
}
