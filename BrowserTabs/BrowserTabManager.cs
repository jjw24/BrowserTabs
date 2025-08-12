using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Automation;
using System.Linq;
using System.Threading.Tasks;

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
                // Use HashSet for O(1) class name checks
                var classNames = new HashSet<string>(ChromiumTabClassNames, StringComparer.Ordinal);

                // Combine conditions: TabItem + known class names
                var tabCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);
                var tabElements = mainWindow.FindAll(TreeScope.Descendants, tabCondition);

                if (tabElements.Count > 0)
                {
                    // Use parallel processing for tab creation
                    var tabList = new BrowserTab?[tabElements.Count];
                    Parallel.For(0, tabElements.Count, i =>
                    {
                        var tabElement = tabElements[i];
                        if (!classNames.Contains(tabElement.Current.ClassName))
                            return;

                        var tab = CreateTabFromElement(tabElement, process, i);
                        if (tab != null)
                        {
                            tabList[i] = tab;
                        }
                    });

                    // Add non-null tabs to result
                    foreach (var tab in tabList)
                    {
                        if (tab != null)
                            tabs.Add(tab);
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

                    // Search for the close button as a child of the tab element
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

                    // Use TreeScope.Children to only look at direct children
                    var closeButtons = tab.AutomationElement.FindAll(TreeScope.Children, closeButtonCondition);
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
            var chromiumNames = new HashSet<string>(ChromiumProcessNames, StringComparer.OrdinalIgnoreCase);
            var windowHandles = new List<(IntPtr hwnd, uint pid)>();

            // First, collect all window handles and process IDs
            NativeMethods.EnumWindows((hwnd, lParam) =>
            {
                uint pid;
                NativeMethods.GetWindowThreadProcessId(hwnd, out pid);
                windowHandles.Add((hwnd, pid));
                return true;
            }, IntPtr.Zero);

            // Now process in parallel
            Parallel.ForEach(windowHandles, window =>
            {
                try
                {
                    var process = Process.GetProcessById((int)window.pid);
                    if (chromiumNames.Contains(process.ProcessName))
                    {
                        int length = NativeMethods.GetWindowTextLength(window.hwnd);
                        if (length > 0)
                        {
                            lock (browserWindows)
                            {
                                browserWindows.Add((window.hwnd, (int)window.pid));
                            }
                        }
                    }
                }
                catch (ArgumentException)
                {
                    // Process might have exited, ignore
                }
            });

            return browserWindows;
        }
    }
}
