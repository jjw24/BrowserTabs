using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Automation;
using System.Threading.Tasks;
using System.Collections.Concurrent;

namespace BrowserTabs
{
    /// <summary>
    /// Retrieve, activate, and close tabs in Chromium-based browsers.
    /// </summary>
    public class BrowserTabManager
    {
        private static readonly HashSet<string> ChromiumProcessNames = new(
            new[] { "msedge", "chrome", "brave", "vivaldi", "opera", "chromium" },
            StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Retrieves all open tabs from all Chromium-based browser windows.
        /// </summary>
        /// <returns>List of BrowserTab objects representing each open tab.</returns>
        public static List<BrowserTab> GetChromiumTabs()
        {
            var tabBag = new ConcurrentBag<BrowserTab>();
            try
            {
                var browserWindows = GetAllChromiumWindows();

                Parallel.ForEach(
                    browserWindows,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    window =>
                {
                    try
                    {
                        var process = Process.GetProcessById(window.processId);
                        var mainWindow = AutomationElement.FromHandle(window.hwnd);
                        if (mainWindow != null)
                        {
                            foreach (var tab in GetTabsFromWindow(mainWindow, process))
                                tabBag.Add(tab);
                        }
                    }
                    catch (ArgumentException)
                    {
                        // Process might have exited, ignore
                    }
                });
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error getting Chromium tabs: {ex}");
            }
            return new List<BrowserTab>(tabBag);
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
                // Use AndCondition to filter by both ControlType and ClassName
                var tabCondition = new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                    new OrCondition(
                        new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeTab"), // Microsoft Edge
                        new PropertyCondition(AutomationElement.ClassNameProperty, "Tab") // Chrome, Brave, Vivaldi, Opera, Chromium
                    )
                );

                // Restrict to TreeScope.Children if possible for speed
                AutomationElementCollection tabElements = mainWindow.FindAll(TreeScope.Descendants, tabCondition);

                int count = tabElements.Count;
                if (count > 0)
                {
                    // Use partitioner for better parallelism
                    var rangePartitioner = Partitioner.Create(0, count);
                    var tabList = new BrowserTab[count];

                    Parallel.ForEach(rangePartitioner, range =>
                    {
                        for (int i = range.Item1; i < range.Item2; i++)
                        {
                            var tabElement = tabElements[i];
                            var tab = CreateTabFromElement(tabElement, process, i);
                            if (tab != null)
                                tabList[i] = tab;
                        }
                    });

                    // Add non-null tabs to result
                    for (int i = 0; i < count; i++)
                    {
                        if (tabList[i] != null)
                            tabs.Add(tabList[i]);
                    }
                }
            }
            catch (ElementNotAvailableException ex)
            {
                Console.Error.WriteLine($"Element not available: {ex}");
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
                var name = tabElement.GetCurrentPropertyValue(AutomationElement.NameProperty) as string;
                if (string.IsNullOrEmpty(name) || name == "New Tab" || name.Contains("about:blank"))
                    return null;

                var isSelected = false;
                try
                {
                    var selectionPattern = tabElement.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                    isSelected = selectionPattern?.Current.IsSelected ?? false;
                }
                catch (InvalidOperationException)
                {
                    // Pattern not available
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
            catch (ElementNotAvailableException ex)
            {
                Console.Error.WriteLine($"Element not available: {ex}");
                return null;
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
                    int idx = title.IndexOf(suffix, StringComparison.Ordinal);
                    if (idx >= 0)
                        return title.Substring(0, idx);
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
            var browserWindows = new ConcurrentBag<(IntPtr, int)>();
            var windowHandles = new List<(IntPtr hwnd, uint pid)>();

            // Collect all window handles and process IDs in a background thread
            Task.Run(() =>
            {
                NativeMethods.EnumWindows((hwnd, lParam) =>
                {
                    uint pid;
                    NativeMethods.GetWindowThreadProcessId(hwnd, out pid);
                    windowHandles.Add((hwnd, pid));
                    return true;
                }, IntPtr.Zero);
            }).Wait();

            Parallel.ForEach(
                windowHandles,
                new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                window =>
            {
                try
                {
                    var process = Process.GetProcessById((int)window.pid);
                    if (ChromiumProcessNames.Contains(process.ProcessName))
                    {
                        int length = NativeMethods.GetWindowTextLength(window.hwnd);
                        if (length > 0)
                        {
                            browserWindows.Add((window.hwnd, (int)window.pid));
                        }
                    }
                }
                catch (ArgumentException)
                {
                    // Process might have exited, ignore
                }
            });

            return new List<(IntPtr, int)>(browserWindows);
        }
    }
}
