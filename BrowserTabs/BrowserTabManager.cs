using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Runtime.InteropServices;

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
        /// Retrieves all open tabs from all Chromium-based browser windows,
        /// using a unified logic to avoid duplicate or separate calls.
        /// </summary>
        /// <returns>List of BrowserTab objects representing each open tab.</returns>
        public static List<BrowserTab> GetAllChromiumTabs()
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
                            if (mainWindow is null)
                                return;

                            var tabs = !IsWindowMinimized(window.hwnd)
                                ? GetTabsFromWindow(mainWindow, process)
                                : GetTabsFromWindowMinimized(mainWindow, process, window.hwnd);

                            for (int i = 0; i < tabs.Count; i++)
                                tabBag.Add(tabs[i]);
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
                var tabCondition =
                    new OrCondition(
                        new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeTab"), // Microsoft Edge
                        new PropertyCondition(AutomationElement.ClassNameProperty, "Tab") // Chrome, Brave, Vivaldi, Opera, Chromium
                );

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
        /// Retrieves all tabs from a minimized browser window, specifically handling Edge's minimized tabs.
        /// This is currently unable to retrieve tabs from minimized windows in Chrome, and possibly other browsers like Brave, etc.
        /// </summary>
        /// <param name="mainWindow">AutomationElement representing the browser window.</param>
        /// <param name="process">Process object for the browser.</param>
        /// <param name="hwnd">The window handle for restoring the window before activating the tab.</param>
        /// <returns>List of BrowserTab objects found in the minimized window.</returns>
        private static List<BrowserTab> GetTabsFromWindowMinimized(AutomationElement mainWindow, Process process, IntPtr hwnd)
        {
            var tabBag = new ConcurrentBag<BrowserTab>();
            try
            {
                var elementConditions = new OrCondition(
                    new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeVerticalTabContainerView"),
                    new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeTabStripRegionView")
                );

                var matchedElements = mainWindow.FindAll(TreeScope.Descendants, elementConditions);
                
                if (matchedElements.Count == 0)
                    return new List<BrowserTab>();

                Parallel.ForEach(matchedElements.Cast<AutomationElement>(), element =>
                {
                    var stack = new Stack<AutomationElement>();
                    stack.Push(element);

                    while (stack.Count > 0)
                    {
                        var current = stack.Pop();
                        // Only check children, not all descendants, for performance
                        var child = TreeWalker.RawViewWalker.GetFirstChild(current);
                        while (child != null)
                        {
                            // Check if an Edge tab item
                            if (child.Current.ClassName == "EdgeTab" &&
                                child.Current.ControlType == ControlType.TabItem)
                            {
                                var tab = CreateTabFromElement(child, process, 0, hwnd, isTabMinimized: true);
                                if (tab != null)
                                    tabBag.Add(tab);
                            }
                            stack.Push(child);
                            child = TreeWalker.RawViewWalker.GetNextSibling(child);
                        }
                    }
                });
            }
            catch (ElementNotAvailableException ex)
            {
                Console.Error.WriteLine($"Element not available: {ex}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error getting tabs from minimized window: {ex}");
            }

            return new List<BrowserTab>(tabBag);
        }

        /// <summary>
        /// Creates a BrowserTab object from an AutomationElement representing a tab.
        /// </summary>
        /// <param name="tabElement">AutomationElement for the tab.</param>
        /// <param name="process">Process object for the browser.</param>
        /// <param name="index">Tab index.</param>
        /// <param name="hwnd">The window handle for restoring the window before activating the tab.</param>
        /// <param name="isTabMinimized">Is the tab part of a minimized window</param>
        /// <returns>BrowserTab object or null if invalid.</returns>
        private static BrowserTab? CreateTabFromElement(AutomationElement tabElement, Process process, int index, IntPtr hwnd = default, bool isTabMinimized = false)
        {
            try
            {
                var name = tabElement.GetCurrentPropertyValue(AutomationElement.NameProperty) as string;
                if (string.IsNullOrEmpty(name) || name == "New Tab" || name.Contains("about:blank"))
                    return null;

                return new BrowserTab
                {
                    Id = $"{process.Id}_{index}",
                    Title = name,
                    IsMinimized = isTabMinimized,
                    AutomationElement = tabElement,
                    Hwnd = hwnd,
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

                if (tab.IsMinimized)
                    NativeMethods.ShowWindow(tab.Hwnd, NativeMethods.SW_RESTORE);

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
            if (tab.AutomationElement is null)
            {
                Console.Error.WriteLine("Cannot close tab because AutomationElement is null");
                return false;
            }

            try
            {
                // Window needs to be restored before able to find the close button element
                if (tab.IsMinimized)
                    NativeMethods.ShowWindow(tab.Hwnd, NativeMethods.SW_RESTORE);

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
        /// Finds all top-level Chromium browser windows on the system.
        /// </summary>
        /// <returns>List of tuples containing window handle and process ID.</returns>
        private static List<(IntPtr hwnd, int processId)> GetAllChromiumWindows()
        {
            var browserWindows = new ConcurrentBag<(IntPtr, int)>();
            var windowHandles = new List<(IntPtr hwnd, uint pid)>();

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

        /// <summary>
        /// Determines whether the specified window is minimized.
        /// </summary>
        /// <param name="hwnd">Window handle.</param>
        /// <returns>True if the window is minimized; otherwise, false.</returns>
        private static bool IsWindowMinimized(IntPtr hwnd)
        {
            var placement = new NativeMethods.WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf(typeof(NativeMethods.WINDOWPLACEMENT));
            if (NativeMethods.GetWindowPlacement(hwnd, ref placement))
                return placement.showCmd == NativeMethods.SW_SHOWMINIMIZED;
            return false;
        }
    }
}
