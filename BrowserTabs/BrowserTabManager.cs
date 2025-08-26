using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
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
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <returns>List of BrowserTab objects representing each open tab.</returns>
        public static List<BrowserTab> GetAllChromiumTabs(CancellationToken cancellationToken = default)
        {
            var tabBag = new ConcurrentBag<BrowserTab>();
            var tabBag2 = new ConcurrentBag<BrowserTab>();

            try
            {
                if (cancellationToken.IsCancellationRequested)
                    return new List<BrowserTab>();

                var browserWindows = GetAllChromiumWindows(cancellationToken);

                if (cancellationToken.IsCancellationRequested)
                    return new List<BrowserTab>();

                Parallel.ForEach(
                    browserWindows,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount, CancellationToken = cancellationToken },
                    window =>
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        try
                        {
                            var process = Process.GetProcessById(window.processId);
                            var mainWindow = AutomationElement.FromHandle(window.hwnd);
                            if (mainWindow is null)
                                return;

                            if (process.ProcessName.Equals("msedge", StringComparison.OrdinalIgnoreCase))
                            {
                                cancellationToken.ThrowIfCancellationRequested();

                                var edgeTabs = GetTabsFromWindowMinimized(mainWindow, process, window.hwnd, cancellationToken);
                                for (int i = 0; i < edgeTabs.Count; i++)
                                {
                                    cancellationToken.ThrowIfCancellationRequested();
                                    tabBag.Add(edgeTabs[i]);
                                }
                            }
                            else
                            {
                                cancellationToken.ThrowIfCancellationRequested();

                                var chromiumTabs = GetTabsFromWindow(mainWindow, process, cancellationToken);
                                for (int i = 0; i < chromiumTabs.Count; i++)
                                {
                                    cancellationToken.ThrowIfCancellationRequested();
                                    tabBag.Add(chromiumTabs[i]);
                                }
                            }


                            //var sw1 = Stopwatch.StartNew();
                            //var result1 = GetTabsFromWindowMinimized(mainWindow, process, window.hwnd, cancellationToken);
                            //sw1.Stop();
                            //Console.WriteLine($"GetTabsFromWindow took {sw1.ElapsedMilliseconds} ms");

                            //var sw2 = Stopwatch.StartNew();
                            //var result2 = GetTabsFromWindowMinimizedSolution2(mainWindow, process, window.hwnd, cancellationToken);
                            //sw2.Stop();
                            //Console.WriteLine($"GetTabsFromWindowMinimized took {sw2.ElapsedMilliseconds} ms");

                            //for (int i = 0; i < result1.Count; i++)
                            //{
                            //    cancellationToken.ThrowIfCancellationRequested();

                            //    tabBag.Add(result1[i]);
                            //}

                            //for (int i = 0; i < result2.Count; i++)
                            //{
                            //    cancellationToken.ThrowIfCancellationRequested();

                            //    tabBag2.Add(result2[i]);
                            //}

                            //if (result1.Count > 3 && result2.Count > 3)
                            //{

                            //}
                        }
                        catch (ArgumentException)
                        {
                            // Process might have exited, ignore
                        }
                    });
            }
            catch (OperationCanceledException)
            {
                return new List<BrowserTab>();
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
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <returns>List of BrowserTab objects found in the window.</returns>
        private static List<BrowserTab> GetTabsFromWindow(AutomationElement mainWindow, Process process, CancellationToken cancellationToken)
        {
            var tabs = new List<BrowserTab>();
            try
            {
                if (cancellationToken.IsCancellationRequested)
                    return new List<BrowserTab>();

                var tabCondition =
                    new OrCondition(
                        new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeTab"), // Microsoft Edge
                        new PropertyCondition(AutomationElement.ClassNameProperty, "Tab") // Chrome, Brave, Vivaldi, Opera, Chromium
                );

                AutomationElementCollection tabElements = mainWindow.FindAll(TreeScope.Descendants, tabCondition);

                int count = tabElements.Count;
                if (count > 0)
                {
                    if (cancellationToken.IsCancellationRequested)
                        return new List<BrowserTab>();

                    // Use partitioner for better parallelism
                    var rangePartitioner = Partitioner.Create(0, count);
                    var tabList = new BrowserTab[count];

                    Parallel.ForEach(rangePartitioner, new ParallelOptions { CancellationToken = cancellationToken }, range =>
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        for (int i = range.Item1; i < range.Item2; i++)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            var tabElement = tabElements[i];
                            var tab = CreateTabFromElement(tabElement, process, i);
                            if (tab != null)
                                tabList[i] = tab;
                        }
                    });

                    // Add non-null tabs to result
                    for (int i = 0; i < count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

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
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <returns>List of BrowserTab objects found in the minimized window.</returns>
        private static List<BrowserTab> GetTabsFromWindowMinimized(AutomationElement mainWindow, Process process, IntPtr hwnd, CancellationToken cancellationToken)
        {
            var tabBag = new ConcurrentBag<BrowserTab>();
            try
            {
                if (cancellationToken.IsCancellationRequested)
                    return new List<BrowserTab>();

                var elementConditions = new OrCondition(
                    new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeVerticalTabContainerView"),
                    new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeTabStripRegionView")
                );

                var matchedElements = mainWindow.FindAll(TreeScope.Descendants, elementConditions);

                if (matchedElements.Count == 0)
                    return new List<BrowserTab>();

                if (cancellationToken.IsCancellationRequested)
                    return new List<BrowserTab>();

                Parallel.ForEach(matchedElements.Cast<AutomationElement>(), new ParallelOptions { CancellationToken = cancellationToken }, element =>
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var stack = new Stack<AutomationElement>();
                    stack.Push(element);

                    while (stack.Count > 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        var current = stack.Pop();
                        // Only check children, not all descendants, for performance
                        var child = TreeWalker.RawViewWalker.GetFirstChild(current);
                        while (child != null)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

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

        private static List<BrowserTab> GetTabsFromWindowMinimizedSolution2(AutomationElement mainWindow, Process process, IntPtr hwnd, CancellationToken cancellationToken)
        {
            var tabBag = new ConcurrentBag<BrowserTab>();
            var seenRuntimeIds = new ConcurrentDictionary<string, byte>(StringComparer.Ordinal);
            static string GetRuntimeIdKey(AutomationElement el)
            {
                try
                {
                    var ids = el.GetRuntimeId();
                    if (ids != null && ids.Length > 0)
                        return string.Join("-", ids);
                }
                catch { }
                return string.Empty;
            }
            try
            {
                if (cancellationToken.IsCancellationRequested)
                    return new List<BrowserTab>();

                // Use per-container query first; fallback to raw only when needed
                var strictTabCondition = new AndCondition(
                    new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeTab"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem)
                );

                // Fallback: find candidate containers then query tabs within each
                var containerConditions = new OrCondition(
                    new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeVerticalTabContainerView"),
                    new PropertyCondition(AutomationElement.ClassNameProperty, "EdgeTabStripRegionView")
                );
                var matchedElements = mainWindow.FindAll(TreeScope.Descendants, containerConditions);
                if (matchedElements.Count == 0)
                    return new List<BrowserTab>();

                if (cancellationToken.IsCancellationRequested)
                    return new List<BrowserTab>();

                Parallel.ForEach(matchedElements.Cast<AutomationElement>(), new ParallelOptions { CancellationToken = cancellationToken }, element =>
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var cache = new CacheRequest();
                    cache.Add(AutomationElement.NameProperty);
                    cache.Add(AutomationElement.ClassNameProperty);
                    cache.Add(AutomationElement.ControlTypeProperty);
                    cache.TreeScope = TreeScope.Element | TreeScope.Children;
                    using (cache.Activate())
                    {
                        var tabElements = element.FindAll(TreeScope.Descendants, strictTabCondition);
                        if (tabElements.Count > 0)
                        {
                            int count = tabElements.Count;
                            for (int i = 0; i < count; i++)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                var tabElement = tabElements[i];
                                var key = GetRuntimeIdKey(tabElement);
                                if (!string.IsNullOrEmpty(key))
                                {
                                    if (!seenRuntimeIds.TryAdd(key, 1))
                                        continue;
                                }
                                var tab = CreateTabFromElement(tabElement, process, i, hwnd, isTabMinimized: true);
                                if (tab != null)
                                {
                                    tabBag.Add(tab);
                                }
                            }
                        }
                    }

                    // Shallow raw traversal (limited depth) to catch missed tabs without full deep walk
                    const int maxDepth = 3;
                    var nodeStack = new Stack<(AutomationElement node, int depth)>();
                    nodeStack.Push((element, 0));
                    while (nodeStack.Count > 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var (current, depth) = nodeStack.Pop();
                        if (depth >= maxDepth)
                            continue;
                        var child = TreeWalker.RawViewWalker.GetFirstChild(current);
                        while (child != null)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            try
                            {
                                if (child.Current.ClassName == "EdgeTab" && child.Current.ControlType == ControlType.TabItem)
                                {
                                    var key = GetRuntimeIdKey(child);
                                    if (string.IsNullOrEmpty(key) || seenRuntimeIds.TryAdd(key, 1))
                                    {
                                        var tab = CreateTabFromElement(child, process, 0, hwnd, isTabMinimized: true);
                                        if (tab != null)
                                            tabBag.Add(tab);
                                    }
                                }
                            }
                            catch { }
                            nodeStack.Push((child, depth + 1));
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
        /// Finds all top-level Chromium browser windows on the system.
        /// </summary>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <returns>List of tuples containing window handle and process ID.</returns>
        private static List<(IntPtr hwnd, int processId)> GetAllChromiumWindows(CancellationToken cancellationToken)
        {
            var browserWindows = new ConcurrentBag<(IntPtr, int)>();
            var windowHandles = new List<(IntPtr hwnd, uint pid)>();

            if (cancellationToken.IsCancellationRequested)
                return new List<(IntPtr, int)>();

            Task.Run(() =>
            {
                NativeMethods.EnumWindows((hwnd, lParam) =>
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    uint pid;
                    NativeMethods.GetWindowThreadProcessId(hwnd, out pid);
                    windowHandles.Add((hwnd, pid));

                    return true;

                }, IntPtr.Zero);
            }, cancellationToken).Wait(cancellationToken);

            if (cancellationToken.IsCancellationRequested)
                return new List<(IntPtr, int)>();

            Parallel.ForEach(
                windowHandles,
                new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount, CancellationToken = cancellationToken },
                window =>
                {
                    cancellationToken.ThrowIfCancellationRequested();

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
