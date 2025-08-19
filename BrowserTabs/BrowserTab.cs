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
    }
}
