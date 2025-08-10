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
        /// The URL or page title of the tab.
        /// </summary>
        public string Url { get; set; } = string.Empty;

        /// <summary>
        /// Indicates whether the tab is currently active.
        /// </summary>
        public bool IsActive { get; set; }

        /// <summary>
        /// The index of the tab within its window.
        /// </summary>
        public int TabIndex { get; set; }

        /// <summary>
        /// The UI Automation element representing the tab.
        /// </summary>
        public AutomationElement? AutomationElement { get; set; }

        /// <summary>
        /// The name of the browser process that owns this tab.
        /// </summary>
        public string BrowserName { get; set; } = string.Empty;
    }
}
