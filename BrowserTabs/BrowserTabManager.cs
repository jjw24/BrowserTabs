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
    }
}
