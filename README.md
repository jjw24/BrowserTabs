# BrowserTabs

BrowserTabs is a .NET library for retrieving all open tabs from Chromium-based browsers (such as Microsoft Edge, Google Chrome, Brave, Vivaldi, Opera, and Chromium) on Windows.  

## Features

- Enumerate all open tabs across multiple Chromium-based browsers.
- Retrieve tab title, URL, active status, tab index, and browser name.
- Activate (bring to foreground) or close browser tabs programmatically.
- Supports multi-window scenarios (pop-outs, incognito, etc.).

## Usage

1. Reference the `BrowserTabs` library in your .NET project.
2. Call `BrowserTabManager.GetChromiumTabs()` to get a list of all open tabs.
3. Use `BrowserTabManager.ActivateTab(tab)` or `BrowserTabManager.CloseTab(tab)` to interact with tabs.

````````

# TODO:
1. Handle grouped tabs
2. Fix tab URL extraction logic
3. Optimize tab retrieval logic for better performance
