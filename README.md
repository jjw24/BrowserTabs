# BrowserTabs

BrowserTabs is a .NET library for retrieving all open tabs from Chromium-based browsers (such as Microsoft Edge, Google Chrome, Brave, Vivaldi, Opera, and Chromium) and Firefox-based browsers on Windows.  

## Installation
You can install BrowserTabs via NuGet Package Manager, or

Powershell:
```bash
Install-Package BrowserTabs
```

.Net CLI:
```bash
dotnet add package BrowserTabs
```

## Features

- Enumerate all open tabs across multiple Chromium-based and Firefox-based browsers.
- Retrieve tab title, URL, active status, tab index, and browser name.
- Activate (bring to foreground) or close browser tabs programmatically.
- Supports multi-window scenarios (pop-outs, incognito, etc.).

## Usage

1. Reference the `BrowserTabs` library in your .NET project.
2. Call `BrowserTabManager.GetChromiumTabs()` or `BrowserTabManager.GetFirefoxTabs()` to get a list of all open tabs.
3. Use `BrowserTabManager.ActivateTab(tab)` or `BrowserTabManager.CloseTab(tab)` to interact with tabs.

````````

# TODO:
1. Handle grouped tabs
2. Add proper logging
