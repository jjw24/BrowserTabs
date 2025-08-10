using System;
using System.Runtime.InteropServices;
using System.Text;

namespace BrowserTabs
{
    /// <summary>
    /// Provides P/Invoke signatures for native Windows API methods used for window enumeration and information.
    /// </summary>
    internal static class NativeMethods
    {
        /// <summary>
        /// Delegate for the EnumWindows callback function.
        /// </summary>
        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        /// <summary>
        /// Enumerates all top-level windows on the screen by passing the handle to each window to an application-defined callback function.
        /// </summary>
        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        /// <summary>
        /// Copies the text of the specified window's title bar into a buffer.
        /// </summary>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        /// <summary>
        /// Retrieves the length, in characters, of the specified window's title bar text.
        /// </summary>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        /// <summary>
        /// Retrieves the identifier of the thread and process that created the specified window.
        /// </summary>
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    }
}
