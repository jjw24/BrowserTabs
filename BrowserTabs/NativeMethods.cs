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
        /// <param name="lpEnumFunc">A pointer to an application-defined callback function.</param>
        /// <param name="lParam">An application-defined value to be passed to the callback function.</param>
        /// <returns>True if successful; otherwise, false.</returns>
        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        /// <summary>
        /// Copies the text of the specified window's title bar into a buffer.
        /// </summary>
        /// <param name="hWnd">Handle to the window.</param>
        /// <param name="lpString">Pointer to the buffer that will receive the text.</param>
        /// <param name="nMaxCount">Maximum number of characters to copy.</param>
        /// <returns>The length of the copied string.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        /// <summary>
        /// Retrieves the length, in characters, of the specified window's title bar text.
        /// </summary>
        /// <param name="hWnd">Handle to the window.</param>
        /// <returns>The length of the window's title text.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        /// <summary>
        /// Retrieves the identifier of the thread and process that created the specified window.
        /// </summary>
        /// <param name="hWnd">Handle to the window.</param>
        /// <param name="lpdwProcessId">Receives the process identifier.</param>
        /// <returns>The identifier of the thread that created the window.</returns>
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        /// <summary>
        /// Contains information about the placement of a window on the screen.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public POINT ptMinPosition;
            public POINT ptMaxPosition;
            public RECT rcNormalPosition;
        }

        /// <summary>
        /// Defines the x- and y-coordinates of a point.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        /// <summary>
        /// Defines the coordinates of the upper-left and lower-right corners of a rectangle.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        /// <summary>
        /// Retrieves the show state and the restored, minimized, and maximized positions of the specified window.
        /// </summary>
        /// <param name="hWnd">Handle to the window.</param>
        /// <param name="lpwndpl">Pointer to the WINDOWPLACEMENT structure that receives the information.</param>
        /// <returns>True if successful; otherwise, false.</returns>
        [DllImport("user32.dll")]
        public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        /// <summary>
        /// Shows or hides a window.
        /// </summary>
        /// <param name="hWnd">Handle to the window.</param>
        /// <param name="nCmdShow">Specifies how the window is to be shown.</param>
        /// <returns>True if successful; otherwise, false.</returns>
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        /// <summary>
        /// Value for the showCmd member of WINDOWPLACEMENT indicating the window is minimized.
        /// </summary>
        public const int SW_SHOWMINIMIZED = 2;

        /// <summary>
        /// Value for the nCmdShow parameter of ShowWindow which activates and displays a window.
        /// </summary>
        public const int SW_RESTORE = 9;
    }
}
