using System;
using System.Runtime.InteropServices;
using System.Text;

public class Win32Methods
{
    public enum WindowLongFlags
    {
        // ReSharper disable InconsistentNaming

        GWL_EXSTYLE = -20,
        GWLP_HINSTANCE = -6,
        GWLP_HWNDPARENT = -8,
        GWL_ID = -12,
        GWL_STYLE = -16,
        GWL_USERDATA = -21,
        GWL_WNDPROC = -4,
        DWLP_USER = 0x8,
        DWLP_MSGRESULT = 0x0,
        DWLP_DLGPROC = 0x4
    }

    [Flags]
    public enum Modifiers
    {
        None = 0,
        LShift = 1 << 0,
        RShift = 1 << 1,
        LControl = 1 << 2,
        RControl = 1 << 3,
        LAlt = 1 << 4,
        RAlt = 1 << 5,
        LWin = 1 << 6,
        RWin = 1 << 7,

        Shift = LShift | RShift,
        Control = LControl | RControl,
        Alt = LAlt | RAlt,
        Win = LWin | RWin,

        ControlShift = Shift | Control,
        AltShift = Shift | Alt,
        ControlAlt = Control | Alt,
        ControlAltShift = Shift | Control | Alt
    }

    // lParam bit 30: The previous key state. The value is 1 if the key
    // is down before the message is sent; it is 0 if the key is up.
    public const int HOOK_PREVIOUS_STATE = 1 << 30;

    // lParam bit 31: The transition state. The value is 0 if the key is
    // being pressed and 1 if it is being released.
    public const int HOOK_TRANSITION_STATE = 1 << 31;

    public const int VK_SHIFT = 0x10;
    public const int VK_CONTROL = 0x11;
    public const int VK_ALT = 0x12;

    [Flags]
    public enum SetWindowPosFlags : uint
    {
        // ReSharper disable InconsistentNaming

        /// <summary>
        ///     If the calling thread and the thread that owns the window are attached to different input queues, the system posts
        ///     the request to the thread that owns the window. This prevents the calling thread from blocking its execution while
        ///     other threads process the request.
        /// </summary>
        SWP_ASYNCWINDOWPOS = 0x4000,

        /// <summary>
        ///     Prevents generation of the WM_SYNCPAINT message.
        /// </summary>
        SWP_DEFERERASE = 0x2000,

        /// <summary>
        ///     Draws a frame (defined in the window's class description) around the window.
        /// </summary>
        SWP_DRAWFRAME = 0x0020,

        /// <summary>
        ///     Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to the window, even if
        ///     the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE is sent only when the window's
        ///     size is being changed.
        /// </summary>
        SWP_FRAMECHANGED = 0x0020,

        /// <summary>
        ///     Hides the window.
        /// </summary>
        SWP_HIDEWINDOW = 0x0080,

        /// <summary>
        ///     Does not activate the window. If this flag is not set, the window is activated and moved to the top of either the
        ///     topmost or non-topmost group (depending on the setting of the hWndInsertAfter parameter).
        /// </summary>
        SWP_NOACTIVATE = 0x0010,

        /// <summary>
        ///     Discards the entire contents of the client area. If this flag is not specified, the valid contents of the client
        ///     area are saved and copied back into the client area after the window is sized or repositioned.
        /// </summary>
        SWP_NOCOPYBITS = 0x0100,

        /// <summary>
        ///     Retains the current position (ignores X and Y parameters).
        /// </summary>
        SWP_NOMOVE = 0x0002,

        /// <summary>
        ///     Does not change the owner window's position in the Z order.
        /// </summary>
        SWP_NOOWNERZORDER = 0x0200,

        /// <summary>
        ///     Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to the client area,
        ///     the nonclient area (including the title bar and scroll bars), and any part of the parent window uncovered as a
        ///     result of the window being moved. When this flag is set, the application must explicitly invalidate or redraw any
        ///     parts of the window and parent window that need redrawing.
        /// </summary>
        SWP_NOREDRAW = 0x0008,

        /// <summary>
        ///     Same as the SWP_NOOWNERZORDER flag.
        /// </summary>
        SWP_NOREPOSITION = 0x0200,

        /// <summary>
        ///     Prevents the window from receiving the WM_WINDOWPOSCHANGING message.
        /// </summary>
        SWP_NOSENDCHANGING = 0x0400,

        /// <summary>
        ///     Retains the current size (ignores the cx and cy parameters).
        /// </summary>
        SWP_NOSIZE = 0x0001,

        /// <summary>
        ///     Retains the current Z order (ignores the hWndInsertAfter parameter).
        /// </summary>
        SWP_NOZORDER = 0x0004,

        /// <summary>
        ///     Displays the window.
        /// </summary>
        SWP_SHOWWINDOW = 0x0040

        // ReSharper restore InconsistentNaming
    }

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", EntryPoint = "SetWindowLong", CharSet = CharSet.Auto)]
    private static extern IntPtr SetWindowLongPtr32(HandleRef hWnd, int nIndex, HandleRef dwNewLong);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", CharSet = CharSet.Auto)]
    private static extern IntPtr SetWindowLongPtr64(HandleRef hWnd, int nIndex, HandleRef dwNewLong);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(Win32Helper.EnumWindowsProc lpEnumFunc, ref Win32Helper.SearchData data);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetWindowPos(HandleRef hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy,
        SetWindowPosFlags uFlags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnableWindow(HandleRef hWnd, bool bEnable);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
    public static extern IntPtr GetParent(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(HandleRef hWnd, uint Msg, HandleRef wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    public delegate IntPtr HookProc(int nCode, int wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall,
        SetLastError = true)]
    public static extern IntPtr SetWindowsHookExW(int idHook, HookProc lpfn, IntPtr hmod, int dwThreadId);

    [DllImport("user32.dll")]
    public static extern IntPtr CallNextHookEx(IntPtr hHook, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall,
        SetLastError = true)]
    public static extern bool UnhookWindowsHookEx(IntPtr hHook);

    [DllImport("Kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetModuleHandle(IntPtr lpModuleName);

    [DllImport("Kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetCurrentThreadId();

    [DllImport("user32")]
    public static extern int ToAscii(int uVirtKey, int uScanCode, byte[] lpbKeyState, byte[] lpwTransKey, int fuState);

    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern short GetKeyState(int vKey);

    [DllImport("user32")]
    public static extern int GetKeyboardState(byte[] pbKeyState);

    // This helper static method is required because the 32-bit version of user32.dll does not contain this API
    // (on any versions of Windows), so linking the method will fail at run-time. The bridge dispatches the request
    // to the correct function (GetWindowLong in 32-bit mode and GetWindowLongPtr in 64-bit mode)
    public static IntPtr SetWindowLong(HandleRef hWnd, int nIndex, HandleRef dwNewLong)
    {
        return IntPtr.Size == 4
            ? SetWindowLongPtr32(hWnd, nIndex, dwNewLong)
            : SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
    }

    /// <summary>
    ///     HWND Constants
    /// </summary>
    public static class Hwnd
    {
        public static readonly IntPtr NoTopMost = new IntPtr(-2),
            TopMost = new IntPtr(-1),
            Top = new IntPtr(0),
            Bottom = new IntPtr(1);
    }
}