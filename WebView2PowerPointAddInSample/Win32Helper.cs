using System;
using System.Runtime.InteropServices;
using System.Text;

public class Win32Helper
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, ref SearchData data);

    public const string WordWindowClassName = "OpusApp";
    public const string ExcelWindowClassName = "XLMAIN";
    public const string PowerPointWindowClassName = "PPTFrameClass";

    //http://www.pinvoke.net/default.aspx/user32/EnumWindows.html
    public static IntPtr SearchForWindow(string wndclass, string title)
    {
        var searchData = new SearchData(wndclass, title);
        Win32Methods.EnumWindows(EnumProc, ref searchData);
        return searchData.HWnd;
    }

    [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
    public static extern IntPtr GetParent(IntPtr hWnd);

    public static bool EnumProc(IntPtr hWnd, ref SearchData data)
    {
        var stringBuilder = new StringBuilder(1024);
        Win32Methods.GetClassName(hWnd, stringBuilder, stringBuilder.Capacity);
        if (!stringBuilder.ToString().Equals(data.Wndclass, StringComparison.InvariantCultureIgnoreCase)) return true;

        var title = GetText(hWnd);
        if (title?.Contains(data.Title) == true)
        {
            data.HWnd = hWnd;
            return false;
        }

        return true;
    }

    public static string GetText(IntPtr hWnd)
    {
        // Allocate correct string length first
        var length = Win32Methods.GetWindowTextLength(hWnd);

        var stringBuilder = new StringBuilder(length + 1);
        Win32Methods.GetWindowText(hWnd, stringBuilder, stringBuilder.Capacity);
        return stringBuilder.ToString();
    }


    private static IntPtr GetActiveWindowHandler(dynamic window, string className)
    {
        var windowSignature = Guid.NewGuid().ToString();
        if (window == null) return IntPtr.Zero;

        string oldCaption = window.Caption;

        try
        {
            window.Caption = windowSignature;
            return SearchForWindow(className, windowSignature);
        }
        finally
        {
            window.Caption = oldCaption;
        }
    }

    public static bool SetOwner(HandleRef ownerHandler, HandleRef owneeHandler, bool owneeVisible,
        bool owneeContainsFocus)
    {
        if (ownerHandler.Handle == owneeHandler.Handle) return false;

        // Refresh window
        var flags = Win32Methods.SetWindowPosFlags.SWP_NOSIZE | Win32Methods.SetWindowPosFlags.SWP_NOMOVE |
                    Win32Methods.SetWindowPosFlags.SWP_FRAMECHANGED;
        if (owneeVisible) flags |= Win32Methods.SetWindowPosFlags.SWP_SHOWWINDOW;

        if (!owneeContainsFocus) flags |= Win32Methods.SetWindowPosFlags.SWP_NOACTIVATE;

        return true;
    }

    public class SearchData
    {
        public readonly string Title;
        public readonly string Wndclass;

        public IntPtr HWnd;

        public SearchData(string wndClass, string title)
        {
            Wndclass = wndClass;
            Title = title;
        }
    }
}