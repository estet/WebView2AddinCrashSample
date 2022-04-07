using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;

namespace WebView2PowerPointAddInSample
{
    public partial class Dialog : Form
    {
        private readonly HandleRef _ownerHandler;

        public Dialog()
        {
            InitializeComponent();

            CustomInitialize();
        }

        public Dialog(dynamic ownerWindow)
        {
            IntPtr handler = GetHandler(ownerWindow);
            OwnerWindow = ownerWindow;
            if (handler == IntPtr.Zero)
            {
                var errorMessage = $"Failed to locate parent handler - {Marshal.GetLastWin32Error()}";
                throw new InvalidOperationException(errorMessage);
            }

            _ownerHandler = new HandleRef(ownerWindow, handler);

            InitializeComponent();

            CustomInitialize();

            Closing += OnClosing;
        }

        public dynamic OwnerWindow { get; }

        private IntPtr GetHandler(DocumentWindow ownerWindow)
        {
            return Win32Helper.GetParent(Win32Helper.GetParent(new IntPtr(ownerWindow.HWND)));
        }

        private void CustomInitialize()
        {
            ContainerPanel.Controls.Clear();
            ToolStrip.Renderer = new CustomToolStripRenderer();
            Name = "TemplafyDialog";
        }

        private void OnClosing(object sender, CancelEventArgs cancelEventArgs)
        {
            Win32Methods.EnableWindow(_ownerHandler, true);
        }

        public void Show(bool showAsModal)
        {
            if (showAsModal)
            {
                ShowDialog();
            }
            else
            {
                // We can't use Show(owner) because the the activewindow is not a Control object
                Show();
                if (Win32Helper.SetOwner(_ownerHandler, new HandleRef(this, Handle), Visible, ContainsFocus))
                    Win32Methods.EnableWindow(_ownerHandler, false);

                CenterToParent();
            }
        }

        public void Close(DialogResult dialogResult)
        {
            DialogResult = dialogResult;
            Close();
        }

        private void OnCloseButtonClick(object sender, EventArgs e)
        {
            Close();
        }

        private void OnMaximizeButtonClick(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Normal)
            {
                MaximizedBounds = Screen.FromHandle(Handle).WorkingArea;
                WindowState = FormWindowState.Maximized;
            }
            else
            {
                WindowState = FormWindowState.Normal;
            }
        }

        private void ToolStrip_MouseDown(object sender, MouseEventArgs e)
        {
            // Implementing this using MouseMove and MouseUp makes it unreliable; This will ensure a nice experience when moving the window
            OnMouseDown(e);

            if (e.Button != MouseButtons.Left) return;

            Capture = false;

            const int nonClientLeftButtonDown = 0XA1;

            var message = Message.Create(Handle, nonClientLeftButtonDown, new IntPtr(2), IntPtr.Zero);
            WndProc(ref message);
        }

        public class CustomToolStripRenderer : ToolStripSystemRenderer
        {
            protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
            {
            }
        }
    }
}