using System;
using System.Collections.Generic;
using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;

namespace WebView2PowerPointAddInSample
{
    public class SidePanel
    {
        protected const int DefaultWidth = 420;

        private readonly Application _application;
        private readonly CustomTaskPaneCollection _customTaskPaneCollection;
        private readonly Dictionary<int, TaskPaneItems> _customTaskPaneItems = new Dictionary<int, TaskPaneItems>();

        public SidePanel(Application application, CustomTaskPaneCollection customTaskPaneCollection)
        {
            _application = application;
            _customTaskPaneCollection = customTaskPaneCollection;
        }

        public void CreateWebAppTaskPane(object window)
        {
            var webAppContentControl = CreateReservedInstance();

            var documentWindow = window as DocumentWindow;
            var customTaskPane = _customTaskPaneCollection.Add(webAppContentControl, "taskPaneTitle", documentWindow);

            var autoScaleFactor = GetAutoScaleFactor();
            customTaskPane.Width = Math.Max(DefaultWidth, (int)Math.Floor(DefaultWidth * autoScaleFactor));
            customTaskPane.Visible = true;

            var taskPaneItems = new TaskPaneItems
                { TaskPaneUserControl = webAppContentControl, TaskPane = customTaskPane };
            var windowHandle = GetWindowHandle(documentWindow);
            _customTaskPaneItems.Add(windowHandle, taskPaneItems);

            CreateReservedInstance();
        }

        private int GetWindowHandle(DocumentWindow documentWindow)
        {
            if (documentWindow == null) return _application.HWND;

            return documentWindow.HWND;
        }

        protected WebAppContentControl CreateReservedInstance()
        {
            var defaultUrl = "https://teams.microsoft.com";
            var webAppContentControl = new WebAppContentControl(defaultUrl);

            // Set initial size for browser control to load frontend properly. Since it's docked in side pane, this numbers are not important for VSTO.
            var autoScaleFactor = GetAutoScaleFactor();
            webAppContentControl.Height = (int)(_application.Height * 1.5);
            webAppContentControl.Width = Math.Max(DefaultWidth, (int)Math.Floor(DefaultWidth * autoScaleFactor));
            return webAppContentControl;
        }

        public static double GetAutoScaleFactor()
        {
            return Graphics.FromHwnd(IntPtr.Zero).DpiX / 96.0;
        }
    }

    public class TaskPaneItems
    {
        public WebAppContentControl TaskPaneUserControl { get; set; }

        public CustomTaskPane TaskPane { get; set; }
    }
}