using System;
using Office = Microsoft.Office.Core;

namespace WebView2PowerPointAddInSample
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var webAppContentControl = new WebAppContentControl("", @"
<html><head></head><body><button onclick='onClick()'>Crash me!</button><script type='text/javascript'>async function onClick() {
                const bridge = chrome.webview.hostObjects.bridge;
                bridge.showModalDialog();
            }</script></body></html>");
            var myCustomTaskPane = CustomTaskPanes.Add(webAppContentControl,
                "New Task Pane");

            myCustomTaskPane.DockPosition =
                Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
            myCustomTaskPane.Height = 500;
            myCustomTaskPane.Width = 500;

            myCustomTaskPane.DockPosition =
                Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            myCustomTaskPane.Width = 300;

            myCustomTaskPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}