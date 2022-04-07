using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;

namespace WebView2PowerPointAddInSample
{
    public partial class WebAppContentControl : UserControl
    {
        private readonly string _html;
        private readonly Uri _webAppUri;

        public WebAppContentControl(string url, string html = null)
        {
            _html = html;
            if (!string.IsNullOrEmpty(url)) _webAppUri = new Uri(url);

            InitializeComponent();
            HandleCreated += WebAppContentControl_HandleCreated;
            _customWebBrowserControl.CoreWebView2InitializationCompleted +=
                OnCustomWebBrowserControlOnCoreWebView2InitializationCompleted;
        }

        private async void WebAppContentControl_HandleCreated(object sender, EventArgs e)
        {
            var customWebBrowserUserDataFolder = Path.Combine(Path.GetTempPath(), "OfficeAddins", "Test", "WebView2");
            await _customWebBrowserControl.Initialize(customWebBrowserUserDataFolder);
        }

        private void OnCustomWebBrowserControlOnCoreWebView2InitializationCompleted(object sender,
            CoreWebView2InitializationCompletedEventArgs args)
        {
            if (!args.IsSuccess) return;
            _customWebBrowserControl.CoreWebView2.AddHostObjectToScript("bridge", new Bridge());

            SetWebAppUrl();
        }

        public void SetWebAppUrl()
        {
            if (!string.IsNullOrWhiteSpace(_html))
                _customWebBrowserControl.NavigateToHtml(_html);
            else
                _customWebBrowserControl.InvokeIfRequired(() => _customWebBrowserControl.Navigate(_webAppUri));
        }
    }
}

public static class FormsExtensions
{
    /// <summary>
    ///     Invokes action on a thread that owns the specified control
    /// </summary>
    /// <param name="control"></param>
    /// <param name="action"></param>
    public static void InvokeIfRequired(this Control control, MethodInvoker action)
    {
        if (control != null && control.InvokeRequired)
            control.Invoke(action);
        else
            action();
    }
}