using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;

namespace WebView2PowerPointAddInSample
{
    public interface IWebBrowser
    {
        void AddOrUpdateCookie(WebBrowser.Cookie cookie);
    }

    public class WebBrowser : WebView2, IWebBrowser
    {
        private readonly bool _enableDevTools;
        private readonly bool _enableVerboseLogging;

        public WebBrowser()
        {
            _enableDevTools = true;
            _enableVerboseLogging = true;
            CoreWebView2InitializationCompleted += OnCoreWebView2InitializationCompleted;
            NavigationCompleted += OnNavigationCompleted;
        }

        public CoreWebView2Environment CoreWebView2Environment { get; private set; }

        public void AddOrUpdateCookie(Cookie cookie)
        {
            if (CoreWebView2 != null)
            {
                var coreWebView2Cookie =
                    CoreWebView2.CookieManager.CreateCookie(cookie.Name, cookie.Value, cookie.Domain, cookie.Path);
                coreWebView2Cookie.IsSecure = cookie.Secure;
                coreWebView2Cookie.IsHttpOnly = cookie.HttpOnly;
                CoreWebView2.CookieManager.AddOrUpdateCookie(coreWebView2Cookie);
            }
        }

        private void CoreWebView2OnProcessFailed(object sender, CoreWebView2ProcessFailedEventArgs e)
        {
            Console.WriteLine($"WebView2 Process failed with reason {e.Reason} and exit code {e.ExitCode}");
            Console.WriteLine(
                $"Description of failed process is {e.ProcessDescription} with failed kind {e.ProcessFailedKind}");

            if (e.FrameInfosForFailedProcess != null)
                foreach (var coreWebView2FrameInfo in e.FrameInfosForFailedProcess)
                    Console.WriteLine(
                        $"Frame information Name: {coreWebView2FrameInfo.Name} Source: {coreWebView2FrameInfo.Source}");
        }

        private void OnNavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (_enableVerboseLogging) Console.WriteLine("RestrictedWebBrowser OnNavigationCompleted");

            if (!e.IsSuccess)
                Console.WriteLine($"WebView2 failed to navigate with error {e.WebErrorStatus}");

            if (!e.IsSuccess && e.WebErrorStatus != CoreWebView2WebErrorStatus.OperationCanceled)
                NavigateError?.Invoke(this, e);
        }

        private void OnCoreWebView2InitializationCompleted(object sender,
            CoreWebView2InitializationCompletedEventArgs e)
        {
            if (_enableVerboseLogging)
                Console.WriteLine($"RestrictedWebBrowser InitializationCompleted with result {e.IsSuccess}");

            if (!e.IsSuccess)
            {
                Console.WriteLine("WebView2 initialization failed.");
                InitializationError?.Invoke(this, e);
                return;
            }

            CoreWebView2.ProcessFailed += CoreWebView2OnProcessFailed;
            CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = false;
            CoreWebView2.Settings.IsStatusBarEnabled = false;
            CoreWebView2.Settings.AreDevToolsEnabled = _enableDevTools;
            //CoreWebView2.ClientCertificateRequested += (_, eventArgs) => eventArgs.TryAutoSelectCertificate();
        }

        protected override void Dispose(bool disposing)
        {
            CoreWebView2InitializationCompleted -= OnCoreWebView2InitializationCompleted;
            NavigationCompleted -= OnNavigationCompleted;

            base.Dispose(disposing);
        }

        public event WebBrowserNavigateErrorEventHandler NavigateError;

        public event WebBrowserInitializationErrorEventHandler InitializationError;

        public async Task<object> InvokeScriptAsync(string methodName, object[] methodArguments)
        {
            if (CoreWebView2 == null) return null;

            var methodArgumentsString = string.Join(", ",
                methodArguments?.Select(ma => JsonConvert.SerializeObject(ma)).ToArray() ?? new[] { "" });
            var script = $"window.{methodName}({methodArgumentsString})";
            var scriptResult = await ExecuteScriptAsync(script);
            return JsonConvert.DeserializeObject(scriptResult);
        }

        public async Task Initialize(string userDataFolder)
        {
            if (_enableVerboseLogging) Console.WriteLine("Initializing WebView2");

            // We want to make sure we return on the main thread, since we are accessing the windows forms component.
            if (SynchronizationContext.Current == null)
            {
                if (_enableVerboseLogging)
                    Console.WriteLine("Setting Synchronization Context to main thread");

                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            CoreWebView2Environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
            await EnsureCoreWebView2Async(CoreWebView2Environment);

            if (_enableVerboseLogging) Console.WriteLine("Finished Initializing WebView2");
        }

        public void Navigate(Uri source)
        {
            if (CoreWebView2 == null)
                Source = source;
            else
                CoreWebView2.Navigate(source.AbsoluteUri);
        }

        public void NavigateToHtml(string htmlContent)
        {
            if (CoreWebView2 == null)
                NavigateToString(htmlContent);
            else
                CoreWebView2.NavigateToString(htmlContent);
        }

        public async Task<string> GetCookieInUrlByName(string url, string name)
        {
            var cookies = await CoreWebView2.CookieManager.GetCookiesAsync(url);

            return cookies.FirstOrDefault(c => c.Name == name)?.Value;
        }

        public class Cookie
        {
            public string Name { get; set; }

            public string Value { get; set; }

            public string Domain { get; set; }

            public bool Secure { get; set; }

            public bool HttpOnly { get; set; }

            public string Path { get; set; }
        }
    }

    public delegate void WebBrowserNavigateErrorEventHandler(object sender, CoreWebView2NavigationCompletedEventArgs e);

    public delegate void WebBrowserInitializationErrorEventHandler(object sender,
        CoreWebView2InitializationCompletedEventArgs e);
}