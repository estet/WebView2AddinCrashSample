using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WebView2PowerPointAddInSample
{
    // Bridge and BridgeAnotherClass are C# classes that implement IDispatch and works with AddHostObjectToScript.
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class BridgeAnotherClass
    {
        // Sample property.
        public string Prop { get; set; } = "Example";
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class Bridge
    {
        private Dictionary<int, string> m_dictionary = new Dictionary<int, string>();

        public BridgeAnotherClass AnotherObject { get; set; } = new BridgeAnotherClass();

        // Sample indexed property.
        [IndexerName("Items")]
        public string this[int index]
        {
            get => m_dictionary[index];
            set => m_dictionary[index] = value;
        }

        public string Func(string param)
        {
            return "Example: " + param;
        }

        public void ShowModalDialog()
        {
            var dialog = new Dialog();
            var webAppContentControl = new WebAppContentControl("https://avetemplafy.sharepoint.com");
            webAppContentControl.Dock = DockStyle.Fill;
            dialog.ContainerPanel.Controls.Add(webAppContentControl);
            try
            {
                dialog.ShowDialog();
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }
        }
    }
}