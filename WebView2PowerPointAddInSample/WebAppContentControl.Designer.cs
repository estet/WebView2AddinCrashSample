using System.Windows.Forms;

namespace WebView2PowerPointAddInSample
{
    partial class WebAppContentControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            this.webBrowserPanel = new System.Windows.Forms.Panel();
            this._customWebBrowserControl = new WebBrowser();

            ((System.ComponentModel.ISupportInitialize)(this._customWebBrowserControl)).BeginInit();
            this.SuspendLayout();
            this._customWebBrowserControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this._customWebBrowserControl.Location = new System.Drawing.Point(0, 0);
            this._customWebBrowserControl.MinimumSize = new System.Drawing.Size(250, 20);
            this._customWebBrowserControl.Name = "_customWebBrowserControl";
            this._customWebBrowserControl.Size = new System.Drawing.Size(568, 544);
            this._customWebBrowserControl.TabIndex = 0;

            // 
            // webBrowserPanel
            // 
            this.webBrowserPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                    | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
            this.webBrowserPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.webBrowserPanel.BackColor = System.Drawing.Color.White;
            this.webBrowserPanel.Controls.Add(this._customWebBrowserControl);
            this.webBrowserPanel.Location = new System.Drawing.Point(0, 0);
            this.webBrowserPanel.Name = "templafyLoadingPanel";
            this.webBrowserPanel.Size = new System.Drawing.Size(757, 670);
            this.webBrowserPanel.TabIndex = 3;


            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.webBrowserPanel);
            webBrowserPanel.Controls.SetChildIndex(_customWebBrowserControl, 2);
            this.webBrowserPanel.ResumeLayout(false);
            this.webBrowserPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(_customWebBrowserControl)).EndInit();

            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private WebBrowser _customWebBrowserControl;
        private Panel webBrowserPanel;
        #endregion
    }
}
