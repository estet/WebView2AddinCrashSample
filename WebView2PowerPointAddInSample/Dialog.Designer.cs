namespace WebView2PowerPointAddInSample
{
    partial class Dialog
    {
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Dialog));
            this.ToolStrip = new System.Windows.Forms.ToolStrip();
            this.CloseButton = new System.Windows.Forms.ToolStripButton();
            this.MaximizeButton = new System.Windows.Forms.ToolStripButton();
            this.ContainerPanel = new System.Windows.Forms.Panel();
            this.ToolStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // ToolStrip
            // 
            this.ToolStrip.AutoSize = false;
            this.ToolStrip.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(121)))), ((int)(((byte)(255)))));
            this.ToolStrip.GripMargin = new System.Windows.Forms.Padding(0);
            this.ToolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.CloseButton,
            this.MaximizeButton});
            this.ToolStrip.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow;
            this.ToolStrip.Location = new System.Drawing.Point(1, 1);
            this.ToolStrip.Name = "ToolStrip";
            this.ToolStrip.Padding = new System.Windows.Forms.Padding(0);
            this.ToolStrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;
            this.ToolStrip.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ToolStrip.Size = new System.Drawing.Size(798, 25);
            this.ToolStrip.TabIndex = 5;
            this.ToolStrip.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ToolStrip_MouseDown);
            // 
            // CloseButton
            // 
            this.CloseButton.AutoSize = false;
            this.CloseButton.BackColor = System.Drawing.Color.Transparent;
            this.CloseButton.Text = "x";
            this.CloseButton.ImageTransparentColor = System.Drawing.Color.White;
            this.CloseButton.Margin = new System.Windows.Forms.Padding(0);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(25, 20);
            this.CloseButton.Click += new System.EventHandler(this.OnCloseButtonClick);
            // 
            // MaximizeButton
            // 
            this.MaximizeButton.AutoSize = false;
            this.MaximizeButton.BackColor = System.Drawing.Color.Transparent;
            this.MaximizeButton.ForeColor = System.Drawing.Color.White;
            this.MaximizeButton.Text = "_";
            this.MaximizeButton.Margin = new System.Windows.Forms.Padding(0);
            this.MaximizeButton.Name = "MaximizeButton";
            this.MaximizeButton.Size = new System.Drawing.Size(25, 20);
            this.MaximizeButton.Click += new System.EventHandler(this.OnMaximizeButtonClick);
            // 
            // ContainerPanel
            // 
            this.ContainerPanel.BackColor = System.Drawing.Color.Transparent;
            this.ContainerPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ContainerPanel.Location = new System.Drawing.Point(1, 26);
            this.ContainerPanel.Margin = new System.Windows.Forms.Padding(0);
            this.ContainerPanel.Name = "ContainerPanel";
            this.ContainerPanel.Size = new System.Drawing.Size(798, 598);
            this.ContainerPanel.TabIndex = 7;
            // 
            // TemplafyDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(121)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(800, 625);
            this.Controls.Add(this.ContainerPanel);
            this.Controls.Add(this.ToolStrip);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.IsMdiContainer = true;
            this.Name = "Dialog";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Test";
            this.ToolStrip.ResumeLayout(false);
            this.ToolStrip.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ToolStripButton CloseButton;
        private System.Windows.Forms.ToolStripButton MaximizeButton;
        public System.Windows.Forms.ToolStrip ToolStrip;
        public System.Windows.Forms.Panel ContainerPanel;
    }
}