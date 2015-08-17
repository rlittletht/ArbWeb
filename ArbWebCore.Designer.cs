namespace ArbWeb
    {
    partial class ArbWebCore
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
            {
                System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ArbWebCore));
                this.m_axWebBrowser1 = new AxSHDocVw.AxWebBrowser();
                ((System.ComponentModel.ISupportInitialize)(this.m_axWebBrowser1)).BeginInit();
                this.SuspendLayout();
                // 
                // m_axWebBrowser1
                // 
                this.m_axWebBrowser1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                            | System.Windows.Forms.AnchorStyles.Left)
                            | System.Windows.Forms.AnchorStyles.Right)));
                this.m_axWebBrowser1.Enabled = true;
                this.m_axWebBrowser1.Location = new System.Drawing.Point(12, 12);
                this.m_axWebBrowser1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("m_axWebBrowser1.OcxState")));
                this.m_axWebBrowser1.Size = new System.Drawing.Size(787, 560);
                this.m_axWebBrowser1.TabIndex = 11;
                this.m_axWebBrowser1.Visible = false;
                this.m_axWebBrowser1.BeforeNavigate2 += new AxSHDocVw.DWebBrowserEvents2_BeforeNavigate2EventHandler(this.BeforeNav2);
                this.m_axWebBrowser1.DownloadComplete += new System.EventHandler(this.DownloadComplete);
                this.m_axWebBrowser1.DownloadBegin += new System.EventHandler(this.DownloadBegin);
                this.m_axWebBrowser1.NavigateComplete2 += new AxSHDocVw.DWebBrowserEvents2_NavigateComplete2EventHandler(this.NavComplete2);
                this.m_axWebBrowser1.DocumentComplete += new AxSHDocVw.DWebBrowserEvents2_DocumentCompleteEventHandler(this.TriggerDocumentDone);
                // 
                // ArbWebCore
                // 
                this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.ClientSize = new System.Drawing.Size(811, 584);
                this.Controls.Add(this.m_axWebBrowser1);
                this.Name = "ArbWebCore";
                this.Text = "ArbWeb Diagnostics";
                ((System.ComponentModel.ISupportInitialize)(this.m_axWebBrowser1)).EndInit();
                this.ResumeLayout(false);

            }

        #endregion

        private AxSHDocVw.AxWebBrowser m_axWebBrowser1;
        }
    }