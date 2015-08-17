using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using AxSHDocVw;
using mshtml;
using StatusBox;

namespace ArbWeb
    {
    partial class ArbWebControl
        {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
                ComponentResourceManager resources = new ComponentResourceManager(typeof(ArbWebControl));
                this.m_wbc = new AxWebBrowser();
                ((ISupportInitialize)(this.m_wbc)).BeginInit();
                this.SuspendLayout();
                // 
                // m_axWebBrowser1
                // 
                this.m_wbc.Anchor = ((AnchorStyles)((((AnchorStyles.Top | AnchorStyles.Bottom)
                            | AnchorStyles.Left)
                            | AnchorStyles.Right)));
                this.m_wbc.Enabled = true;
                this.m_wbc.Location = new Point(12, 12);
                this.m_wbc.OcxState = ((AxHost.State)(resources.GetObject("m_axWebBrowser1.OcxState")));
                this.m_wbc.Size = new Size(787, 560);
                this.m_wbc.TabIndex = 11;
                this.m_wbc.Visible = false;
                this.m_wbc.BeforeNavigate2 += new DWebBrowserEvents2_BeforeNavigate2EventHandler(this.BeforeNav2);
                this.m_wbc.DownloadComplete += new EventHandler(this.DownloadComplete);
                this.m_wbc.DownloadBegin += new EventHandler(this.DownloadBegin);
                this.m_wbc.NavigateComplete2 += new DWebBrowserEvents2_NavigateComplete2EventHandler(this.NavComplete2);
                this.m_wbc.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(this.TriggerDocumentDone);
                // 
                // ArbWebCore
                // 
                this.AutoScaleDimensions = new SizeF(6F, 13F);
                this.AutoScaleMode = AutoScaleMode.Font;
                this.ClientSize = new Size(811, 584);
                this.Controls.Add(this.m_wbc);
                this.Name = "ArbWebControl";
                this.Text = "ArbWeb Diagnostics";
                ((ISupportInitialize)(this.m_wbc)).EndInit();
                this.ResumeLayout(false);

            }

        #endregion

        private AxWebBrowser m_wbc;
        }
    }