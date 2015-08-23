namespace ArbWeb
{
    partial class ArbWebControl
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
            this.m_wbc = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // m_wbc
            // 
            this.m_wbc.Dock = System.Windows.Forms.DockStyle.Fill;
            this.m_wbc.Location = new System.Drawing.Point(0, 0);
            this.m_wbc.MinimumSize = new System.Drawing.Size(20, 20);
            this.m_wbc.Name = "m_wbc";
            this.m_wbc.Size = new System.Drawing.Size(943, 596);
            this.m_wbc.TabIndex = 0;
            this.m_wbc.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.TriggerDocumentDone);
            // 
            // ArbWebControl2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(943, 596);
            this.Controls.Add(this.m_wbc);
            this.Name = "ArbWebControl2";
            this.Text = "ArbWebControl2";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.WebBrowser m_wbc;
    }
}