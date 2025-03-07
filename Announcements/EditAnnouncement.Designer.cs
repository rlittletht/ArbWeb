namespace ArbWeb.Announcements
{
    partial class EditAnnouncement
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox txtContent;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox m_showToAssigners;
        private System.Windows.Forms.CheckBox m_showToContacts;
        private System.Windows.Forms.ComboBox m_officialsFilter;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox m_announcementID;
        private System.Windows.Forms.Label label3;

        private void InitializeComponent()
        {
            this.txtContent = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.m_showToAssigners = new System.Windows.Forms.CheckBox();
            this.m_showToContacts = new System.Windows.Forms.CheckBox();
            this.m_officialsFilter = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.m_announcementID = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtContent
            // 
            this.txtContent.Location = new System.Drawing.Point(13, 34);
            this.txtContent.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtContent.Multiline = true;
            this.txtContent.Name = "txtContent";
            this.txtContent.Size = new System.Drawing.Size(715, 602);
            this.txtContent.TabIndex = 1;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(420, 707);
            this.btnSave.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(150, 35);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(578, 707);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(150, 35);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(118, 20);
            this.label1.TabIndex = 4;
            this.label1.Text = "Announcement";
            // 
            // m_showToAssigners
            // 
            this.m_showToAssigners.AutoSize = true;
            this.m_showToAssigners.Location = new System.Drawing.Point(364, 675);
            this.m_showToAssigners.Name = "m_showToAssigners";
            this.m_showToAssigners.Size = new System.Drawing.Size(167, 24);
            this.m_showToAssigners.TabIndex = 5;
            this.m_showToAssigners.Text = "Show to Assigners";
            this.m_showToAssigners.UseVisualStyleBackColor = true;
            // 
            // m_showToContacts
            // 
            this.m_showToContacts.AutoSize = true;
            this.m_showToContacts.Location = new System.Drawing.Point(548, 675);
            this.m_showToContacts.Name = "m_showToContacts";
            this.m_showToContacts.Size = new System.Drawing.Size(161, 24);
            this.m_showToContacts.TabIndex = 6;
            this.m_showToContacts.Text = "Show to Contacts";
            this.m_showToContacts.UseVisualStyleBackColor = true;
            // 
            // m_officialsFilter
            // 
            this.m_officialsFilter.FormattingEnabled = true;
            this.m_officialsFilter.Items.AddRange(new object[] {
            "None",
            "All Official",
            "Active Officials",
            "Ready Officials",
            "Not-Ready Officials"});
            this.m_officialsFilter.Location = new System.Drawing.Point(158, 672);
            this.m_officialsFilter.Name = "m_officialsFilter";
            this.m_officialsFilter.Size = new System.Drawing.Size(189, 28);
            this.m_officialsFilter.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 675);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(127, 20);
            this.label2.TabIndex = 8;
            this.label2.Text = "Show to Officials";
            // 
            // m_announcementID
            // 
            this.m_announcementID.Location = new System.Drawing.Point(158, 645);
            this.m_announcementID.Name = "m_announcementID";
            this.m_announcementID.ReadOnly = true;
            this.m_announcementID.Size = new System.Drawing.Size(189, 26);
            this.m_announcementID.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 648);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(139, 20);
            this.label3.TabIndex = 10;
            this.label3.Text = "Announcement ID";
            // 
            // EditAnnouncement
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(741, 758);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.m_announcementID);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.m_officialsFilter);
            this.Controls.Add(this.m_showToContacts);
            this.Controls.Add(this.m_showToAssigners);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.txtContent);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "EditAnnouncement";
            this.Text = "Edit Announcement";
            this.ResumeLayout(false);
            this.PerformLayout();

        }


    }
}