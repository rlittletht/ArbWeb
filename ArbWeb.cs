using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Net;
using System.IO;
using Microsoft.Win32;
using AxSHDocVw;
using StatusBox;
using mshtml;
using System.Text.RegularExpressions;
using Microsoft.Office;
using System.Runtime.InteropServices;
using Outlook=Microsoft.Office.Interop.Outlook;
using Excel=Microsoft.Office.Interop.Excel;

namespace ArbWeb
{
	/// <summary>
	/// Summary description for AwMainForm.
	/// </summary>
	public class AwMainForm : System.Windows.Forms.Form
	{
		// Top level web pages
		private const string _s_Home = "http://www.arbitersports.com"; // ok2010
		private const string _s_Assigning = "https://www.arbitersports.com/Assigner/Games/NewGamesView.aspx"; // ok2010
		private const string _s_RanksEdit = "https://www.arbitersports.com/Assigner/RanksEdit.aspx"; // ok2010
		private const string _s_AddUser = "https://www.arbitersports.com/Assigner/UserAdd.aspx?userTypeID=3"; // ok2010u
		private const string _s_OfficialsView = "https://www.arbitersports.com/Assigner/OfficialsView.aspx"; // ok2010u
		// Direct access web pages
		private const string _s_Assigning_PrintAddress = "https://www.arbitersports.com/Assigner/Games/Print.aspx?filterID="; // ok2010
		private const string _s_EditUser_MiscFields = "https://www.arbitersports.com/Official/MiscFieldsEdit.aspx?userID="; // ok2010
		private const string _s_EditUser = "https://www.arbitersports.com/Official/OfficialEdit.aspx?userID="; // ok2010u

		// Home links
//		private const string _s_Home_Anchor_Login = "ctl00$SignInButton"; // ctl00$ucMiniLogin$SignInButton"; // ok2010
		//private const string _s_Home_Input_Email = "ctl00$EmailTextbox"; // ctl00$ucMiniLogin$UsernameTextBox"; // ok2010
		//private const string _s_Home_Input_Password = "ctl00$PasswordTextbox"; // ctl00$ucMiniLogin$PasswordTextBox"; // ok2010
		//private const string _s_Home_Button_SignIn = "ctl00$SignInButton"; // ctl00$ucMiniLogin$SignInButton"; // ok2010
		private const string _s_Home_Anchor_Login = "SignInButton"; // ctl00$ucMiniLogin$SignInButton"; // ok2010
		private const string _s_Home_Input_Email = "UserName"; // ctl00$ucMiniLogin$UsernameTextBox"; // ok2010
		private const string _s_Home_Input_Password = "Password"; // ctl00$ucMiniLogin$PasswordTextBox"; // ok2010
		private const string _s_Home_Button_SignIn = "SignInButton"; // ctl00$ucMiniLogin$SignInButton"; // ok2010

		private const string _sid_Home_Div_PnlAccounts = "ctl00_ContentHolder_pgeDefault_conDefault_pnlAccounts"; // ok2010
		private const string _sid_Home_Anchor_ActingLink = "ctl00_ActingLink"; // ok2010

		// Assigning (games view) links
		private const string _s_Assigning_Select_Filters = "ctl00$ContentHolder$pgeGamesView$conGamesView$ddlSavedFilters"; // ok2010
		private const string _s_Assigning_Reports_Select_Format = "ctl00$ContentHolder$pgePrint$conPrint$ddlFormat"; // ok2010
		private const string _s_Assigning_Reports_Submit_Print = "ctl00$ContentHolder$pgePrint$navPrint$btnBeginPrint"; // ok2010

		// Add User links
		private const string _s_AddUser_Input_FirstName = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$txtFirstName"; // ok2010
		private const string _s_AddUser_Input_LastName = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$txtLastName"; // ok2010
		private const string _s_AddUser_Input_Email = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$txtEmail"; // ok2010
		private const string _s_AddUser_Input_Address1 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$txtAddress1"; // ok2010
		private const string _s_AddUser_Input_Address2 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$txtAddress2"; // ok2010
		private const string _s_AddUser_Input_City = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$txtCity"; // ok2010
		private const string _s_AddUser_Input_State = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$txtState"; // ok2010
		private const string _s_AddUser_Input_Zip = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$txtPostalCode"; // ok2010
													 
		private const string _s_AddUser_Input_PhoneNum1 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl00$txtPhone"; // ok2010a
		private const string _s_AddUser_Input_PhoneNum2 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl01$txtPhone"; // ok2010a
		private const string _s_AddUser_Input_PhoneNum3 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl02$txtPhone"; // ok2010a
		private const string _s_AddUser_Input_PhoneType1 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl00$ddlPhone"; // ok2010a
		private const string _s_AddUser_Input_PhoneType2 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl01$ddlPhone"; // ok2010a
		private const string _s_AddUser_Input_PhoneType3 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl02$ddlPhone"; // ok2010a
		private const string _sid_AddUser_Input_PhoneType1 = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_rptPhones_ctl00_ddlPhone"; // ok2010a
		private const string _sid_AddUser_Input_PhoneType2 = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_rptPhones_ctl01_ddlPhone"; // ok2010a
		private const string _sid_AddUser_Input_PhoneType3 = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_rptPhones_ctl02_ddlPhone"; // ok2010a

		private const string _sid_MiscFields_Button_Save = "ctl00_ContentHolder_pgeMiscFieldsEdit_navMiscFieldsEdit_btnSave"; // ok2010u
		private const string _sid_MiscFields_Button_Cancel = "ctl00_ContentHolder_pgeMiscFieldsEdit_navMiscFieldsEdit_lnkCancel"; // ok2010u

		// phone types are Home, Work, Fax, Cellular, Pager, Security, Other
		private const string _sid_AddUser_Button_Next = "ctl00_ContentHolder_pgeUserAdd_navUserAdd_btnNext"; // ok2010
		private const string _sid_AddUser_Input_Address1 = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_txtAddress1"; // ok2010
		private const string _sid_AddUser_Input_IsActive = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_chkIsActive"; // ok2010
		private const string _sid_AddUser_Button_Cancel = "ctl00_ContentHolder_pgeUserAdd_navUserAdd_lnkCancel"; // ok2010
														   
		// OfficialsView page
		private const string _s_Page_OfficialsView = "https://www.arbitersports.com/Assigner/OfficialsView.aspx"; // ok2010
		private const string _s_OfficialsView_Select_Filter = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$ddlFilter"; // ok2010
		private const string _sid_OfficialsView_PrintRoster = "ctl00_ContentHolder_pgeOfficialsView_sbrReports_tskPrint"; // ok2010u
    	private const string _sid_OfficialsView_ContentTable = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_dgOfficials"; // ok2013

        private const string _s_OfficialsView_PaginationHrefPostbackSubstr = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$dgOfficials$ctl204$ctl"; // ok2014

		// OfficialsEdit page
		private const string _sid_OfficialsEdit_Button_Save = "ctl00_ContentHolder_pgeOfficialEdit_navOfficialEdit_btnSave"; // ok2010u
		private const string _sid_OfficialsEdit_Button_Cancel = "ctl00_ContentHolder_pgeOfficialEdit_navOfficialEdit_lnkCancel"; // ok2010u

		// Roster print links
		private const string _sid_RosterPrint_MergeStyle = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkMerge"; // ok2010u 

		private const string _s_RosterPrint_OfficialNumber = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$chkOfficialNumber"; // ok2010
		private const string _s_RosterPrint_DateJoined = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$chkDateJoined"; // ok2010
		private const string _s_RosterPrint_MiscFields = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$chkMiscFields"; // ok2010
		private const string _s_RosterPrint_NonPublicPhone = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$chkIncludeNonPublic"; // ok2010
		private const string _s_RosterPrint_NonPublicAddress = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$chkIncludeNonPublicAddress"; // ok2010

		private const string _sid_RosterPrint_OfficialNumber = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkOfficialNumber"; // ok2010
		private const string _sid_RosterPrint_DateJoined = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkDateJoined"; // ok2010
		private const string _sid_RosterPrint_MiscFields = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkMiscFields"; // ok2010
		private const string _sid_RosterPrint_NonPublicPhone = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkIncludeNonPublic"; // ok2010
		private const string _sid_RosterPrint_NonPublicAddress = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkIncludeNonPublicAddress"; // ok2010
																  
		private const string _sid_RosterPrint_BeginPrint = "ctl00_ContentHolder_pgeOfficialsView_navOfficialsView_btnBeginPrint"; // ok2010

		// Ranks edit page
		private const string _s_RanksEdit_Select_PosNames = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$ddlPosNames"; // ok2010
		private const string _s_RanksEdit_Checkbox_Active = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$chkActive"; // ok2010
		private const string _s_RanksEdit_Checkbox_Rank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$chkRank"; // ok2010
		private const string _s_RanksEdit_Select_NotRanked = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$lstNotRanked"; // ok2010
		private const string _s_RanksEdit_Select_Ranked = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$lstRanked"; // ok2010

		private const string _s_RanksEdit_Button_Unrank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$btnUnrank"; // ok2010
		private const string _s_RanksEdit_Button_ReRank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$btnReRank"; // ok2010
		private const string _s_RanksEdit_Button_Rank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$btnRank"; // ok2010
		private const string _s_RanksEdit_Input_Rank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$txtRank"; // ok2010


	    private System.Windows.Forms.Button m_pbDownloadGames;

	    private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.MenuItem menuItem1;
        private TextBox m_ebGameFile;
        private Label label2;
        private Label label3;
        private TextBox m_ebUserID;
        private Label label4;
        private TextBox m_ebPassword;
        object Zero = 0;
        object EmptyString = "";
        private Label label5;
        private TextBox m_ebRoster;
        private Button button1;
        private GroupBox groupBox2;
        private RichTextBox m_recStatus;
        private Button m_pbGenCounts;
        private Label label1;
        private TextBox m_ebOutputFile;
        private CheckBox m_cbIncludeCanceled;
        private Button m_pbUploadRoster;
        private Button m_pbOpenSlots;
        private Label m_lblSearchCriteria;
        private CheckBox m_cbShowBrowser;
        private Label label6;
		
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
        private Label label7;
        private Label label8;
        private Label label9;
        private DateTimePicker m_dtpStart;
        private DateTimePicker m_dtpEnd;
        private Label label10;
        private TextBox m_ebFilter;
        private CheckBox m_cbOpenSlotDetail;
        private ComboBox m_cbxProfile;
        private Label label11;
        private Button button2;
        private Label label12;
        private CheckBox m_cbFilterSport;
        private CheckedListBox m_chlbxSports;
        private CheckedListBox m_chlbxSportLevels;
        private CheckBox m_cbFilterLevel;
        private CheckBox m_cbTestEmail;
        private TextBox m_ebGameCopy;
        private TextBox m_ebRosterCopy;
        private Label label13;
        private Label label14;
        private CheckBox m_cbTestOnly;
        private Button m_pbReload;
        private CheckBox m_cbRankOnly;
        private Label label15;
        private Label label16;
        private TextBox m_ebGameOutput;
        private Button m_pbGenGames;
        private CheckBox m_cbAddOfficialsOnly;
        private Button m_pbBrowseGameFile;
        private Button m_pbBrowseGameFile2;
        private Button m_pbBrowseRoster;
        private Button m_pbBrowseRoster2;
        private Button m_pbBrowseGamesReport;
        private Button m_pbBrowseAnalysis;
        private Button button3;
        private TextBox m_ebAffiliationIndex;
        private CheckBox m_cbFuzzyTimes;
        private CheckBox m_cbDatePivot;
        private CheckBox m_cbSplitSports;
        private CheckBox m_cbLogToFile;
        private Label label18;
        private CheckedListBox m_chlbxRoster;

		private StatusBox.StatusRpt m_srpt;

		private void ThrowIfNot(bool f, string s)
		{
			if (!f)
				throw new Exception(s);
		}

		private void EH_RenderHeadingLine(object sender, System.Windows.Forms.PaintEventArgs e)
		{
			Label lbl = (Label)sender;
			string s = (string)lbl.Tag;

			SizeF sf = e.Graphics.MeasureString(s, lbl.Font);
			int nWidth = (int)sf.Width;
			int nHeight = (int)sf.Height;

			e.Graphics.DrawString(s, lbl.Font, new SolidBrush(Color.SlateBlue), 0, 0);// new System.Drawing.Point(0, (lbl.Width - nWidth) / 2));
			e.Graphics.DrawLine(new Pen(new SolidBrush(Color.Gray), 1), 6 + nWidth + 1, (nHeight / 2), lbl.Width, (nHeight / 2));

		}

		ReHistory.ReHistElt[] m_rgreheProfile;
		ReHistory.ReHistElt[] m_rgrehe;
		ReHistory m_rehProfile;
		ReHistory m_reh;
		ArbWebCore m_awc;
		bool m_fDontUpdateProfile;

        public void EnableAdminFunctions()
        {
            bool fAdmin = (String.Compare(System.Environment.MachineName, "obelix", true) == 0);
            m_pbUploadRoster.Enabled = fAdmin;
        }
        public AwMainForm()
		{
            //
			// Required for Windows Form Designer support
			//
			m_plCursor = new List<Cursor>();

			InitializeComponent();

			m_srpt = new StatusBox.StatusRpt(m_recStatus);
			m_awc = new ArbWebCore(m_srpt);
			m_fDontUpdateProfile = true;
			RegistryKey rk = Registry.CurrentUser.OpenSubKey("Software\\Thetasoft\\ArbWeb");

            if (rk == null)
                rk = Registry.CurrentUser.CreateSubKey("Software\\Thetasoft\\ArbWeb");
                
			string []rgs = rk.GetSubKeyNames();

			foreach(string s in rgs)
				{
				m_cbxProfile.Items.Add(s);
				}
			m_rgreheProfile = new ReHistory.ReHistElt[] 
				{ 
					new ReHistory.ReHistElt("Login", ReHistory.Type.Str, m_ebUserID, ""), 
					new ReHistory.ReHistElt("Password", ReHistory.Type.Str, m_ebPassword, ""), 
					new ReHistory.ReHistElt("GameFile", ReHistory.Type.Str, m_ebGameFile, ""), 
					new ReHistory.ReHistElt("Roster", ReHistory.Type.Str, m_ebRoster, ""),
					new ReHistory.ReHistElt("GameFileCopy", ReHistory.Type.Str, m_ebGameCopy, ""), 
					new ReHistory.ReHistElt("RosterCopy", ReHistory.Type.Str, m_ebRosterCopy, ""),
					new ReHistory.ReHistElt("GameOutput", ReHistory.Type.Str, m_ebGameOutput, ""),
					new ReHistory.ReHistElt("OutputFile", ReHistory.Type.Str, m_ebOutputFile, ""),
					new ReHistory.ReHistElt("IncludeCanceled", ReHistory.Type.Bool, m_cbIncludeCanceled, 0),
					new ReHistory.ReHistElt("ShowBrowser", ReHistory.Type.Bool, m_cbShowBrowser, 0), 
					new ReHistory.ReHistElt("LastSlotStartDate", ReHistory.Type.Dttm, m_dtpStart, ""),
					new ReHistory.ReHistElt("LastSlotEndDate", ReHistory.Type.Dttm, m_dtpEnd, ""),
					new ReHistory.ReHistElt("LastOpenSlotDetail", ReHistory.Type.Bool, m_cbOpenSlotDetail, 0),
					new ReHistory.ReHistElt("LastGroupTimeSlots", ReHistory.Type.Bool, m_cbFuzzyTimes, 0),
					new ReHistory.ReHistElt("LastTestEmail", ReHistory.Type.Bool, m_cbTestEmail, 0),
					new ReHistory.ReHistElt("AddOfficialsOnly", ReHistory.Type.Bool, m_cbAddOfficialsOnly, 0),
					new ReHistory.ReHistElt("AfiliationIndex", ReHistory.Type.Int, m_ebAffiliationIndex, 0),
					new ReHistory.ReHistElt("LastSplitSports", ReHistory.Type.Bool, m_cbSplitSports, 0),
					new ReHistory.ReHistElt("LastPivotDate", ReHistory.Type.Bool, m_cbDatePivot, 0),
					new ReHistory.ReHistElt("LastLogToFile", ReHistory.Type.Bool, m_cbLogToFile, 0),
				};

			m_rgrehe = new ReHistory.ReHistElt[]
				{
					new ReHistory.ReHistElt("LastProfile", ReHistory.Type.Str, m_cbxProfile, "")
				};

            SetupLogToFile();

			m_reh = new ReHistory(m_rgrehe, "Software\\Thetasoft\\ArbWeb", "root");
			m_reh.Load();


			m_rehProfile = new ReHistory(m_rgreheProfile, String.Format("Software\\Thetasoft\\ArbWeb\\{0}", m_cbxProfile.Text), m_cbxProfile.Text);

			// load MRU from registry
			m_rehProfile.Load();
			m_fDontUpdateProfile = false;
			EnableControls();

			if (m_cbShowBrowser.Checked)
				m_awc.Show();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.Windows.Forms.Label label17;
            this.m_pbDownloadGames = new System.Windows.Forms.Button();
            this.contextMenu1 = new System.Windows.Forms.ContextMenu();
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.m_ebGameFile = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.m_ebUserID = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.m_ebPassword = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.m_ebRoster = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.m_recStatus = new System.Windows.Forms.RichTextBox();
            this.m_pbGenCounts = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.m_ebOutputFile = new System.Windows.Forms.TextBox();
            this.m_cbIncludeCanceled = new System.Windows.Forms.CheckBox();
            this.m_pbUploadRoster = new System.Windows.Forms.Button();
            this.m_pbOpenSlots = new System.Windows.Forms.Button();
            this.m_lblSearchCriteria = new System.Windows.Forms.Label();
            this.m_cbShowBrowser = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.m_dtpStart = new System.Windows.Forms.DateTimePicker();
            this.m_dtpEnd = new System.Windows.Forms.DateTimePicker();
            this.label10 = new System.Windows.Forms.Label();
            this.m_ebFilter = new System.Windows.Forms.TextBox();
            this.m_cbOpenSlotDetail = new System.Windows.Forms.CheckBox();
            this.m_cbxProfile = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.m_cbFilterSport = new System.Windows.Forms.CheckBox();
            this.m_chlbxSports = new System.Windows.Forms.CheckedListBox();
            this.m_chlbxSportLevels = new System.Windows.Forms.CheckedListBox();
            this.m_cbFilterLevel = new System.Windows.Forms.CheckBox();
            this.m_cbTestEmail = new System.Windows.Forms.CheckBox();
            this.m_ebGameCopy = new System.Windows.Forms.TextBox();
            this.m_ebRosterCopy = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.m_cbTestOnly = new System.Windows.Forms.CheckBox();
            this.m_pbReload = new System.Windows.Forms.Button();
            this.m_cbRankOnly = new System.Windows.Forms.CheckBox();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.m_ebGameOutput = new System.Windows.Forms.TextBox();
            this.m_pbGenGames = new System.Windows.Forms.Button();
            this.m_cbAddOfficialsOnly = new System.Windows.Forms.CheckBox();
            this.m_pbBrowseGameFile = new System.Windows.Forms.Button();
            this.m_pbBrowseGameFile2 = new System.Windows.Forms.Button();
            this.m_pbBrowseRoster = new System.Windows.Forms.Button();
            this.m_pbBrowseRoster2 = new System.Windows.Forms.Button();
            this.m_pbBrowseGamesReport = new System.Windows.Forms.Button();
            this.m_pbBrowseAnalysis = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.m_ebAffiliationIndex = new System.Windows.Forms.TextBox();
            this.m_cbFuzzyTimes = new System.Windows.Forms.CheckBox();
            this.m_cbDatePivot = new System.Windows.Forms.CheckBox();
            this.m_cbSplitSports = new System.Windows.Forms.CheckBox();
            this.m_cbLogToFile = new System.Windows.Forms.CheckBox();
            this.label18 = new System.Windows.Forms.Label();
            this.m_chlbxRoster = new System.Windows.Forms.CheckedListBox();
            label17 = new System.Windows.Forms.Label();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label17
            // 
            label17.AutoSize = true;
            label17.Location = new System.Drawing.Point(331, 166);
            label17.Name = "label17";
            label17.Size = new System.Drawing.Size(103, 13);
            label17.TabIndex = 76;
            label17.Text = "Affiliation Field Index";
            // 
            // m_pbDownloadGames
            // 
            this.m_pbDownloadGames.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbDownloadGames.Location = new System.Drawing.Point(617, 28);
            this.m_pbDownloadGames.Name = "m_pbDownloadGames";
            this.m_pbDownloadGames.Size = new System.Drawing.Size(110, 24);
            this.m_pbDownloadGames.TabIndex = 14;
            this.m_pbDownloadGames.Text = "Download Games";
            this.m_pbDownloadGames.Click += new System.EventHandler(this.DoDownloadGames);
            // 
            // contextMenu1
            // 
            this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem1});
            this.contextMenu1.Popup += new System.EventHandler(this.contextMenu1_Popup);
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 0;
            this.menuItem1.Text = "Paste";
            // 
            // m_ebGameFile
            // 
            this.m_ebGameFile.Location = new System.Drawing.Point(76, 55);
            this.m_ebGameFile.Name = "m_ebGameFile";
            this.m_ebGameFile.Size = new System.Drawing.Size(208, 20);
            this.m_ebGameFile.TabIndex = 18;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 13);
            this.label2.TabIndex = 19;
            this.label2.Text = "Game File";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 32);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(43, 13);
            this.label3.TabIndex = 21;
            this.label3.Text = "User ID";
            // 
            // m_ebUserID
            // 
            this.m_ebUserID.Location = new System.Drawing.Point(76, 29);
            this.m_ebUserID.Name = "m_ebUserID";
            this.m_ebUserID.Size = new System.Drawing.Size(109, 20);
            this.m_ebUserID.TabIndex = 20;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(192, 32);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 13);
            this.label4.TabIndex = 23;
            this.label4.Text = "Password";
            // 
            // m_ebPassword
            // 
            this.m_ebPassword.Location = new System.Drawing.Point(252, 29);
            this.m_ebPassword.Name = "m_ebPassword";
            this.m_ebPassword.Size = new System.Drawing.Size(99, 20);
            this.m_ebPassword.TabIndex = 22;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(310, 62);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(63, 13);
            this.label5.TabIndex = 25;
            this.label5.Text = "Ump Roster";
            // 
            // m_ebRoster
            // 
            this.m_ebRoster.Location = new System.Drawing.Point(379, 59);
            this.m_ebRoster.Name = "m_ebRoster";
            this.m_ebRoster.Size = new System.Drawing.Size(175, 20);
            this.m_ebRoster.TabIndex = 24;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(617, 52);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(110, 24);
            this.button1.TabIndex = 26;
            this.button1.Text = "Get FULL Roster";
            this.button1.Click += new System.EventHandler(this.DoDownloadRoster);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.m_recStatus);
            this.groupBox2.Location = new System.Drawing.Point(8, 524);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(719, 157);
            this.groupBox2.TabIndex = 27;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Status";
            // 
            // m_recStatus
            // 
            this.m_recStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_recStatus.Location = new System.Drawing.Point(6, 19);
            this.m_recStatus.Name = "m_recStatus";
            this.m_recStatus.Size = new System.Drawing.Size(707, 132);
            this.m_recStatus.TabIndex = 0;
            this.m_recStatus.Text = "";
            // 
            // m_pbGenCounts
            // 
            this.m_pbGenCounts.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbGenCounts.Location = new System.Drawing.Point(617, 215);
            this.m_pbGenCounts.Name = "m_pbGenCounts";
            this.m_pbGenCounts.Size = new System.Drawing.Size(110, 27);
            this.m_pbGenCounts.TabIndex = 28;
            this.m_pbGenCounts.Text = "Gen Analysis";
            this.m_pbGenCounts.Click += new System.EventHandler(this.DoGenCounts);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 215);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 30;
            this.label1.Text = "Output File";
            // 
            // m_ebOutputFile
            // 
            this.m_ebOutputFile.Location = new System.Drawing.Point(76, 212);
            this.m_ebOutputFile.Name = "m_ebOutputFile";
            this.m_ebOutputFile.Size = new System.Drawing.Size(208, 20);
            this.m_ebOutputFile.TabIndex = 29;
            // 
            // m_cbIncludeCanceled
            // 
            this.m_cbIncludeCanceled.AutoSize = true;
            this.m_cbIncludeCanceled.Location = new System.Drawing.Point(314, 215);
            this.m_cbIncludeCanceled.Name = "m_cbIncludeCanceled";
            this.m_cbIncludeCanceled.Size = new System.Drawing.Size(109, 17);
            this.m_cbIncludeCanceled.TabIndex = 31;
            this.m_cbIncludeCanceled.Text = "Include Canceled";
            this.m_cbIncludeCanceled.UseVisualStyleBackColor = true;
            this.m_cbIncludeCanceled.Click += new System.EventHandler(this.DoGenericInvalGc);
            // 
            // m_pbUploadRoster
            // 
            this.m_pbUploadRoster.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbUploadRoster.Location = new System.Drawing.Point(617, 499);
            this.m_pbUploadRoster.Name = "m_pbUploadRoster";
            this.m_pbUploadRoster.Size = new System.Drawing.Size(110, 24);
            this.m_pbUploadRoster.TabIndex = 32;
            this.m_pbUploadRoster.Text = "Upload Roster";
            this.m_pbUploadRoster.Click += new System.EventHandler(this.DoUploadRoster);
            // 
            // m_pbOpenSlots
            // 
            this.m_pbOpenSlots.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbOpenSlots.Location = new System.Drawing.Point(617, 260);
            this.m_pbOpenSlots.Name = "m_pbOpenSlots";
            this.m_pbOpenSlots.Size = new System.Drawing.Size(110, 27);
            this.m_pbOpenSlots.TabIndex = 33;
            this.m_pbOpenSlots.Text = "Calc Slots";
            this.m_pbOpenSlots.Click += new System.EventHandler(this.CalcOpenSlots);
            // 
            // m_lblSearchCriteria
            // 
            this.m_lblSearchCriteria.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_lblSearchCriteria.Location = new System.Drawing.Point(5, 9);
            this.m_lblSearchCriteria.Name = "m_lblSearchCriteria";
            this.m_lblSearchCriteria.Size = new System.Drawing.Size(716, 16);
            this.m_lblSearchCriteria.TabIndex = 34;
            this.m_lblSearchCriteria.Tag = "Shared Configuration";
            this.m_lblSearchCriteria.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // m_cbShowBrowser
            // 
            this.m_cbShowBrowser.AutoSize = true;
            this.m_cbShowBrowser.Location = new System.Drawing.Point(15, 106);
            this.m_cbShowBrowser.Name = "m_cbShowBrowser";
            this.m_cbShowBrowser.Size = new System.Drawing.Size(111, 17);
            this.m_cbShowBrowser.TabIndex = 35;
            this.m_cbShowBrowser.Text = "Show Diagnostics";
            this.m_cbShowBrowser.UseVisualStyleBackColor = true;
            this.m_cbShowBrowser.Click += new System.EventHandler(this.ChangeShowBrowser);
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label6.Location = new System.Drawing.Point(5, 190);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(716, 19);
            this.label6.TabIndex = 36;
            this.label6.Tag = "Games Worked Analysis";
            this.label6.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label7.Location = new System.Drawing.Point(5, 242);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(716, 19);
            this.label7.TabIndex = 37;
            this.label7.Tag = "Open Slot Reporting";
            this.label7.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(12, 264);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(55, 13);
            this.label8.TabIndex = 40;
            this.label8.Text = "Start Date";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(298, 264);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(52, 13);
            this.label9.TabIndex = 41;
            this.label9.Text = "End Date";
            // 
            // m_dtpStart
            // 
            this.m_dtpStart.CustomFormat = "ddd MMM dd, yyyy";
            this.m_dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.m_dtpStart.Location = new System.Drawing.Point(76, 260);
            this.m_dtpStart.Name = "m_dtpStart";
            this.m_dtpStart.Size = new System.Drawing.Size(154, 20);
            this.m_dtpStart.TabIndex = 42;
            this.m_dtpStart.Value = new System.DateTime(2008, 5, 4, 0, 0, 0, 0);
            // 
            // m_dtpEnd
            // 
            this.m_dtpEnd.CustomFormat = "ddd MMM dd, yyyy";
            this.m_dtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.m_dtpEnd.Location = new System.Drawing.Point(356, 260);
            this.m_dtpEnd.Name = "m_dtpEnd";
            this.m_dtpEnd.Size = new System.Drawing.Size(154, 20);
            this.m_dtpEnd.TabIndex = 43;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(418, 410);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(79, 13);
            this.label10.TabIndex = 45;
            this.label10.Text = "Misc Field Filter";
            // 
            // m_ebFilter
            // 
            this.m_ebFilter.Location = new System.Drawing.Point(503, 407);
            this.m_ebFilter.Name = "m_ebFilter";
            this.m_ebFilter.Size = new System.Drawing.Size(129, 20);
            this.m_ebFilter.TabIndex = 44;
            // 
            // m_cbOpenSlotDetail
            // 
            this.m_cbOpenSlotDetail.AutoSize = true;
            this.m_cbOpenSlotDetail.Location = new System.Drawing.Point(421, 309);
            this.m_cbOpenSlotDetail.Name = "m_cbOpenSlotDetail";
            this.m_cbOpenSlotDetail.Size = new System.Drawing.Size(139, 17);
            this.m_cbOpenSlotDetail.TabIndex = 46;
            this.m_cbOpenSlotDetail.Text = "Include game/slot detail";
            this.m_cbOpenSlotDetail.UseVisualStyleBackColor = true;
            this.m_cbOpenSlotDetail.CheckedChanged += new System.EventHandler(this.HandleSlotDetailChecked);
            // 
            // m_cbxProfile
            // 
            this.m_cbxProfile.FormattingEnabled = true;
            this.m_cbxProfile.Location = new System.Drawing.Point(406, 28);
            this.m_cbxProfile.Name = "m_cbxProfile";
            this.m_cbxProfile.Size = new System.Drawing.Size(143, 21);
            this.m_cbxProfile.TabIndex = 47;
            this.m_cbxProfile.SelectedIndexChanged += new System.EventHandler(this.ChangeProfile);
            this.m_cbxProfile.Leave += new System.EventHandler(this.OnProfileLeave);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(364, 32);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(36, 13);
            this.label11.TabIndex = 48;
            this.label11.Text = "Profile";
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.Location = new System.Drawing.Point(617, 324);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(110, 27);
            this.button2.TabIndex = 49;
            this.button2.Text = "Create Email";
            this.button2.Click += new System.EventHandler(this.GenOpenSlotsMail);
            // 
            // label12
            // 
            this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label12.Location = new System.Drawing.Point(31, 290);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(690, 19);
            this.label12.TabIndex = 50;
            this.label12.Tag = "Open slot email generation";
            this.label12.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // m_cbFilterSport
            // 
            this.m_cbFilterSport.AutoSize = true;
            this.m_cbFilterSport.Location = new System.Drawing.Point(34, 309);
            this.m_cbFilterSport.Name = "m_cbFilterSport";
            this.m_cbFilterSport.Size = new System.Drawing.Size(141, 17);
            this.m_cbFilterSport.TabIndex = 51;
            this.m_cbFilterSport.Text = "Include/filter sport count";
            this.m_cbFilterSport.UseVisualStyleBackColor = true;
            this.m_cbFilterSport.CheckedChanged += new System.EventHandler(this.HandleSportChecked);
            // 
            // m_chlbxSports
            // 
            this.m_chlbxSports.CheckOnClick = true;
            this.m_chlbxSports.FormattingEnabled = true;
            this.m_chlbxSports.Location = new System.Drawing.Point(44, 329);
            this.m_chlbxSports.Name = "m_chlbxSports";
            this.m_chlbxSports.Size = new System.Drawing.Size(164, 94);
            this.m_chlbxSports.TabIndex = 52;
            this.m_chlbxSports.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.DoSportLevelFilter);
            // 
            // m_chlbxSportLevels
            // 
            this.m_chlbxSportLevels.CheckOnClick = true;
            this.m_chlbxSportLevels.FormattingEnabled = true;
            this.m_chlbxSportLevels.Location = new System.Drawing.Point(243, 329);
            this.m_chlbxSportLevels.Name = "m_chlbxSportLevels";
            this.m_chlbxSportLevels.Size = new System.Drawing.Size(164, 94);
            this.m_chlbxSportLevels.TabIndex = 54;
            // 
            // m_cbFilterLevel
            // 
            this.m_cbFilterLevel.AutoSize = true;
            this.m_cbFilterLevel.Location = new System.Drawing.Point(233, 309);
            this.m_cbFilterLevel.Name = "m_cbFilterLevel";
            this.m_cbFilterLevel.Size = new System.Drawing.Size(166, 17);
            this.m_cbFilterLevel.TabIndex = 53;
            this.m_cbFilterLevel.Text = "Include/filter sport level count";
            this.m_cbFilterLevel.UseVisualStyleBackColor = true;
            this.m_cbFilterLevel.CheckedChanged += new System.EventHandler(this.HandleSportLevelChecked);
            // 
            // m_cbTestEmail
            // 
            this.m_cbTestEmail.AutoSize = true;
            this.m_cbTestEmail.Location = new System.Drawing.Point(594, 357);
            this.m_cbTestEmail.Name = "m_cbTestEmail";
            this.m_cbTestEmail.Size = new System.Drawing.Size(117, 17);
            this.m_cbTestEmail.TabIndex = 55;
            this.m_cbTestEmail.Text = "Generate test email";
            this.m_cbTestEmail.UseVisualStyleBackColor = true;
            // 
            // m_ebGameCopy
            // 
            this.m_ebGameCopy.Location = new System.Drawing.Point(76, 78);
            this.m_ebGameCopy.Name = "m_ebGameCopy";
            this.m_ebGameCopy.Size = new System.Drawing.Size(208, 20);
            this.m_ebGameCopy.TabIndex = 56;
            // 
            // m_ebRosterCopy
            // 
            this.m_ebRosterCopy.Location = new System.Drawing.Point(379, 82);
            this.m_ebRosterCopy.Name = "m_ebRosterCopy";
            this.m_ebRosterCopy.Size = new System.Drawing.Size(175, 20);
            this.m_ebRosterCopy.TabIndex = 57;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(11, 81);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(59, 13);
            this.label13.TabIndex = 58;
            this.label13.Text = "  (Working)";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(311, 85);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(59, 13);
            this.label14.TabIndex = 59;
            this.label14.Text = "  (Working)";
            // 
            // m_cbTestOnly
            // 
            this.m_cbTestOnly.AutoSize = true;
            this.m_cbTestOnly.Location = new System.Drawing.Point(132, 106);
            this.m_cbTestOnly.Name = "m_cbTestOnly";
            this.m_cbTestOnly.Size = new System.Drawing.Size(178, 17);
            this.m_cbTestOnly.TabIndex = 60;
            this.m_cbTestOnly.Text = "TEST Only (perform 1 operation)";
            this.m_cbTestOnly.UseVisualStyleBackColor = true;
            // 
            // m_pbReload
            // 
            this.m_pbReload.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbReload.Location = new System.Drawing.Point(617, 525);
            this.m_pbReload.Name = "m_pbReload";
            this.m_pbReload.Size = new System.Drawing.Size(110, 24);
            this.m_pbReload.TabIndex = 61;
            this.m_pbReload.Text = "Load Data";
            this.m_pbReload.Click += new System.EventHandler(this.m_pbReload_Click);
            // 
            // m_cbRankOnly
            // 
            this.m_cbRankOnly.AutoSize = true;
            this.m_cbRankOnly.Location = new System.Drawing.Point(313, 106);
            this.m_cbRankOnly.Name = "m_cbRankOnly";
            this.m_cbRankOnly.Size = new System.Drawing.Size(132, 17);
            this.m_cbRankOnly.TabIndex = 62;
            this.m_cbRankOnly.Text = "Upload Rankings Only";
            this.m_cbRankOnly.UseVisualStyleBackColor = true;
            // 
            // label15
            // 
            this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label15.Location = new System.Drawing.Point(5, 141);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(716, 19);
            this.label15.TabIndex = 63;
            this.label15.Tag = "Game Reporting";
            this.label15.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(12, 166);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(58, 13);
            this.label16.TabIndex = 65;
            this.label16.Text = "Output File";
            // 
            // m_ebGameOutput
            // 
            this.m_ebGameOutput.Location = new System.Drawing.Point(76, 163);
            this.m_ebGameOutput.Name = "m_ebGameOutput";
            this.m_ebGameOutput.Size = new System.Drawing.Size(208, 20);
            this.m_ebGameOutput.TabIndex = 64;
            // 
            // m_pbGenGames
            // 
            this.m_pbGenGames.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbGenGames.Location = new System.Drawing.Point(617, 163);
            this.m_pbGenGames.Name = "m_pbGenGames";
            this.m_pbGenGames.Size = new System.Drawing.Size(110, 27);
            this.m_pbGenGames.TabIndex = 66;
            this.m_pbGenGames.Text = "Games Report";
            this.m_pbGenGames.Click += new System.EventHandler(this.DoGamesReport);
            // 
            // m_cbAddOfficialsOnly
            // 
            this.m_cbAddOfficialsOnly.AutoSize = true;
            this.m_cbAddOfficialsOnly.Location = new System.Drawing.Point(452, 106);
            this.m_cbAddOfficialsOnly.Name = "m_cbAddOfficialsOnly";
            this.m_cbAddOfficialsOnly.Size = new System.Drawing.Size(149, 17);
            this.m_cbAddOfficialsOnly.TabIndex = 67;
            this.m_cbAddOfficialsOnly.Text = "Upload New Officials Only";
            this.m_cbAddOfficialsOnly.UseVisualStyleBackColor = true;
            // 
            // m_pbBrowseGameFile
            // 
            this.m_pbBrowseGameFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbBrowseGameFile.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_pbBrowseGameFile.Location = new System.Drawing.Point(293, 61);
            this.m_pbBrowseGameFile.Name = "m_pbBrowseGameFile";
            this.m_pbBrowseGameFile.Size = new System.Drawing.Size(23, 15);
            this.m_pbBrowseGameFile.TabIndex = 68;
            this.m_pbBrowseGameFile.Tag = ArbWeb.AwMainForm.FNC.GameFile;
            this.m_pbBrowseGameFile.Text = "...";
            this.m_pbBrowseGameFile.Click += new System.EventHandler(this.DoBrowseOpen);
            // 
            // m_pbBrowseGameFile2
            // 
            this.m_pbBrowseGameFile2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbBrowseGameFile2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_pbBrowseGameFile2.Location = new System.Drawing.Point(294, 83);
            this.m_pbBrowseGameFile2.Name = "m_pbBrowseGameFile2";
            this.m_pbBrowseGameFile2.Size = new System.Drawing.Size(23, 15);
            this.m_pbBrowseGameFile2.TabIndex = 69;
            this.m_pbBrowseGameFile2.Tag = ArbWeb.AwMainForm.FNC.GameFile2;
            this.m_pbBrowseGameFile2.Text = "...";
            this.m_pbBrowseGameFile2.Click += new System.EventHandler(this.DoBrowseOpen);
            // 
            // m_pbBrowseRoster
            // 
            this.m_pbBrowseRoster.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbBrowseRoster.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_pbBrowseRoster.Location = new System.Drawing.Point(563, 64);
            this.m_pbBrowseRoster.Name = "m_pbBrowseRoster";
            this.m_pbBrowseRoster.Size = new System.Drawing.Size(23, 15);
            this.m_pbBrowseRoster.TabIndex = 70;
            this.m_pbBrowseRoster.Tag = ArbWeb.AwMainForm.FNC.RosterFile;
            this.m_pbBrowseRoster.Text = "...";
            this.m_pbBrowseRoster.Click += new System.EventHandler(this.DoBrowseOpen);
            // 
            // m_pbBrowseRoster2
            // 
            this.m_pbBrowseRoster2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbBrowseRoster2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_pbBrowseRoster2.Location = new System.Drawing.Point(563, 87);
            this.m_pbBrowseRoster2.Name = "m_pbBrowseRoster2";
            this.m_pbBrowseRoster2.Size = new System.Drawing.Size(23, 15);
            this.m_pbBrowseRoster2.TabIndex = 71;
            this.m_pbBrowseRoster2.Tag = ArbWeb.AwMainForm.FNC.RosterFile2;
            this.m_pbBrowseRoster2.Text = "...";
            this.m_pbBrowseRoster2.Click += new System.EventHandler(this.DoBrowseOpen);
            // 
            // m_pbBrowseGamesReport
            // 
            this.m_pbBrowseGamesReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbBrowseGamesReport.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_pbBrowseGamesReport.Location = new System.Drawing.Point(294, 169);
            this.m_pbBrowseGamesReport.Name = "m_pbBrowseGamesReport";
            this.m_pbBrowseGamesReport.Size = new System.Drawing.Size(23, 15);
            this.m_pbBrowseGamesReport.TabIndex = 72;
            this.m_pbBrowseGamesReport.Tag = ArbWeb.AwMainForm.FNC.ReportFile;
            this.m_pbBrowseGamesReport.Text = "...";
            this.m_pbBrowseGamesReport.Click += new System.EventHandler(this.DoBrowseOpen);
            // 
            // m_pbBrowseAnalysis
            // 
            this.m_pbBrowseAnalysis.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbBrowseAnalysis.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_pbBrowseAnalysis.Location = new System.Drawing.Point(294, 215);
            this.m_pbBrowseAnalysis.Name = "m_pbBrowseAnalysis";
            this.m_pbBrowseAnalysis.Size = new System.Drawing.Size(23, 15);
            this.m_pbBrowseAnalysis.TabIndex = 73;
            this.m_pbBrowseAnalysis.Tag = ArbWeb.AwMainForm.FNC.AnalysisFile;
            this.m_pbBrowseAnalysis.Text = "...";
            this.m_pbBrowseAnalysis.Click += new System.EventHandler(this.DoBrowseOpen);
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Location = new System.Drawing.Point(617, 76);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(110, 24);
            this.button3.TabIndex = 74;
            this.button3.Text = "Get Quick Roster";
            this.button3.Click += new System.EventHandler(this.DoDownloadQuickRoster);
            // 
            // m_ebAffiliationIndex
            // 
            this.m_ebAffiliationIndex.Location = new System.Drawing.Point(443, 163);
            this.m_ebAffiliationIndex.Name = "m_ebAffiliationIndex";
            this.m_ebAffiliationIndex.Size = new System.Drawing.Size(32, 20);
            this.m_ebAffiliationIndex.TabIndex = 77;
            this.m_ebAffiliationIndex.Text = "0";
            this.m_ebAffiliationIndex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // m_cbFuzzyTimes
            // 
            this.m_cbFuzzyTimes.AutoSize = true;
            this.m_cbFuzzyTimes.Location = new System.Drawing.Point(421, 327);
            this.m_cbFuzzyTimes.Name = "m_cbFuzzyTimes";
            this.m_cbFuzzyTimes.Size = new System.Drawing.Size(107, 17);
            this.m_cbFuzzyTimes.TabIndex = 78;
            this.m_cbFuzzyTimes.Text = "Group Time Slots";
            this.m_cbFuzzyTimes.UseVisualStyleBackColor = true;
            // 
            // m_cbDatePivot
            // 
            this.m_cbDatePivot.AutoSize = true;
            this.m_cbDatePivot.Location = new System.Drawing.Point(421, 345);
            this.m_cbDatePivot.Name = "m_cbDatePivot";
            this.m_cbDatePivot.Size = new System.Drawing.Size(91, 17);
            this.m_cbDatePivot.TabIndex = 79;
            this.m_cbDatePivot.Text = "Pivot on Date";
            this.m_cbDatePivot.UseVisualStyleBackColor = true;
            // 
            // m_cbSplitSports
            // 
            this.m_cbSplitSports.AutoSize = true;
            this.m_cbSplitSports.Location = new System.Drawing.Point(421, 363);
            this.m_cbSplitSports.Name = "m_cbSplitSports";
            this.m_cbSplitSports.Size = new System.Drawing.Size(156, 17);
            this.m_cbSplitSports.TabIndex = 80;
            this.m_cbSplitSports.Text = "Automatic Softball/Baseball";
            this.m_cbSplitSports.UseVisualStyleBackColor = true;
            // 
            // m_cbLogToFile
            // 
            this.m_cbLogToFile.AutoSize = true;
            this.m_cbLogToFile.Location = new System.Drawing.Point(604, 106);
            this.m_cbLogToFile.Name = "m_cbLogToFile";
            this.m_cbLogToFile.Size = new System.Drawing.Size(72, 17);
            this.m_cbLogToFile.TabIndex = 81;
            this.m_cbLogToFile.Text = "Log to file";
            this.m_cbLogToFile.UseVisualStyleBackColor = true;
            this.m_cbLogToFile.CheckedChanged += new System.EventHandler(this.m_cbLogToFile_CheckedChanged);
            // 
            // label18
            // 
            this.label18.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label18.Location = new System.Drawing.Point(31, 430);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(690, 19);
            this.label18.TabIndex = 82;
            this.label18.Tag = "Open slot email generation";
            // 
            // m_chlbxRoster
            // 
            this.m_chlbxRoster.CheckOnClick = true;
            this.m_chlbxRoster.FormattingEnabled = true;
            this.m_chlbxRoster.Location = new System.Drawing.Point(44, 441);
            this.m_chlbxRoster.Name = "m_chlbxRoster";
            this.m_chlbxRoster.Size = new System.Drawing.Size(164, 64);
            this.m_chlbxRoster.TabIndex = 83;
            // 
            // AwMainForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(735, 693);
            this.Controls.Add(this.m_chlbxRoster);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.m_cbLogToFile);
            this.Controls.Add(this.m_cbSplitSports);
            this.Controls.Add(this.m_cbDatePivot);
            this.Controls.Add(this.m_cbFuzzyTimes);
            this.Controls.Add(this.m_ebAffiliationIndex);
            this.Controls.Add(label17);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.m_pbBrowseAnalysis);
            this.Controls.Add(this.m_pbBrowseGamesReport);
            this.Controls.Add(this.m_pbBrowseRoster2);
            this.Controls.Add(this.m_pbBrowseRoster);
            this.Controls.Add(this.m_pbBrowseGameFile2);
            this.Controls.Add(this.m_pbBrowseGameFile);
            this.Controls.Add(this.m_cbAddOfficialsOnly);
            this.Controls.Add(this.m_pbGenGames);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.m_ebGameOutput);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.m_cbRankOnly);
            this.Controls.Add(this.m_pbReload);
            this.Controls.Add(this.m_cbTestOnly);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.m_ebRosterCopy);
            this.Controls.Add(this.m_ebGameCopy);
            this.Controls.Add(this.m_cbTestEmail);
            this.Controls.Add(this.m_chlbxSportLevels);
            this.Controls.Add(this.m_cbFilterLevel);
            this.Controls.Add(this.m_chlbxSports);
            this.Controls.Add(this.m_cbFilterSport);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.m_cbxProfile);
            this.Controls.Add(this.m_cbOpenSlotDetail);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.m_ebFilter);
            this.Controls.Add(this.m_dtpEnd);
            this.Controls.Add(this.m_dtpStart);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.m_cbShowBrowser);
            this.Controls.Add(this.m_lblSearchCriteria);
            this.Controls.Add(this.m_pbOpenSlots);
            this.Controls.Add(this.m_pbUploadRoster);
            this.Controls.Add(this.m_cbIncludeCanceled);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.m_ebOutputFile);
            this.Controls.Add(this.m_pbGenCounts);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.m_ebRoster);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.m_ebPassword);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.m_ebUserID);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.m_ebGameFile);
            this.Controls.Add(this.m_pbDownloadGames);
            this.Name = "AwMainForm";
            this.Text = "AwMainForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DoSaveState);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new AwMainForm());
		}
		
		void AdjustRectForAspect(Rectangle rectOrig, ref Rectangle rectTarget)
		{
			double dfH = 1.0;
			double dfV = 1.0;

			dfH = ((double)rectTarget.Width) / ((double)rectOrig.Width);
			dfV = ((double)rectTarget.Height) / ((double)rectOrig.Height);

			double df = Math.Min(dfH, dfV);

			Rectangle rectNew = new Rectangle(rectTarget.X, rectTarget.Y, (int)((double)rectOrig.Width * df), (int)((double)rectOrig.Height * df));

			// now adjust rectTarget to have rectNew centered in it
			rectTarget.X += (rectTarget.Width - rectNew.Width) / 2;
			rectTarget.Y += (rectTarget.Height - rectNew.Height) / 2;
			rectTarget.Width = rectNew.Width;
			rectTarget.Height = rectNew.Height;
		}

		void CenterForRect(Rectangle rectFrame, ref Rectangle rectImage)
		{
			rectImage.X = rectFrame.X + (rectFrame.Width - rectImage.Width) / 2;
		}

		public void Draw(Graphics gr, Rectangle rectItem, Image img)
		{
			int ypDraw = rectItem.Y;

			// draw the item frame
			Pen penDark = new Pen(new SolidBrush(SystemColors.ControlDarkDark), 1);
			Pen penLight = new Pen(new SolidBrush(SystemColors.ControlLightLight), 1);

			gr.FillRectangle(new SolidBrush(SystemColors.Control), rectItem);

			gr.DrawLine(penDark, rectItem.X, rectItem.Y, rectItem.X + rectItem.Width, rectItem.Y);
			gr.DrawLine(penLight, rectItem.X + rectItem.Width, rectItem.Y, rectItem.X + rectItem.Width, rectItem.Y + rectItem.Height);
			gr.DrawLine(penDark, rectItem.X, rectItem.Y, rectItem.X, rectItem.Y + rectItem.Height);
			gr.DrawLine(penLight, rectItem.X, rectItem.Y + rectItem.Height, rectItem.X + rectItem.Width, rectItem.Y + rectItem.Height);

			// ypDraw += ypThumbnailMargin;

//			Rectangle rectImage = m_pbx.ClientRectangle; // m_rectThumbnail;
//			CenterForRect(rectImage, ref rectItem);

			gr.DrawImage(img, rectItem);

			// ypDraw += pexo.ypThumbnail + ypThumbnailMargin;
		}

		private void m_pbLoadValues_Click(object sender, System.EventArgs e)
		{
		}

	    private class PGL
	    {
	        public class OFI
	        {
	            public string sOfficialID;
	            public string sEmail;

	            public OFI()
	            {
	                sOfficialID = null;
	                sEmail = null;
	            }
	        };

	        public PGL()
	        {
	            plofi = new List<OFI>();
	        }
	        public List<OFI> plofi;
//	        public List<string> rgsLinks;
//	        public List<string> rgsData;
	        public int iCur;
	    };
			
        bool m_fLoggedIn;

		RST m_rst;
		GenCounts m_gc;

		private bool FSetCheckboxControlVal(IHTMLDocument2 oDoc2, bool fChecked, string sName)
		{
			IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("input");
            
			foreach (IHTMLInputElement ihie in hec)
				{
				if (String.Compare(ihie.name, sName, true) == 0)
					{
					if (ihie.@checked == fChecked)
						return false;

					ihie.@checked = fChecked;
					return true;
					}
				}
			return false;
		}

        private static bool FSetInputControlText(IHTMLDocument2 oDoc2, string sName, string sValue, bool fCheck)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("input");
            string sT = null;
            bool fNeedSave = false;
            foreach (IHTMLInputElement ihie in hec)
            {
                if (String.Compare(ihie.name, sName, true) == 0)
                {
                    if (fCheck)
                    {
                        sT = ihie.value;
                        if (sT == null)
                            sT = "";
                        if (String.Compare(sValue, sT) != 0)
                            fNeedSave = true;
                    }
                    else
                    {
                        fNeedSave = true;
                    }
                    ihie.value = sValue;
                }
            }
            return fNeedSave;
        }

		/* M P  G E T  S E L E C T  V A L U E S */
		/*----------------------------------------------------------------------------
			%%Function: MpGetSelectValues
			%%Qualified: ArbWeb.AwMainForm.MpGetSelectValues
			%%Contact: rlittle

            for a given <select name=$sName><option value=$sValue>$sText</option>...
         
            Find the given sName select object. Then add a mapping of
            $sText -> $sValue to a dictionary and return it.
		----------------------------------------------------------------------------*/
		Dictionary<string, string> MpGetSelectValues(IHTMLDocument2 oDoc2, string sName)
        
        {
            IHTMLElementCollection hec;
			
            Dictionary<string, string> mp = new Dictionary<string, string>();

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");
            foreach (HTMLSelectElementClass ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if (mp.ContainsKey(ihoe.text))
                            m_srpt.AddMessage(String.Format("How strange!  '{0}' shows up more than once as a position", ihoe.text), StatusBox.StatusRpt.MSGT.Warning);
                        else
                            mp.Add(ihoe.text, ihoe.value);
						}
                    }
                }
            return mp;
        }

        // if fValueIsValue == false, then sValue is the "text" of the option control
	    static bool FSelectMultiSelectOption(IHTMLDocument2 oDoc2, string sName, string sValue, bool fValueIsValue)
		{
			IHTMLElementCollection hec;

			hec = (IHTMLElementCollection)oDoc2.all.tags("select");

			foreach (HTMLSelectElementClass ihie in hec)
				{
				if (String.Compare(ihie.name, sName, true) == 0)
					{
					foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
						{
                        if ((fValueIsValue && ihoe.value == sValue) ||
                            (!fValueIsValue && ihoe.text == sValue))
							{
							ihoe.selected = true;
							return true;
							}
						}
					}
				}
			return false;
		}

		static bool FResetMultiSelectOptions(IHTMLDocument2 oDoc2, string sName)
		{
			IHTMLElementCollection hec;

			hec = (IHTMLElementCollection)oDoc2.all.tags("select");

			foreach (HTMLSelectElementClass ihie in hec)
				{
				if (String.Compare(ihie.name, sName, true) == 0)
					{
					foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
						{
						ihoe.selected = false;
						}
					}
				}
			return true;
		}

        string SGetFilterID(IHTMLDocument2 oDoc2, string sName, string sValue)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");
            foreach (HTMLSelectElementClass ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach (IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if (ihoe.text == sValue)
                            {
                            return ihoe.value;
                            }
                        }
                    }
                }
            return null;
        }

        private bool FSetSelectControlText(IHTMLDocument2 oDoc2, string sName, string sValue, bool fCheck)
        {
            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection)oDoc2.all.tags("select");
            bool fNeedSave = false;
            foreach (IHTMLSelectElement ihie in hec)
                {
                if (String.Compare(ihie.name, sName, true) == 0)
                    {
                    foreach(IHTMLOptionElement ihoe in (IHTMLElementCollection)ihie.tags("option"))
                        {
                        if (ihoe.text == sValue)
                            {
							// value is already set...
							if (ihie.value == ihoe.value)
								return false;
                            ihoe.selected = true;
                            object dummy = null;
                            IHTMLDocument4 oDoc4 = (IHTMLDocument4)oDoc2;
                            object eventObj = oDoc4.CreateEventObject(ref dummy);
                            HTMLSelectElementClass hsec = ihie as HTMLSelectElementClass;
                            hsec.FireEvent("onchange", ref eventObj);
                            return true;
                            }
                        }
                    }
                }
            return fNeedSave;
        }

#if no
		private bool FClickControlName(IHTMLDocument2 oDoc2, string sTag, string sName)
		{
			// find an sTag with name sName
			IHTMLElementCollection hec;
			hec = (IHTMLElementCollection)oDoc2.all.tags(sTag);

			if (sTag.ToUpper() == "a")
				{
				foreach (IHTMLAnchorElement 
				}
		}

#endif // no

        private static bool FClickControl(ArbWebCore awc, IHTMLDocument2 oDoc2, string sId)
        {
//			m_srpt.AddMessage("Before clickcontrol: "+sId);
            ((IHTMLElement)(oDoc2.all.item(sId, 0))).click();
//			m_srpt.AddMessage("After clickcontrol");
            return awc.FWaitForNavFinish();
        }

		private bool FClickControlNoWait(IHTMLDocument2 oDoc2, string sId)
		{
//			m_srpt.AddMessage("Before clickcontrol: "+sId);
			((IHTMLElement)(oDoc2.all.item(sId, 0))).click();
//			m_srpt.AddMessage("After clickcontrol");
			return true;
		}
        bool FCheckForControl(IHTMLDocument2 oDoc2, string sId)
        {
            if (oDoc2.all.item(sId, 0) != null)
                return true;
                
            return false;
        }

		string SGetControlValue(IHTMLDocument2 oDoc2, string sId)
		{
			if (FCheckForControl(oDoc2, sId))
				return (string)((IHTMLInputElement)oDoc2.all.item(sId, 0)).value;
			return null;
		}

        private void EnsureLoggedIn()
        {
            IHTMLDocument2 oDoc2;

            if (m_fLoggedIn == false)
                {
				m_srpt.AddMessage("Logging in...");
				m_srpt.PushLevel();
                // login to arbiter
                // nav to the main arbiter login page
                if (!m_awc.FNavToPage(_s_Home))
                    throw (new Exception("could not navigate to arbiter homepage!"));

                // if this control is already there, then we were auto-logged in...
                oDoc2 = m_awc.Document2;
                if (!FCheckForControl(oDoc2, _sid_Home_Anchor_ActingLink))
                    {
                    IHTMLDocument oDoc = m_awc.Document;
                    IHTMLDocument3 oDoc3 = m_awc.Document3;

					FSetInputControlText(oDoc2, _s_Home_Input_Email, m_ebUserID.Text, false);
                    FSetInputControlText(oDoc2, _s_Home_Input_Password, m_ebPassword.Text, false);

                    m_awc.ResetNav();
                    FClickControl(m_awc, oDoc2, _s_Home_Button_SignIn);
                    m_awc.FWaitForNavFinish();
                    }
                oDoc2 = m_awc.Document2;
                
                int count = 0;

                oDoc2 = m_awc.Document2;
                bool fToggledBrowser = false;
                
                while (count < 100 && (FCheckForControl(oDoc2, _sid_Home_Div_PnlAccounts) || !FCheckForControl(oDoc2, _sid_Home_Anchor_ActingLink)))
                    {
                    if (m_cbShowBrowser.Checked == false)
                        {
                        m_cbShowBrowser.Checked = true;
                        ChangeShowBrowser(null, null);
                        fToggledBrowser = true;
                        }
                        
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(100);
                    oDoc2 = m_awc.Document2;
                    }
                    
                if (fToggledBrowser)
                    {
                    m_cbShowBrowser.Checked = false;
                    ChangeShowBrowser(null, null);
                    }                    
                    
                oDoc2 = m_awc.Document2;
                if (!FCheckForControl(oDoc2, _sid_Home_Anchor_ActingLink))
                    MessageBox.Show("Login failed for arbiter.net!");
                else
                    m_fLoggedIn = true;
				m_srpt.PopLevel();
				m_srpt.AddMessage("Completed login.");
                }
        }
        
        private string SBuildTempFilename(string sBaseName)
        {
            return String.Format("{0}\\{1}{2}.xls", Environment.GetEnvironmentVariable("Temp"),
                                 sBaseName,
                                 System.Guid.NewGuid().ToString());
        }


		private void DownloadGames()
		{
			m_srpt.AddMessage("Starting games download...");
			m_srpt.PushLevel();
            string sTempFile = SBuildTempFilename("temp");

		    sTempFile = DownloadGamesToFile(sTempFile);
		    HandleDownloadGames(sTempFile);

			System.IO.File.Delete(sTempFile);

//          m_awc.PopNewWindow3Delegate();

            // set the view to "all games"
//            fNavDone = false;
//            FSetSelectControlText(oDoc2, "pgeGamesView$conGamesView$ddlSavedFilters", "All Games", false);
//            FWaitForNavFinish();
            
            // ok, now we have all games selected...
            // time to try to download a report
            
            m_srpt.PopLevel();
			m_srpt.AddMessage("Completed downloading games.");
		    return;
		}

	    private string DownloadGamesToFile(string sTempFile)
	    {
	        IHTMLDocument2 oDoc2 = null;

	        EnsureLoggedIn();

	        int count = 0;
	        string sFilter = null;

	        while (count < 2)
	            {
	            // ok, now we're at the main assigner page...
	            if (!m_awc.FNavToPage(_s_Assigning))
	                throw (new Exception("could not navigate to games view"));

	            oDoc2 = m_awc.Document2;
	            sFilter = SGetFilterID(oDoc2, _s_Assigning_Select_Filters, "All Games");
	            if (sFilter != null)
	                break;

	            count++;
	            }

	        if (sFilter == null)
	            throw (new Exception("there is no 'all games' filter"));

	        // now set that filter

	        m_awc.ResetNav();
	        FSetSelectControlText(oDoc2, _s_Assigning_Select_Filters, "All Games", false);
	        m_awc.FWaitForNavFinish();

	        if (!m_awc.FNavToPage(_s_Assigning_PrintAddress + sFilter))
	            throw (new Exception("could not navigate to the reports page!"));

	        // setup the file formats and go!

	        oDoc2 = m_awc.Document2;
	        FSetSelectControlText(oDoc2, _s_Assigning_Reports_Select_Format, "Excel Worksheet Format (.xls)", false);


	        System.Windows.Forms.Clipboard.SetText(sTempFile);

	        m_awc.ResetNav();
	        //          m_awc.PushNewWindow3Delegate(new DWebBrowserEvents2_NewWindow3EventHandler(DownloadGamesNewWindowDelegate));

	        FClickControlNoWait(oDoc2, _s_Assigning_Reports_Submit_Print);

	        MessageBox.Show(
	            String.Format(
	                "Please download the games file to {0}. This path is on the clipboard, so you can just paste it into the file/save dialog when you click Save.\n\nWhen the download is complete, click OK.",
	                sTempFile), "ArbWeb", MessageBoxButtons.OK);
	        return sTempFile;
	    }

    	void NavigateOfficialsPageAllOfficials()
		{
			EnsureLoggedIn();

			ThrowIfNot(m_awc.FNavToPage(_s_Page_OfficialsView), "Couldn't nav to officials view!");
			m_awc.FWaitForNavFinish();
    		
			// from the officials view, make sure we are looking at active officials
			m_awc.ResetNav();
			IHTMLDocument2 oDoc2 = m_awc.Document2;

			FSetSelectControlText(oDoc2, _s_OfficialsView_Select_Filter, "All Officials", true);
			m_awc.FWaitForNavFinish();
		}

		/* P O P U L A T E  P G L  F R O M  P A G E  C O R E */
		/*----------------------------------------------------------------------------
			%%Function: PopulatePglFromPageCore
			%%Qualified: ArbWeb.AwMainForm.PopulatePglFromPageCore
			%%Contact: rlittle

			Return a PGL (page of links) from the give sUrl.  

				rx3 is a match for either the link name or the link text
				rx4 is a match for the link name always (will supercede rx3)
				rxData, if set, is the match that will be used to populat the rgsData

			on exit, rgsLinkNames, rgsLinks, and (optionally) rgsData will be
			populated in the pglLinks

            we need to collect information from two separate places in the DOM --
            the Official Name (Last, First) will be in an anchor linking to the
            offical page (which we can get the official ID from).  Then we have the
            email address from which we can get the actual email address.
         
            because of this, we collect the email first, then note that we
            are looking for the official ID.  essentially, a state machine.. (albeit 2 
            states)
		----------------------------------------------------------------------------*/
		void PopulatePglFromPageCore(PGL pgl, IHTMLDocument2 oDoc)
		{
			IHTMLElementCollection links = oDoc.links;
            Regex rx3 = new Regex("OfficialEdit.aspx\\?userID=.*");
            Regex rxData = new Regex("mailto:.*");

			bool fLookingForEmail = false;

			// build up a list of probable index links
			foreach (HTMLAnchorElementClass link in links)
			    {
				string sLinkName = link.nameProp;
				string sLinkText = link.innerText;
				string sLinkTarget = link.href;

				if (sLinkText == null)
					sLinkText = "";

				if (sLinkName == null)
					sLinkName = link.innerText.Substring(1, link.innerText.Length - 1);

				if (rxData != null && sLinkTarget != null && rxData.IsMatch(sLinkTarget))
					{
					if (fLookingForEmail)
					    {
					    // adjust the top item in plofi...
					    pgl.plofi[pgl.plofi.Count - 1].sEmail = sLinkTarget;
					    fLookingForEmail = false;
					    }
					else
					    {
					    m_srpt.AddMessage("Found (" + sLinkTarget + ") when not looking for email!", StatusRpt.MSGT.Error);
					    }
					}

				if (rx3.IsMatch(sLinkName))
					{
                    PGL.OFI ofi = new PGL.OFI();
                    
                    ofi.sEmail = "";
                    ofi.sOfficialID = link.href;
                    pgl.plofi.Add(ofi);

                    fLookingForEmail = true;
					}

			    }	
			pgl.iCur = 0;
			}

        private bool FAppendToFile(string sFile, string sName, string sEmail, List<string> plsMisc)
		{
            StreamWriter sw = new StreamWriter(sFile, true, System.Text.Encoding.Default);

            sEmail = sEmail.Substring(sEmail.IndexOf(":") + 1);

            if (sw == null)
                return false;

            string sFirst, sLast;

            if (sName.IndexOf(" ") == -1)
                {
                sFirst = sName;
                sLast = "";
                }
            else
                {
                sFirst = sName.Substring(0, sName.IndexOf(" "));
                sLast = sName.Substring(sName.IndexOf(" ") + 1);
                }

            sw.Write("\"{0}\",\"{1}\",\"{2}\"", sFirst, sLast, sEmail);
            foreach(string s in plsMisc)
                sw.Write(",\"{0}\"", s);
                
            sw.WriteLine();
            sw.Close();
            return true;
		}

		/* U P D A T E  M I S C */
		/*----------------------------------------------------------------------------
			%%Function: UpdateMisc
			%%Qualified: ArbWeb.AwMainForm.UpdateMisc
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		void UpdateMisc(PGL pgl, RST rst, ref RSTE rste)
		{
			// ok, nav to the page and scrape
//			m_awc.ResetNav();
//			ThrowIfNot(FClickControl(m_awc.Document2, "pgeOfficialEdit_cmnRelatedData_tskEditMiscFields"), "couldn't find misc fields link");
//          m_awc.FWaitForNavFinish();
			
			if (!m_awc.FNavToPage(_s_EditUser_MiscFields + pgl.plofi[pgl.iCur].sOfficialID))
				{
				throw(new Exception("could not navigate to the officials page"));
				}

            IHTMLDocument oDoc = m_awc.Document;
			IHTMLDocument2 oDoc2 = m_awc.Document2;
			IHTMLDocument3 oDoc3 = m_awc.Document3;

			IHTMLElementCollection hec;

			// misc field info.  every text input field is a misc field we want to save
			hec = (IHTMLElementCollection)oDoc2.all.tags("input");
			List<string> plsValue = new List<string>(); 
			string sValue = null;
			bool fNeedSave = false;
			
			foreach (IHTMLInputElement ihie in hec)
				{
				if (String.Compare(ihie.type, "text", true) == 0)
					{
					// cool, extract the value
					sValue = ihie.value;
					if (sValue == null)
						sValue = "";
					if (rst != null)
						{
						// check to see if it matches what we have
						// find a match on email address first
						List<string> plsMisc = rst.PlsLookupEmail(pgl.plofi[pgl.iCur].sEmail);

						if (plsMisc != null)
						    {
						    if (plsMisc.Count <= plsValue.Count 
						        && ihie.value != null
						        && ihie.value.Length > 0)
						        {
						        // null means empty which replaces non-empty
						        ihie.value = "";
						        fNeedSave = true;
						        }
						    else if (plsMisc.Count > plsValue.Count 
							         && String.Compare(plsMisc[plsValue.Count], sValue, true/*ignoreCase*/) != 0)
							    {
							    ihie.value = plsMisc[plsValue.Count];
							    fNeedSave = true;
							    }
							}
						}
					plsValue.Add(sValue);
					// don't break here -- just get the next misc value...
					}
				}

			if (fNeedSave)
				{
				m_srpt.AddMessage(String.Format("Updating misc info...", pgl.plofi[pgl.iCur].sEmail));
				m_awc.ResetNav();
				ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_MiscFields_Button_Save), "Couldn't find save button");
									
				m_awc.FWaitForNavFinish();
				}
            else
                {
				m_awc.ResetNav();
                ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_MiscFields_Button_Cancel), "Couldn't find cancel button");
//				((IHTMLElement)(oDoc2.all.item("pgeMiscFieldsEdit_navMiscFieldsEdit_btnSave", 0))).click();
				m_awc.FWaitForNavFinish();
				}
                
			if (sValue == null || plsValue.Count == 0)
				throw new Exception("couldn't extract misc field for official");

			rste.m_plsMisc = plsValue;
		}

		void MatchAssignText(IHTMLInputElement ihie, RST rst, string sMatch, string sRsteMatched, ref string sAssign, ref bool fNeedSave, ref bool fFailAssign)
		{
			if (ihie.name.Contains(sMatch))
				{
				sAssign = ihie.value;

				if (sAssign == null)
					sAssign = "";

				if (rst != null)
					{
					// check to see if it matches what we have
					// find a match on email address first
					if (sRsteMatched != null 
						&& String.Compare(sRsteMatched, sAssign, true/*ignoreCase*/) != 0)
						{
						if (ihie.disabled)
						    {
						    fFailAssign = true;
						    }
						else
						    {
						    ihie.value = sRsteMatched;
						    fNeedSave = true;
						    }
						}
					}
				}
		}

		/* U P D A T E  I N F O */
		/*----------------------------------------------------------------------------
			%%Function: UpdateInfo
			%%Qualified: ArbWeb.AwMainForm.UpdateInfo
			%%Contact: rlittle

			when we leave, if rst was null, then rste will have the values as we
			fetched from arbiter
		----------------------------------------------------------------------------*/
		void UpdateInfo(PGL pgl, RST rst, ref RSTE rste, bool fMarkOnly)
        {
            bool fNeedSave = false;
            bool fFailUpdate = false;

            RSTE rsteMatch = null;

			if (rst != null)
                rsteMatch = rst.RsteLookupEmail(pgl.plofi[pgl.iCur].sEmail);

			if (rsteMatch == null)
				rsteMatch = new RSTE();	// just to get nulls filled in to the member variables
			else
				rsteMatch.Marked = true;

			if (fMarkOnly)
				return;

			// ok, nav to the page and scrape
            if (!m_awc.FNavToPage(_s_EditUser + pgl.plofi[pgl.iCur].sOfficialID))
				{
				throw(new Exception("could not navigate to the officials page"));
				}

			IHTMLDocument oDoc = m_awc.Document;
			IHTMLDocument2 oDoc2 = m_awc.Document2;
			IHTMLDocument3 oDoc3 = m_awc.Document3;

			IHTMLElementCollection hec;

			hec = (IHTMLElementCollection)oDoc2.all.tags("input");

			foreach (IHTMLInputElement ihie in hec)
				{
				if (String.Compare(ihie.type, "checkbox", true) == 0)
					{
					// checkboxes are either ready or active
					if (ihie.name.Contains("Active"))
					    rste.m_fActive = String.Compare(ihie.value, "on", true) == 0;
					else if (ihie.name.Contains("Ready"))
						rste.m_fReady = String.Compare(ihie.value, "on", true) == 0;
					}

				if (String.Compare(ihie.type, "text", true) == 0)
					{
					if (ihie.name.Contains("Email"))
						{
//						if (ihie.value == null && rste.m_sEmail != null && rste.m_sEmail != "")
						    // continue;

						if (ihie.value != null && rste.m_sEmail != null)
							{
							if (String.Compare(ihie.value, rste.m_sEmail, true) !=  0)
								throw new Exception("email addresses don't match!");
							}
						else
							{
							m_srpt.AddMessage(String.Format("NULL Email address for {0},{1}", rste.m_sFirst, rste.m_sLast), StatusBox.StatusRpt.MSGT.Error);
							}
						}
					MatchAssignText(ihie, rst, "FirstName", rsteMatch.m_sFirst, ref rste.m_sFirst, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "LastName", rsteMatch.m_sLast, ref rste.m_sLast, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "Address1", rsteMatch.m_sAddress1, ref rste.m_sAddress1, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "Address2", rsteMatch.m_sAddress2, ref rste.m_sAddress2, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "City", rsteMatch.m_sCity, ref rste.m_sCity, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "State", rsteMatch.m_sState, ref rste.m_sState, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "PostalCode", rsteMatch.m_sZip, ref rste.m_sZip, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "OfficialNumber", rsteMatch.m_sOfficialNumber, ref rste.m_sOfficialNumber, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "DateOfBirth", rsteMatch.m_sDateOfBirth, ref rste.m_sDateOfBirth, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "DateJoined", rsteMatch.m_sDateJoined, ref rste.m_sDateJoined, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "GamesPerDay", rsteMatch.m_sGamesPerDay, ref rste.m_sGamesPerDay, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "GamesPerWeek", rsteMatch.m_sGamesPerWeek, ref rste.m_sGamesPerWeek, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "GamesTotal", rsteMatch.m_sTotalGames, ref rste.m_sTotalGames, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "WaitMinutes", rsteMatch.m_sWaitMinutes, ref rste.m_sWaitMinutes, ref fNeedSave, ref fFailUpdate);
					}
				}

            if (fFailUpdate)
                {
                m_srpt.AddMessage(String.Format("FAILED to update some general info!  '{0}' was read only", pgl.plofi[pgl.iCur].sEmail), StatusBox.StatusRpt.MSGT.Error);
                }      
			if (fNeedSave)
				{
				m_srpt.AddMessage(String.Format("Updating general info...", pgl.plofi[pgl.iCur].sEmail));
                m_awc.ResetNav();
                ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_OfficialsEdit_Button_Save), "couldn't find save button");
				m_awc.FWaitForNavFinish();
				}
			else
				{
				m_awc.ResetNav();
				ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_OfficialsEdit_Button_Cancel), "Couldn't find cancel button!");
				m_awc.FWaitForNavFinish();
				}
            if (rste.m_sAddress1 == null || rste.m_sAddress2 == null || rste.m_sCity == null || rste.m_sDateJoined == null
                || rste.m_sDateOfBirth == null || rste.m_sEmail == null || rste.m_sFirst == null || rste.m_sGamesPerDay == null 
                || rste.m_sGamesPerWeek == null || rste.m_sLast == null || rste.m_sOfficialNumber == null
                || rste.m_sState == null || rste.m_sTotalGames == null || rste.m_sWaitMinutes == null
                || rste.m_sZip == null)
                {
				throw new Exception("couldn't extract one more more fields from official info");
				}
		}

        static void VisitRankCallbackUpload(RST rst, string sRankPosition, Dictionary<string, int> mpRanked, Dictionary<string, string> mpRankedId, ArbWebCore awc, StatusBox.StatusRpt srpt)
		{
			IHTMLDocument2 oDoc2;
			oDoc2 = awc.Document2;

            List<string> plsUnrank;
            Dictionary<int, List<string>> mpRank;
            Dictionary<int, List<string>> mpRerank;

            BuildRankingJobs(rst, sRankPosition, mpRanked, out plsUnrank, out mpRank, out mpRerank);

            // at this point, we have a list of jobs to do.

			// first, unrank everyone that needs unranked
			if (plsUnrank.Count > 0)
				{
				FResetMultiSelectOptions(oDoc2, _s_RanksEdit_Select_Ranked);
				foreach (string s in plsUnrank)
					{
					if (!FSelectMultiSelectOption(oDoc2, _s_RanksEdit_Select_Ranked, mpRankedId[s], true))
						throw new Exception("couldn't select an official for unranking!");
					}

				// now, do the unrank
				awc.ResetNav();
				FClickControl(awc, oDoc2, _s_RanksEdit_Button_Unrank);
				awc.FWaitForNavFinish();
				oDoc2 = awc.Document2;
				}

			// now, let's rerank the folks that need to be re-ranked
			// we will do this once for every new rank we are setting
			foreach (int nRank in mpRerank.Keys)
				{
				FResetMultiSelectOptions(oDoc2, _s_RanksEdit_Select_Ranked);
				foreach (string s in mpRerank[nRank])
					{
					if (!FSelectMultiSelectOption(oDoc2, _s_RanksEdit_Select_Ranked, mpRankedId[s], true))
						throw new Exception("couldn't select an official for reranking!");
					}
				FSetInputControlText(oDoc2, _s_RanksEdit_Input_Rank, nRank.ToString(), false);

				// now, rank'em
				awc.ResetNav();
				FClickControl(awc, oDoc2, _s_RanksEdit_Button_ReRank);
				awc.FWaitForNavFinish();
				oDoc2 = awc.Document2;
				}

			// finally, let's rank the folks that weren't ranked before

			foreach (int nRank in mpRank.Keys)
				{
				FResetMultiSelectOptions(oDoc2, _s_RanksEdit_Select_NotRanked);
				foreach (string s in mpRank[nRank])
					{
					if (!FSelectMultiSelectOption(oDoc2, _s_RanksEdit_Select_NotRanked, s, false))
						srpt.AddMessage(String.Format("Could not select an official for ranking: {0}", s),
										  StatusRpt.MSGT.Error);
					// throw new Exception("couldn't select an official for ranking!");
					}

				FSetInputControlText(oDoc2, _s_RanksEdit_Input_Rank, nRank.ToString(), false);

				// now, rank'em
				awc.ResetNav();
				FClickControl(awc, oDoc2, _s_RanksEdit_Button_Rank);
				awc.FWaitForNavFinish();
				oDoc2 = awc.Document2;
				}
		}

        // make the rankings on the page match the rankings in our roster
	    private static void BuildRankingJobs(
            RST rst, 
            string sRankPosition, 
            Dictionary<string, int> mpRanked, 
            out List<string> plsUnrank, // officials that need to be unranked
            out Dictionary<int, List<string>> mpRank, // officials that need to be ranked
            out Dictionary<int, List<string>> mpRerank) // officials that need to be re-ranked
	    {
	        List<RSTE> plrste = rst.Plrste;

	        // there are 3 things we can potentially do-
	        //  1) unrank
	        //  2) rank
	        //  3) re-rank

	        // all of these are most optimally done by multi-selecting and 
	        // doing like-item things together
	        // 
	        // so, we will collect all the stuff together

	        // just keep a list of officials to unrank.
	        // for rank and re-rank, we want a mapping of (rank -> list of officials)

	        plsUnrank = new List<string>();
	        mpRank = new Dictionary<int, List<string>>();
	        mpRerank = new Dictionary<int, List<string>>();

	        // first, unrank any officials that should now become unranked
	        foreach (RSTE rste in plrste)
	            {
	            string sReversed = String.Format("{0}, {1}", rste.m_sLast, rste.m_sFirst);

	            if (!rste.FRanked(sRankPosition))
	                {
	                if (mpRanked.ContainsKey(sReversed))
	                    {
	                    // need to unrank
	                    plsUnrank.Add(sReversed);
	                    }
	                // else, we're cool..we're both unranked
	                }
	            else
	                {
	                int nRank = rste.Rank(sRankPosition);

	                // see if we need to rank or rerank
	                if (mpRanked.ContainsKey(sReversed))
	                    {
	                    // may need to rerank
	                    if (mpRanked[sReversed] != nRank)
	                        {
	                        // need to rerank
	                        if (!mpRerank.ContainsKey(nRank))
	                            mpRerank.Add(nRank, new List<string>());

	                        mpRerank[nRank].Add(sReversed);
	                        }
	                    }
	                else
	                    {
	                    // need to rank
	                    if (!mpRank.ContainsKey(nRank))
	                        mpRank.Add(nRank, new List<string>());

	                    mpRank[nRank].Add(sReversed);
	                    }
	                }
	            }
	    }

	    static void VisitRankCallbackDownload(RST rst, string sRank, Dictionary<string, int> mpRanked, Dictionary<string, string> mpRankedId, ArbWebCore awc, StatusBox.StatusRpt srpt)
		{
			// don't do anything with unranked
			// just add the rankings
			foreach (string s in mpRanked.Keys)
				rst.FAddRanking(s, sRank, mpRanked[s]);
		}

    	delegate void VisitRankCallback(RST rst, string sRank, Dictionary<string, int> mpRanked, Dictionary<string, string> mpRankedId, ArbWebCore awc, StatusBox.StatusRpt srpt);

		void HandleRankings(RST rst, ref RST rstBuilding)
		{
		    if (rst != null && rst.PlsRankings == null)
		        return;
		        
			NavigateArbiterRankings();

		    IHTMLDocument2 oDoc2;

			oDoc2 = m_awc.Document2;

			Dictionary<string, string> mpRankFilter = MpGetSelectValues(oDoc2, _s_RanksEdit_Select_PosNames);
			List<string> plsRankings = PlsRankingsBuildFromRst(rst, rstBuilding, mpRankFilter);

		    if (rst == null)
		        VisitRankings(plsRankings, mpRankFilter, VisitRankCallbackDownload, rstBuilding, false/*fVerbose*/);
    		else
		        VisitRankings(plsRankings, mpRankFilter, VisitRankCallbackUpload, rst, true/*fVerbose*/);
		}

	    /* V I S I T  R A N K I N G S */
	    /*----------------------------------------------------------------------------
	    	%%Function: VisitRankings
	    	%%Qualified: ArbWeb.AwMainForm.VisitRankings
	    	%%Contact: rlittle
         
            Visit a rankings page. Used for both upload and download, with the
            callback interface used to differentiate up/down.
	    ----------------------------------------------------------------------------*/
	    private void VisitRankings(List<string> plsRankedPositions, IDictionary<string, string> mpRankFilter, VisitRankCallback pfnVrc, RST rstParam, bool fVerboseLog)
	    {
            // now, navigate to every ranked positions' page and either fetch or sync every
	        // official
	        m_srpt.LogData("Visit Rankings", 1, StatusRpt.MSGT.Header);
	        m_srpt.LogData("plsRankedPositions:", 2, StatusRpt.MSGT.Body, plsRankedPositions);

	        foreach (string sRankPosition in plsRankedPositions)
	            {
	            m_srpt.AddMessage(String.Format("Processing ranks for {0}...", sRankPosition));

	            if (!FNavigateToRankPosition(mpRankFilter, sRankPosition))
	                {
	                m_srpt.AddMessage("Ranks for position '{0}' do not exist on Arbiter!  Skipping...",
	                    StatusBox.StatusRpt.MSGT.Error);
	                continue;
                    }

	            IHTMLDocument2 oDoc2 = m_awc.Document2;
	            // m_awc.RefreshPage();

	            Dictionary<string, int> mpRanked;
	            Dictionary<string, string> mpRankedId;

	            BuildRankingMapFromPage(oDoc2, sRankPosition, out mpRanked, out mpRankedId);

                m_srpt.LogData("Rankings built: mpRanked:", 4, StatusRpt.MSGT.Body, mpRanked);
                m_srpt.LogData("Rankings built: mpRankedId:", 4, StatusRpt.MSGT.Body, mpRankedId);

	            pfnVrc(rstParam, sRankPosition, mpRanked, mpRankedId, m_awc, m_srpt);

	            if (fVerboseLog)
	                {
                    m_awc.RefreshPage();

	                Dictionary<string, int> mpRankedCheck;
	                Dictionary<string, string> mpRankedIdCheck;

	                BuildRankingMapFromPage(oDoc2, sRankPosition, out mpRankedCheck, out mpRankedIdCheck);

                    List<string> plsUnrank;
                    Dictionary<int, List<string>> mpRank;
                    Dictionary<int, List<string>> mpRerank;
	                BuildRankingJobs(rstParam, sRankPosition, mpRankedCheck, out plsUnrank, out mpRank, out mpRerank);

	                if (plsUnrank.Count != 0)
	                    m_srpt.LogData("plsUnrank not empty: ", 1, StatusRpt.MSGT.Error, plsUnrank);
	                else
	                    m_srpt.LogData("plsUnrank empty after upload", 4, StatusRpt.MSGT.Header);

                    if (mpRank.Count != 0)
                        m_srpt.LogData("mpRank not empty: ", 1, StatusRpt.MSGT.Error, mpRank);
                    else
                        m_srpt.LogData("mpRank empty after upload", 4, StatusRpt.MSGT.Header);
                    if (mpRerank.Count != 0)
                        m_srpt.LogData("mpRerank not empty: ", 1, StatusRpt.MSGT.Error, mpRerank);
                    else
                        m_srpt.LogData("mpRerank empty after upload", 4, StatusRpt.MSGT.Header);
                
                
                }
	            }
	    }

	    private bool FNavigateToRankPosition(IDictionary<string, string> mpRankFilter, string sRankPosition)
	    {
// try to navigate to the page
	        if (!mpRankFilter.ContainsKey(sRankPosition))
	            return false;

	        // make sure we have the right checkbox states 
	        // (Show unranked only = false, Show Active only = false)
	        FSetCheckboxControlVal(m_awc.Document2, false, _s_RanksEdit_Checkbox_Active);
	        FSetCheckboxControlVal(m_awc.Document2, false, _s_RanksEdit_Checkbox_Rank);

	        m_awc.ResetNav();
	        FSetSelectControlText(m_awc.Document2, _s_RanksEdit_Select_PosNames, sRankPosition, false);
	        m_awc.FWaitForNavFinish();
	        return true;
	    }

	    private void BuildRankingMapFromPage(IHTMLDocument2 oDoc2, string sRankPosition, out Dictionary<string, int> mpRanked, out Dictionary<string, string> mpRankedId)
	    {
	        List<string> plsUnranked = new List<string>();
	        mpRanked = new Dictionary<string, int>();
	        mpRankedId = new Dictionary<string, string>();

	        Dictionary<string, string> mpT;

	        // unranked officials
	        mpT = MpGetSelectValues(oDoc2, _s_RanksEdit_Select_NotRanked);

	        foreach (string s in mpT.Keys)
	            plsUnranked.Add(s);

	        // ranked officials
	        mpT = MpGetSelectValues(oDoc2, _s_RanksEdit_Select_Ranked);

	        foreach (string s in mpT.Keys)
	            {
	            int iColon = s.IndexOf(":");
	            if (iColon == -1)
	                throw new Exception("bad format for ranked official on arbiter!");

	            int nRank = Int32.Parse(s.Substring(0, iColon));

	            int iStart = iColon + 1;
	            while (Char.IsWhiteSpace(s.Substring(iStart, 1)[0]))
	                iStart++;

	            string sRankKey = s.Substring(iStart);
	            if (!mpRanked.ContainsKey(sRankKey))
	                mpRanked.Add(sRankKey, nRank);
	            else
	                {
	                m_srpt.AddMessage(
	                    String.Format("Duplicate key {0} adding rank {1} to rank {2}", sRankKey, nRank, sRankPosition),
	                    StatusRpt.MSGT.Error);
	                }

	            if (!mpRankedId.ContainsKey(sRankKey))
	                mpRankedId.Add(sRankKey, mpT[s]);
	            else
	                {
	                m_srpt.AddMessage(
	                    String.Format("Duplicate key {0} adding rankid {1} to rank {2}", sRankKey, mpT[s], sRankPosition),
	                    StatusRpt.MSGT.Error);
	                }
	            }
	    }

	    private void VisitRosterRankUploaded(RST rst, string sRank, Dictionary<string, int> mpRanked, IHTMLDocument2 oDoc2,
	                                         Dictionary<string, string> mpRankedId)
	    {
	    }

	    private static void VisitRosterRankDownload(RST rstBuilding, Dictionary<string, int> mpRanked, string sRank)
	    {
	    }

	    private static List<string> PlsRankingsBuildFromRst(RST rst, RST rstBuilding, Dictionary<string, string> mpRankFilter)
	    {
	        List<string> plsRankings;
	        if (rst == null)
	            {
	            // now, build up our plsRankedPositions
	            plsRankings = new List<string>();

	            foreach (string s in mpRankFilter.Keys)
	                plsRankings.Add(s);

	            rstBuilding.PlsRankings = plsRankings;
	            }
	        else
	            plsRankings = rst.PlsRankings;
	        return plsRankings;
	    }

	    private void NavigateArbiterRankings()
	    {
	        m_awc.ResetNav();
	        if (!m_awc.FNavToPage(_s_RanksEdit))
	            throw (new Exception("could not navigate to the bulk rankings page"));
	        m_awc.FWaitForNavFinish();
	    }

	    void AddOfficials(List<RSTE> plrsteNew)
		{
			foreach (RSTE rste in plrsteNew)
				{
				// add the official rste
                m_srpt.AddMessage(String.Format("Adding official '{0}', {1}", rste.Name, rste.m_sEmail), StatusBox.StatusRpt.MSGT.Body);
				m_srpt.PushLevel();

				// go to the add user page
				m_awc.ResetNav();
				if (!m_awc.FNavToPage(_s_AddUser))
					{
					throw(new Exception("could not navigate to the add user page"));
					}
				m_awc.FWaitForNavFinish(); 

				IHTMLDocument2 oDoc2;

				oDoc2 = m_awc.Document2;

				ThrowIfNot(FSetInputControlText(oDoc2, _s_AddUser_Input_FirstName, rste.m_sFirst, false/*fCheck*/), "Failed to find first name control");
				ThrowIfNot(FSetInputControlText(oDoc2, _s_AddUser_Input_LastName, rste.m_sLast, false/*fCheck*/), "Failed to find last name control");
				ThrowIfNot(FSetInputControlText(oDoc2, _s_AddUser_Input_Email, rste.m_sEmail, false/*fCheck*/), "Failed to find email control");

				m_awc.ResetNav();
				ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_AddUser_Button_Next), "Can't click next button on adduser");
				m_awc.FWaitForNavFinish(); 

				// we are either adding a new user, or a user that arbiter already knows
				// about...
				// 
				if (!FCheckForControl(oDoc2, _sid_AddUser_Input_Address1))
					{
                    m_srpt.AddMessage(String.Format("Email {0} already in use", rste.m_sEmail), StatusBox.StatusRpt.MSGT.Warning);

					// this email is member of another group.  we can't change their personal info
					// do a quick sanity match to make sure this is the same user
					string sText = oDoc2.body.innerText;
					string sPrefix = "is already being used in the system by ";
					int iFirst = sText.IndexOf(sPrefix);

					ThrowIfNot(iFirst > 0, "Failed hierarchy on assumed 'in use' email name");
					iFirst += sPrefix.Length;

					int iLast = sText.IndexOf(". Click", iFirst);
					ThrowIfNot(iLast > iFirst, "couldn't find the end of the users name on 'in use' email page");

					string sName = sText.Substring(iFirst, iLast - iFirst);
					if (String.Compare(sName, rste.Name, true/*ignoreCase*/) != 0)
						{
						if (MessageBox.Show(String.Format("Trying to add office {0} and found a mismatch with existing official {1}, with email {2}", rste.Name, sName, rste.m_sEmail), "ArbWeb", MessageBoxButtons.YesNo) != DialogResult.Yes)
							{
							// ok, then just cancel...
							m_awc.ResetNav();
							ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_AddUser_Button_Cancel), "Can't click cancel button on adduser");
							m_awc.FWaitForNavFinish(); 
							continue;
							}
						}
					// cool, just go on...
					m_awc.ResetNav();
					ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_AddUser_Button_Next), "Can't click next button on adduser");
					m_awc.FWaitForNavFinish(); 

					// sigh, now we're being asked whether we want to add local personal info.  of course
					// we don't since it will be thrown away when they choose to join our group!

					// but make sure that we're really on that page...
					sText = oDoc2.body.innerText;
					ThrowIfNot(sText.IndexOf("as a fully integrated user") > 0, "Didn't find the confirmation text on 'personal info' portion of existing user sequence");

					// cool, let's just move on again...
					m_awc.ResetNav();
					ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_AddUser_Button_Next), "Can't click next button on adduser");
					m_awc.FWaitForNavFinish(); 

					// now fallthrough to the "Official's info" page handling, which is common
					}
				else
					{
					// if there's an address control, then this is a brand new official
					ThrowIfNot(FSetInputControlText(oDoc2, _s_AddUser_Input_Address1, rste.m_sAddress1, false/*fCheck*/), "Failed to find address1 control");
					ThrowIfNot(FSetInputControlText(oDoc2, _s_AddUser_Input_Address1, rste.m_sAddress2, false/*fCheck*/), "Failed to find address2 control");
					ThrowIfNot(FSetInputControlText(oDoc2, _s_AddUser_Input_City, rste.m_sCity, false/*fCheck*/), "Failed to find city control");
					ThrowIfNot(FSetInputControlText(oDoc2, _s_AddUser_Input_State, rste.m_sState, false/*fCheck*/), "Failed to find state control");
					ThrowIfNot(FSetInputControlText(oDoc2, _s_AddUser_Input_Zip, rste.m_sZip, false/*fCheck*/), "Failed to find zip control");

					m_awc.ResetNav();
					ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_AddUser_Button_Next), "Can't click next button on adduser");
					m_awc.FWaitForNavFinish(); 

					// fallthrough to the common handling below
					}

				// now we are on the last add official page
				// the only thing that *might* be interesting on this page is the active button which is
				// not checked by default...
				ThrowIfNot(FCheckForControl(oDoc2, _sid_AddUser_Input_IsActive), "bad hierarchy in add user.  expected screen with 'active' checkbox, didn't find it.");

				// don't worry about Active for now...Just click next again
				m_awc.ResetNav();
				ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_AddUser_Button_Next), "Can't click next button on adduser");
				m_awc.FWaitForNavFinish(); 

				// and now we're on the finish page.  oddly enough, the finish button has the "Cancel" ID
				ThrowIfNot(String.Compare("Finish", SGetControlValue(oDoc2, _sid_AddUser_Button_Cancel)) == 0, "Finish screen didn't have a finish button");

				m_awc.ResetNav();
				ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_AddUser_Button_Cancel), "Can't click finish/cancel button on adduser");
				m_awc.FWaitForNavFinish(); 
				m_srpt.PopLevel();
				// and now we're back somewhere (probably officials edit page)
				// continue to the next one...
				}
			// and that's it...simple n'est pas?
		}

        // Update the "last login" value.  since we are scraping the screen for this, we have to deal with pagination
        static void VOPC_UpdateLastAccess(AwMainForm awf, IHTMLDocument2 oDoc2, Object o)
        {
            RST rstBuilding = (RST)o;

            awf.UpdateLastAccessFromDoc(rstBuilding, oDoc2);
        }

        // object could be RST or PGL
        delegate void VisitOfficialsPageCallback(AwMainForm awf, IHTMLDocument2 oDoc2, Object o);
    	
        void ProcessAllOfficialPages(VisitOfficialsPageCallback vopc, Object o)
		{
			NavigateOfficialsPageAllOfficials();

            IHTMLDocument2 oDoc2 = m_awc.Document2;

            // first, get the first pages and callback

            vopc(this, oDoc2, o);

            // figure out how many pages we have
            // find all of the <a> tags with an href that targets a pagination postback
            IHTMLElementCollection ihec = (IHTMLElementCollection)oDoc2.all.tags("a");
            List<string> plsHrefs = new List<string>();

            foreach (IHTMLAnchorElement iha in ihec)
                {
                if (iha.href != null && iha.href.Contains(_s_OfficialsView_PaginationHrefPostbackSubstr))
                    {
                    // we can't just remember this element because we will be navigating around.  instead we will
                    // just remember the entire target so we can find it again
                    plsHrefs.Add(iha.href);
                    }
                }

            // now, we are going to navigate to each page by finding and clicking each pagination link in turn
            foreach (string sHref in plsHrefs)
                {
                ihec = (IHTMLElementCollection) oDoc2.all.tags("a");
                foreach (IHTMLAnchorElement iha in ihec)
                    {
                    if (String.Compare(iha.href, sHref, true /*ignoreCase*/) == 0)
                        {
                        // now we need to click on the navigation
                        ((IHTMLElement)iha).click();
                        m_awc.FWaitForNavFinish();
                        oDoc2 = m_awc.Document2;

                        vopc(this, oDoc2, o);
                        break; // done processing the element collection -- have to process the next one for the next doc
                        }
                    }

                }
		}

        
        // Assuming we are on the core officials page...
        private void UpdateLastAccessFromDoc(RST rstBuilding, IHTMLDocument2 oDoc2)
	    {
	        IHTMLTable ihtbl;

	        // misc field info.  every text input field is a misc field we want to save
	        ihtbl = (IHTMLTable) oDoc2.all.item(_sid_OfficialsView_ContentTable, 0);

	        foreach (IHTMLTableRow ihtr in ihtbl.rows)
	            {
	            IHTMLElement iheEmail = (IHTMLElement) ihtr.cells.item(3);
	            IHTMLElement iheSignedIn = (IHTMLElement) ihtr.cells.item(4);

	            if (iheEmail == null || iheSignedIn == null)
	                continue;

	            string sEmail = iheEmail.innerText;
	            string sSignedIn = iheSignedIn.innerText;

	            RSTE rste = rstBuilding.RsteLookupEmail(sEmail);
	            if (rste == null)
	                {
	                m_srpt.AddMessage(
	                    String.Format("Lookup failed during ProcessAllOfficialPages for official '{0}'({1})",
	                        ((IHTMLElement) ihtr.cells.item(2)).innerText, sEmail), StatusBox.StatusRpt.MSGT.Error);
	                continue;
	                }

	            m_srpt.AddMessage(String.Format("Updating last access for official '{0}', {1}", rste.Name, sSignedIn),
	                StatusBox.StatusRpt.MSGT.Body);
	            rste.m_sLastSignin = sSignedIn;
	            }
	    }


	    /* D O  C O R E  R O S T E R  U P D A T E */
		/*----------------------------------------------------------------------------
			%%Function: DoCoreRosterUpdate
			%%Qualified: ArbWeb.AwMainForm.DoCoreRosterUpdate
			%%Contact: rlittle

			Do the core roster updating.  We are being given the list of links on
			the official's edit page, the roster that we are uploading (if any),
			and a list of officials to limit our handling to (this is used when 
			we just added new officials and we just want to update their info/misc
			fields...)
			
		----------------------------------------------------------------------------*/
		void DoCoreRosterUpdate(PGL pgl, RST rst, RST rstBuilding, List<RSTE> plrsteLimit)
		{
			pgl.iCur = 0;
            Dictionary<string, bool> mpOfficials = new Dictionary<string, bool>();
            
            if (plrsteLimit != null)
                {
				foreach (RSTE rsteCheck in plrsteLimit)
				    mpOfficials.Add("MAILTO:" + rsteCheck.m_sEmail.ToUpper(), true);
				}

            while (pgl.iCur < pgl.plofi.Count && (rst == null || m_cbRankOnly.Checked == false) && pgl.iCur < pgl.plofi.Count)
				{
				if (rst == null
                    || (rst.PlsLookupEmail(pgl.plofi[pgl.iCur].sEmail) != null
                        && pgl.plofi[pgl.iCur].sEmail.Length != 0))
					{
					RSTE rste = new RSTE();
					bool fMarkOnly = false;

                    rste.SetEmail((string)pgl.plofi[pgl.iCur].sEmail);
                    m_srpt.AddMessage(String.Format("Processing roster info for {0}...", pgl.plofi[pgl.iCur].sEmail));

                    if (m_cbAddOfficialsOnly.Checked && plrsteLimit == null)
                        fMarkOnly = true;
                    
                    if (plrsteLimit != null)
                        {
                        if (!mpOfficials.ContainsKey(((string) pgl.plofi[pgl.iCur].sEmail.ToUpper())))
                            {
                            pgl.iCur++;
                            continue; // it doesn't match an official in the "limit-to" list.
                            }
                        fMarkOnly = false;  // we want to process this one.
                        }                            

                    if (!fMarkOnly)
                        UpdateMisc(pgl, rst, ref rste);

                    // don't call UpdateInfo on a newly added official
                    if (plrsteLimit == null && (rst == null || !rst.IsQuick))
                        UpdateInfo(pgl, rst, ref rste, fMarkOnly);

					if (rst == null)
						{
						rstBuilding.Add(rste);
//                        rste.AppendToFile(sOutFile, m_rgsRankings);
						// at this point, we have the name and the affiliation
						//						if (!FAppendToFile(sOutFile, sName, (string)pgl.rgsData[pgl.iCur], plsValue))
						//							throw new Exception("couldn't append to the file!");
						}
					else
						{
						RSTE rsteT = rst.RsteLookupEmail(pgl.plofi[pgl.iCur].sEmail);

						if (rsteT != null)
							rsteT.Marked = true;
						}

					if (m_cbTestOnly.Checked)
						{
						break;
						}
					}

				pgl.iCur++;
				}
		}

        static void VOPC_PopulatePgl(AwMainForm awf, IHTMLDocument2 oDoc2, Object o)
        {
            awf.PopulatePglFromPageCore((PGL)o, oDoc2);
        }

		PGL PglGetOfficialsFromWeb()
		{
			EnsureLoggedIn();
			int i;

		    PGL pgl = new PGL();

            ProcessAllOfficialPages(VOPC_PopulatePgl, pgl);

//            NavigateOfficialsPageAllOfficials();

//            IHTMLDocument2 oDoc = m_awc.Document2;

//			PopulatePglFromPageCore(pgl, oDoc);
			// for now, assume that all the officials fit on the same screen!!

			// if there are no links, then we aren't logged in yet
            if (pgl.plofi.Count == 0)
				{
				throw(new Exception("Not logged in after EnsureLoggedIn()!!"));
				}

			// ok, now grab the userIDs and put those in the pgl
			i = 0;
            while (i < pgl.plofi.Count)
                {
                string s = (string) pgl.plofi[i].sOfficialID;

				string sID = s.Substring(s.IndexOf("=") + 1);
                pgl.plofi[i].sOfficialID = sID;
				i++;
				}

			return pgl;
		}

		/* H A N D L E  R O S T E R */
		/*----------------------------------------------------------------------------
			%%Function: HandleRoster
			%%Qualified: ArbWeb.AwMainForm.HandleRoster
			%%Contact: rlittle

			If rst == null, then we're downloading the roster.  Otherwise, we are
			uploading
		----------------------------------------------------------------------------*/
		void HandleRoster(RST rst, string sOutFile)
        {
            RST rstBuilding = null;
		    PGL pgl;

			// we're not going to write the roster out until the end now...

			if (rst == null)
				rstBuilding = new RST();

			pgl = PglGetOfficialsFromWeb();
			DoCoreRosterUpdate(pgl, rst, rstBuilding, null/*plrsteLimit*/);

    		if (rstBuilding != null)
				{
				// get the last login date from the officials main page
				NavigateOfficialsPageAllOfficials();
				ProcessAllOfficialPages(VOPC_UpdateLastAccess, rstBuilding);
				}

			if (rst != null)
				{
				List<RSTE> plrsteUnmarked = rst.PlrsteUnmarked();

				// we might have some officials left "unmarked".  These need to be added
				
				// at this point, all the officials have either been marked or need to 
				// be added

				if (plrsteUnmarked.Count > 0)
				    {
					if (MessageBox.Show(String.Format("There are {0} new officials.  Add these officials?", plrsteUnmarked.Count), "ArbWeb", MessageBoxButtons.YesNo) == DialogResult.Yes)
					    {
					    AddOfficials(plrsteUnmarked);
						// now we have to reload the page of links and do the whole thing again (updating info, etc)
						// so we get the misc fields updated.  Then fall through to the rankings and do everyone at
						// once
						pgl = PglGetOfficialsFromWeb(); 	// refresh to get new officials
						DoCoreRosterUpdate(pgl, rst, null/*rstBuilding*/, plrsteUnmarked);
						// now we can fall through to our core ranking handling...
					    }
					}
				}

			// now, do the rankings.  this is easiest done in the bulk rankings tool...
			HandleRankings(rst, ref rstBuilding);
			// lastly, if we're downloading, then output the roster

			if (rst == null)
				rstBuilding.WriteRoster(sOutFile);

			if (m_cbTestOnly.Checked)
				{
				MessageBox.Show("Stopping after 1 roster item");
				}
			}

		void HandleDownloadGames(string sFile)
		{
			object missing = System.Type.Missing;
			Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

			Microsoft.Office.Interop.Excel.Workbook wkb;

			wkb = app.Workbooks.Open(sFile, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);

			string sOutFile = "";
			string sPrefix = "";

			if (m_ebGameFile.Text.Length < 1)
				{
				sOutFile = String.Format("{0}", Environment.GetEnvironmentVariable("temp"));
				}
			else
				{
				sOutFile = System.IO.Path.GetDirectoryName(m_ebGameFile.Text);
				string[] rgs;
				if (m_ebGameFile.Text.Length > 5 && sOutFile.Length > 0)
					{
					rgs = GenCounts.RexHelper.RgsMatch(m_ebGameFile.Text.Substring(sOutFile.Length + 1), "([.*])games");
					if (rgs != null && rgs.Length > 0 && rgs[0] != null)
						sPrefix = rgs[0];
					}
				}


			sOutFile = String.Format("{0}\\{2}games_{1:MM}{1:dd}{1:yy}_{1:HH}{1:mm}.csv", sOutFile, DateTime.Now, sPrefix);

			if (wkb != null)
				{
				wkb.SaveAs(sOutFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
				wkb.Close(0, missing, missing);
				}
			app.Quit();
			app = null;                
			m_ebGameFile.Text = sOutFile;
			System.IO.File.Delete(m_ebGameCopy.Text);
			System.IO.File.Copy(sOutFile, m_ebGameCopy.Text);
		}

        void DownloadGamesNewWindowDelegate(object sender, DWebBrowserEvents2_NewWindow3Event e)
		{
            // at this point, e.bstrUrlContext has the URL to the XLS schedule file!!!
            WebClient wc = new WebClient();

            string sFile = String.Format("{0}\\{1}", Environment.GetEnvironmentVariable("temp"), System.Guid.NewGuid().ToString());
            wc.DownloadFile(e.bstrUrl, sFile);
			HandleDownloadGames(sFile);
            System.IO.File.Delete(sFile);
		}

		void DownloadQuickRosterNewWindowDelegate(object sender, DWebBrowserEvents2_NewWindow3Event e)
		{
			// at this point, e.bstrUrlContext has the URL to the CSV schedule file!!!
			WebClient wc = new WebClient();
			object missing = System.Type.Missing;

			// copy the file directly to the output filenames
			string sFile = m_ebRoster.Text;
			wc.DownloadFile(e.bstrUrl, sFile);
		}
#if no
		private void TriggerDocumentDone(object sender, AxSHDocVw.DWebBrowserEvents2_DocumentCompleteEvent e)
		{
			fNavDone = true;
		}
		bool m_fInternalBrowserChange = false;

		private void ShowBrowserStateChange(object sender, System.EventArgs e) {
			if (!m_fInternalBrowserChange)
				m_axWebBrowser1.Visible = true;
			
		}
#endif // no
		private void contextMenu1_Popup(object sender, System.EventArgs e) {
		
		}


	    private void DoSaveStateCore()
		{
			m_rehProfile.Save();
			m_reh.Save();
		}

		private void DoSaveState(object sender, FormClosingEventArgs e)
        {
			DoSaveStateCore();
        }

        private void DoDownloadGames(object sender, EventArgs e)
        {
            DownloadGames();
        }

        List<Cursor> m_plCursor;
        
        void PushCursor(Cursor crs)
        {
            m_plCursor.Add(this.Cursor);
            this.Cursor = crs;
        }
        
        void PopCursor()
        {
            if (m_plCursor.Count > 0)
                {
                this.Cursor = m_plCursor[m_plCursor.Count - 1];
                m_plCursor.RemoveAt(m_plCursor.Count - 1);
                }
        }

		string SBuildRosterFilename()
		{
			string sOutFile;
            string sPrefix = "";
            
			if (m_ebRoster.Text.Length < 1)
				{
				sOutFile = String.Format("{0}", Environment.GetEnvironmentVariable("temp"));
				}
			else
				{
				sOutFile = System.IO.Path.GetDirectoryName(m_ebRoster.Text);
				string[] rgs;
				if (m_ebRoster.Text.Length > 5 && sOutFile.Length > 0)
					{
					rgs = GenCounts.RexHelper.RgsMatch(m_ebRoster.Text.Substring(sOutFile.Length + 1), "([.*])roster");
					if (rgs != null && rgs.Length > 0 && rgs[0] != null)
						sPrefix = rgs[0];
					}
				}

			sOutFile = String.Format("{0}{2}\\roster_{1:MM}{1:dd}{1:yy}_{1:HH}{1:mm}.csv", sOutFile, DateTime.Now, sPrefix);
			return sOutFile;
		}

		private void DoDownloadQuickRoster(object sender, EventArgs e)
		{
			m_srpt.AddMessage("Starting Quick Roster download...");
			m_srpt.PushLevel();
			PushCursor(Cursors.WaitCursor);

		    string sOutFile = SBuildRosterFilename();

			m_ebRoster.Text = sOutFile;

			//string sTempFile = "C:\\Users\\rlittle\\AppData\\Local\\Temp\\temp3c92cb56-0b95-41c0-8eb5-37387bacf4f6.csv";
			string sTempFile = SRosterFileDownload();

//			m_awc.PopSaveToFile();
//			m_awc.PopNewWindow3Delegate();

			// now, get the rankings and update the last access date

			RST rstBuilding = new RST();

			rstBuilding.ReadRoster(sTempFile);

    		ProcessAllOfficialPages(VOPC_UpdateLastAccess, rstBuilding);
			HandleRankings(null, ref rstBuilding);
			rstBuilding.WriteRoster(m_ebRoster.Text);
//			System.IO.File.Copy(sTempFile, m_ebRoster.Text);
			System.IO.File.Delete(m_ebRosterCopy.Text);
			System.IO.File.Copy(sOutFile, m_ebRosterCopy.Text);

			PopCursor();
			m_srpt.PopLevel();

			m_srpt.AddMessage("Completed Quick Roster download.");
		}

	    private string SRosterFileDownload()
	    {
	        IHTMLDocument2 oDoc2;
// navigate to the officials page...
	        EnsureLoggedIn();

	        ThrowIfNot(m_awc.FNavToPage(_s_Page_OfficialsView), "Couldn't nav to officials view!");
	        m_awc.FWaitForNavFinish();

	        oDoc2 = m_awc.Document2;

	        // from the officials view, make sure we are looking at active officials
	        m_awc.ResetNav();
	        FSetSelectControlText(oDoc2, _s_OfficialsView_Select_Filter, "All Officials", true);
	        m_awc.FWaitForNavFinish();

	        oDoc2 = m_awc.Document2;
	        // now we have all officials showing.  download the report

	        // sometimes running the javascript takes a while, but the page isn't busy
	        int cTry = 3;
	        while (cTry > 0)
	            {
	            m_awc.ResetNav();
	            m_awc.ReportNavState("Before click on PrintRoster: ");
	            ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_OfficialsView_PrintRoster), "Can't click on roster control");
	            m_awc.FWaitForNavFinish();

	            oDoc2 = m_awc.Document2;
	            if (FCheckForControl(oDoc2, _sid_RosterPrint_MergeStyle))
	                break;

	            cTry--;
	            }

	        // now we are on the PrintRoster screen

	        // clicking on the Merge Style control will cause a page refresh
	        m_awc.ResetNav();
	        ThrowIfNot(FClickControl(m_awc, oDoc2, _sid_RosterPrint_MergeStyle), "Can't click on roster control");
	        m_awc.FWaitForNavFinish();

	        oDoc2 = m_awc.Document2;

	        ThrowIfNot(FCheckForControl(oDoc2, _sid_RosterPrint_DateJoined),
	                   "Couldn't find expected control on roster print config!");

	        // check a whole bunch of config checkboxes
	        FSetCheckboxControlVal(oDoc2, true, _s_RosterPrint_DateJoined);
	        FSetCheckboxControlVal(oDoc2, true, _s_RosterPrint_OfficialNumber);
	        FSetCheckboxControlVal(oDoc2, true, _s_RosterPrint_MiscFields);
	        FSetCheckboxControlVal(oDoc2, true, _s_RosterPrint_NonPublicPhone);
	        FSetCheckboxControlVal(oDoc2, true, _s_RosterPrint_NonPublicAddress);

	        m_awc.ResetNav();
//    		m_awc.PushNewWindow3Delegate(new DWebBrowserEvents2_NewWindow3EventHandler(DownloadQuickRosterNewWindowDelegate));
//          m_awc.PushSaveToFile(sOutFile);


	        ((IHTMLElement) (oDoc2.all.item(_sid_RosterPrint_BeginPrint, 0))).click();

	        string sTempFile = String.Format("{0}\\temp{1}.csv", Environment.GetEnvironmentVariable("Temp"),
	                                         System.Guid.NewGuid().ToString());
	        System.Windows.Forms.Clipboard.SetText(sTempFile);
	        MessageBox.Show(
	            String.Format(
	                "Please download the roster to {0}. This path is on the clipboard, so you can just past it into the file/save dialog when you click Save.\n\nWhen the download is complete, click OK.",
	                sTempFile), "ArbWeb", MessageBoxButtons.OK);
	        return sTempFile;
	    }

	    private void DoDownloadRoster(object sender, EventArgs e)
        {
			m_srpt.AddMessage("Starting FULL Roster download...");
			m_srpt.PushLevel();

	        PushCursor(Cursors.WaitCursor);
			string sOutFile = SBuildRosterFilename();

			m_ebRoster.Text = sOutFile;

            HandleRoster(null, sOutFile);
            PopCursor();
			m_srpt.PopLevel();
			System.IO.File.Delete(m_ebRosterCopy.Text);
			System.IO.File.Copy(sOutFile, m_ebRosterCopy.Text);
			m_srpt.AddMessage("Completed FULL Roster download.");
        }

		private void InvalRoster()
		{
			m_rst = null;
		}

		private RST RstEnsure(string sInFile)
		{
			if (m_rst != null)
				return m_rst;

			m_rst = new RST();

			m_rst.ReadRoster(sInFile);
			return m_rst;
		}

		private void DoUploadRoster(object sender, EventArgs e)
		{
			m_srpt.AddMessage("Starting Roster upload...");
			m_srpt.PushLevel();

			string sInFile = m_ebRosterCopy.Text;

			RST rst = RstEnsure(sInFile);

			if (rst.IsQuick && (!m_cbRankOnly.Checked || !rst.HasRankings))
				{
//				MessageBox.Show("Cannot upload a quick roster.  Please perform a full roster download before uploading.\n\nIf you want to upload rankings only, please check 'Upload Rankings Only'");
//    			m_srpt.PopLevel();
                m_srpt.AddMessage("Detected QuickRoster...", StatusBox.StatusRpt.MSGT.Warning);
                }

			HandleRoster(rst, sInFile);
			m_srpt.PopLevel();
			m_srpt.AddMessage("Completed Roster upload.");
		}

		private void InvalGameCount()
		{
			m_gc = null;
		}

		private GenCounts GcEnsure(string sRoster, string sGameFile, bool fIncludeCanceled)
		{
			if (m_gc != null)
				return m_gc;

			GenCounts gc = new GenCounts(m_srpt);

			gc.DoGenCounts(sRoster, sGameFile, fIncludeCanceled, Int32.Parse(m_ebAffiliationIndex.Text));
			m_gc = gc;
			return gc;
		}

        private void DoGenCounts(object sender, EventArgs e)
		{
			m_srpt.AddMessage(String.Format("Generating analysis ({0})...", m_ebOutputFile.Text));
			m_srpt.PushLevel();

			GenCounts gc = GcEnsure(m_ebRosterCopy.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);

			gc.GenReport(m_ebOutputFile.Text);
			m_srpt.PopLevel();
			m_srpt.AddMessage("Analysis complete.");
		}

        private string SHtmlReadFile(string sFile)
        {
            string sHtml = "";
            TextReader tr = new StreamReader(sFile);
            string sLine;
                        
            while ((sLine = tr.ReadLine()) != null)
                sHtml += " " + sLine;
                
            tr.Close();
            return sHtml;
        }    
            
        private void ChangeShowBrowser(object sender, EventArgs e)
        {
            if (m_cbShowBrowser.Checked)
                m_awc.Show();
            else
                m_awc.Hide();
        }

        private void DoGenericInvalGc(object sender, EventArgs e)
        {
            InvalGameCount();
        }

        private void ChangeProfile(object sender, EventArgs e)
		{
			if (!m_fDontUpdateProfile)
				{
				DoSaveStateCore();
				m_rehProfile = new ReHistory(m_rgreheProfile, String.Format("Software\\Thetasoft\\ArbWeb\\{0}", m_cbxProfile.Text), m_cbxProfile.Text);
				m_rehProfile.Load();
				}
		}


        private void OnProfileLeave(object sender, EventArgs e)
        {
            if (m_fDontUpdateProfile)
                return;
               
			if (m_rehProfile.FMatchesTag(m_cbxProfile.Text))
				return;

			// otherwise, this is a new profile
			ChangeProfile(sender, e);
        }

		void EnableControls()
		{
			m_chlbxSports.Enabled = m_cbFilterSport.Checked;
			m_chlbxSportLevels.Enabled = m_cbFilterLevel.Checked;
			m_cbFilterLevel.Enabled = m_cbFilterSport.Checked;
			
			EnableAdminFunctions();
			
		}

        private void HandleSportChecked(object sender, EventArgs e)
        {
			EnableControls();
        }

        private void HandleSportLevelChecked(object sender, EventArgs e)
        {
			EnableControls();
        }

        private void HandleSlotDetailChecked(object sender, EventArgs e)
        {

        }

		string[] RgsFromChlbx(bool fUse, CheckedListBox chlbx)
        {
            return RgsFromChlbx(fUse, chlbx, -1, false, null, false);
        }
        
		string[] RgsFromChlbxSport(bool fUse, CheckedListBox chlbx, string sSport, bool fMatch)
		{
			return RgsFromChlbx(fUse, chlbx, -1, false, sSport, fMatch);
		}
		        
		string[] RgsFromChlbx(bool fUse, CheckedListBox chlbx, int iForceToggle, bool fForceOn, string sSport, bool fMatch)
		{
		    string sSport2 = sSport == "Softball" ? "SB" : sSport;

			if (!fUse && sSport == null)
				return null;

            int c = chlbx.CheckedItems.Count;
            
			if (!fUse)
				c = chlbx.Items.Count;
			            
            if (iForceToggle != -1)
                {
                if (fForceOn)
                    c++;
                else
                    c--;
                }
                
			string[] rgs = new string[c];
			int i = 0;

			if (!fUse)
				{
				int iT = 0;

				for (i = 0; i < c; i++)
					{
					rgs[iT] = (string)chlbx.Items[i];
					if (sSport != null)
						{
						if ((rgs[iT].IndexOf(sSport) >= 0 && fMatch)
						    || (rgs[iT].IndexOf(sSport) == -1 && !fMatch)
						    || (rgs[iT].IndexOf(sSport2) >= 0 && fMatch)
						    || (rgs[iT].IndexOf(sSport2) == -1 && !fMatch))
						    {
						    iT++;
						    }
						}
					else
						{
						iT++;
						}
					}
				if (iT < c)
					Array.Resize(ref rgs, iT);

				return rgs;
				}

			i = 0;
            foreach (int iChecked in chlbx.CheckedIndices)
                {
                if (iChecked == iForceToggle)
                    continue;
                rgs[i] = (string)chlbx.Items[iChecked];
				if (sSport != null)
					{
					if ((rgs[i].IndexOf(sSport) >= 0 && fMatch)
						|| (rgs[i].IndexOf(sSport) == -1 && !fMatch))
						{
						i++;
						}
					}
				else
					{
					i++;
					}
                }
            if (fForceOn && iForceToggle != -1)
                rgs[i++] = (string)chlbx.Items[iForceToggle];

			if (i < c)
				Array.Resize(ref rgs, i);

			return rgs;

		}

		private void GenOpenSlotsMail(object sender, EventArgs e)
		{
			GenCounts gc = GcEnsure(m_ebRosterCopy.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
			string sTempFile = String.Format("{0}\\temp{1}.htm", Environment.GetEnvironmentVariable("Temp"), System.Guid.NewGuid().ToString());
            RST rst = RstEnsure(m_ebRosterCopy.Text);
            
            string sBcc = m_cbTestEmail.Checked ? "" : rst.SBuildAddressLine(m_ebFilter.Text); ;

			Outlook.Application appOlk = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");

			if (appOlk == null)
				{
				MessageBox.Show("No running instance of outlook!");
				return;
				}

			Outlook.MailItem oNote = (Outlook.MailItem)appOlk.CreateItem(Outlook.OlItemType.olMailItem);

			oNote.To = "rlittle@thetasoft.com";
			oNote.BCC = sBcc;
			oNote.Subject = "This is a test";
			oNote.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            oNote.HTMLBody = "<html><style>\r\n*#myId {\ncolor:Blue;\n}\n</style><body><p>Put your preamble here...</p>";

		    if (m_cbSplitSports.Checked)
		        {
		        string[] rgs;

		        oNote.HTMLBody += "<h1>Baseball open slots</h1>";
		        rgs = RgsFromChlbxSport(m_cbFilterSport.Checked, m_chlbxSports, "Softball", false);
		        gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
                                      rgs, RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
		        oNote.HTMLBody += SHtmlReadFile(sTempFile) + "<h1>Softball Open Slots</h1>";
		        rgs = RgsFromChlbxSport(m_cbFilterSport.Checked, m_chlbxSports, "Softball", true);
		        gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
                                      rgs, RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
		        oNote.HTMLBody += SHtmlReadFile(sTempFile);
		        }
		    else
		        {
		        gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
		                              RgsFromChlbx(m_cbFilterSport.Checked, m_chlbxSports),
		                              RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
		        oNote.HTMLBody += SHtmlReadFile(sTempFile);
		        }
		    oNote.Display(true);

			appOlk = null;
			System.IO.File.Delete(sTempFile);            
		}


		void UpdateChlbxFromRgs(CheckedListBox chlbx, string[] rgsSource, string[] rgsChecked, string[] rgsFilterPrefix, bool fCheckAll)
		{
			chlbx.Items.Clear();
			SortedList<string, int> mp = Utils.PlsUniqueFromRgs(rgsChecked);
			
			foreach (string s in rgsSource)
				{
				bool fSkip = false;

				if (rgsFilterPrefix != null)
					{
					fSkip = true;
					foreach (string sPrefix in rgsFilterPrefix)
						{
						if (s.Length > sPrefix.Length && String.Compare(s.Substring(0, sPrefix.Length), sPrefix, true/*ignoreCase*/) == 0)
							{
							fSkip = false;
							break;
							}
						}
					}
				if (fSkip)
					continue;

				CheckState cs;

				if (fCheckAll || mp.ContainsKey(s))
					cs = CheckState.Checked;
				else
					cs = CheckState.Unchecked;

				int i = chlbx.Items.Add(s, cs);
				}
		}

	    private SlotAggr m_saOpenSlots;
        private void CalcOpenSlots(object sender, EventArgs e)
        {
			GenCounts gc = GcEnsure(m_ebRosterCopy.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
			RST rst = RstEnsure(m_ebRosterCopy.Text);

            m_saOpenSlots = gc.CalcOpenSlots(m_dtpStart.Value, m_dtpEnd.Value);

			// update regenerate the listboxes...
			string[] rgsSports = RgsFromChlbx(true, m_chlbxSports);
			string[] rgsSportLevels = RgsFromChlbx(true, m_chlbxSportLevels);

			bool fCheckAllSports = false;
			bool fCheckAllSportLevels = false;

			if (rgsSports.Length == 0 && m_chlbxSports.Items.Count == 0)
				fCheckAllSports = true;

			if (rgsSports.Length == 0 && m_chlbxSportLevels.Items.Count == 0)
				fCheckAllSportLevels = true;

            UpdateChlbxFromRgs(m_chlbxSports, gc.GetOpenSlotSports(m_saOpenSlots), rgsSports, null, fCheckAllSports);
            UpdateChlbxFromRgs(m_chlbxSportLevels, gc.GetOpenSlotSportLevels(m_saOpenSlots), rgsSportLevels, fCheckAllSports ? null : rgsSports, fCheckAllSportLevels);
        }

        private void DoSportLevelFilter(object sender, ItemCheckEventArgs e)
        {
            GenCounts gc = GcEnsure(m_ebRosterCopy.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
            string[] rgsSports = RgsFromChlbx(true, m_chlbxSports, e.Index, e.CurrentValue != CheckState.Checked, null, false);
            string[] rgsSportLevels = RgsFromChlbx(true, m_chlbxSportLevels);
            UpdateChlbxFromRgs(m_chlbxSportLevels, gc.GetOpenSlotSportLevels(m_saOpenSlots), rgsSportLevels, rgsSports, false);
        }

        private void m_pbReload_Click(object sender, EventArgs e)
		{
            InvalRoster();
            InvalGameCount();
		}

        private void DoGamesReport(object sender, EventArgs e)
		{
			m_srpt.AddMessage(String.Format("Generating games report ({0})...", m_ebGameOutput.Text));
			m_srpt.PushLevel();
			GenCounts gc = GcEnsure(m_ebRosterCopy.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
			RST rst = RstEnsure(m_ebRosterCopy.Text);

            gc.GenGamesReport(m_ebGameOutput.Text);
			m_srpt.PopLevel();
			m_srpt.AddMessage("Games report complete.");
		}

        enum FNC // FileName Control
            {
            GameFile,
            GameFile2,
            RosterFile,
            RosterFile2,
            ReportFile,
            AnalysisFile
            };
        
        TextBox EbFromFnc(FNC fnc)
        {
            switch (fnc)
                {
                case FNC.GameFile:
                    return m_ebGameFile;
                case FNC.GameFile2:
                    return m_ebGameCopy;
                case FNC.RosterFile:
                    return m_ebRoster;
                case FNC.RosterFile2:
                    return m_ebRosterCopy;
                case FNC.AnalysisFile:
                    return m_ebOutputFile;
                case FNC.ReportFile:
                    return m_ebGameOutput;
                }
            return null;
        }
        
        private void DoBrowseOpen(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            TextBox eb = EbFromFnc((FNC)(((Button)sender).Tag));
            
            ofd.InitialDirectory = Path.GetDirectoryName(eb.Text);
            if (ofd.ShowDialog() == DialogResult.OK)
                {
                eb.Text = ofd.FileName;
                }
        }

        private void m_cbLogToFile_CheckedChanged(object sender, EventArgs e)
        {
            SetupLogToFile();
        }

	    private void SetupLogToFile()
	    {
	        if (m_cbLogToFile.Checked)
	            {
	            m_srpt.AttachLogfile(SBuildTempFilename("arblog"));
	            m_srpt.SetLogLevel(5);
	            m_srpt.SetFilter(StatusRpt.MSGT.Body);
	            }
	        else
	            {
	            m_srpt.SetLogLevel(0);
	            m_srpt.SetFilter(StatusRpt.MSGT.Error);
	            }
	    }
	}

	public class ReHistory
		{
		public enum Type
			{
			Bool = 1,
			Str = 2,
			Int = 3,
			Dttm = 4,
			};
	
		public struct ReHistElt
			{
			public string sRePath;
			public Type type;
			public object oref;
			public object oDefault;
	
			public ReHistElt(string s, Type typ, object o, object oDef)
				{
				sRePath = s;
				type = typ;
				oref = o;
				oDefault = oDef;
				}
			};
	
		ReHistElt[] m_rgrehe;
		string m_sRoot;
		string m_sTag;

		public bool FMatchesTag(string sTag)
		{
			return String.Compare(sTag, m_sTag) == 0;
		}

		public ReHistory(ReHistElt[] rgrehe, string sReRoot, string sTag)
		{
			m_rgrehe = rgrehe;
			m_sRoot = sReRoot;
			m_sTag = sTag;
		}
	
		public void Save()
        {
			RegistryKey rk = Registry.CurrentUser.OpenSubKey(m_sRoot, true);
	
			if (rk == null)
				{
				rk = Registry.CurrentUser.CreateSubKey(m_sRoot);
				if (rk == null)
					return;
				}
	
			foreach (ReHistElt rehe in m_rgrehe)
				{
				object oVal;
	
				if (rehe.oref is System.Windows.Forms.TextBox)
					{
					oVal = ((System.Windows.Forms.TextBox)rehe.oref).Text;
					}
				else if (rehe.oref is System.Windows.Forms.ComboBox)
					{
					oVal = ((System.Windows.Forms.ComboBox)rehe.oref).Text;
					}
				else if (rehe.oref is System.Windows.Forms.ListBox)
					{
					oVal = ((System.Windows.Forms.ListBox)rehe.oref).Text;
					}
				else if (rehe.oref is System.Windows.Forms.CheckBox)
					{
					oVal = ((System.Windows.Forms.CheckBox)rehe.oref).Checked;
					}
				else if (rehe.oref is System.Windows.Forms.DateTimePicker)
					{
					oVal = ((System.Windows.Forms.DateTimePicker)rehe.oref).Value;
					}
				else
					{
					throw(new Exception("Unkonwn control type in ReHistory.Save"));
					}
	
				int nT;
				
				switch (rehe.type)
					{
					case Type.Dttm:
						DateTime dttm = (DateTime)oVal;
						rk.SetValue(rehe.sRePath, dttm.ToString());
						break;
					case Type.Str:
						string sT = (string)oVal;
						rk.SetValue(rehe.sRePath, sT);
						break;
					case Type.Bool:
						nT = ((bool)oVal) ? 1 : 0;
						rk.SetValue(rehe.sRePath, nT, RegistryValueKind.DWord);
						break;
					case Type.Int:
						nT = Int32.Parse((string)oVal);
						rk.SetValue(rehe.sRePath, nT, RegistryValueKind.DWord);
						break;
					}
				}
        }
     
		public void Load()
		{
			RegistryKey rk = Registry.CurrentUser.OpenSubKey(m_sRoot);
			string sVal = "";
			DateTime dttmVal = DateTime.Today;
			int nVal = 0;
			bool fVal = false;
            int nT;
            string sT;
            
			if (rk == null)
				return;

			foreach (ReHistElt rehe in m_rgrehe)
				{
				switch (rehe.type)
					{
					case Type.Dttm:
						sT = (string)rk.GetValue(rehe.sRePath, rehe.oDefault);
						sVal = sT;
						try
						{
							dttmVal = DateTime.Parse(sT);
						} catch
						{
						}
						break;
					case Type.Str:
						sT = (string)rk.GetValue(rehe.sRePath, rehe.oDefault);
						sVal = sT;
						try
						{
						nVal = Int32.Parse(sT);
						} catch
						{
						nVal = 0;
						}
						break;
					case Type.Bool:
						nT = (int)rk.GetValue(rehe.sRePath, rehe.oDefault);
						fVal = (nT != 0 ? true : false);
						break;
					case Type.Int:
                        nT = (int)rk.GetValue(rehe.sRePath, rehe.oDefault);
						sVal = nT.ToString();
						nVal = nT;
						break;
					}

				if (rehe.oref is System.Windows.Forms.TextBox)
					{
					((System.Windows.Forms.TextBox)rehe.oref).Text = sVal;
					}
				else if (rehe.oref is System.Windows.Forms.ListBox)
					{
					((System.Windows.Forms.ListBox)rehe.oref).Text = sVal;
					}
				else if (rehe.oref is System.Windows.Forms.ComboBox)
					{
					((System.Windows.Forms.ComboBox)rehe.oref).Text = sVal;
					}
				else if (rehe.oref is System.Windows.Forms.CheckBox)
					{
					((System.Windows.Forms.CheckBox)rehe.oref).Checked = fVal;
					}
				else if (rehe.oref is System.Windows.Forms.DateTimePicker)
					{
					((System.Windows.Forms.DateTimePicker)rehe.oref).Value = dttmVal;
					}
				else
					{
					throw(new Exception("Unkonwn control type in ReHistory.Save"));
					}

				}
		}
	}

}
