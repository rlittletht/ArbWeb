using System;
using System.Diagnostics;
using System.Drawing;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using StatusBox;
using mshtml;
using System.Runtime.InteropServices;
using Outlook=Microsoft.Office.Interop.Outlook;
using System.Threading.Tasks;
using TCore.Settings;

namespace ArbWeb
{
	/// <summary>
	/// Summary description for AwMainForm.
	/// </summary>
	public partial class AwMainForm : System.Windows.Forms.Form
    {
        #region ArbiterStrings

        // Top level web pages
		private const string _s_Home = "https://www.arbitersports.com/Shared/SignIn/Signin.aspx"; // ok2010
		private const string _s_Assigning = "https://www.arbitersports.com/Assigner/Games/NewGamesView.aspx"; // ok2010
		private const string _s_RanksEdit = "https://www.arbitersports.com/Assigner/RanksEdit.aspx"; // ok2010
		private const string _s_AddUser = "https://www.arbitersports.com/Assigner/UserAdd.aspx?userTypeID=3"; // ok2010u
		private const string _s_OfficialsView = "https://www.arbitersports.com/Assigner/OfficialsView.aspx"; // ok2010u
	    private const string _s_Announcements = "https://www.arbitersports.com/Shared/AnnouncementsEdit.aspx"; // ok2015

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
		private const string _s_Home_Input_Email = "ctl00$EmailTextbox"; // ctl00$ucMiniLogin$UsernameTextBox"; // ok2015
		private const string _s_Home_Input_Password = "ctl00$PasswordTextbox"; // ctl00$ucMiniLogin$PasswordTextBox"; // ok2015
		private const string _s_Home_Button_SignIn = "ctl00$SignInButton"; // ctl00$ucMiniLogin$SignInButton"; // ok2015

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

        // Announcements page
	    private const string _s_Announcements_Button_Edit_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$";   // the control will be "ctl##"
	    private const string _s_Announcements_Button_Edit_Suffix = "$btnEdit";
	    private const string _s_Announcements_Checkbox_Assigners_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$"; // the control will be "ctl##"
        private const string _s_Announcements_Checkbox_Assigners_Suffix = "$chkToAssigners_Edit";
	    private const string _s_Announcements_Checkbox_Contacts_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$"; // the control will be "ctl##"
        private const string _s_Announcements_Checkbox_Contacts_Suffix = "$chkToContacts_Edit";
	    private const string _s_Announcements_Select_Filters_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$"; // the control will be "ctl##"
        private const string _s_Announcements_Select_Filters_Suffix = "$ddlFilters";
	    private const string _s_Announcements_Button_Save_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$"; // the control will be "ctl##"
        private const string _s_Announcements_Button_Save_Suffix = "$btnSave";

	    private const string _s_Announcements_Textarea_Text_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$";
	    private const string _s_Announcements_Textarea_Text_Suffix = "$txtAnnouncement";

        private const string _sid_Announcements_Button_Edit_Prefix = "ctl00_ContentHolder_pgeAnnouncementsEdit_conAnnouncementsEdit_dgAnnouncements_";
	    private const string _sid_Announcements_Button_Edit_Suffix = "_btnEdit";

	    private const string _sid_Announcements_Button_Save_Prefix = "ctl00_ContentHolder_pgeAnnouncementsEdit_conAnnouncementsEdit_dgAnnouncements_";
	    private const string _sid_Announcements_Button_Save_Suffix = "_btnSave";

#endregion

        #region Controls
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
        private TextBox m_ebRosterWorking;
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
        private Button m_pbCreateRosterReport;
        private CheckBox m_cbFilterRank;
        private Button m_pbMailMerge;
        private CheckBox m_cbFutureOnly;
        private Label label19;
        private CheckBox m_cbLaunch;
        private CheckBox m_cbSetArbiterAnnounce;

		private StatusBox.StatusRpt m_srpt;
        #endregion

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

	    private Settings.SettingsElt[] m_rgreheProfile;
        private Settings.SettingsElt[] m_rgrehe;
		// ReHistory.ReHistElt[] m_rgreheProfile;
		// ReHistory.ReHistElt[] m_rgrehe;
	    private Settings m_rehProfile;
	    private Settings m_reh;

//		ReHistory m_rehProfile;
//		ReHistory m_reh;
		ArbWebControl m_awc;
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
			m_awc = new ArbWebControl(m_srpt);
			m_fDontUpdateProfile = true;
			RegistryKey rk = Registry.CurrentUser.OpenSubKey("Software\\Thetasoft\\ArbWeb");

            if (rk == null)
                rk = Registry.CurrentUser.CreateSubKey("Software\\Thetasoft\\ArbWeb");
                
			string []rgs = rk.GetSubKeyNames();

			foreach(string s in rgs)
				{
				m_cbxProfile.Items.Add(s);
				}
			m_rgreheProfile = new[] 
				{ 
					new Settings.SettingsElt("Login", Settings.Type.Str, m_ebUserID, ""), 
					new Settings.SettingsElt("Password", Settings.Type.Str, m_ebPassword, ""), 
					new Settings.SettingsElt("GameFile", Settings.Type.Str, m_ebGameFile, ""), 
					new Settings.SettingsElt("Roster", Settings.Type.Str, m_ebRoster, ""),
					new Settings.SettingsElt("GameFileCopy", Settings.Type.Str, m_ebGameCopy, ""), 
					new Settings.SettingsElt("RosterCopy", Settings.Type.Str, m_ebRosterWorking, ""),
					new Settings.SettingsElt("GameOutput", Settings.Type.Str, m_ebGameOutput, ""),
					new Settings.SettingsElt("OutputFile", Settings.Type.Str, m_ebOutputFile, ""),
					new Settings.SettingsElt("IncludeCanceled", Settings.Type.Bool, m_cbIncludeCanceled, 0),
					new Settings.SettingsElt("ShowBrowser", Settings.Type.Bool, m_cbShowBrowser, 0), 
					new Settings.SettingsElt("LastSlotStartDate", Settings.Type.Dttm, m_dtpStart, ""),
					new Settings.SettingsElt("LastSlotEndDate", Settings.Type.Dttm, m_dtpEnd, ""),
					new Settings.SettingsElt("LastOpenSlotDetail", Settings.Type.Bool, m_cbOpenSlotDetail, 0),
					new Settings.SettingsElt("LastGroupTimeSlots", Settings.Type.Bool, m_cbFuzzyTimes, 0),
					new Settings.SettingsElt("LastTestEmail", Settings.Type.Bool, m_cbTestEmail, 0),
					new Settings.SettingsElt("AddOfficialsOnly", Settings.Type.Bool, m_cbAddOfficialsOnly, 0),
					new Settings.SettingsElt("AfiliationIndex", Settings.Type.Int, m_ebAffiliationIndex, 0),
					new Settings.SettingsElt("LastSplitSports", Settings.Type.Bool, m_cbSplitSports, 0),
					new Settings.SettingsElt("LastPivotDate", Settings.Type.Bool, m_cbDatePivot, 0),
					new Settings.SettingsElt("LastLogToFile", Settings.Type.Bool, m_cbLogToFile, 0),
					new Settings.SettingsElt("FilterMailMergeByRank", Settings.Type.Bool, m_cbFilterRank, 0),
					new Settings.SettingsElt("DownloadOnlyFutureGames", Settings.Type.Bool, m_cbFutureOnly, 0),
                    new Settings.SettingsElt("LaunchMailMergeDoc", Settings.Type.Bool, m_cbLaunch, 0),
                    new Settings.SettingsElt("SetArbiterAnnouncement", Settings.Type.Bool, m_cbSetArbiterAnnounce, 0),
				};

			m_rgrehe = new []
				{
					new Settings.SettingsElt("LastProfile", Settings.Type.Str, m_cbxProfile, "")
				};

            SetupLogToFile();

			m_reh = new Settings(m_rgrehe, "Software\\Thetasoft\\ArbWeb", "root");
			m_reh.Load();


			m_rehProfile = new Settings(m_rgreheProfile, String.Format("Software\\Thetasoft\\ArbWeb\\{0}", m_cbxProfile.Text), m_cbxProfile.Text);

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
            this.m_ebRosterWorking = new System.Windows.Forms.TextBox();
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
            this.m_pbCreateRosterReport = new System.Windows.Forms.Button();
            this.m_cbFilterRank = new System.Windows.Forms.CheckBox();
            this.m_pbMailMerge = new System.Windows.Forms.Button();
            this.m_cbFutureOnly = new System.Windows.Forms.CheckBox();
            this.label19 = new System.Windows.Forms.Label();
            this.m_cbLaunch = new System.Windows.Forms.CheckBox();
            this.m_cbSetArbiterAnnounce = new System.Windows.Forms.CheckBox();
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
            this.m_pbDownloadGames.Location = new System.Drawing.Point(589, 28);
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
            this.button1.Location = new System.Drawing.Point(589, 52);
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
            this.groupBox2.Location = new System.Drawing.Point(8, 637);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(691, 157);
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
            this.m_recStatus.Size = new System.Drawing.Size(679, 132);
            this.m_recStatus.TabIndex = 0;
            this.m_recStatus.Text = "";
            // 
            // m_pbGenCounts
            // 
            this.m_pbGenCounts.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbGenCounts.Location = new System.Drawing.Point(589, 215);
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
            this.m_pbUploadRoster.Location = new System.Drawing.Point(476, 607);
            this.m_pbUploadRoster.Name = "m_pbUploadRoster";
            this.m_pbUploadRoster.Size = new System.Drawing.Size(110, 24);
            this.m_pbUploadRoster.TabIndex = 32;
            this.m_pbUploadRoster.Text = "Upload Roster";
            this.m_pbUploadRoster.Click += new System.EventHandler(this.DoUploadRoster);
            // 
            // m_pbOpenSlots
            // 
            this.m_pbOpenSlots.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbOpenSlots.Location = new System.Drawing.Point(589, 264);
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
            this.m_lblSearchCriteria.Size = new System.Drawing.Size(688, 16);
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
            this.label6.Size = new System.Drawing.Size(688, 19);
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
            this.label7.Size = new System.Drawing.Size(688, 19);
            this.label7.TabIndex = 37;
            this.label7.Tag = "Open Slot Reporting";
            this.label7.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(12, 268);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(55, 13);
            this.label8.TabIndex = 40;
            this.label8.Text = "Start Date";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(298, 268);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(52, 13);
            this.label9.TabIndex = 41;
            this.label9.Text = "End Date";
            // 
            // m_dtpStart
            // 
            this.m_dtpStart.CustomFormat = "ddd MMM dd, yyyy";
            this.m_dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.m_dtpStart.Location = new System.Drawing.Point(76, 264);
            this.m_dtpStart.Name = "m_dtpStart";
            this.m_dtpStart.Size = new System.Drawing.Size(154, 20);
            this.m_dtpStart.TabIndex = 42;
            this.m_dtpStart.Value = new System.DateTime(2008, 5, 4, 0, 0, 0, 0);
            // 
            // m_dtpEnd
            // 
            this.m_dtpEnd.CustomFormat = "ddd MMM dd, yyyy";
            this.m_dtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.m_dtpEnd.Location = new System.Drawing.Point(356, 264);
            this.m_dtpEnd.Name = "m_dtpEnd";
            this.m_dtpEnd.Size = new System.Drawing.Size(154, 20);
            this.m_dtpEnd.TabIndex = 43;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(41, 434);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(79, 13);
            this.label10.TabIndex = 45;
            this.label10.Text = "Misc Field Filter";
            // 
            // m_ebFilter
            // 
            this.m_ebFilter.Location = new System.Drawing.Point(126, 431);
            this.m_ebFilter.Name = "m_ebFilter";
            this.m_ebFilter.Size = new System.Drawing.Size(129, 20);
            this.m_ebFilter.TabIndex = 44;
            // 
            // m_cbOpenSlotDetail
            // 
            this.m_cbOpenSlotDetail.AutoSize = true;
            this.m_cbOpenSlotDetail.Location = new System.Drawing.Point(421, 313);
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
            this.button2.Location = new System.Drawing.Point(576, 354);
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
            this.label12.Location = new System.Drawing.Point(31, 294);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(662, 19);
            this.label12.TabIndex = 50;
            this.label12.Tag = "Slot reporting";
            this.label12.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // m_cbFilterSport
            // 
            this.m_cbFilterSport.AutoSize = true;
            this.m_cbFilterSport.Location = new System.Drawing.Point(34, 313);
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
            this.m_chlbxSports.Location = new System.Drawing.Point(44, 333);
            this.m_chlbxSports.Name = "m_chlbxSports";
            this.m_chlbxSports.Size = new System.Drawing.Size(164, 94);
            this.m_chlbxSports.TabIndex = 52;
            this.m_chlbxSports.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.DoSportLevelFilter);
            // 
            // m_chlbxSportLevels
            // 
            this.m_chlbxSportLevels.CheckOnClick = true;
            this.m_chlbxSportLevels.FormattingEnabled = true;
            this.m_chlbxSportLevels.Location = new System.Drawing.Point(243, 333);
            this.m_chlbxSportLevels.Name = "m_chlbxSportLevels";
            this.m_chlbxSportLevels.Size = new System.Drawing.Size(164, 94);
            this.m_chlbxSportLevels.TabIndex = 54;
            // 
            // m_cbFilterLevel
            // 
            this.m_cbFilterLevel.AutoSize = true;
            this.m_cbFilterLevel.Location = new System.Drawing.Point(233, 313);
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
            this.m_cbTestEmail.Location = new System.Drawing.Point(576, 313);
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
            // m_ebRosterWorking
            // 
            this.m_ebRosterWorking.Location = new System.Drawing.Point(379, 82);
            this.m_ebRosterWorking.Name = "m_ebRosterWorking";
            this.m_ebRosterWorking.Size = new System.Drawing.Size(175, 20);
            this.m_ebRosterWorking.TabIndex = 57;
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
            this.m_cbTestOnly.Location = new System.Drawing.Point(227, 106);
            this.m_cbTestOnly.Name = "m_cbTestOnly";
            this.m_cbTestOnly.Size = new System.Drawing.Size(130, 17);
            this.m_cbTestOnly.TabIndex = 60;
            this.m_cbTestOnly.Text = "TEST Only (1 op only)";
            this.m_cbTestOnly.UseVisualStyleBackColor = true;
            // 
            // m_pbReload
            // 
            this.m_pbReload.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbReload.Location = new System.Drawing.Point(589, 607);
            this.m_pbReload.Name = "m_pbReload";
            this.m_pbReload.Size = new System.Drawing.Size(110, 24);
            this.m_pbReload.TabIndex = 61;
            this.m_pbReload.Text = "Load Data";
            this.m_pbReload.Click += new System.EventHandler(this.m_pbReload_Click);
            // 
            // m_cbRankOnly
            // 
            this.m_cbRankOnly.AutoSize = true;
            this.m_cbRankOnly.Location = new System.Drawing.Point(363, 106);
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
            this.label15.Size = new System.Drawing.Size(688, 19);
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
            this.m_pbGenGames.Location = new System.Drawing.Point(589, 163);
            this.m_pbGenGames.Name = "m_pbGenGames";
            this.m_pbGenGames.Size = new System.Drawing.Size(110, 27);
            this.m_pbGenGames.TabIndex = 66;
            this.m_pbGenGames.Text = "Games Report";
            this.m_pbGenGames.Click += new System.EventHandler(this.DoGamesReport);
            // 
            // m_cbAddOfficialsOnly
            // 
            this.m_cbAddOfficialsOnly.AutoSize = true;
            this.m_cbAddOfficialsOnly.Location = new System.Drawing.Point(502, 106);
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
            this.m_pbBrowseGameFile.Location = new System.Drawing.Point(257, 61);
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
            this.m_pbBrowseGameFile2.Location = new System.Drawing.Point(258, 83);
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
            this.m_pbBrowseRoster.Location = new System.Drawing.Point(527, 64);
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
            this.m_pbBrowseRoster2.Location = new System.Drawing.Point(527, 87);
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
            this.m_pbBrowseGamesReport.Location = new System.Drawing.Point(258, 169);
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
            this.m_pbBrowseAnalysis.Location = new System.Drawing.Point(258, 215);
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
            this.button3.Location = new System.Drawing.Point(589, 76);
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
            this.m_cbFuzzyTimes.Location = new System.Drawing.Point(421, 331);
            this.m_cbFuzzyTimes.Name = "m_cbFuzzyTimes";
            this.m_cbFuzzyTimes.Size = new System.Drawing.Size(107, 17);
            this.m_cbFuzzyTimes.TabIndex = 78;
            this.m_cbFuzzyTimes.Text = "Group Time Slots";
            this.m_cbFuzzyTimes.UseVisualStyleBackColor = true;
            // 
            // m_cbDatePivot
            // 
            this.m_cbDatePivot.AutoSize = true;
            this.m_cbDatePivot.Location = new System.Drawing.Point(576, 331);
            this.m_cbDatePivot.Name = "m_cbDatePivot";
            this.m_cbDatePivot.Size = new System.Drawing.Size(91, 17);
            this.m_cbDatePivot.TabIndex = 79;
            this.m_cbDatePivot.Text = "Pivot on Date";
            this.m_cbDatePivot.UseVisualStyleBackColor = true;
            // 
            // m_cbSplitSports
            // 
            this.m_cbSplitSports.AutoSize = true;
            this.m_cbSplitSports.Location = new System.Drawing.Point(421, 351);
            this.m_cbSplitSports.Name = "m_cbSplitSports";
            this.m_cbSplitSports.Size = new System.Drawing.Size(156, 17);
            this.m_cbSplitSports.TabIndex = 80;
            this.m_cbSplitSports.Text = "Automatic Softball/Baseball";
            this.m_cbSplitSports.UseVisualStyleBackColor = true;
            // 
            // m_cbLogToFile
            // 
            this.m_cbLogToFile.AutoSize = true;
            this.m_cbLogToFile.Location = new System.Drawing.Point(654, 106);
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
            this.label18.Location = new System.Drawing.Point(31, 490);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(662, 19);
            this.label18.TabIndex = 82;
            this.label18.Tag = "Site roster report";
            this.label18.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // m_chlbxRoster
            // 
            this.m_chlbxRoster.CheckOnClick = true;
            this.m_chlbxRoster.FormattingEnabled = true;
            this.m_chlbxRoster.Location = new System.Drawing.Point(44, 508);
            this.m_chlbxRoster.Name = "m_chlbxRoster";
            this.m_chlbxRoster.Size = new System.Drawing.Size(363, 64);
            this.m_chlbxRoster.TabIndex = 83;
            // 
            // m_pbCreateRosterReport
            // 
            this.m_pbCreateRosterReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbCreateRosterReport.Location = new System.Drawing.Point(589, 542);
            this.m_pbCreateRosterReport.Name = "m_pbCreateRosterReport";
            this.m_pbCreateRosterReport.Size = new System.Drawing.Size(110, 27);
            this.m_pbCreateRosterReport.TabIndex = 84;
            this.m_pbCreateRosterReport.Text = "Create Roster";
            this.m_pbCreateRosterReport.Click += new System.EventHandler(this.GenSiteRosterReport);
            // 
            // m_cbFilterRank
            // 
            this.m_cbFilterRank.AutoSize = true;
            this.m_cbFilterRank.Location = new System.Drawing.Point(424, 430);
            this.m_cbFilterRank.Name = "m_cbFilterRank";
            this.m_cbFilterRank.Size = new System.Drawing.Size(86, 17);
            this.m_cbFilterRank.TabIndex = 85;
            this.m_cbFilterRank.Text = "Filter by rank";
            this.m_cbFilterRank.UseVisualStyleBackColor = true;
            // 
            // m_pbMailMerge
            // 
            this.m_pbMailMerge.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbMailMerge.Location = new System.Drawing.Point(583, 464);
            this.m_pbMailMerge.Name = "m_pbMailMerge";
            this.m_pbMailMerge.Size = new System.Drawing.Size(110, 27);
            this.m_pbMailMerge.TabIndex = 86;
            this.m_pbMailMerge.Text = "Gen Help";
            this.m_pbMailMerge.Click += new System.EventHandler(this.GenMailMergeMail);
            // 
            // m_cbFutureOnly
            // 
            this.m_cbFutureOnly.AutoSize = true;
            this.m_cbFutureOnly.Location = new System.Drawing.Point(129, 106);
            this.m_cbFutureOnly.Name = "m_cbFutureOnly";
            this.m_cbFutureOnly.Size = new System.Drawing.Size(92, 17);
            this.m_cbFutureOnly.TabIndex = 87;
            this.m_cbFutureOnly.Text = "Future Games";
            this.m_cbFutureOnly.UseVisualStyleBackColor = true;
            // 
            // label19
            // 
            this.label19.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label19.Location = new System.Drawing.Point(31, 585);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(662, 19);
            this.label19.TabIndex = 88;
            this.label19.Tag = "Data/Upload Operations";
            // 
            // m_cbLaunch
            // 
            this.m_cbLaunch.AutoSize = true;
            this.m_cbLaunch.Location = new System.Drawing.Point(424, 451);
            this.m_cbLaunch.Name = "m_cbLaunch";
            this.m_cbLaunch.Size = new System.Drawing.Size(103, 17);
            this.m_cbLaunch.TabIndex = 89;
            this.m_cbLaunch.Text = "Launch MMDoc";
            this.m_cbLaunch.UseVisualStyleBackColor = true;
            // 
            // m_cbSetArbiterAnnounce
            // 
            this.m_cbSetArbiterAnnounce.AutoSize = true;
            this.m_cbSetArbiterAnnounce.Location = new System.Drawing.Point(527, 430);
            this.m_cbSetArbiterAnnounce.Name = "m_cbSetArbiterAnnounce";
            this.m_cbSetArbiterAnnounce.Size = new System.Drawing.Size(127, 17);
            this.m_cbSetArbiterAnnounce.TabIndex = 90;
            this.m_cbSetArbiterAnnounce.Text = "Set Arbiter Announce";
            this.m_cbSetArbiterAnnounce.UseVisualStyleBackColor = true;
            // 
            // AwMainForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(707, 806);
            this.Controls.Add(this.m_cbSetArbiterAnnounce);
            this.Controls.Add(this.m_cbLaunch);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.m_cbFutureOnly);
            this.Controls.Add(this.m_pbMailMerge);
            this.Controls.Add(this.m_cbFilterRank);
            this.Controls.Add(this.m_pbCreateRosterReport);
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
            this.Controls.Add(this.m_ebRosterWorking);
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
		
        bool m_fLoggedIn;

		Roster m_rst;
		CountsData m_gc;

        /* E N S U R E  L O G G E D  I N */
        /*----------------------------------------------------------------------------
        	%%Function: EnsureLoggedIn
        	%%Qualified: ArbWeb.AwMainForm.EnsureLoggedIn
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
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
                if (!ArbWebControl.FCheckForControl(oDoc2, _sid_Home_Anchor_ActingLink))
                    {
                    IHTMLDocument oDoc = m_awc.Document;
                    IHTMLDocument3 oDoc3 = m_awc.Document3;

					ArbWebControl.FSetInputControlText(oDoc2, _s_Home_Input_Email, m_ebUserID.Text, false);
                    ArbWebControl.FSetInputControlText(oDoc2, _s_Home_Input_Password, m_ebPassword.Text, false);

                    m_awc.ResetNav();
                    m_awc.FClickControl(oDoc2, _s_Home_Button_SignIn);
                    m_awc.FWaitForNavFinish();
                    }
                oDoc2 = m_awc.Document2;
                
                int count = 0;

                oDoc2 = m_awc.Document2;
                bool fToggledBrowser = false;
                
                while (count < 100 && (ArbWebControl.FCheckForControl(oDoc2, _sid_Home_Div_PnlAccounts) || !ArbWebControl.FCheckForControl(oDoc2, _sid_Home_Anchor_ActingLink)))
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
                if (!ArbWebControl.FCheckForControl(oDoc2, _sid_Home_Anchor_ActingLink))
                    MessageBox.Show("Login failed for arbiter.net!");
                else
                    m_fLoggedIn = true;
				m_srpt.PopLevel();
				m_srpt.AddMessage("Completed login.");
                }
        }
        
        /* S  B U I L D  T E M P  F I L E N A M E */
        /*----------------------------------------------------------------------------
        	%%Function: SBuildTempFilename
        	%%Qualified: ArbWeb.AwMainForm.SBuildTempFilename
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private string SBuildTempFilename(string sBaseName, string sExt)
        {
            return String.Format("{0}\\{1}{2}.{3}", Environment.GetEnvironmentVariable("Temp"),
                                 sBaseName,
                                 System.Guid.NewGuid().ToString(),
                                 sExt);
        }


	    /* D O  S A V E  S T A T E  C O R E */
	    /*----------------------------------------------------------------------------
	    	%%Function: DoSaveStateCore
	    	%%Qualified: ArbWeb.AwMainForm.DoSaveStateCore
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
	    private void DoSaveStateCore()
		{
			m_rehProfile.Save();
			m_reh.Save();
		}

		/* D O  S A V E  S T A T E */
		/*----------------------------------------------------------------------------
			%%Function: DoSaveState
			%%Qualified: ArbWeb.AwMainForm.DoSaveState
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
		private void DoSaveState(object sender, FormClosingEventArgs e)
        {
			DoSaveStateCore();
        }

        /* D O  D O W N L O A D  G A M E S */
        /*----------------------------------------------------------------------------
        	%%Function: DoDownloadGames
        	%%Qualified: ArbWeb.AwMainForm.DoDownloadGames
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoDownloadGames(object sender, EventArgs e)
        {
            DownloadGames();
        }

        List<Cursor> m_plCursor;
        
        /* P U S H  C U R S O R */
        /*----------------------------------------------------------------------------
        	%%Function: PushCursor
        	%%Qualified: ArbWeb.AwMainForm.PushCursor
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void PushCursor(Cursor crs)
        {
            m_plCursor.Add(this.Cursor);
            this.Cursor = crs;
        }
        
        /* P O P  C U R S O R */
        /*----------------------------------------------------------------------------
        	%%Function: PopCursor
        	%%Qualified: ArbWeb.AwMainForm.PopCursor
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void PopCursor()
        {
            if (m_plCursor.Count > 0)
                {
                this.Cursor = m_plCursor[m_plCursor.Count - 1];
                m_plCursor.RemoveAt(m_plCursor.Count - 1);
                }
        }


	    /* D O  D O W N L O A D  R O S T E R */
	    /*----------------------------------------------------------------------------
	    	%%Function: DoDownloadRoster
	    	%%Qualified: ArbWeb.AwMainForm.DoDownloadRoster
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
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
			System.IO.File.Delete(m_ebRosterWorking.Text);
			System.IO.File.Copy(sOutFile, m_ebRosterWorking.Text);
			m_srpt.AddMessage("Completed FULL Roster download.");
        }


		/* D O  U P L O A D  R O S T E R */
		/*----------------------------------------------------------------------------
			%%Function: DoUploadRoster
			%%Qualified: ArbWeb.AwMainForm.DoUploadRoster
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
		private void DoUploadRoster(object sender, EventArgs e)
		{
			m_srpt.AddMessage("Starting Roster upload...");
			m_srpt.PushLevel();

			string sInFile = m_ebRosterWorking.Text;

			Roster rst = RstEnsure(sInFile);

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

        /* D O  G E N  C O U N T S */
        /*----------------------------------------------------------------------------
        	%%Function: DoGenCounts
        	%%Qualified: ArbWeb.AwMainForm.DoGenCounts
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoGenCounts(object sender, EventArgs e)
		{
			m_srpt.AddMessage(String.Format("Generating analysis ({0})...", m_ebOutputFile.Text));
			m_srpt.PushLevel();

			CountsData gc = GcEnsure(m_ebRosterWorking.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);

			gc.GenAnalysis(m_ebOutputFile.Text);
			m_srpt.PopLevel();
			m_srpt.AddMessage("Analysis complete.");
		}

        /* S  H T M L  R E A D  F I L E */
        /*----------------------------------------------------------------------------
        	%%Function: SHtmlReadFile
        	%%Qualified: ArbWeb.AwMainForm.SHtmlReadFile
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
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
            
        /* C H A N G E  S H O W  B R O W S E R */
        /*----------------------------------------------------------------------------
        	%%Function: ChangeShowBrowser
        	%%Qualified: ArbWeb.AwMainForm.ChangeShowBrowser
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void ChangeShowBrowser(object sender, EventArgs e)
        {
            if (m_cbShowBrowser.Checked)
                m_awc.Show();
            else
                m_awc.Hide();
        }

        /* D O  G E N E R I C  I N V A L  G C */
        /*----------------------------------------------------------------------------
        	%%Function: DoGenericInvalGc
        	%%Qualified: ArbWeb.AwMainForm.DoGenericInvalGc
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoGenericInvalGc(object sender, EventArgs e)
        {
            InvalGameCount();
        }

        /* C H A N G E  P R O F I L E */
        /*----------------------------------------------------------------------------
        	%%Function: ChangeProfile
        	%%Qualified: ArbWeb.AwMainForm.ChangeProfile
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void ChangeProfile(object sender, EventArgs e)
		{
			if (!m_fDontUpdateProfile)
				{
				DoSaveStateCore();
				m_rehProfile = new Settings(m_rgreheProfile, String.Format("Software\\Thetasoft\\ArbWeb\\{0}", m_cbxProfile.Text), m_cbxProfile.Text);
				m_rehProfile.Load();
				}
		}


        /* O N  P R O F I L E  L E A V E */
        /*----------------------------------------------------------------------------
        	%%Function: OnProfileLeave
        	%%Qualified: ArbWeb.AwMainForm.OnProfileLeave
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void OnProfileLeave(object sender, EventArgs e)
        {
            if (m_fDontUpdateProfile)
                return;
               
			if (m_rehProfile.FMatchesTag(m_cbxProfile.Text))
				return;

			// otherwise, this is a new profile
			ChangeProfile(sender, e);
        }

		/* E N A B L E  C O N T R O L S */
		/*----------------------------------------------------------------------------
			%%Function: EnableControls
			%%Qualified: ArbWeb.AwMainForm.EnableControls
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
		void EnableControls()
		{
			m_chlbxSports.Enabled = m_cbFilterSport.Checked;
			m_chlbxSportLevels.Enabled = m_cbFilterLevel.Checked;
			m_cbFilterLevel.Enabled = m_cbFilterSport.Checked;
			
			EnableAdminFunctions();
			
		}

        /* H A N D L E  S P O R T  C H E C K E D */
        /*----------------------------------------------------------------------------
        	%%Function: HandleSportChecked
        	%%Qualified: ArbWeb.AwMainForm.HandleSportChecked
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void HandleSportChecked(object sender, EventArgs e)
        {
			EnableControls();
        }

        /* H A N D L E  S P O R T  L E V E L  C H E C K E D */
        /*----------------------------------------------------------------------------
        	%%Function: HandleSportLevelChecked
        	%%Qualified: ArbWeb.AwMainForm.HandleSportLevelChecked
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void HandleSportLevelChecked(object sender, EventArgs e)
        {
			EnableControls();
        }

        /* H A N D L E  S L O T  D E T A I L  C H E C K E D */
        /*----------------------------------------------------------------------------
        	%%Function: HandleSlotDetailChecked
        	%%Qualified: ArbWeb.AwMainForm.HandleSlotDetailChecked
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void HandleSlotDetailChecked(object sender, EventArgs e)
        {

        }


	    /* G E N  S I T E  R O S T E R  R E P O R T */
	    /*----------------------------------------------------------------------------
	    	%%Function: GenSiteRosterReport
	    	%%Qualified: ArbWeb.AwMainForm.GenSiteRosterReport
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
	    private void GenSiteRosterReport(object sender, EventArgs e)
	    {
	        CountsData gc = GcEnsure(m_ebRosterWorking.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
	        string sTempFile = String.Format("{0}\\temp{1}.doc", Environment.GetEnvironmentVariable("Temp"),
	                                         System.Guid.NewGuid().ToString());
	        Roster rst = RstEnsure(m_ebRosterWorking.Text);

	        gc.GenSiteRosterResport(sTempFile, rst, ArbWebControl.RgsFromChlbx(true, m_chlbxRoster), m_dtpStart.Value, m_dtpEnd.Value);
            // launch word with the file
	        Process.Start(sTempFile);
	        // System.IO.File.Delete(sTempFile);
	    }

        static string BuildAnnName(string sPrefix, string sSuffix, string sCtl)
        {
            return String.Format("{0}{1}{2}", sPrefix, sCtl, sSuffix);
        }
	    /* G E N  M A I L  M E R G E  M A I L */
	    /*----------------------------------------------------------------------------
	    	%%Function: GenMailMergeMail
	    	%%Qualified: ArbWeb.AwMainForm.GenMailMergeMail
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
	    private void GenMailMergeMail(object sender, EventArgs e)
	    {
	        CountsData gc = GcEnsure(m_ebRosterWorking.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
	        Roster rst = RstEnsure(m_ebRosterWorking.Text);
	        m_srpt.AddMessage("Generating mail merge documents...", StatusRpt.MSGT.Header, false);

	        // first, generate the mailmerge source csv file.  this is either the entire roster, or just the folks 
	        // rated for the sports we are filtered to
	        GameData.GameSlots gms = gc.GamesFromFilter(ArbWebControl.RgsFromChlbx(m_cbFilterSport.Checked, m_chlbxSports),
	                                                    ArbWebControl.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), false, m_saOpenSlots);

	        Roster rstFiltered;

	        if (m_cbFilterRank.Checked)
	            rstFiltered = rst.FilterByRanks(gms.RequiredRanks());
	        else
	            rstFiltered = rst;

	        string sCsvTemp = SBuildTempFilename("MailMergeRoster", "csv");
	        StreamWriter sw = new StreamWriter(sCsvTemp, false, System.Text.Encoding.Default);

	        sw.WriteLine("email,firstname,lastname");
	        foreach (RosterEntry rste in rstFiltered.Plrste)
	            {
	            sw.WriteLine("{0},{1},{2}", rste.Email, rste.First, rste.m_sLast);
	            }
	        sw.Flush();
	        sw.Close();

	        // ok, now create the mailmerge .docx
	        string sTempName;
	        string sArbiterHelpNeeded;

	        OOXML.CreateMailMergeDoc("mailmergedoc.docx", sTempName = SBuildTempFilename("mailmergedoc", "docx"), sCsvTemp, gms, out sArbiterHelpNeeded);

	        System.Windows.Forms.Clipboard.SetText(sArbiterHelpNeeded);
	        if (m_cbLaunch.Checked)
	            {
	            m_srpt.AddMessage("Done, launching document...", StatusRpt.MSGT.Header, false);
	            System.Diagnostics.Process.Start(sTempName);
	            }
	        if (m_cbSetArbiterAnnounce.Checked)
	            SetArbiterAnnounce(sArbiterHelpNeeded);
	    }

        void SetArbiterAnnounce(string sArbiterHelpNeeded)
        {
			m_srpt.AddMessage("Starting Announcement Set...");
			m_srpt.PushLevel();

            EnsureLoggedIn();
            ThrowIfNot(m_awc.FNavToPage(_s_Announcements), "Couldn't nav to announcements page!");
            m_awc.FWaitForNavFinish();

            // now we need to find the URGENT HELP NEEDED row
            IHTMLDocument2 oDoc2 = m_awc.Document2;
            IHTMLElementCollection hec = (IHTMLElementCollection)oDoc2.all.tags("div");

            string sCtl = null;

            foreach (IHTMLElement he in hec)
                {
                if (he.id == "D9UrgentHelpNeeded")
                    {
                    IHTMLElement heFind = he;
                    while (heFind.tagName.ToLower() != "tr")
                        {
                        heFind = heFind.parentElement;
                        ThrowIfNot(heFind != null, "Can't find HELP announcement");
                        }
                    // ok, go up to the parent TR.
                    // now find one of our controls and get its control number
                    string s = heFind.innerHTML;
                    int ich = s.IndexOf(_s_Announcements_Button_Edit_Prefix);
                    if (ich > 0)
                        {
                        sCtl = s.Substring(ich + _s_Announcements_Button_Edit_Prefix.Length, 5);
                        }
                    break;
                    }
                }

            ThrowIfNot(sCtl != null, "Can't find HELP announcement");

            m_awc.ResetNav();
            string sControl = BuildAnnName(_sid_Announcements_Button_Edit_Prefix, _sid_Announcements_Button_Edit_Suffix, sCtl);

            ThrowIfNot(m_awc.FClickControl(oDoc2, sControl), "Couldn't find edit button");
            m_awc.FWaitForNavFinish();

            // now edit the text
            sControl = BuildAnnName(_s_Announcements_Textarea_Text_Prefix, _s_Announcements_Textarea_Text_Suffix, sCtl);

            ThrowIfNot(ArbWebControl.FSetTextareaControlText(oDoc2, sControl, sArbiterHelpNeeded, true), "Can't set control text");
            m_awc.FWaitForNavFinish();

            sControl = BuildAnnName(_sid_Announcements_Button_Save_Prefix, _sid_Announcements_Button_Save_Suffix, sCtl);

            ThrowIfNot(m_awc.FClickControl(oDoc2, sControl), "Couldn't find save button");
            m_awc.FWaitForNavFinish();

            // and now save it.

            m_srpt.PopLevel();
			m_srpt.AddMessage("Completed Announcement Set.");
        }

	    /* G E N  O P E N  S L O T S  M A I L */
	    /*----------------------------------------------------------------------------
	    	%%Function: GenOpenSlotsMail
	    	%%Qualified: ArbWeb.AwMainForm.GenOpenSlotsMail
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
	    private void GenOpenSlotsMail(object sender, EventArgs e)
		{
			CountsData gc = GcEnsure(m_ebRosterWorking.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
			string sTempFile = String.Format("{0}\\temp{1}.htm", Environment.GetEnvironmentVariable("Temp"), System.Guid.NewGuid().ToString());
            Roster rst = RstEnsure(m_ebRosterWorking.Text);
            
            string sBcc = m_cbTestEmail.Checked ? "" : rst.SBuildAddressLine(m_ebFilter.Text); ;

			Outlook.Application appOlk = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");

			if (appOlk == null)
				{
				MessageBox.Show("No running instance of outlook!");
				return;
				}

	        Outlook.MailItem oNote = appOlk.CreateItem(Outlook.OlItemType.olMailItem);
			// Outlook.MailItem oNote = (Outlook.MailItem)appOlk.CreateItem(Outlook.OlItemType.olMailItem);

			oNote.To = "rlittle@thetasoft.com";
			oNote.BCC = sBcc;
			oNote.Subject = "This is a test";
			oNote.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            oNote.HTMLBody = "<html><style>\r\n*#myId {\ncolor:Blue;\n}\n</style><body><p>Put your preamble here...</p>";

		    if (m_cbSplitSports.Checked)
		        {
		        string[] rgs;

		        oNote.HTMLBody += "<h1>Baseball open slots</h1>";
		        rgs = ArbWebControl.RgsFromChlbxSport(m_cbFilterSport.Checked, m_chlbxSports, "Softball", false);
		        gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
                                      rgs, ArbWebControl.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
		        oNote.HTMLBody += SHtmlReadFile(sTempFile) + "<h1>Softball Open Slots</h1>";
		        rgs = ArbWebControl.RgsFromChlbxSport(m_cbFilterSport.Checked, m_chlbxSports, "Softball", true);
		        gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
                                      rgs, ArbWebControl.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
		        oNote.HTMLBody += SHtmlReadFile(sTempFile);
		        }
		    else
		        {
		        gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
		                              ArbWebControl.RgsFromChlbx(m_cbFilterSport.Checked, m_chlbxSports),
		                              ArbWebControl.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
		        oNote.HTMLBody += SHtmlReadFile(sTempFile);
		        }
		    oNote.Display(true);

			appOlk = null;
			System.IO.File.Delete(sTempFile);            
		}

	    private SlotAggr m_saOpenSlots;
        private async void CalcOpenSlots(object sender, EventArgs e)
        {
            Task<CountsData> taskCalc = new Task<CountsData>(CalcOpenSlotsWork);
            taskCalc.Start();

            CountsData cd = await taskCalc;

            m_srpt.PopLevel();
            m_srpt.AddMessage("Updating listboxes...", StatusRpt.MSGT.Header, false);
			// update regenerate the listboxes...
			string[] rgsSports = ArbWebControl.RgsFromChlbx(true, m_chlbxSports);
			string[] rgsSportLevels = ArbWebControl.RgsFromChlbx(true, m_chlbxSportLevels);

			bool fCheckAllSports = false;
			bool fCheckAllSportLevels = false;

			if (rgsSports.Length == 0 && m_chlbxSports.Items.Count == 0)
				fCheckAllSports = true;

			if (rgsSports.Length == 0 && m_chlbxSportLevels.Items.Count == 0)
				fCheckAllSportLevels = true;

            ArbWebControl.UpdateChlbxFromRgs(m_chlbxSports, cd.GetOpenSlotSports(m_saOpenSlots), rgsSports, null, fCheckAllSports);
            ArbWebControl.UpdateChlbxFromRgs(m_chlbxSportLevels, cd.GetOpenSlotSportLevels(m_saOpenSlots), rgsSportLevels, fCheckAllSports ? null : rgsSports, fCheckAllSportLevels);
            string[] rgsRosterSites = ArbWebControl.RgsFromChlbx(true, m_chlbxRoster);

            ArbWebControl.UpdateChlbxFromRgs(m_chlbxRoster, cd.GetSiteRosterSites(m_saOpenSlots), rgsRosterSites, null, false);
            m_srpt.PopLevel();
        }

	    private CountsData CalcOpenSlotsWork()
	    {
	        m_srpt.AddMessage("Calculating slot data...", StatusRpt.MSGT.Header, false);

	        CountsData gc = GcEnsure(m_ebRosterWorking.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
	        Roster rst = RstEnsure(m_ebRosterWorking.Text);

	        m_srpt.PopLevel();
	        m_srpt.AddMessage("Calculating open slots...", StatusRpt.MSGT.Header, false);
	        m_saOpenSlots = gc.CalcOpenSlots(m_dtpStart.Value, m_dtpEnd.Value);
	        return gc;
	    }

	    private void DoSportLevelFilter(object sender, ItemCheckEventArgs e)
        {
            CountsData gc = GcEnsure(m_ebRosterWorking.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
            string[] rgsSports = ArbWebControl.RgsFromChlbx(true, m_chlbxSports, e.Index, e.CurrentValue != CheckState.Checked, null, false);
            string[] rgsSportLevels = ArbWebControl.RgsFromChlbx(true, m_chlbxSportLevels);
            ArbWebControl.UpdateChlbxFromRgs(m_chlbxSportLevels, gc.GetOpenSlotSportLevels(m_saOpenSlots), rgsSportLevels, rgsSports, false);
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
			CountsData gc = GcEnsure(m_ebRosterWorking.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
			Roster rst = RstEnsure(m_ebRosterWorking.Text);

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
                    return m_ebRosterWorking;
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
	            m_srpt.AttachLogfile(SBuildTempFilename("arblog", "log"));
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

}
