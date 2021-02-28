using System;
using System.Diagnostics;
using System.Drawing;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using Microsoft.Win32;
using StatusBox;
using mshtml;
using System.Runtime.InteropServices;
using TCore.CmdLine;
using Outlook=Microsoft.Office.Interop.Outlook;
using System.Threading.Tasks;
using TCore.Settings;
using TCore.UI;
using TCore.Util;

namespace ArbWeb
{
    public interface IAppContext
    {
        StatusRpt StatusReport { get; }
        ArbWebControl WebControl { get; }
        void EnsureLoggedIn();
        Profile Profile { get; }
        void ThrowIfNot(bool f, string s);
    }

    /// <summary>
    /// Summary description for AwMainForm.
    /// </summary>
    public partial class AwMainForm : System.Windows.Forms.Form, TCore.CmdLine.ICmdLineDispatch, IAppContext
    {
        public AwMainForm() { } // for unit tests only

        #region Controls

        private System.Windows.Forms.Button m_pbDownloadGames;

        private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.MenuItem menuItem1;
        object Zero = 0;
        object EmptyString = "";
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
        private Button m_pbReload;
        private CheckBox m_cbRankOnly;
        private Label label15;
        private Label label16;
        private TextBox m_ebGameOutput;
        private Button m_pbGenGames;
        private CheckBox m_cbAddOfficialsOnly;
        private Button m_pbBrowseGamesReport;
        private Button m_pbBrowseAnalysis;
        private Button button3;
        private TextBox m_ebAffiliationIndex;
        private CheckBox m_cbFuzzyTimes;
        private CheckBox m_cbDatePivot;
        private CheckBox m_cbSplitSports;
        private Label label18;
        private CheckedListBox m_chlbxRoster;
        private Button m_pbCreateRosterReport;
        private CheckBox m_cbFilterRank;
        private Button m_pbMailMerge;
        private CheckBox m_cbFutureOnly;
        private Label label19;
        private CheckBox m_cbLaunch;
        private CheckBox m_cbSetArbiterAnnounce;
        private Button m_pbEditProfile;
        private Button m_pbAddProfile;
        private ComboBox m_cbxGameFilter;
        private Button m_pbRefreshGameFilters;
        private Button pbTestDownload;
        private CheckBox m_cbSkipContactDownload;
        private StatusBox.StatusRpt m_srpt;

        #endregion

        public StatusRpt StatusReport => m_srpt;

        public void ThrowIfNot(bool f, string s)
        {
            if (!f)
                throw new Exception(s);
        }

        private bool m_fAutomateUpdateHelp = false;
        private List<string> m_plsAutomateIncludeSport = new List<string>();
        private string m_sAutomateDateStart = null;
        private string m_sAutomateDateEnd = null;
        private bool m_fForceFutureGames = false;

        public Profile Profile => m_pr;

        private Settings.SettingsElt[] m_rgrehe;
        private Settings m_reh;

        ArbWebControl m_awc;
        public ArbWebControl WebControl => m_awc;

        bool m_fDontUpdateProfile;
        bool m_fLoggedIn;
        Roster m_rst;
        CountsData m_gc;
        private bool m_fAutomating = false;

        #region Top Level Program Flow

        /*----------------------------------------------------------------------------
        	%%Function: FDispatchCmdLineSwitch
        	%%Qualified: ArbWeb.AwMainForm.FDispatchCmdLineSwitch
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public bool FDispatchCmdLineSwitch(CmdLineSwitch cls, string sParam, object oClient, out string sError)
        {
            sError = null;

            if (cls.Switch.Length == 1)
                {
                switch (cls.Switch[0])
                    {
                    case 'H':
                        m_fAutomateUpdateHelp = true;
                        break;
                    case 'F':
                        m_plsAutomateIncludeSport.Add(sParam);
                        break;
                    case 'f':
                        m_fForceFutureGames = true;
                        break;
                    default:
                        sError = String.Format("Unknown arg: '{0}'", cls.Switch);
                        return false;
                    }

                return true;
                }

            if (cls.Switch == "DS")
                m_sAutomateDateStart = sParam;
            else if (cls.Switch == "DE")
                m_sAutomateDateEnd = sParam;
            else
                {
                sError = String.Format("Unknown arg: '{0}'", cls.Switch);
                return false;
                }

            return true;
        }


        private StringBuilder m_sbUsage = null;

        /*----------------------------------------------------------------------------
        	%%Function: AppendUsageString
        	%%Qualified: ArbWeb.AwMainForm.AppendUsageString
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void AppendUsageString(string s)
        {
            m_sbUsage.AppendLine(s);
        }

        #endregion

        #region Form Support

        /* E N A B L E  A D M I N  F U N C T I O N S */
        /*----------------------------------------------------------------------------
        	%%Function: EnableAdminFunctions
        	%%Qualified: ArbWeb.AwMainForm.EnableAdminFunctions
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public void EnableAdminFunctions()
        {
            bool fAdmin = (String.Compare(System.Environment.MachineName, "dogmatix", true) == 0);
            m_pbUploadRoster.Enabled = fAdmin;
        }
        /* A W  M A I N  F O R M */
        /*----------------------------------------------------------------------------
        	%%Function: AwMainForm
        	%%Qualified: ArbWeb.AwMainForm.AwMainForm
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public AwMainForm(string[] rgsCmdLine)
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

            string[] rgs = rk.GetSubKeyNames();

            foreach (string s in rgs)
                {
                m_cbxProfile.Items.Add(s);
                }

            m_rgrehe = new[]
                           {
                           new Settings.SettingsElt("LastProfile", Settings.Type.Str, m_cbxProfile, "")
                           };

            m_reh = new Settings(m_rgrehe, "Software\\Thetasoft\\ArbWeb", "root");
            m_reh.Load();

            m_pr = new Profile();
            m_pr.Load(m_cbxProfile.Text);
            SetupLogToFile();

            SetUIForProfile(m_pr);

            // load MRU from registry
            m_fDontUpdateProfile = false;
            EnableControls();

            if (m_cbShowBrowser.Checked)
                m_awc.Show();

            CmdLineConfig clcfg = new CmdLineConfig(new CmdLineSwitch[]
                                                        {
                                                        new CmdLineSwitch("H", true /*fToggle*/, false /*fReq*/,
                                                                          "Update arbiter HELP NEEDED (includes downloading games, calculating slots). Requires -DS and -DE.",
                                                                          "help announce", null),
                                                        new CmdLineSwitch("DS", false /*fToggle*/, false /*fReq*/, "Start date for slot calculation (required if -H specified)",
                                                                          "date start", null),
                                                        new CmdLineSwitch("DE", false /*fToggle*/, false /*fReq*/, "End date for slot calculation (required if -H specified)",
                                                                          "date end", null),
                                                        new CmdLineSwitch("F", false /*fToggle*/, false /*fReq*/, "Check this item in the Game/Slot filter", "Sport filter",
                                                                          null),
                                                        new CmdLineSwitch("f", true /*fToggle*/, false /*fReq*/, "Force the games download to only download future games",
                                                                          "Future Games Only", null),
                                                        });

            CmdLine cl = new CmdLine(clcfg);
            string sError = null;

            if (rgsCmdLine != null && rgsCmdLine.Length > 0)
                m_srpt.AddMessage(String.Format("Commandline args: {0} {1}", rgsCmdLine.Length, rgsCmdLine[0]));

            if (!cl.FParse(rgsCmdLine, this, null, out sError) || (m_fAutomateUpdateHelp && (m_sAutomateDateEnd == null || m_sAutomateDateStart == null)))
                {
                m_sbUsage = new StringBuilder();

                cl.Usage(AppendUsageString);
                MessageBox.Show(String.Format("Command Line error: {0}\n{1}", sError, m_sbUsage.ToString()), "ArbWeb");
                m_fAutomating = true;
                Close();
                }

            if (rgsCmdLine != null && rgsCmdLine.Length > 0)
                {
                m_fAutomating = true;

                if (m_fAutomateUpdateHelp)
                    {
                    DateTime dttmStart = DateTime.Parse(m_sAutomateDateStart);
                    DateTime dttmEnd = DateTime.Parse(m_sAutomateDateEnd);

                    m_cbLaunch.Checked = false;
                    m_cbSetArbiterAnnounce.Checked = true;
                    m_dtpStart.Value = dttmStart;
                    m_dtpEnd.Value = dttmEnd;
                    if (m_fForceFutureGames)
                        m_cbFutureOnly.Checked = true;

                    QueueUIOp(new DelayedUIOpDel(HandleDownloadGamesClick), new object[] {null, null});
                    QueueUIOp(new DelayedUIOpDel(CalcOpenSlots), new object[] {null, null});
                    QueueUIOp(new DelayedUIOpDel(DoCheckSportListboxes), new object[] {null, null});
                    QueueUIOp(new DelayedUIOpDel(GenMailMergeMail), new object[] {null, null});
                    QueueUIOp(new DelayedUIOpDel(DoExitApp), new object[] {null, null});

                    DoPendingQueueUIOp();
                    }
                }
        }

        /*----------------------------------------------------------------------------
        	%%Function: DoCheckSportListboxes
        	%%Qualified: ArbWeb.AwMainForm.DoCheckSportListboxes
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoCheckSportListboxes(object sender, EventArgs e)
        {
            if (m_plsAutomateIncludeSport.Count == 0)
                return;

            m_cbFilterSport.Checked = true;
            EnableControls();

            // first, uncheck everone
            for (int i = 0; i < m_chlbxSports.Items.Count; i++)
                {
                m_chlbxSports.SetItemChecked(i, false);
                }

            foreach (string s in m_plsAutomateIncludeSport)
                {
                // find the item in the listbox
                for (int i = 0; i < m_chlbxSports.Items.Count; i++)
                    {
                    string sItem = (string) m_chlbxSports.Items[i];

                    if (sItem.Contains(s))
                        m_chlbxSports.SetItemChecked(i, true);
                    }
                }

            DoPendingQueueUIOp();
        }

        void DoExitApp(object sender, EventArgs e)
        {
            this.Close();
        }

        #region Queued UI Op

        delegate void DelayedUIOpDel(object sender, EventArgs e);

        struct QueuedUIOp
        {
            public DelayedUIOpDel del;
            public object[] rgo;

            public QueuedUIOp(DelayedUIOpDel del, object[] rgo)
            {
                this.del = del;
                this.rgo = rgo;
            }
        }

        private delegate void QueueUIOpDel(DelayedUIOpDel del, object[] rgo);

        void DoQueueUIOp(DelayedUIOpDel del, object[] rgo)
        {
            QueuedUIOp qui = new QueuedUIOp(del, rgo);

            plqui.Add(qui);
        }

        void QueueUIOp(DelayedUIOpDel del, object[] rgo)
        {
            if (this.InvokeRequired)
                this.Invoke(new QueueUIOpDel(DoQueueUIOp), new object[] {del, rgo});
            else
                DoQueueUIOp(del, rgo);
        }

        private delegate void DoPendingQueueUIOpDel();

        void DoDoPendingQueueUIOp()
        {
            if (plqui.Count > 0)
                {
                QueuedUIOp qui = plqui[0];

                plqui.RemoveAt(0);
                qui.del(qui.rgo[0], (EventArgs) qui.rgo[1]);
                }
        }

        void DoPendingQueueUIOp()
        {
            if (this.InvokeRequired)
                this.Invoke(new DoPendingQueueUIOpDel(DoDoPendingQueueUIOp));
            else
                DoDoPendingQueueUIOp();
        }



        private List<QueuedUIOp> plqui = new List<QueuedUIOp>();


        #endregion


        #region Cursors

        List<Cursor> m_plCursor;

        private delegate void PushCursorDel(Cursor crs);

        /*----------------------------------------------------------------------------
        	%%Function: DoPushCursor
        	%%Qualified: ArbWeb.AwMainForm.DoPushCursor
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoPushCursor(Cursor crs)
        {
            m_plCursor.Add(this.Cursor);
            this.Cursor = crs;
        }


        /* P U S H  C U R S O R */
        /*----------------------------------------------------------------------------
        	%%Function: PushCursor
        	%%Qualified: ArbWeb.AwMainForm.PushCursor
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void PushCursor(Cursor crs)
        {
            if (this.InvokeRequired)
                this.BeginInvoke(new PushCursorDel(DoPushCursor), crs);
            else
                DoPushCursor(crs);
        }

        private delegate void PopCursorDel();

        /*----------------------------------------------------------------------------
        	%%Function: DoPopCursor
        	%%Qualified: ArbWeb.AwMainForm.DoPopCursor
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void DoPopCursor()
        {
            if (m_plCursor.Count > 0)
                {
                this.Cursor = m_plCursor[m_plCursor.Count - 1];
                m_plCursor.RemoveAt(m_plCursor.Count - 1);
                }
        }
        /* P O P  C U R S O R */
        /*----------------------------------------------------------------------------
        	%%Function: PopCursor
        	%%Qualified: ArbWeb.AwMainForm.PopCursor
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void PopCursor()
        {
            if (this.InvokeRequired)
                this.BeginInvoke(new PopCursorDel(DoPopCursor));
            else
                DoPopCursor();
        }

        #endregion

        #endregion

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                {
                if (components != null)
                    {
                    components.Dispose();
                    }
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
			System.Windows.Forms.Label label17;
			this.m_pbDownloadGames = new System.Windows.Forms.Button();
			this.contextMenu1 = new System.Windows.Forms.ContextMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
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
			this.m_pbReload = new System.Windows.Forms.Button();
			this.m_cbRankOnly = new System.Windows.Forms.CheckBox();
			this.label15 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.m_ebGameOutput = new System.Windows.Forms.TextBox();
			this.m_pbGenGames = new System.Windows.Forms.Button();
			this.m_cbAddOfficialsOnly = new System.Windows.Forms.CheckBox();
			this.m_pbBrowseGamesReport = new System.Windows.Forms.Button();
			this.m_pbBrowseAnalysis = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.m_ebAffiliationIndex = new System.Windows.Forms.TextBox();
			this.m_cbFuzzyTimes = new System.Windows.Forms.CheckBox();
			this.m_cbDatePivot = new System.Windows.Forms.CheckBox();
			this.m_cbSplitSports = new System.Windows.Forms.CheckBox();
			this.label18 = new System.Windows.Forms.Label();
			this.m_chlbxRoster = new System.Windows.Forms.CheckedListBox();
			this.m_pbCreateRosterReport = new System.Windows.Forms.Button();
			this.m_cbFilterRank = new System.Windows.Forms.CheckBox();
			this.m_pbMailMerge = new System.Windows.Forms.Button();
			this.m_cbFutureOnly = new System.Windows.Forms.CheckBox();
			this.label19 = new System.Windows.Forms.Label();
			this.m_cbLaunch = new System.Windows.Forms.CheckBox();
			this.m_cbSetArbiterAnnounce = new System.Windows.Forms.CheckBox();
			this.m_pbEditProfile = new System.Windows.Forms.Button();
			this.m_pbAddProfile = new System.Windows.Forms.Button();
			this.m_cbxGameFilter = new System.Windows.Forms.ComboBox();
			this.m_pbRefreshGameFilters = new System.Windows.Forms.Button();
			this.pbTestDownload = new System.Windows.Forms.Button();
			this.m_cbSkipContactDownload = new System.Windows.Forms.CheckBox();
			label17 = new System.Windows.Forms.Label();
			this.groupBox2.SuspendLayout();
			this.SuspendLayout();
			// 
			// label17
			// 
			label17.AutoSize = true;
			label17.Location = new System.Drawing.Point(530, 243);
			label17.Name = "label17";
			label17.Size = new System.Drawing.Size(155, 20);
			label17.TabIndex = 76;
			label17.Text = "Affiliation Field Index";
			// 
			// m_pbDownloadGames
			// 
			this.m_pbDownloadGames.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.m_pbDownloadGames.Location = new System.Drawing.Point(985, 113);
			this.m_pbDownloadGames.Name = "m_pbDownloadGames";
			this.m_pbDownloadGames.Size = new System.Drawing.Size(176, 35);
			this.m_pbDownloadGames.TabIndex = 14;
			this.m_pbDownloadGames.Text = "Download Games";
			this.m_pbDownloadGames.Click += new System.EventHandler(this.HandleDownloadGamesClick);
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
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.Location = new System.Drawing.Point(985, 38);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(176, 35);
			this.button1.TabIndex = 26;
			this.button1.Text = "Get Contacts";
			this.button1.Click += new System.EventHandler(this.HandleDownloadContactsClick);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.m_recStatus);
			this.groupBox2.Location = new System.Drawing.Point(13, 926);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1148, 229);
			this.groupBox2.TabIndex = 27;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Status";
			// 
			// m_recStatus
			// 
			this.m_recStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.m_recStatus.Location = new System.Drawing.Point(10, 28);
			this.m_recStatus.Name = "m_recStatus";
			this.m_recStatus.Size = new System.Drawing.Size(1132, 193);
			this.m_recStatus.TabIndex = 0;
			this.m_recStatus.Text = "";
			// 
			// m_pbGenCounts
			// 
			this.m_pbGenCounts.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.m_pbGenCounts.Location = new System.Drawing.Point(985, 314);
			this.m_pbGenCounts.Name = "m_pbGenCounts";
			this.m_pbGenCounts.Size = new System.Drawing.Size(176, 40);
			this.m_pbGenCounts.TabIndex = 28;
			this.m_pbGenCounts.Text = "Gen Analysis";
			this.m_pbGenCounts.Click += new System.EventHandler(this.DoGenCounts);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(19, 314);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(87, 20);
			this.label1.TabIndex = 30;
			this.label1.Text = "Output File";
			// 
			// m_ebOutputFile
			// 
			this.m_ebOutputFile.Location = new System.Drawing.Point(122, 310);
			this.m_ebOutputFile.Name = "m_ebOutputFile";
			this.m_ebOutputFile.Size = new System.Drawing.Size(332, 26);
			this.m_ebOutputFile.TabIndex = 29;
			// 
			// m_cbIncludeCanceled
			// 
			this.m_cbIncludeCanceled.AutoSize = true;
			this.m_cbIncludeCanceled.Location = new System.Drawing.Point(502, 314);
			this.m_cbIncludeCanceled.Name = "m_cbIncludeCanceled";
			this.m_cbIncludeCanceled.Size = new System.Drawing.Size(158, 24);
			this.m_cbIncludeCanceled.TabIndex = 31;
			this.m_cbIncludeCanceled.Text = "Include Canceled";
			this.m_cbIncludeCanceled.UseVisualStyleBackColor = true;
			this.m_cbIncludeCanceled.Click += new System.EventHandler(this.DoGenericInvalGc);
			// 
			// m_pbUploadRoster
			// 
			this.m_pbUploadRoster.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.m_pbUploadRoster.Location = new System.Drawing.Point(805, 887);
			this.m_pbUploadRoster.Name = "m_pbUploadRoster";
			this.m_pbUploadRoster.Size = new System.Drawing.Size(176, 35);
			this.m_pbUploadRoster.TabIndex = 32;
			this.m_pbUploadRoster.Text = "Upload Roster";
			this.m_pbUploadRoster.Click += new System.EventHandler(this.HandleUploadRosterClick);
			// 
			// m_pbOpenSlots
			// 
			this.m_pbOpenSlots.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.m_pbOpenSlots.Location = new System.Drawing.Point(985, 386);
			this.m_pbOpenSlots.Name = "m_pbOpenSlots";
			this.m_pbOpenSlots.Size = new System.Drawing.Size(176, 39);
			this.m_pbOpenSlots.TabIndex = 33;
			this.m_pbOpenSlots.Text = "Calc Slots";
			this.m_pbOpenSlots.Click += new System.EventHandler(this.CalcOpenSlots);
			// 
			// m_lblSearchCriteria
			// 
			this.m_lblSearchCriteria.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.m_lblSearchCriteria.Location = new System.Drawing.Point(8, 13);
			this.m_lblSearchCriteria.Name = "m_lblSearchCriteria";
			this.m_lblSearchCriteria.Size = new System.Drawing.Size(1144, 24);
			this.m_lblSearchCriteria.TabIndex = 34;
			this.m_lblSearchCriteria.Tag = "Shared Configuration";
			this.m_lblSearchCriteria.Paint += new System.Windows.Forms.PaintEventHandler(this.RenderHeadingLine);
			// 
			// m_cbShowBrowser
			// 
			this.m_cbShowBrowser.AutoSize = true;
			this.m_cbShowBrowser.Location = new System.Drawing.Point(24, 155);
			this.m_cbShowBrowser.Name = "m_cbShowBrowser";
			this.m_cbShowBrowser.Size = new System.Drawing.Size(162, 24);
			this.m_cbShowBrowser.TabIndex = 35;
			this.m_cbShowBrowser.Text = "Show Diagnostics";
			this.m_cbShowBrowser.UseVisualStyleBackColor = true;
			this.m_cbShowBrowser.Click += new System.EventHandler(this.ChangeShowBrowser);
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.label6.Location = new System.Drawing.Point(8, 278);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(1144, 27);
			this.label6.TabIndex = 36;
			this.label6.Tag = "Games Worked Analysis";
			this.label6.Paint += new System.Windows.Forms.PaintEventHandler(this.RenderHeadingLine);
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.label7.Location = new System.Drawing.Point(8, 354);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(1144, 27);
			this.label7.TabIndex = 37;
			this.label7.Tag = "Open Slot Reporting";
			this.label7.Paint += new System.Windows.Forms.PaintEventHandler(this.RenderHeadingLine);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(19, 392);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(83, 20);
			this.label8.TabIndex = 40;
			this.label8.Text = "Start Date";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(477, 392);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(77, 20);
			this.label9.TabIndex = 41;
			this.label9.Text = "End Date";
			// 
			// m_dtpStart
			// 
			this.m_dtpStart.CustomFormat = "ddd MMM dd, yyyy";
			this.m_dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.m_dtpStart.Location = new System.Drawing.Point(122, 386);
			this.m_dtpStart.Name = "m_dtpStart";
			this.m_dtpStart.Size = new System.Drawing.Size(246, 26);
			this.m_dtpStart.TabIndex = 42;
			this.m_dtpStart.Value = new System.DateTime(2008, 5, 4, 0, 0, 0, 0);
			// 
			// m_dtpEnd
			// 
			this.m_dtpEnd.CustomFormat = "ddd MMM dd, yyyy";
			this.m_dtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.m_dtpEnd.Location = new System.Drawing.Point(570, 386);
			this.m_dtpEnd.Name = "m_dtpEnd";
			this.m_dtpEnd.Size = new System.Drawing.Size(246, 26);
			this.m_dtpEnd.TabIndex = 43;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(66, 634);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(118, 20);
			this.label10.TabIndex = 45;
			this.label10.Text = "Misc Field Filter";
			// 
			// m_ebFilter
			// 
			this.m_ebFilter.Location = new System.Drawing.Point(202, 630);
			this.m_ebFilter.Name = "m_ebFilter";
			this.m_ebFilter.Size = new System.Drawing.Size(206, 26);
			this.m_ebFilter.TabIndex = 44;
			// 
			// m_cbOpenSlotDetail
			// 
			this.m_cbOpenSlotDetail.AutoSize = true;
			this.m_cbOpenSlotDetail.Location = new System.Drawing.Point(674, 457);
			this.m_cbOpenSlotDetail.Name = "m_cbOpenSlotDetail";
			this.m_cbOpenSlotDetail.Size = new System.Drawing.Size(202, 24);
			this.m_cbOpenSlotDetail.TabIndex = 46;
			this.m_cbOpenSlotDetail.Text = "Include game/slot detail";
			this.m_cbOpenSlotDetail.UseVisualStyleBackColor = true;
			// 
			// m_cbxProfile
			// 
			this.m_cbxProfile.FormattingEnabled = true;
			this.m_cbxProfile.Location = new System.Drawing.Point(90, 41);
			this.m_cbxProfile.Name = "m_cbxProfile";
			this.m_cbxProfile.Size = new System.Drawing.Size(228, 28);
			this.m_cbxProfile.TabIndex = 47;
			this.m_cbxProfile.SelectedIndexChanged += new System.EventHandler(this.ChangeProfile);
			this.m_cbxProfile.Leave += new System.EventHandler(this.OnProfileLeave);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(22, 47);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(53, 20);
			this.label11.TabIndex = 48;
			this.label11.Text = "Profile";
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.Location = new System.Drawing.Point(965, 517);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(176, 40);
			this.button2.TabIndex = 49;
			this.button2.Text = "Create Email";
			this.button2.Click += new System.EventHandler(this.GenOpenSlotsMail);
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.label12.Location = new System.Drawing.Point(50, 430);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(1102, 27);
			this.label12.TabIndex = 50;
			this.label12.Tag = "Slot reporting";
			this.label12.Paint += new System.Windows.Forms.PaintEventHandler(this.RenderHeadingLine);
			// 
			// m_cbFilterSport
			// 
			this.m_cbFilterSport.AutoSize = true;
			this.m_cbFilterSport.Location = new System.Drawing.Point(54, 457);
			this.m_cbFilterSport.Name = "m_cbFilterSport";
			this.m_cbFilterSport.Size = new System.Drawing.Size(205, 24);
			this.m_cbFilterSport.TabIndex = 51;
			this.m_cbFilterSport.Text = "Include/filter sport count";
			this.m_cbFilterSport.UseVisualStyleBackColor = true;
			this.m_cbFilterSport.CheckedChanged += new System.EventHandler(this.HandleSportChecked);
			// 
			// m_chlbxSports
			// 
			this.m_chlbxSports.CheckOnClick = true;
			this.m_chlbxSports.FormattingEnabled = true;
			this.m_chlbxSports.Location = new System.Drawing.Point(70, 487);
			this.m_chlbxSports.Name = "m_chlbxSports";
			this.m_chlbxSports.Size = new System.Drawing.Size(263, 96);
			this.m_chlbxSports.TabIndex = 52;
			this.m_chlbxSports.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.DoSportLevelFilter);
			// 
			// m_chlbxSportLevels
			// 
			this.m_chlbxSportLevels.CheckOnClick = true;
			this.m_chlbxSportLevels.FormattingEnabled = true;
			this.m_chlbxSportLevels.Location = new System.Drawing.Point(389, 487);
			this.m_chlbxSportLevels.Name = "m_chlbxSportLevels";
			this.m_chlbxSportLevels.Size = new System.Drawing.Size(262, 96);
			this.m_chlbxSportLevels.TabIndex = 54;
			// 
			// m_cbFilterLevel
			// 
			this.m_cbFilterLevel.AutoSize = true;
			this.m_cbFilterLevel.Location = new System.Drawing.Point(373, 457);
			this.m_cbFilterLevel.Name = "m_cbFilterLevel";
			this.m_cbFilterLevel.Size = new System.Drawing.Size(240, 24);
			this.m_cbFilterLevel.TabIndex = 53;
			this.m_cbFilterLevel.Text = "Include/filter sport level count";
			this.m_cbFilterLevel.UseVisualStyleBackColor = true;
			this.m_cbFilterLevel.CheckedChanged += new System.EventHandler(this.HandleSportLevelChecked);
			// 
			// m_cbTestEmail
			// 
			this.m_cbTestEmail.AutoSize = true;
			this.m_cbTestEmail.Location = new System.Drawing.Point(922, 457);
			this.m_cbTestEmail.Name = "m_cbTestEmail";
			this.m_cbTestEmail.Size = new System.Drawing.Size(175, 24);
			this.m_cbTestEmail.TabIndex = 55;
			this.m_cbTestEmail.Text = "Generate test email";
			this.m_cbTestEmail.UseVisualStyleBackColor = true;
			// 
			// m_pbReload
			// 
			this.m_pbReload.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.m_pbReload.Location = new System.Drawing.Point(985, 887);
			this.m_pbReload.Name = "m_pbReload";
			this.m_pbReload.Size = new System.Drawing.Size(176, 35);
			this.m_pbReload.TabIndex = 61;
			this.m_pbReload.Text = "Load Data";
			this.m_pbReload.Click += new System.EventHandler(this.DoReloadClick);
			// 
			// m_cbRankOnly
			// 
			this.m_cbRankOnly.AutoSize = true;
			this.m_cbRankOnly.Location = new System.Drawing.Point(363, 155);
			this.m_cbRankOnly.Name = "m_cbRankOnly";
			this.m_cbRankOnly.Size = new System.Drawing.Size(192, 24);
			this.m_cbRankOnly.TabIndex = 62;
			this.m_cbRankOnly.Text = "Upload Rankings Only";
			this.m_cbRankOnly.UseVisualStyleBackColor = true;
			// 
			// label15
			// 
			this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.label15.Location = new System.Drawing.Point(8, 206);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(1144, 28);
			this.label15.TabIndex = 63;
			this.label15.Tag = "Game Reporting";
			this.label15.Paint += new System.Windows.Forms.PaintEventHandler(this.RenderHeadingLine);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(19, 243);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(87, 20);
			this.label16.TabIndex = 65;
			this.label16.Text = "Output File";
			// 
			// m_ebGameOutput
			// 
			this.m_ebGameOutput.Location = new System.Drawing.Point(122, 238);
			this.m_ebGameOutput.Name = "m_ebGameOutput";
			this.m_ebGameOutput.Size = new System.Drawing.Size(332, 26);
			this.m_ebGameOutput.TabIndex = 64;
			// 
			// m_pbGenGames
			// 
			this.m_pbGenGames.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.m_pbGenGames.Location = new System.Drawing.Point(985, 238);
			this.m_pbGenGames.Name = "m_pbGenGames";
			this.m_pbGenGames.Size = new System.Drawing.Size(176, 40);
			this.m_pbGenGames.TabIndex = 66;
			this.m_pbGenGames.Text = "Games Report";
			this.m_pbGenGames.Click += new System.EventHandler(this.DoGamesReport);
			// 
			// m_cbAddOfficialsOnly
			// 
			this.m_cbAddOfficialsOnly.AutoSize = true;
			this.m_cbAddOfficialsOnly.Location = new System.Drawing.Point(578, 155);
			this.m_cbAddOfficialsOnly.Name = "m_cbAddOfficialsOnly";
			this.m_cbAddOfficialsOnly.Size = new System.Drawing.Size(216, 24);
			this.m_cbAddOfficialsOnly.TabIndex = 67;
			this.m_cbAddOfficialsOnly.Text = "Upload New Officials Only";
			this.m_cbAddOfficialsOnly.UseVisualStyleBackColor = true;
			// 
			// m_pbBrowseGamesReport
			// 
			this.m_pbBrowseGamesReport.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.m_pbBrowseGamesReport.Location = new System.Drawing.Point(456, 242);
			this.m_pbBrowseGamesReport.Name = "m_pbBrowseGamesReport";
			this.m_pbBrowseGamesReport.Size = new System.Drawing.Size(37, 22);
			this.m_pbBrowseGamesReport.TabIndex = 72;
			this.m_pbBrowseGamesReport.Tag = ArbWeb.EditProfile.FNC.ReportFile;
			this.m_pbBrowseGamesReport.Text = "...";
			this.m_pbBrowseGamesReport.Click += new System.EventHandler(this.DoBrowseOpen);
			// 
			// m_pbBrowseAnalysis
			// 
			this.m_pbBrowseAnalysis.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.m_pbBrowseAnalysis.Location = new System.Drawing.Point(456, 314);
			this.m_pbBrowseAnalysis.Name = "m_pbBrowseAnalysis";
			this.m_pbBrowseAnalysis.Size = new System.Drawing.Size(37, 22);
			this.m_pbBrowseAnalysis.TabIndex = 73;
			this.m_pbBrowseAnalysis.Tag = ArbWeb.EditProfile.FNC.AnalysisFile;
			this.m_pbBrowseAnalysis.Text = "...";
			this.m_pbBrowseAnalysis.Click += new System.EventHandler(this.DoBrowseOpen);
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button3.Location = new System.Drawing.Point(985, 75);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(176, 35);
			this.button3.TabIndex = 74;
			this.button3.Text = "Get Quick Roster";
			this.button3.Click += new System.EventHandler(this.HandleDownloadQuickRosterClick);
			// 
			// m_ebAffiliationIndex
			// 
			this.m_ebAffiliationIndex.Location = new System.Drawing.Point(709, 238);
			this.m_ebAffiliationIndex.Name = "m_ebAffiliationIndex";
			this.m_ebAffiliationIndex.Size = new System.Drawing.Size(51, 26);
			this.m_ebAffiliationIndex.TabIndex = 77;
			this.m_ebAffiliationIndex.Text = "0";
			this.m_ebAffiliationIndex.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// m_cbFuzzyTimes
			// 
			this.m_cbFuzzyTimes.AutoSize = true;
			this.m_cbFuzzyTimes.Location = new System.Drawing.Point(674, 484);
			this.m_cbFuzzyTimes.Name = "m_cbFuzzyTimes";
			this.m_cbFuzzyTimes.Size = new System.Drawing.Size(158, 24);
			this.m_cbFuzzyTimes.TabIndex = 78;
			this.m_cbFuzzyTimes.Text = "Group Time Slots";
			this.m_cbFuzzyTimes.UseVisualStyleBackColor = true;
			// 
			// m_cbDatePivot
			// 
			this.m_cbDatePivot.AutoSize = true;
			this.m_cbDatePivot.Location = new System.Drawing.Point(922, 484);
			this.m_cbDatePivot.Name = "m_cbDatePivot";
			this.m_cbDatePivot.Size = new System.Drawing.Size(130, 24);
			this.m_cbDatePivot.TabIndex = 79;
			this.m_cbDatePivot.Text = "Pivot on Date";
			this.m_cbDatePivot.UseVisualStyleBackColor = true;
			// 
			// m_cbSplitSports
			// 
			this.m_cbSplitSports.AutoSize = true;
			this.m_cbSplitSports.Location = new System.Drawing.Point(674, 513);
			this.m_cbSplitSports.Name = "m_cbSplitSports";
			this.m_cbSplitSports.Size = new System.Drawing.Size(230, 24);
			this.m_cbSplitSports.TabIndex = 80;
			this.m_cbSplitSports.Text = "Automatic Softball/Baseball";
			this.m_cbSplitSports.UseVisualStyleBackColor = true;
			// 
			// label18
			// 
			this.label18.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.label18.Location = new System.Drawing.Point(50, 716);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(1102, 28);
			this.label18.TabIndex = 82;
			this.label18.Tag = "Site roster report";
			this.label18.Click += new System.EventHandler(this.label18_Click);
			this.label18.Paint += new System.Windows.Forms.PaintEventHandler(this.RenderHeadingLine);
			// 
			// m_chlbxRoster
			// 
			this.m_chlbxRoster.CheckOnClick = true;
			this.m_chlbxRoster.FormattingEnabled = true;
			this.m_chlbxRoster.Location = new System.Drawing.Point(70, 742);
			this.m_chlbxRoster.Name = "m_chlbxRoster";
			this.m_chlbxRoster.Size = new System.Drawing.Size(581, 73);
			this.m_chlbxRoster.TabIndex = 83;
			// 
			// m_pbCreateRosterReport
			// 
			this.m_pbCreateRosterReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.m_pbCreateRosterReport.Location = new System.Drawing.Point(985, 792);
			this.m_pbCreateRosterReport.Name = "m_pbCreateRosterReport";
			this.m_pbCreateRosterReport.Size = new System.Drawing.Size(176, 40);
			this.m_pbCreateRosterReport.TabIndex = 84;
			this.m_pbCreateRosterReport.Text = "Create Roster";
			this.m_pbCreateRosterReport.Click += new System.EventHandler(this.GenSiteRosterReport);
			// 
			// m_cbFilterRank
			// 
			this.m_cbFilterRank.AutoSize = true;
			this.m_cbFilterRank.Location = new System.Drawing.Point(678, 628);
			this.m_cbFilterRank.Name = "m_cbFilterRank";
			this.m_cbFilterRank.Size = new System.Drawing.Size(125, 24);
			this.m_cbFilterRank.TabIndex = 85;
			this.m_cbFilterRank.Text = "Filter by rank";
			this.m_cbFilterRank.UseVisualStyleBackColor = true;
			// 
			// m_pbMailMerge
			// 
			this.m_pbMailMerge.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.m_pbMailMerge.Location = new System.Drawing.Point(976, 678);
			this.m_pbMailMerge.Name = "m_pbMailMerge";
			this.m_pbMailMerge.Size = new System.Drawing.Size(176, 40);
			this.m_pbMailMerge.TabIndex = 86;
			this.m_pbMailMerge.Text = "Gen Help";
			this.m_pbMailMerge.Click += new System.EventHandler(this.GenMailMergeMail);
			// 
			// m_cbFutureOnly
			// 
			this.m_cbFutureOnly.AutoSize = true;
			this.m_cbFutureOnly.Location = new System.Drawing.Point(206, 155);
			this.m_cbFutureOnly.Name = "m_cbFutureOnly";
			this.m_cbFutureOnly.Size = new System.Drawing.Size(138, 24);
			this.m_cbFutureOnly.TabIndex = 87;
			this.m_cbFutureOnly.Text = "Future Games";
			this.m_cbFutureOnly.UseVisualStyleBackColor = true;
			// 
			// label19
			// 
			this.label19.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.label19.Location = new System.Drawing.Point(50, 855);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(1102, 28);
			this.label19.TabIndex = 88;
			this.label19.Tag = "Data/Upload Operations";
			this.label19.Paint += new System.Windows.Forms.PaintEventHandler(this.RenderHeadingLine);
			// 
			// m_cbLaunch
			// 
			this.m_cbLaunch.AutoSize = true;
			this.m_cbLaunch.Location = new System.Drawing.Point(678, 659);
			this.m_cbLaunch.Name = "m_cbLaunch";
			this.m_cbLaunch.Size = new System.Drawing.Size(147, 24);
			this.m_cbLaunch.TabIndex = 89;
			this.m_cbLaunch.Text = "Launch MMDoc";
			this.m_cbLaunch.UseVisualStyleBackColor = true;
			// 
			// m_cbSetArbiterAnnounce
			// 
			this.m_cbSetArbiterAnnounce.AutoSize = true;
			this.m_cbSetArbiterAnnounce.Location = new System.Drawing.Point(843, 628);
			this.m_cbSetArbiterAnnounce.Name = "m_cbSetArbiterAnnounce";
			this.m_cbSetArbiterAnnounce.Size = new System.Drawing.Size(188, 24);
			this.m_cbSetArbiterAnnounce.TabIndex = 90;
			this.m_cbSetArbiterAnnounce.Text = "Set Arbiter Announce";
			this.m_cbSetArbiterAnnounce.UseVisualStyleBackColor = true;
			// 
			// m_pbEditProfile
			// 
			this.m_pbEditProfile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.m_pbEditProfile.Location = new System.Drawing.Point(390, 37);
			this.m_pbEditProfile.Name = "m_pbEditProfile";
			this.m_pbEditProfile.Size = new System.Drawing.Size(176, 35);
			this.m_pbEditProfile.TabIndex = 91;
			this.m_pbEditProfile.Text = "Edit Profile";
			this.m_pbEditProfile.Click += new System.EventHandler(this.EditProfile);
			// 
			// m_pbAddProfile
			// 
			this.m_pbAddProfile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.m_pbAddProfile.Location = new System.Drawing.Point(577, 38);
			this.m_pbAddProfile.Name = "m_pbAddProfile";
			this.m_pbAddProfile.Size = new System.Drawing.Size(176, 35);
			this.m_pbAddProfile.TabIndex = 92;
			this.m_pbAddProfile.Text = "Add Profile";
			this.m_pbAddProfile.Click += new System.EventHandler(this.AddProfile);
			// 
			// m_cbxGameFilter
			// 
			this.m_cbxGameFilter.FormattingEnabled = true;
			this.m_cbxGameFilter.Location = new System.Drawing.Point(674, 113);
			this.m_cbxGameFilter.Name = "m_cbxGameFilter";
			this.m_cbxGameFilter.Size = new System.Drawing.Size(208, 28);
			this.m_cbxGameFilter.TabIndex = 93;
			// 
			// m_pbRefreshGameFilters
			// 
			this.m_pbRefreshGameFilters.Font = new System.Drawing.Font("Segoe UI Symbol", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.m_pbRefreshGameFilters.Location = new System.Drawing.Point(886, 107);
			this.m_pbRefreshGameFilters.Margin = new System.Windows.Forms.Padding(0);
			this.m_pbRefreshGameFilters.Name = "m_pbRefreshGameFilters";
			this.m_pbRefreshGameFilters.Padding = new System.Windows.Forms.Padding(3, 0, 0, 0);
			this.m_pbRefreshGameFilters.Size = new System.Drawing.Size(47, 44);
			this.m_pbRefreshGameFilters.TabIndex = 94;
			this.m_pbRefreshGameFilters.Text = "🔁";
			this.m_pbRefreshGameFilters.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.m_pbRefreshGameFilters.UseVisualStyleBackColor = true;
			this.m_pbRefreshGameFilters.Click += new System.EventHandler(this.RefreshGameFilters);
			// 
			// pbTestDownload
			// 
			this.pbTestDownload.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.pbTestDownload.Location = new System.Drawing.Point(133, 80);
			this.pbTestDownload.Name = "pbTestDownload";
			this.pbTestDownload.Size = new System.Drawing.Size(176, 35);
			this.pbTestDownload.TabIndex = 95;
			this.pbTestDownload.Text = "Test Download";
			this.pbTestDownload.Click += new System.EventHandler(this.DoTestDownload);
			// 
			// m_cbSkipContactDownload
			// 
			this.m_cbSkipContactDownload.AutoSize = true;
			this.m_cbSkipContactDownload.Location = new System.Drawing.Point(826, 155);
			this.m_cbSkipContactDownload.Name = "m_cbSkipContactDownload";
			this.m_cbSkipContactDownload.Size = new System.Drawing.Size(151, 24);
			this.m_cbSkipContactDownload.TabIndex = 96;
			this.m_cbSkipContactDownload.Text = "Skip Contact DL";
			this.m_cbSkipContactDownload.UseVisualStyleBackColor = true;
			// 
			// AwMainForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(8, 19);
			this.ClientSize = new System.Drawing.Size(1174, 1173);
			this.Controls.Add(this.m_cbSkipContactDownload);
			this.Controls.Add(this.pbTestDownload);
			this.Controls.Add(this.m_pbRefreshGameFilters);
			this.Controls.Add(this.m_cbxGameFilter);
			this.Controls.Add(this.m_pbAddProfile);
			this.Controls.Add(this.m_pbEditProfile);
			this.Controls.Add(this.m_cbSetArbiterAnnounce);
			this.Controls.Add(this.m_cbLaunch);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.m_cbFutureOnly);
			this.Controls.Add(this.m_pbMailMerge);
			this.Controls.Add(this.m_cbFilterRank);
			this.Controls.Add(this.m_pbCreateRosterReport);
			this.Controls.Add(this.m_chlbxRoster);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.m_cbSplitSports);
			this.Controls.Add(this.m_cbDatePivot);
			this.Controls.Add(this.m_cbFuzzyTimes);
			this.Controls.Add(this.m_ebAffiliationIndex);
			this.Controls.Add(label17);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.m_pbBrowseAnalysis);
			this.Controls.Add(this.m_pbBrowseGamesReport);
			this.Controls.Add(this.m_cbAddOfficialsOnly);
			this.Controls.Add(this.m_pbGenGames);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.m_ebGameOutput);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.m_cbRankOnly);
			this.Controls.Add(this.m_pbReload);
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
			this.Controls.Add(this.m_pbDownloadGames);
			this.Name = "AwMainForm";
			this.Text = "AwMainForm";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DoSaveState);
			this.Load += new System.EventHandler(this.AwMainForm_Load);
			this.Move += new System.EventHandler(this.AwMainForm_Move);
			this.groupBox2.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] rgsCmdLine)
        {
            Application.Run(new AwMainForm(rgsCmdLine));
        }


        #region Operation Implementation

        /* D O  D O W N L O A D  G A M E S */
        /*----------------------------------------------------------------------------
        	%%Function: DoDownloadGames
        	%%Qualified: ArbWeb.AwMainForm.DoDownloadGames
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void HandleDownloadGamesClick(object sender, EventArgs e)
        {
            DoDownloadGames();
        }


        private void HandleDownloadContactsClick(object sender, EventArgs e)
        {
            DoDownloadContacts();
        }

        /* D O  D O W N L O A D  R O S T E R */
        /*----------------------------------------------------------------------------
	    	%%Function: DoDownloadRoster
	    	%%Qualified: ArbWeb.AwMainForm.DoDownloadRoster
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
        private void HandleDownloadRosterClick(object sender, EventArgs e)
        {
            DoDownloadRoster();
        }

        /* D O  D O W N L O A D  Q U I C K  R O S T E R */
        /*----------------------------------------------------------------------------
        	%%Function: DoDownloadQuickRoster
        	%%Qualified: ArbWeb.AwMainForm.DoDownloadQuickRoster
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void HandleDownloadQuickRosterClick(object sender, EventArgs e)
        {
            DoDownloadQuickRoster();
        }


        delegate void HandleRosterDel(Roster rst, string sInFile, Roster rstServer, HandleGenericRoster.HandleRosterPostUpdateDelegate hrpu);


        /* D O  U P L O A D  R O S T E R */
        /*----------------------------------------------------------------------------
			%%Function: DoUploadRoster
			%%Qualified: ArbWeb.AwMainForm.DoUploadRoster
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        private void HandleUploadRosterClick(object sender, EventArgs e)
        {
            DoRosterUpload();
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

            CountsData gc = GcEnsure(m_pr.RosterWorking, m_pr.GameCopy, m_cbIncludeCanceled.Checked);

            gc.GenAnalysis(m_ebOutputFile.Text);
            m_srpt.PopLevel();
            m_srpt.AddMessage("Analysis complete.");
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

        /* G E N  S I T E  R O S T E R  R E P O R T */
        /*----------------------------------------------------------------------------
            %%Function: GenSiteRosterReport
            %%Qualified: ArbWeb.AwMainForm.GenSiteRosterReport
            %%Contact: rlittle
            
        ----------------------------------------------------------------------------*/
        private void GenSiteRosterReport(object sender, EventArgs e)
        {
            DoGenSiteRosterReport();
        }


        /* G E N  M A I L  M E R G E  M A I L */
        /*----------------------------------------------------------------------------
            %%Function: GenMailMergeMail
            %%Qualified: ArbWeb.AwMainForm.GenMailMergeMail
            %%Contact: rlittle
            
        ----------------------------------------------------------------------------*/
        private void GenMailMergeMail(object sender, EventArgs e)
        {
            DoGenMailMergeAndAnnouce();
        }

        /* G E N  O P E N  S L O T S  M A I L */
        /*----------------------------------------------------------------------------
            %%Function: GenOpenSlotsMail
            %%Qualified: ArbWeb.AwMainForm.GenOpenSlotsMail
            %%Contact: rlittle
            
        ----------------------------------------------------------------------------*/
        private void GenOpenSlotsMail(object sender, EventArgs e)
        {
            DoGenOpenSlotsMail();
        }

        /* C A L C  O P E N  S L O T S */
        /*----------------------------------------------------------------------------
        	%%Function: CalcOpenSlots
        	%%Qualified: ArbWeb.AwMainForm.CalcOpenSlots
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void CalcOpenSlots(object sender, EventArgs e)
        {
            DoCalcOpenSlots();
        }


        /* D O  S P O R T  L E V E L  F I L T E R */
        /*----------------------------------------------------------------------------
            %%Function: DoSportLevelFilter
            %%Qualified: ArbWeb.AwMainForm.DoSportLevelFilter
            %%Contact: rlittle
            
        ----------------------------------------------------------------------------*/
        private void DoSportLevelFilter(object sender, ItemCheckEventArgs e)
        {
            CountsData gc = GcEnsure(m_pr.RosterWorking, m_pr.GameCopy, m_cbIncludeCanceled.Checked);
            string[] rgsSports = ArbWebControl.RgsFromChlbx(true, m_chlbxSports, e.Index, e.CurrentValue != CheckState.Checked, null, false);
            string[] rgsSportLevels = ArbWebControl.RgsFromChlbx(true, m_chlbxSportLevels);
            ArbWebControl.UpdateChlbxFromRgs(m_chlbxSportLevels, gc.GetOpenSlotSportLevels(m_saOpenSlots), rgsSportLevels, rgsSports, false);
        }

        /* M _ P B  R E L O A D _  C L I C K */
        /*----------------------------------------------------------------------------
        	%%Function: m_pbReload_Click
        	%%Qualified: ArbWeb.AwMainForm.m_pbReload_Click
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoReloadClick(object sender, EventArgs e)
        {
            InvalRoster();
            InvalGameCount();
        }

        /* D O  G A M E S  R E P O R T */
        /*----------------------------------------------------------------------------
        	%%Function: DoGamesReport
        	%%Qualified: ArbWeb.AwMainForm.DoGamesReport
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoGamesReport(object sender, EventArgs e)
        {
            m_srpt.AddMessage(String.Format("Generating games report ({0})...", m_ebGameOutput.Text));
            m_srpt.PushLevel();
            CountsData gc = GcEnsure(m_pr.RosterWorking, m_pr.GameCopy, m_cbIncludeCanceled.Checked);
            Roster rst = RstEnsure(m_pr.RosterWorking);

            gc.GenGamesReport(m_ebGameOutput.Text);
            m_srpt.PopLevel();
            m_srpt.AddMessage("Games report complete.");
        }

        /* M _ C B  L O G  T O  F I L E _  C H E C K E D  C H A N G E D */
        /*----------------------------------------------------------------------------
        	%%Function: m_cbLogToFile_CheckedChanged
        	%%Qualified: ArbWeb.AwMainForm.m_cbLogToFile_CheckedChanged
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void m_cbLogToFile_CheckedChanged(object sender, EventArgs e)
        {
            SetupLogToFile();
        }

        /* E B  F R O M  F N C */
        /*----------------------------------------------------------------------------
        	%%Function: EbFromFnc
        	%%Qualified: ArbWeb.AwMainForm.EbFromFnc
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        TextBox EbFromFnc(EditProfile.FNC fnc)
        {
            switch (fnc)
                {
                case ArbWeb.EditProfile.FNC.AnalysisFile:
                    return m_ebOutputFile;
                case ArbWeb.EditProfile.FNC.ReportFile:
                    return m_ebGameOutput;
                }

            return null;
        }

        /* D O  B R O W S E  O P E N */
        /*----------------------------------------------------------------------------
        	%%Function: DoBrowseOpen
        	%%Qualified: ArbWeb.AwMainForm.DoBrowseOpen
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoBrowseOpen(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            TextBox eb = EbFromFnc((EditProfile.FNC) (((Button) sender).Tag));

            ofd.InitialDirectory = Path.GetDirectoryName(eb.Text);
            if (ofd.ShowDialog() == DialogResult.OK)
                {
                eb.Text = ofd.FileName;
                }
        }

        private void RefreshGameFilters(object sender, EventArgs e)
        {
            EnsureLoggedIn();

            Dictionary<string, string> mpFilters = MpFetchGameFilters();
            SetGameFiltersFromEnumerable(m_cbxGameFilter, mpFilters.Keys);
            m_pr.GameFilters = mpFilters.Keys.ToArray();
            m_pr.GameFilter = (string) m_cbxGameFilter.SelectedItem;
        }

        #endregion

        #region Support Functions

        private void SetGameFiltersFromEnumerable(ComboBox cbx, IEnumerable<string> iens, string sNewFilter = null)
        {
            string sCurFilter = sNewFilter;
            if (sCurFilter == null && m_cbxGameFilter.SelectedItem != null)
                {
                sCurFilter = (string) m_cbxGameFilter.SelectedItem;
                }

            if (sCurFilter == null)
                sCurFilter = "All Games";

            m_cbxGameFilter.Items.Clear();
            if (iens == null)
                return;

            int iSelected = -1;
            int i = 0;
            foreach (string s in iens)
                {
                m_cbxGameFilter.Items.Add(s);
                if (string.Compare(s, sCurFilter, true /*fIgnoreCase*/) == 0)
                    iSelected = i;
                i++;
                }

            if (iSelected != -1)
                m_cbxGameFilter.SelectedIndex = iSelected;
        }

        /* S E T U P  L O G  T O  F I L E */
        /*----------------------------------------------------------------------------
            %%Function: SetupLogToFile
            %%Qualified: ArbWeb.AwMainForm.SetupLogToFile
            %%Contact: rlittle
            
        ----------------------------------------------------------------------------*/
        private void SetupLogToFile()
        {
            if (m_pr.LogToFile)
                {
                m_srpt.AttachLogfile(Filename.SBuildTempFilename("arblog", "log"));
                m_srpt.SetLogLevel(5);
                m_srpt.SetFilter(StatusRpt.MSGT.Body);
                }
            else
                {
                m_srpt.SetLogLevel(0);
                m_srpt.SetFilter(StatusRpt.MSGT.Error);
                }
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

        #endregion

        #region Profiles

        private Profile m_pr;

        void SetUIForProfile(Profile pr)
        {
            m_ebGameOutput.Text = pr.GameOutput;
            m_ebOutputFile.Text = pr.OutputFile;
            m_cbIncludeCanceled.Checked = pr.IncludeCanceled;
            m_cbShowBrowser.Checked = pr.ShowBrowser;
            try
                {
                m_dtpStart.Value = pr.Start;
                }
            catch
                {
                m_dtpStart.Value = DateTime.Today;
                }

            try
                {
                m_dtpEnd.Value = pr.End;
                }
            catch
                {
                m_dtpEnd.Value = DateTime.Today;
                }

            m_cbOpenSlotDetail.Checked = pr.OpenSlotDetail;
            m_cbFuzzyTimes.Checked = pr.FuzzyTimes;
            m_cbTestEmail.Checked = pr.TestEmail;
            m_cbAddOfficialsOnly.Checked = pr.AddOfficialsOnly;
            m_ebAffiliationIndex.Text = pr.AffiliationIndex;
            m_cbSplitSports.Checked = pr.SplitSports;
            m_cbDatePivot.Checked = pr.DatePivot;
            m_cbFilterRank.Checked = pr.FilterRank;
            m_cbFutureOnly.Checked = pr.FutureOnly;
            m_cbLaunch.Checked = pr.Launch;
            m_cbSetArbiterAnnounce.Checked = pr.SetArbiterAnnounce;
            SetGameFiltersFromEnumerable(m_cbxGameFilter, pr.GameFilters, pr.GameFilter);
            if (pr.IsWindowPosSet(pr.MainWindow))
                {
                this.StartPosition = FormStartPosition.Manual;
                // this.Bounds = pr.MainWindow;
                }
            else
                {
                this.StartPosition = FormStartPosition.WindowsDefaultLocation;
                }

            if (pr.IsWindowPosSet(pr.DiagWindow))
                {
                m_awc.StartPosition = FormStartPosition.Manual;
                m_awc.Bounds = pr.DiagWindow;
                }
            else
                {
                m_awc.StartPosition = FormStartPosition.WindowsDefaultLocation;
                }
        }

        void UpdateProfileFromUI(Profile pr)
	    {
	        pr.GameOutput = m_ebGameOutput.Text;
	        pr.OutputFile = m_ebOutputFile.Text;
	        pr.IncludeCanceled = m_cbIncludeCanceled.Checked;
	        pr.ShowBrowser = m_cbShowBrowser.Checked;
	        pr.Start = m_dtpStart.Value;
	        pr.End = m_dtpEnd.Value;
	        pr.OpenSlotDetail = m_cbOpenSlotDetail.Checked;
	        pr.FuzzyTimes = m_cbFuzzyTimes.Checked;
	        pr.TestEmail = m_cbTestEmail.Checked;
	        pr.AddOfficialsOnly = m_cbAddOfficialsOnly.Checked;
	        pr.AffiliationIndex = m_ebAffiliationIndex.Text;
	        pr.SplitSports = m_cbSplitSports.Checked;
	        pr.DatePivot = m_cbDatePivot.Checked;
	        pr.FilterRank = m_cbFilterRank.Checked;
	        pr.FutureOnly = m_cbFutureOnly.Checked;
	        pr.Launch = m_cbLaunch.Checked;
	        pr.SetArbiterAnnounce = m_cbSetArbiterAnnounce.Checked;
	        pr.GameFilter = (string) m_cbxGameFilter.SelectedItem;
            // don't worry about setting GameFilters -- we already set that when we populated it.
            pr.MainWindow = this.Bounds;
            pr.DiagWindow = m_awc.Bounds;
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
				m_pr = new Profile();
                m_pr.Load(m_cbxProfile.Text);
                SetUIForProfile(m_pr);
				}
		}

	    /* D O  S A V E  S T A T E  C O R E */
	    /*----------------------------------------------------------------------------
	    	%%Function: DoSaveStateCore
	    	%%Qualified: ArbWeb.AwMainForm.DoSaveStateCore
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
	    private void DoSaveStateCore()
		{
	        if (!m_fAutomating)
	            {
	            UpdateProfileFromUI(m_pr);
	            m_pr.Save();
	            m_reh.Save();
	            }
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
               
            if (m_pr.FCompareProfile(m_cbxProfile.Text))
				return;

			// otherwise, this is a new profile
			ChangeProfile(sender, e);
        }

        private void EditProfile(object sender, EventArgs e)
        {
            ArbWeb.EditProfile.FShowEditProfile(m_pr);
        }

        private void AddProfile(object sender, EventArgs e)
        {
            string sProfileName = ArbWeb.EditProfile.AddProfile();
            m_cbxProfile.Items.Add(sProfileName);
            m_cbxProfile.SelectedIndex = m_cbxProfile.Items.Count - 1;
            ChangeProfile(null, null);
        }

        #endregion

        private void DoTestDownload(object sender, EventArgs e)
        {
            var x = m_awc.Handle;
            // let's make sure the webbrowser handle is created

            m_srpt.LogData("Testing Download", 3, StatusRpt.MSGT.Header);

            Task tskDownloadTest = new Task(() => TestDownload());

            tskDownloadTest.Start();
            // DownloadGames();
        }

        private void AwMainForm_Move(object sender, EventArgs e)
        {
            m_srpt.LogData("Moving",10,StatusRpt.MSGT.Body);
        }

        private void AwMainForm_Load(object sender, EventArgs e)
        {

        }

        private void RenderHeadingLine(object sender, PaintEventArgs e)
        {
            RenderSupp.RenderHeadingLine(sender, e);
        }

		private void label18_Click(object sender, EventArgs e)
		{

		}
	}

}
