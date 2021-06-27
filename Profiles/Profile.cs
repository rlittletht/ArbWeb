using System;
using System.Drawing;
using TCore.Settings;

namespace ArbWeb
{
    public class Profile
    {
        private Settings.SettingsElt[] m_rgreheProfile;
        private Settings m_ehProfile;

        private string m_sUserID;
        private string m_sPassword;
        private string m_sGameFile;
        private string m_sRoster;
        private string m_sGameCopy;
        private string m_sRosterWorking;
        private string m_sGameOutput;
        private string m_sOutputFile;
        private bool m_fIncludeCanceled;
        private bool m_fShowBrowser;
        private DateTime m_dttmStart;
        private DateTime m_dttmEnd;
        private bool m_fOpenSlotDetail;
        private bool m_fFuzzyTimes;
        private bool m_fTestEmail;
        private bool m_fLogToFile;
        private bool m_fFilterRank;
        private bool m_fAddOfficialsOnly;
        private string m_sAffiliationIndex;
        private bool m_fSplitSports;
        private bool m_fDatePivot;
        private bool m_fFutureOnly;
        private bool m_fLaunch;
        private bool m_fSetArbiterAnnounce;
        private bool m_fTestOnly;
        private bool m_fSkipZ;
        private bool m_fDownloadRosterOnUpload;
        private string m_sProfileName;
        private string[] m_rgsGameFilters;
        private string m_sGameFilter;
        private int m_nLogLevel;
        private string m_sContacts;
        private string m_sContactsWorking;
        private Rectangle m_rectMainWindow;
        private Rectangle m_rectDiagWindow;
        private Rectangle m_rectProfileWindow;

        public Rectangle MainWindow { get { return m_rectMainWindow; } set { m_rectMainWindow = value; } }
        public Rectangle DiagWindow { get { return m_rectDiagWindow; } set { m_rectDiagWindow = value; } }
        public Rectangle ProfileWindow { get { return m_rectProfileWindow; } set { m_rectProfileWindow = value; } }

        public bool IsWindowPosSet(Rectangle rect)
        {
            return rect.Width != 0;
        }

        public string[] BaseballSchedFiles { get; set; }
        public string[] SoftballSchedFiles { get; set; }
        public string SchedSpoSite { get; set; }
        public string SchedSpoSubsite { get; set; }
        public string SchedDownloadFolder { get; set; }
        public string SchedWorkingFolder { get; set; }
        public bool NoHonorificRanks { get; set; }

        public int LogLevel {  get { return m_nLogLevel;} set { m_nLogLevel = value; } }
        public string[] GameFilters { get { return m_rgsGameFilters; } set { m_rgsGameFilters = value; } }
        public string GameFilter { get { return m_sGameFilter; } set { m_sGameFilter = value; } }
        public string UserID { get { return m_sUserID; } set { m_sUserID = value; } } 
        public string Password  { get { return m_sPassword; } set { m_sPassword = value; } }
        public string GameFile { get { return m_sGameFile; }  set { m_sGameFile = value; } }
        public string Contacts { get { return m_sContacts; } set { m_sContacts = value; } }
        public string ContactsWorking { get { return m_sContactsWorking; } set { m_sContactsWorking = value; } }
        public string Roster  { get { return m_sRoster; } set { m_sRoster = value; } }
        public string GameCopy { get { return m_sGameCopy; } set { m_sGameCopy = value; } } 
        public string RosterWorking { get { return m_sRosterWorking; } set { m_sRosterWorking = value; } } 
        public string GameOutput { get { return m_sGameOutput; } set { m_sGameOutput = value; } }
        public string OutputFile { get { return m_sOutputFile; } set { m_sOutputFile = value; } }
        public bool IncludeCanceled { get { return m_fIncludeCanceled; } set { m_fIncludeCanceled = value; } }
        public bool ShowBrowser { get { return m_fShowBrowser; } set { m_fShowBrowser = value; } } 
        public DateTime Start { get { return m_dttmStart; } set { m_dttmStart = value; } } 
        public DateTime End { get { return m_dttmEnd; } set { m_dttmEnd = value; } }
        public bool OpenSlotDetail { get { return m_fOpenSlotDetail; } set { m_fOpenSlotDetail = value; } }
        public bool FuzzyTimes { get { return m_fFuzzyTimes; } set { m_fFuzzyTimes = value; } }
        public bool TestEmail { get { return m_fTestEmail; } set { m_fTestEmail = value; } } 
        public bool LogToFile { get { return m_fLogToFile; } set { m_fLogToFile = value; } }
        public bool FilterRank { get { return m_fFilterRank; } set { m_fFilterRank = value; } } 
        public bool AddOfficialsOnly { get { return m_fAddOfficialsOnly; } set { m_fAddOfficialsOnly = value; } }
        public string AffiliationIndex { get { return m_sAffiliationIndex; } set { m_sAffiliationIndex = value; } } 
        public bool SplitSports { get { return m_fSplitSports; } set { m_fSplitSports = value; } }
        public bool DatePivot { get { return m_fDatePivot; } set { m_fDatePivot = value; } }
        public bool FutureOnly { get { return m_fFutureOnly; } set { m_fFutureOnly = value; } }
        public bool Launch { get { return m_fLaunch; } set { m_fLaunch = value; } }
        public bool SetArbiterAnnounce { get { return m_fSetArbiterAnnounce; } set { m_fSetArbiterAnnounce = value; } }
        public string ProfileName {  get { return m_sProfileName; } set { m_sProfileName = value; } }
        public bool TestOnly {  get { return m_fTestOnly;  } set { m_fTestOnly = value; } }
        public bool SkipZ {  get { return m_fSkipZ; } set { m_fSkipZ = value; } }
        public bool DownloadRosterOnUpload {  get { return m_fDownloadRosterOnUpload; } set { m_fDownloadRosterOnUpload = value; } }

        public Profile()
        {
            m_rgreheProfile = new[]
                                  {
                                  new Settings.SettingsElt("Login", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("Password", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("GameFile", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("Contacts", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("Roster", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("GameFileCopy", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("RosterCopy", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("ContactsCopy", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("GameOutput", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("OutputFile", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("IncludeCanceled", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("ShowBrowser", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("LastSlotStartDate", Settings.Type.Dttm, "", ""),
                                  new Settings.SettingsElt("LastSlotEndDate", Settings.Type.Dttm, "", ""),
                                  new Settings.SettingsElt("LastOpenSlotDetail", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("LastGroupTimeSlots", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("LastTestEmail", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("AddOfficialsOnly", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("AfiliationIndex", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastSplitSports", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("LastPivotDate", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("LastLogToFile", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("FilterMailMergeByRank", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("DownloadOnlyFutureGames", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("LaunchMailMergeDoc", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("SetArbiterAnnouncement", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("TestOnly", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("SkipZSports", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("DownloadRosterOnUpload", Settings.Type.Bool, false, 0),
                                  new Settings.SettingsElt("GameFiltersCache", Settings.Type.StrArray, new string[] { }, new string[] { }),
                                  new Settings.SettingsElt("LastGameFilter", Settings.Type.Str, "", "All Games"),
                                  new Settings.SettingsElt("LastLogLevel", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("NoHonorificRanks", Settings.Type.Bool, false, 0),

                                  new Settings.SettingsElt("SchedSpoSite", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("SchedSpoSubsite", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("SchedDownloadFolder", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("SchedWorkingFolder", Settings.Type.Str, "", ""),
                                  new Settings.SettingsElt("SoftballScheduleFiles", Settings.Type.StrArray, new string[] { }, new string[] { }),
                                  new Settings.SettingsElt("BaseballScheduleFiles", Settings.Type.StrArray, new string[] { }, new string[] { }),
                                  new Settings.SettingsElt("LastMainWindowPos.Top", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastMainWindowPos.Left", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastMainWindowPos.Width", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastMainWindowPos.Height", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastDiagWindowPos.Top", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastDiagWindowPos.Left", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastDiagWindowPos.Width", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastDiagWindowPos.Height", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastProfileWindowPos.Top", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastProfileWindowPos.Left", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastProfileWindowPos.Width", Settings.Type.Int, 0, 0),
                                  new Settings.SettingsElt("LastProfileWindowPos.Height", Settings.Type.Int, 0, 0),
                                  };
        }

        void LoadDataFromSettings()
        {
            m_sUserID = m_ehProfile.SValue("Login");
            m_sPassword = m_ehProfile.SValue("Password");
            m_sGameFile = m_ehProfile.SValue("GameFile");
            m_sRoster = m_ehProfile.SValue("Roster");
            m_sContacts = m_ehProfile.SValue("Contacts");
            m_sGameCopy = m_ehProfile.SValue("GameFileCopy");
            m_sRosterWorking = m_ehProfile.SValue("RosterCopy");
            m_sContactsWorking = m_ehProfile.SValue("ContactsCopy");
            m_sGameOutput = m_ehProfile.SValue("GameOutput");
            m_sOutputFile = m_ehProfile.SValue("OutputFile");
            m_fIncludeCanceled = m_ehProfile.FValue("IncludeCanceled");
            m_fShowBrowser = m_ehProfile.FValue("ShowBrowser");
            m_dttmStart = m_ehProfile.DttmValue("LastSlotStartDate");
            m_dttmEnd = m_ehProfile.DttmValue("LastSlotEndDate");
            m_fOpenSlotDetail = m_ehProfile.FValue("LastOpenSlotDetail");
            m_fFuzzyTimes = m_ehProfile.FValue("LastGroupTimeSlots");
            m_fTestEmail = m_ehProfile.FValue("LastTestEmail");
            m_fAddOfficialsOnly = m_ehProfile.FValue("AddOfficialsOnly");
            m_sAffiliationIndex = m_ehProfile.NValue("AfiliationIndex").ToString();
            m_fSplitSports = m_ehProfile.FValue("LastSplitSports");
            m_fDatePivot = m_ehProfile.FValue("LastPivotDate");
            m_fLogToFile = m_ehProfile.FValue("LastLogToFile");
            m_fFilterRank = m_ehProfile.FValue("FilterMailMergeByRank");
            m_fFutureOnly = m_ehProfile.FValue("DownloadOnlyFutureGames");
            m_fLaunch = m_ehProfile.FValue("LaunchMailMergeDoc");
            m_fSetArbiterAnnounce = m_ehProfile.FValue("SetArbiterAnnouncement");
            m_fTestOnly = m_ehProfile.FValue("TestOnly");
            m_fSkipZ = m_ehProfile.FValue("SkipZSports");
            m_fDownloadRosterOnUpload = m_ehProfile.FValue("DownloadRosterOnUpload");
            m_rgsGameFilters = m_ehProfile.RgsValue("GameFiltersCache");
            m_sGameFilter = m_ehProfile.SValue("LastGameFilter");
            m_nLogLevel = m_ehProfile.NValue("LastLogLevel");

            NoHonorificRanks = m_ehProfile.FValue("NoHonorificRanks");
            SchedSpoSite = m_ehProfile.SValue("SchedSpoSite");
            SchedSpoSubsite = m_ehProfile.SValue("SchedSpoSubsite");
            BaseballSchedFiles = m_ehProfile.RgsValue("BaseballScheduleFiles");
            SoftballSchedFiles = m_ehProfile.RgsValue("SoftballScheduleFiles");
            SchedDownloadFolder = m_ehProfile.SValue("SchedDownloadFolder");
            SchedWorkingFolder = m_ehProfile.SValue("SchedWorkingFolder");

            m_rectMainWindow = new Rectangle(m_ehProfile.NValue("LastMainWindowPos.Left"),
                                             m_ehProfile.NValue("LastMainWindowPos.Top"),
                                             m_ehProfile.NValue("LastMainWindowPos.Width"),
                                             m_ehProfile.NValue("LastMainWindowPos.Height"));
            m_rectDiagWindow = new Rectangle(m_ehProfile.NValue("LastDiagWindowPos.Left"),
                                             m_ehProfile.NValue("LastDiagWindowPos.Top"),
                                             m_ehProfile.NValue("LastDiagWindowPos.Width"),
                                             m_ehProfile.NValue("LastDiagWindowPos.Height"));
            m_rectProfileWindow = new Rectangle(m_ehProfile.NValue("LastProfileWindowPos.Left"),
                                             m_ehProfile.NValue("LastProfileWindowPos.Top"),
                                             m_ehProfile.NValue("LastProfileWindowPos.Width"),
                                             m_ehProfile.NValue("LastProfileWindowPos.Height"));
        }

        void SetSettingsFromData()
        {
            m_ehProfile.SetSValue("Login", m_sUserID);
            m_ehProfile.SetSValue("Password", m_sPassword);
            m_ehProfile.SetSValue("GameFile", m_sGameFile);
            m_ehProfile.SetSValue("Roster", m_sRoster);
            m_ehProfile.SetSValue("Contacts", m_sContacts);
            m_ehProfile.SetSValue("GameFileCopy", m_sGameCopy);
            m_ehProfile.SetSValue("RosterCopy", m_sRosterWorking);
            m_ehProfile.SetSValue("ContactsCopy", m_sContactsWorking);
            m_ehProfile.SetSValue("GameOutput", m_sGameOutput);
            m_ehProfile.SetSValue("OutputFile", m_sOutputFile);
            m_ehProfile.SetFValue("IncludeCanceled", m_fIncludeCanceled);
            m_ehProfile.SetFValue("ShowBrowser", m_fShowBrowser);
            m_ehProfile.SetDttmValue("LastSlotStartDate", m_dttmStart);
            m_ehProfile.SetDttmValue("LastSlotEndDate", m_dttmEnd);
            m_ehProfile.SetFValue("LastOpenSlotDetail", m_fOpenSlotDetail);
            m_ehProfile.SetFValue("LastGroupTimeSlots", m_fFuzzyTimes);
            m_ehProfile.SetFValue("LastTestEmail", m_fTestEmail);
            m_ehProfile.SetFValue("AddOfficialsOnly", m_fAddOfficialsOnly);
            m_ehProfile.SetNValue("AfiliationIndex", m_sAffiliationIndex);
            m_ehProfile.SetFValue("LastSplitSports", m_fSplitSports);
            m_ehProfile.SetFValue("LastPivotDate", m_fDatePivot);
            m_ehProfile.SetFValue("LastLogToFile", m_fLogToFile);
            m_ehProfile.SetFValue("FilterMailMergeByRank", m_fFilterRank);
            m_ehProfile.SetFValue("DownloadOnlyFutureGames", m_fFutureOnly);
            m_ehProfile.SetFValue("LaunchMailMergeDoc", m_fLaunch);
            m_ehProfile.SetFValue("SetArbiterAnnouncement", m_fSetArbiterAnnounce);
            m_ehProfile.SetFValue("TestOnly", m_fTestOnly);
            m_ehProfile.SetFValue("SkipZSports", m_fSkipZ);
            m_ehProfile.SetFValue("DownloadRosterOnUpload", m_fDownloadRosterOnUpload);
            m_ehProfile.SetRgsValue("GameFiltersCache", m_rgsGameFilters);
            m_ehProfile.SetSValue("LastGameFilter", m_sGameFilter);
            m_ehProfile.SetNValue("LastLogLevel", m_nLogLevel);
            m_ehProfile.SetFValue("NoHonorificRanks", NoHonorificRanks);

            m_ehProfile.SetSValue("SchedSpoSite", SchedSpoSite);
            m_ehProfile.SetSValue("SchedSpoSubsite", SchedSpoSubsite);
            m_ehProfile.SetSValue("SchedDownloadFolder", SchedDownloadFolder);
            m_ehProfile.SetSValue("SchedWorkingFolder", SchedWorkingFolder);
            m_ehProfile.SetRgsValue("BaseballScheduleFiles", BaseballSchedFiles);
            m_ehProfile.SetRgsValue("SoftballScheduleFiles", SoftballSchedFiles);

            m_ehProfile.SetNValue("LastMainWindowPos.Left", m_rectMainWindow.Left);
            m_ehProfile.SetNValue("LastMainWindowPos.Top", m_rectMainWindow.Top);
            m_ehProfile.SetNValue("LastMainWindowPos.Width", m_rectMainWindow.Width);
            m_ehProfile.SetNValue("LastMainWindowPos.Height", m_rectMainWindow.Height);
            m_ehProfile.SetNValue("LastDiagWindowPos.Left", m_rectDiagWindow.Left);
            m_ehProfile.SetNValue("LastDiagWindowPos.Top", m_rectDiagWindow.Top);
            m_ehProfile.SetNValue("LastDiagWindowPos.Width", m_rectDiagWindow.Width);
            m_ehProfile.SetNValue("LastDiagWindowPos.Height", m_rectDiagWindow.Height);
            m_ehProfile.SetNValue("LastProfileWindowPos.Left", m_rectProfileWindow.Left);
            m_ehProfile.SetNValue("LastProfileWindowPos.Top", m_rectProfileWindow.Top);
            m_ehProfile.SetNValue("LastProfileWindowPos.Width", m_rectProfileWindow.Width);
            m_ehProfile.SetNValue("LastProfileWindowPos.Height", m_rectProfileWindow.Height);
        }

        public void Load(string sProfileName)
        {
            string sRoot = $"Software\\Thetasoft\\ArbWeb\\{sProfileName}";
            m_ehProfile = new Settings(m_rgreheProfile, sRoot, sRoot);
            m_sProfileName = sProfileName;
            m_ehProfile.Load();
            LoadDataFromSettings();
        }

        public bool FCompareProfile(string sProfile)
        {
            return String.Compare(m_sProfileName, sProfile, StringComparison.OrdinalIgnoreCase) == 0;
        }

        public void Save()
        {
            if (m_ehProfile == null)
                {
                // must be a new profile
                string sRoot = $"Software\\Thetasoft\\ArbWeb\\{m_sProfileName}";
                m_ehProfile = new Settings(m_rgreheProfile, sRoot, sRoot);
                }

            SetSettingsFromData();
            m_ehProfile.Save();
        }
    }
}