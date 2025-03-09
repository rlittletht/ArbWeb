using System;
using System.Drawing;
using TCore.Settings;

namespace ArbWeb
{
    public class Profile
    {
        private Settings.SettingsElt[] m_rgreheProfile;
        private Settings m_ehProfile;

        public bool IsWindowPosSet(Rectangle rect)
        {
            return rect.Width != 0;
        }

        [Setting("BaseballScheduleFiles", new string[] { }, new string[] { })]
        public string[] BaseballSchedFiles { get; set; }

        [Setting("SoftballScheduleFiles", new string[] { }, new string[] { })]
        public string[] SoftballSchedFiles { get; set; }

        [Setting("SchedSpoSite", "", "")] public string SchedSpoSite { get; set; }
        [Setting("SchedSpoSubsite", "", "")] public string SchedSpoSubsite { get; set; }

        [Setting("SchedDownloadFolder", "", "")]
        public string SchedDownloadFolder { get; set; }

        [Setting("SchedWorkingFolder", "", "")]
        public string SchedWorkingFolder { get; set; }

        [Setting("NoHonorificRanks", false, 0)]
        public bool NoHonorificRanks { get; set; }

        [Setting("LastLogLevel", 0, 0)] public int LogLevel { get; set; }

        [Setting("GameFiltersCache", new string[] { }, new string[] { })]
        public string[] GameFilters { get; set; }

        [Setting("LastGameFilter", "", "All Games")]
        public string GameFilter { get; set; }

        [Setting("Login", "", "")] public string UserID { get; set; }
        [Setting("Password", "", "")] public string Password { get; set; }
        [Setting("GameFile", "", "")] public string GameFile { get; set; }
        [Setting("Contacts", "", "")] public string Contacts { get; set; }
        [Setting("ContactsCopy", "", "")] public string ContactsWorking { get; set; }
        [Setting("Roster", "", "")] public string Roster { get; set; }
        [Setting("GameFileCopy", "", "")] public string GameCopy { get; set; }
        [Setting("RosterCopy", "", "")] public string RosterWorking { get; set; }
        [Setting("Announcements", "", "")] public string Announcements { get; set; }
        [Setting("AnnouncementsCopy", "", "")] public string AnnouncementsWorking { get; set; }
        [Setting("GameOutput", "", "")] public string GameOutput { get; set; }
        [Setting("OutputFile", "", "")] public string OutputFile { get; set; }
        [Setting("IncludeCanceled", false, 0)] public bool IncludeCanceled { get; set; }
        [Setting("ShowBrowser", false, 0)] public bool ShowBrowser { get; set; }
        [Setting("LastSlotStartDate", "", "")] public DateTime Start { get; set; }
        [Setting("LastSlotEndDate", "", "")] public DateTime End { get; set; }
        [Setting("MergeCsv", "", "")] public string MergeCsv { get; set; }
        [Setting("MergeCsvCopy", "", "")] public string MergeCsvWorking { get; set; }

        [Setting("LastOpenSlotDetail", false, 0)]
        public bool OpenSlotDetail { get; set; }

        [Setting("LastGroupTimeSlots", false, 0)]
        public bool FuzzyTimes { get; set; }

        [Setting("LastTestEmail", false, 0)] public bool TestEmail { get; set; }
        [Setting("LastLogToFile", false, 0)] public bool LogToFile { get; set; }

        [Setting("FilterMailMergeToAllStars", false, 0)]
        public bool FilterAllStarsOnly{ get; set; }

        [Setting("FilterMailMergeByRank", false, 0)]
        public bool FilterRank { get; set; }

        [Setting("AddOfficialsOnly", false, 0)]
        public bool AddOfficialsOnly { get; set; }

        [Setting("AfiliationIndex", 0, 0)] public int AffiliationIndex { get; set; }
        [Setting("LastSplitSports", false, 0)] public bool SplitSports { get; set; }
        [Setting("LastPivotDate", false, 0)] public bool DatePivot { get; set; }

        [Setting("LaunchMailMergeDoc", false, 0)]
        public bool Launch { get; set; }

        [Setting("SetArbiterAnnouncement", false, 0)]
        public bool SetArbiterAnnounce { get; set; }

        [Setting("TestOnly", false, 0)] public bool TestOnly { get; set; }
        [Setting("SkipZSports", false, 0)] public bool SkipZ { get; set; }

        [Setting("DownloadRosterOnUpload", false, 0)]
        public bool DownloadRosterOnUpload { get; set; }

        [Setting("LastMainWindowPos.Top", 0, 0)]
        public int MainWindowTop { get; set; }

        [Setting("LastMainWindowPos.Left", 0, 0)]
        public int MainWindowLeft { get; set; }

        [Setting("LastMainWindowPos.Width", 0, 0)]
        public int MainWindowWidth { get; set; }

        [Setting("LastMainWindowPos.Height", 0, 0)]
        public int MainWindowHeight { get; set; }

        [Setting("LastProfileWindowPos.Top", 0, 0)]
        public int ProfileWindowTop { get; set; }

        [Setting("LastProfileWindowPos.Left", 0, 0)]
        public int ProfileWindowLeft { get; set; }

        [Setting("LastProfileWindowPos.Width", 0, 0)]
        public int ProfileWindowWidth { get; set; }

        [Setting("LastProfileWindowPos.Height", 0, 0)]
        public int ProfileWindowHeight { get; set; }

        [Setting("AllowAdvancedArbiterFunctions", false, 0)]
        public bool AllowAdvancedArbiterFunctions { get; set; }

        public Rectangle MainWindow { get; set; }
        public Rectangle ProfileWindow { get; set; }

        public string ProfileName { get; set; }

        public Profile()
        {
            m_rgreheProfile = Settings.SettingsElt.CreateSettings<Profile>();
        }

        void LoadDataFromSettings()
        {
            m_ehProfile.SynchronizeGetValues(this);

            MainWindow = new Rectangle(MainWindowLeft, MainWindowTop, MainWindowWidth, MainWindowHeight);
            ProfileWindow = new Rectangle(ProfileWindowLeft, ProfileWindowTop, ProfileWindowWidth, ProfileWindowHeight);
        }

        void SetSettingsFromData()
        {
            MainWindowLeft = MainWindow.Left;
            MainWindowTop = MainWindow.Top;
            MainWindowWidth = MainWindow.Width;
            MainWindowHeight = MainWindow.Height;

            ProfileWindowLeft = ProfileWindow.Left;
            ProfileWindowTop = ProfileWindow.Top;
            ProfileWindowWidth = ProfileWindow.Width;
            ProfileWindowHeight = ProfileWindow.Height;

            m_ehProfile.SynchronizeSetValues(this);
        }

        public void Load(string sProfileName)
        {
            string sRoot = $"Software\\Thetasoft\\ArbWeb\\{sProfileName}";
            m_ehProfile = new Settings(m_rgreheProfile, sRoot, sRoot);
            ProfileName = sProfileName;
            m_ehProfile.Load();
            LoadDataFromSettings();
        }

        public bool FCompareProfile(string sProfile)
        {
            return String.Compare(ProfileName, sProfile, StringComparison.OrdinalIgnoreCase) == 0;
        }

        public void Save()
        {
            if (m_ehProfile == null)
            {
                // must be a new profile
                string sRoot = $"Software\\Thetasoft\\ArbWeb\\{ProfileName}";
                m_ehProfile = new Settings(m_rgreheProfile, sRoot, sRoot);
            }

            SetSettingsFromData();
            m_ehProfile.Save();
        }
    }
}
