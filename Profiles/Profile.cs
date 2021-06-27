using System;
using System.Drawing;
using TCore.Settings;

namespace ArbWeb
{
	public class Profile
	{
		private Settings.SettingsElt[] m_rgreheProfile;
		private Settings m_ehProfile;

		private DateTime m_dttmStart;
		private DateTime m_dttmEnd;


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

		public int LogLevel { get; set; }
		public string[] GameFilters { get; set; }
		public string GameFilter { get; set; }
		public string UserID { get; set; }
		public string Password { get; set; }
		public string GameFile { get; set; }
		public string Contacts { get; set; }
		public string ContactsWorking { get; set; }
		public string Roster { get; set; }
		public string GameCopy { get; set; }
		public string RosterWorking { get; set; }
		public string GameOutput { get; set; }
		public string OutputFile { get; set; }
		public bool IncludeCanceled { get; set; }
		public bool ShowBrowser { get; set; }
		public DateTime Start { get; set; }
		public DateTime End { get; set; }
		public bool OpenSlotDetail { get; set; }
		public bool FuzzyTimes { get; set; }
		public bool TestEmail { get; set; }
		public bool LogToFile { get; set; }
		public bool FilterRank { get; set; }
		public bool AddOfficialsOnly { get; set; }
		public string AffiliationIndex { get; set; }
		public bool SplitSports { get; set; }
		public bool DatePivot { get; set; }
		public bool FutureOnly { get; set; }
		public bool Launch { get; set; }
		public bool SetArbiterAnnounce { get; set; }
		public string ProfileName { get; set; }
		public bool TestOnly { get; set; }
		public bool SkipZ { get; set; }
		public bool DownloadRosterOnUpload { get; set; }
		public Rectangle MainWindow { get; set; }
		public Rectangle ProfileWindow { get; set; }

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
				new Settings.SettingsElt("GameFiltersCache", Settings.Type.StrArray, new string[] { },
					new string[] { }),
				new Settings.SettingsElt("LastGameFilter", Settings.Type.Str, "", "All Games"),
				new Settings.SettingsElt("LastLogLevel", Settings.Type.Int, 0, 0),
				new Settings.SettingsElt("NoHonorificRanks", Settings.Type.Bool, false, 0),

				new Settings.SettingsElt("SchedSpoSite", Settings.Type.Str, "", ""),
				new Settings.SettingsElt("SchedSpoSubsite", Settings.Type.Str, "", ""),
				new Settings.SettingsElt("SchedDownloadFolder", Settings.Type.Str, "", ""),
				new Settings.SettingsElt("SchedWorkingFolder", Settings.Type.Str, "", ""),
				new Settings.SettingsElt("SoftballScheduleFiles", Settings.Type.StrArray, new string[] { },
					new string[] { }),
				new Settings.SettingsElt("BaseballScheduleFiles", Settings.Type.StrArray, new string[] { },
					new string[] { }),
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
			UserID = m_ehProfile.SValue("Login");
			Password = m_ehProfile.SValue("Password");
			GameFile = m_ehProfile.SValue("GameFile");
			Roster = m_ehProfile.SValue("Roster");
			Contacts = m_ehProfile.SValue("Contacts");
			GameCopy = m_ehProfile.SValue("GameFileCopy");
			RosterWorking = m_ehProfile.SValue("RosterCopy");
			ContactsWorking = m_ehProfile.SValue("ContactsCopy");
			GameOutput = m_ehProfile.SValue("GameOutput");
			OutputFile = m_ehProfile.SValue("OutputFile");
			IncludeCanceled = m_ehProfile.FValue("IncludeCanceled");
			ShowBrowser = m_ehProfile.FValue("ShowBrowser");
			m_dttmStart = m_ehProfile.DttmValue("LastSlotStartDate");
			m_dttmEnd = m_ehProfile.DttmValue("LastSlotEndDate");
			OpenSlotDetail = m_ehProfile.FValue("LastOpenSlotDetail");
			FuzzyTimes = m_ehProfile.FValue("LastGroupTimeSlots");
			TestEmail = m_ehProfile.FValue("LastTestEmail");
			AddOfficialsOnly = m_ehProfile.FValue("AddOfficialsOnly");
			AffiliationIndex = m_ehProfile.NValue("AfiliationIndex").ToString();
			SplitSports = m_ehProfile.FValue("LastSplitSports");
			DatePivot = m_ehProfile.FValue("LastPivotDate");
			LogToFile = m_ehProfile.FValue("LastLogToFile");
			FilterRank = m_ehProfile.FValue("FilterMailMergeByRank");
			FutureOnly = m_ehProfile.FValue("DownloadOnlyFutureGames");
			Launch = m_ehProfile.FValue("LaunchMailMergeDoc");
			SetArbiterAnnounce = m_ehProfile.FValue("SetArbiterAnnouncement");
			TestOnly = m_ehProfile.FValue("TestOnly");
			SkipZ = m_ehProfile.FValue("SkipZSports");
			DownloadRosterOnUpload = m_ehProfile.FValue("DownloadRosterOnUpload");
			GameFilters = m_ehProfile.RgsValue("GameFiltersCache");
			GameFilter = m_ehProfile.SValue("LastGameFilter");
			LogLevel = m_ehProfile.NValue("LastLogLevel");

			NoHonorificRanks = m_ehProfile.FValue("NoHonorificRanks");
			SchedSpoSite = m_ehProfile.SValue("SchedSpoSite");
			SchedSpoSubsite = m_ehProfile.SValue("SchedSpoSubsite");
			BaseballSchedFiles = m_ehProfile.RgsValue("BaseballScheduleFiles");
			SoftballSchedFiles = m_ehProfile.RgsValue("SoftballScheduleFiles");
			SchedDownloadFolder = m_ehProfile.SValue("SchedDownloadFolder");
			SchedWorkingFolder = m_ehProfile.SValue("SchedWorkingFolder");

			MainWindow = new Rectangle(m_ehProfile.NValue("LastMainWindowPos.Left"),
				m_ehProfile.NValue("LastMainWindowPos.Top"),
				m_ehProfile.NValue("LastMainWindowPos.Width"),
				m_ehProfile.NValue("LastMainWindowPos.Height"));
			ProfileWindow = new Rectangle(m_ehProfile.NValue("LastProfileWindowPos.Left"),
				m_ehProfile.NValue("LastProfileWindowPos.Top"),
				m_ehProfile.NValue("LastProfileWindowPos.Width"),
				m_ehProfile.NValue("LastProfileWindowPos.Height"));
		}

		void SetSettingsFromData()
		{
			m_ehProfile.SetSValue("Login", UserID);
			m_ehProfile.SetSValue("Password", Password);
			m_ehProfile.SetSValue("GameFile",               GameFile);
			m_ehProfile.SetSValue("Roster",                 Roster);
			m_ehProfile.SetSValue("Contacts",               Contacts);
			m_ehProfile.SetSValue("GameFileCopy",           GameCopy);
			m_ehProfile.SetSValue("RosterCopy",             RosterWorking);
			m_ehProfile.SetSValue("ContactsCopy",           ContactsWorking);
			m_ehProfile.SetSValue("GameOutput",             GameOutput);
			m_ehProfile.SetSValue("OutputFile",             OutputFile);
			m_ehProfile.SetFValue("IncludeCanceled",        IncludeCanceled);
			m_ehProfile.SetFValue("ShowBrowser",            ShowBrowser);
			m_ehProfile.SetDttmValue("LastSlotStartDate",   m_dttmStart);
			m_ehProfile.SetDttmValue("LastSlotEndDate",     m_dttmEnd);
			m_ehProfile.SetFValue("LastOpenSlotDetail",     OpenSlotDetail);
			m_ehProfile.SetFValue("LastGroupTimeSlots",     FuzzyTimes);
			m_ehProfile.SetFValue("LastTestEmail",          TestEmail);
			m_ehProfile.SetFValue("AddOfficialsOnly",       AddOfficialsOnly);
			m_ehProfile.SetNValue("AfiliationIndex",        AffiliationIndex);
			m_ehProfile.SetFValue("LastSplitSports",        SplitSports);
			m_ehProfile.SetFValue("LastPivotDate",          DatePivot);
			m_ehProfile.SetFValue("LastLogToFile",          LogToFile);
			m_ehProfile.SetFValue("FilterMailMergeByRank",  FilterRank);
			m_ehProfile.SetFValue("DownloadOnlyFutureGames",FutureOnly);
			m_ehProfile.SetFValue("LaunchMailMergeDoc",     Launch);
			m_ehProfile.SetFValue("SetArbiterAnnouncement", SetArbiterAnnounce);
			m_ehProfile.SetFValue("TestOnly",               TestOnly);
			m_ehProfile.SetFValue("SkipZSports",            SkipZ);
			m_ehProfile.SetFValue("DownloadRosterOnUpload", DownloadRosterOnUpload);
			m_ehProfile.SetRgsValue("GameFiltersCache",     GameFilters);
			m_ehProfile.SetSValue("LastGameFilter",         GameFilter);
			m_ehProfile.SetNValue("LastLogLevel",           LogLevel);
			m_ehProfile.SetFValue("NoHonorificRanks", NoHonorificRanks);

			m_ehProfile.SetSValue("SchedSpoSite", SchedSpoSite);
			m_ehProfile.SetSValue("SchedSpoSubsite", SchedSpoSubsite);
			m_ehProfile.SetSValue("SchedDownloadFolder", SchedDownloadFolder);
			m_ehProfile.SetSValue("SchedWorkingFolder", SchedWorkingFolder);
			m_ehProfile.SetRgsValue("BaseballScheduleFiles", BaseballSchedFiles);
			m_ehProfile.SetRgsValue("SoftballScheduleFiles", SoftballSchedFiles);

			m_ehProfile.SetNValue("LastMainWindowPos.Left",     MainWindow.Left);
			m_ehProfile.SetNValue("LastMainWindowPos.Top",      MainWindow.Top);
			m_ehProfile.SetNValue("LastMainWindowPos.Width",    MainWindow.Width);
			m_ehProfile.SetNValue("LastMainWindowPos.Height",   MainWindow.Height);
			m_ehProfile.SetNValue("LastProfileWindowPos.Left",  ProfileWindow.Left);
			m_ehProfile.SetNValue("LastProfileWindowPos.Top",   ProfileWindow.Top);
			m_ehProfile.SetNValue("LastProfileWindowPos.Width", ProfileWindow.Width);
			m_ehProfile.SetNValue("LastProfileWindowPos.Height",ProfileWindow.Height);
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