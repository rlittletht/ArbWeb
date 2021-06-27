using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using TCore.StatusBox;

namespace ArbWeb.SPO
{
	// maintains the offline copies of the schedules
	public class Offline
	{
		private IAppContext m_appContext;
		
		public Offline() { }
		public Offline(IAppContext appContext)
		{
			m_appContext = appContext;
		}

		public async Task DownloadAllSchedules()
		{
			await m_appContext.SpoInterop().EnsureLoggedIn();

			string spoSite = m_appContext.Profile.SchedSpoSite;
			string spoSubsite = m_appContext.Profile.SchedSpoSubsite;

			m_appContext.StatusReport.AddMessage("Starting download all schedules...", MSGT.Header);

			m_appContext.StatusReport.AddMessage("Downloading Baseball Schedules...", MSGT.Header);

			foreach (string sched in m_appContext.Profile.BaseballSchedFiles)
				await DownloadScheduleFile(sched, spoSite, spoSubsite, "Baseball");
			
			m_appContext.StatusReport.AddMessage("Finished Baseball.", MSGT.Header);

			m_appContext.StatusReport.AddMessage("Downloading Softball Schedules...", MSGT.Header);
			
			foreach (string sched in m_appContext.Profile.SoftballSchedFiles)
				await DownloadScheduleFile(sched, spoSite, spoSubsite, "Softball");

			m_appContext.StatusReport.AddMessage("Finished Softball.", MSGT.Header);
			m_appContext.StatusReport.AddMessage("Finished Schedules.", MSGT.Header);
			m_appContext.StatusReport.PopLevel();
			m_appContext.StatusReport.PopLevel();
		}

		/*----------------------------------------------------------------------------
			%%Function: DownloadScheduleFile
			%%Qualified: ArbWeb.SPO.Offline.DownloadScheduleFile
		----------------------------------------------------------------------------*/
		private async Task DownloadScheduleFile(string sched, string spoSite, string spoSubsite, string sport)
		{
			string sTargetFile;
			string sLatestFile;

			(sTargetFile, sLatestFile) = MakeDownloadPaths(m_appContext.Profile.SchedDownloadFolder, m_appContext.Profile.SchedWorkingFolder, sport, sched);
			
			m_appContext.StatusReport.AddMessage($"Downloading {sched}...", MSGT.Body);
			string sDir = Path.GetDirectoryName(sTargetFile);

			if (!Directory.Exists(sDir))
				Directory.CreateDirectory(sDir);

			sDir = Path.GetDirectoryName(sLatestFile);

			if (!Directory.Exists(sDir))
				Directory.CreateDirectory(sDir);
			
			await m_appContext.SpoInterop().DownloadFile(spoSite, spoSubsite, sched, sTargetFile);

			File.Delete(sLatestFile);
			File.Copy(sTargetFile, sLatestFile);
		}

		/*----------------------------------------------------------------------------
			%%Function: GetSchedulesAvailableForDiff
			%%Qualified: ArbWeb.SPO.Offline.GetSchedulesAvailableForDiff
		----------------------------------------------------------------------------*/
		public IEnumerable<string> GetSchedulesAvailableForDiff()
		{
			string sDownload, sLatest;
			List<string> schedules = new List<string>();
			
			foreach (string sched in m_appContext.Profile.BaseballSchedFiles)
			{
				(sDownload, sLatest) = MakeDownloadPaths(
					m_appContext.Profile.SchedDownloadFolder,
					m_appContext.Profile.SchedWorkingFolder,
					"Baseball",
					sched);

				if (File.Exists(sLatest))
					schedules.Add(sLatest);
			}
			foreach (string sched in m_appContext.Profile.SoftballSchedFiles)
			{
				(sDownload, sLatest) = MakeDownloadPaths(
					m_appContext.Profile.SchedDownloadFolder,
					m_appContext.Profile.SchedWorkingFolder,
					"Softball",
					sched);

				if (File.Exists(sLatest))
					schedules.Add(sLatest);
			}
			
			return schedules;
		}
		
		/*----------------------------------------------------------------------------
			%%Function: MakeDatedFilename
			%%Qualified: ArbWeb.SPO.Offline.MakeDatedFilename
		----------------------------------------------------------------------------*/
		static string MakeDatedFilename(string leaf, string suffix, DateTime dateTime, string extension = null)
		{
			string sRoot = Path.GetFileNameWithoutExtension(leaf);
			
			if (string.IsNullOrEmpty(suffix))
				suffix = "";
			
			if (extension == null)
				extension = Path.GetExtension(leaf);
			
			return $"{sRoot}{suffix}_{dateTime:yyMMdd}_{dateTime:HHmm}{extension}";
		}

		/*----------------------------------------------------------------------------
			%%Function: MakeLatestFilename
			%%Qualified: ArbWeb.SPO.Offline.MakeLatestFilename
		----------------------------------------------------------------------------*/
		static string MakeLatestFilename(string leaf, string suffix, string extension = null)
		{
			string sRoot = Path.GetFileNameWithoutExtension(leaf);
			if (extension == null)
				extension = Path.GetExtension(leaf);

			if (string.IsNullOrEmpty(suffix))
				suffix = "";

			return $"{sRoot}{suffix}_Latest{extension}";
		}
		
		/*----------------------------------------------------------------------------
			%%Function: MakeDownloadPath
			%%Qualified: ArbWeb.SPO.Offline.MakeDownloadPath
		----------------------------------------------------------------------------*/
		public static (string, string) MakeDownloadPaths(string downloadDir, string sLatestDir, string sSport, string spoPath)
		{
			string sLeaf = Path.GetFileName(spoPath);
			string sDated = MakeDatedFilename(sLeaf, null, DateTime.Now);
			string sLatest = MakeLatestFilename(sLeaf, null);

			return (Path.Combine(downloadDir, sSport, sDated),
					Path.Combine(sLatestDir, sSport, sLatest));
		}
		
		/*----------------------------------------------------------------------------
			%%Function: MakeDiffPaths
			%%Qualified: ArbWeb.SPO.Offline.MakeDiffPaths
		----------------------------------------------------------------------------*/
		public static (string, string) MakeDiffPaths(string downloadDir, string sLatestDir, string sSport, string spoPath)
		{
			string sLeaf = Path.GetFileName(spoPath);
			string sDated = MakeDatedFilename(sLeaf, "_Diff", DateTime.Now, ".csv");
			string sLatest = MakeLatestFilename(sLeaf, "Diff", ".csv");

			return (Path.Combine(downloadDir, sSport, sDated),
				Path.Combine(sLatestDir, sSport, sLatest));
		}

		/*----------------------------------------------------------------------------
			%%Function: GetSpoLeafName
			%%Qualified: ArbWeb.SPO.Offline.GetSpoLeafName
		----------------------------------------------------------------------------*/
		static string GetSpoLeafName(string spoPath)
		{
			return Path.GetFileName(spoPath);
		}

		[TestCase("Test\\", "")]
		[TestCase("Test/", "")]
		[TestCase("With Space\\MyLeaf.xlsx", "MyLeaf.xlsx")]
		[TestCase("With Space/MyLeaf.xlsx", "MyLeaf.xlsx")]
		[TestCase("Test/MyLeaf.xlsx", "MyLeaf.xlsx")]
		[TestCase("Test\\MyLeaf.xlsx", "MyLeaf.xlsx")]
		[TestCase("MyLeaf.xlsx", "MyLeaf.xlsx")]
		[Test]
		public static void TestGetSpoLeafName(string spoPath, string sExpected)
		{
			Assert.AreEqual(sExpected, GetSpoLeafName(spoPath));
		}

		[TestCase("leaf.xlsx", "1/1/2001 9:00", "leaf_010101_0900.xlsx")]
		[TestCase("leaf.xlsx", "1/1/2001 13:00", "leaf_010101_1300.xlsx")]
		[TestCase("leaf.xlsx", "12/1/2001 9:00", "leaf_011201_0900.xlsx")]
		[Test]
		public static void TestMakeDatedFilename(string sLeafIn, string sDateTimeIn, string sExpected)
		{
			Assert.AreEqual(sExpected, MakeDatedFilename(sLeafIn, null, DateTime.Parse(sDateTimeIn)));
		}
	}
}
