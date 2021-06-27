using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using ArbWeb.Games;
using NUnit.Framework;
using TCore.StatusBox;
using TCore.WebControl;

namespace ArbWeb.Reports
{
	public class SiteRosterReport
	{
		private IAppContext m_appContext;

		public SiteRosterReport() { } // for unit tests
		/*----------------------------------------------------------------------------
			%%Function:SiteRosterReport
			%%Qualified:ArbWeb.Reports.SiteRosterReport.SiteRosterReport
		----------------------------------------------------------------------------*/
		public SiteRosterReport(IAppContext appContext)
		{
			m_appContext = appContext;
		}

		/*----------------------------------------------------------------------------
			%%Function:DoGenSiteRosterReport
			%%Qualified:ArbWeb.Reports.SiteRosterReport.DoGenSiteRosterReport
		----------------------------------------------------------------------------*/
		public void DoGenSiteRosterReport(CountsData gc, Roster rst, string []rgsRoster, DateTime dttmStart, DateTime dttmEnd)
		{
			string sTempFile = $"{Environment.GetEnvironmentVariable("Temp")}\\temp{System.Guid.NewGuid().ToString()}.doc";

			gc.GenSiteRosterReport(sTempFile, rst, rgsRoster, dttmStart, dttmEnd, m_appContext.Profile.NoHonorificRanks);
			// launch word with the file
			Process.Start(sTempFile);
			// System.IO.File.Delete(sTempFile);
		}

        public static void GenSiteRosterReport(ScheduleGames games, string sReportFile, ArbWeb.Roster rst, string[] rgsRosterFilter, DateTime dttmStart, DateTime dttmEnd, bool noHonorificRanks)
        {
            StreamWriter sw = new StreamWriter(sReportFile, false, Encoding.Default);
            SortedList<string, int> plsSiteShort = Utils.PlsUniqueFromRgs(rgsRosterFilter);
            List<GameSlot> plgm = new List<GameSlot>();

            sw.WriteLine("<html>");
            sw.WriteLine("<head><style>");
            sw.WriteLine(".rosterOuter, .rosterInner, td.rosterInnerName { margin-left: 5pt; font-family: 'Calibri'; font-size: 11pt; border-collapse: collapse; border-spacing: 0pt; }");
            sw.WriteLine("td.rosterOuter { border-top: .5pt solid black; border-bottom: .5pt solid black; }");
            sw.WriteLine("table.rosterInner { width: 100%;}");
            sw.WriteLine("table.rosterOuter { width: 100%; max-width: 7.5in;}");
            sw.WriteLine("td.rosterInner, td.rosterInnerName { padding-left: 5pt; padding-right: 5pt;} td.rosterInnerName { }");
            sw.WriteLine("@page sec1 { margin: .25in; } div.sec1 { page:sec1} ");
            sw.WriteLine("</style></head>");

            sw.WriteLine("<div class='sec1'><table class='rosterOuter'>");

            // now we want to create a custom sort, so consultants sort to the top of the time slot/game site

            // first, figure out the site name key
            //   We have to deal with:
            //     DateTime/Foo Field #1/Game 1
            //     DateTime/Foo Field, Foo Field/ Consultants
            //     DateTime/Bar Field, Bar Field/Consultants
            //     DateTime/Bar Field, North Field/Game 1

            // (note how #1 sorts before Foo Field, but North Field sorts after Bar Field. This gives us an inconsistent
            // consultant sort order.

            // we want these to become (note new sort order).
            //     DateTime/Foo Field Consult/Foo Field, Foo Field/ Consultants
            //     DateTime/Foo Field Game/Foo Field #1/Game 1
            //     DateTime/Bar Field Consult/Bar Field, Bar Field/Consultants
            //     DateTime/Bar Field Game/Bar Field, North Field/Game 1

            // we need to figure out the common root between fields (and we can't rely on punctuation...)
            SortedList<string, GameSlot> mpgames = new SortedList<string, GameSlot>();
            SortedSet<string> plsSites = new SortedSet<string>();

            foreach (GameSlot gm in games.SortedGameSlotsByGameNumber)
            {
                if (!plsSites.Contains(gm.Site))
                    plsSites.Add(gm.Site);
            }

            Dictionary<string, string> mpSiteRoot = MapCommonRootsFromList(plsSites);

            // now, build a new game list
            foreach (GameSlot gm in games.SortedGameSlotsByGameNumber)
            {
                string sType;
                string sSite;

                if (mpSiteRoot.ContainsKey(gm.Site))
                {
                    sSite = mpSiteRoot[gm.Site];
                }
                else
                {
                    sSite = gm.Site;
                }

                if (gm.SportLevel.ToUpper().Contains("CONSULTANT"))
                    sType = $"{sSite}Consult";
                else
                    sType = $"{sSite}Game";

                mpgames.Add(
                    $"{gm.Dttm.ToString("yyyyMMdd:HH:mm")}_{sType}_{gm.Site}_{gm.Sport}_{gm.Level}_{gm.GameNum}_{mpgames.Count}", gm);
            }
            // at this point we are ready to generate the report

            bool fHeader = false;
            foreach (GameSlot gm in mpgames.Values)
            {
                if (gm.Dttm < dttmStart || gm.Dttm > dttmEnd)
                    continue;

                if (rgsRosterFilter != null)
                    if (!plsSiteShort.ContainsKey(gm.SiteShort))
                        continue;

                if (plgm.Count > 0 && plgm[0].GameNum != gm.GameNum)
                {
                    WriteGameRoster(sw, plgm, fHeader, rst, mpSiteRoot, noHonorificRanks);
                    plgm.Clear();
                    fHeader = false;
                }

                plgm.Add(gm);
                if (gm.SportLevel.ToUpper().Contains("CONSULTANT"))
                    fHeader = true;
            }

            if (plgm.Count > 0)
                WriteGameRoster(sw, plgm, fHeader, rst, mpSiteRoot, noHonorificRanks);

            sw.WriteLine("</table></div></html>");
            sw.Close();
        }

        private static void WriteGameRoster(StreamWriter sw, List<GameSlot> plgm, bool fHeader, ArbWeb.Roster rst, Dictionary<string, string> mpSiteRoot, bool noHonorificRanks)
        {
	        HashSet<string> honorifics = new HashSet<string>();

	        if (noHonorificRanks)
	        {
		        honorifics.Add("HP");
		        honorifics.Add("1B");
		        honorifics.Add("2B");
		        honorifics.Add("3B");
		        honorifics.Add("LINE");
	        }


	        string sBackground = "";

            if (fHeader)
            {
                sBackground = " style='background: #c0c0c0'";
            }
            sw.WriteLine($"<tr{sBackground}>");
            sw.WriteLine($"<td class='rosterOuter'>{plgm[0].GameNum}");
            sw.WriteLine($"<td class='rosterOuter'>{plgm[0].Dttm.ToString("ddd M/dd")}");
            sw.WriteLine($"<td class='rosterOuter'>{plgm[0].Dttm.ToString("h:mm tt")}");
            sw.WriteLine($"<td class='rosterOuter'>{plgm[0].SportLevel}");
            if (fHeader)
            {
                sw.WriteLine($"<td class='rosterOuter'>{mpSiteRoot[plgm[0].Site]}");
            }
            else
            {
                sw.WriteLine($"<td class='rosterOuter'>{plgm[0].Site}");
            }

            sw.WriteLine($"<td class='rosterOuter'>{plgm[0].Home}");
            sw.WriteLine($"<td class='rosterOuter'>{plgm[0].Away}");
            sw.WriteLine($"<tr{sBackground}><td colspan='7' class='rosterOuter'>");
            sw.WriteLine($"<table {sBackground} class='rosterInner'>");
            foreach (GameSlot gm in plgm)
            {
                sw.WriteLine("<tr>");
                if (gm.Open)
                {
                    if (!gm.Sport.Contains("Admin"))
                    {
                        sw.WriteLine($"<td class='rosterInner'>{gm.Pos}");
                        sw.WriteLine("<td colspan='4'>&nbsp;");
                    }
                }
                else
                {
                    RosterEntry rste = rst.RsteLookupEmail(gm.Email);
                    int nBaseRank;
                    string sName;
                    string sOtherRanks;
                    string sPhone;

                    if (rste == null)
                    {
                        sName = "<unknown>";
                        sOtherRanks = "";
                        nBaseRank = 0;
                        sPhone = "";
                    }
                    else
                    {
                        sPhone = rste.CellPhone;
                        sName = rste.Name;
                        nBaseRank = rste.Rank($"{gm.Sport}, {gm.Pos}");
                        sOtherRanks = rste.OtherRanks(gm.Sport, gm.Pos, nBaseRank, honorifics);
                    }

                    sw.WriteLine($"<td class='rosterInner'>{gm.Pos} ({nBaseRank})");
                    sw.WriteLine($"<td class='rosterInnerName'>{sName}");
                    sw.WriteLine($"<td class='rosterInner'>{sPhone}");
                    sw.WriteLine($"<td class='rosterInner'>{sOtherRanks}");
                    sw.WriteLine($"<td class='rosterInner'>{gm.Status}");
                }
            }
            sw.WriteLine("</table>");
        }

        static int CchCommonRoot(string sLeft, string sRight)
        {
            if (sLeft == null || sRight == null)
                return 0;

            int ichMac = Math.Min(sLeft.Length, sRight.Length);
            int ich = 0;

            while (ich < ichMac && sLeft[ich] == sRight[ich])
                ich++;

            return ich;
        }

        [Test]
        [TestCase("foo", "foo", 3)]
        [TestCase("foo", "foobar", 3)]
        [TestCase("foo ", "foo bar", 4)]
        [TestCase("foo", "bar", 0)]
        [TestCase("foo", "ffoo", 1)]
        [TestCase("foo", "bfoo", 0)]
        [TestCase("foo", "fan", 1)]
        [TestCase("", "bfoo", 0)]
        [TestCase(null, "bfoo", 0)]
        public static void TestCchCommonRoot(string sLeft, string sRight, int cchExpected)
        {
            Assert.AreEqual(cchExpected, CchCommonRoot(sLeft, sRight));
        }

        static Dictionary<string, string> MapCommonRootsFromList(SortedSet<string> sset)
        {
            Dictionary<string, string> mpRoots = new Dictionary<string, string>();
            List<string> plsPending = new List<string>();

            string sLast = null;
            int cchRootCur = 0;

            foreach (string s in sset)
            {
                int cch = CchCommonRoot(sLast, s);

                if (cch == 0 || cch < Math.Min(sLast.Length, s.Length) / 3 + 1) // we require the root to be at least 1/3 of the name...
                {
                    // start building a new root
                    if (cchRootCur > 0)
                    {
                        string sRoot = sLast.Substring(0, cchRootCur);

                        foreach (string sToMap in plsPending)
                            mpRoots.Add(sToMap, sRoot);
                    }

                    cchRootCur = 0;
                    sLast = s;
                    plsPending.Clear();
                    plsPending.Add(s);
                }
                else
                {
                    if (cch < cchRootCur || cchRootCur == 0)
                        cchRootCur = cch;
                    plsPending.Add(s);
                    sLast = s;
                }
            }

            // and finally, do one last check
            if (cchRootCur > 0)
            {
                string sRoot = sLast.Substring(0, cchRootCur);

                foreach (string sToMap in plsPending)
                    mpRoots.Add(sToMap, sRoot);
            }

            return mpRoots;
        }

        [Test]
        [TestCase(new String[] { "foo" }, new string[] { })]
        [TestCase(new String[] { "foo", "bar" }, new string[] { })]
        [TestCase(new String[] { "foo", "foo2", "bar" }, new string[] { "foo|foo", "foo2|foo" })]
        [TestCase(new String[] { "foo", "foo2", "bar2", "bar3", "bar4", "boo" }, new string[] { "foo|foo", "foo2|foo", "bar2|bar", "bar3|bar", "bar4|bar" })]
        public static void TestCommonRootsFromList(string[] rgs, string[] rgmapExpected)
        {
            SortedSet<string> sset = new SortedSet<string>(rgs);
            Dictionary<string, string> mpExpected = new Dictionary<string, string>();

            foreach (string s in rgmapExpected)
            {
                string[] rgsMap = s.Split('|');

                mpExpected.Add(rgsMap[0], rgsMap[1]);
            }

            Dictionary<string, string> mpActual = MapCommonRootsFromList(sset);

            Assert.AreEqual(mpExpected, mpActual);
        }

    }
}