using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using NUnit.Framework;
using TCore.StatusBox;

namespace ArbWeb.Games
{
    // ================================================================================
    //  G A M E S  S L O T S
    // ================================================================================
    public class ScheduleGames // GMSS
    {
        public ScheduleGames() { } // just for unit tests

        public SortedList<string, Game> Games { get { return m_plgmSorted; } }

        private SortedList<string, Game> m_plgmSorted;
        private Dictionary<string, Game> m_mpnumgm;

        private List<string> m_plsMiscHeadings;
        private SortedList<string, GameSlot> m_plgmsSorted;
        private SortedList<string, GameSlot> m_plgmsSortedGameNum;
        private StatusBox m_srpt;
        // private Dictionary<string, Dictionary<string, int>> m_mpNameSportLevelCount;
        private Dictionary<string, Sport> m_mpSportSport;
        private List<string> m_plsLegend;

        // NOTE:  This isn't just Team -> Count.  This is also Team-Sport -> Count.
        private Dictionary<string, int> m_mpTeamCount;

        /* S E T  M I S C  H E A D I N G S */
        /*----------------------------------------------------------------------------
            %%Function: SetMiscHeadings
            %%Qualified: ArbWeb.GameData.GameSlots.SetMiscHeadings
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public void SetMiscHeadings(List<string> plsMisc)
        {
            m_plsMiscHeadings = plsMisc;
        }

        /* G A M E  S L O T S */
        /*----------------------------------------------------------------------------
            %%Function: GameSlots
            %%Qualified: ArbWeb.GameData.GameSlots.GameSlots
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public ScheduleGames(StatusBox srpt)
        {
            m_plgmsSorted = new SortedList<string, GameSlot>();
            m_plgmsSortedGameNum = new SortedList<string, GameSlot>();
            m_srpt = srpt;
            m_mpSportSport = new Dictionary<string, Sport>();
            m_plsLegend = new List<string>();
            m_mpTeamCount = new Dictionary<string, int>();
            m_plgmSorted = new SortedList<string, Game>();
            m_mpnumgm = new Dictionary<string, Game>();
            UnitTest();
        }

        /* E N S U R E  S P O R T  L E V E L  P O S */
        /*----------------------------------------------------------------------------
            %%Function: EnsureSportLevelPos
            %%Qualified: ArbWeb.GameData.GameSlots.EnsureSportLevelPos
            %%Contact: rlittle

            Make sure that the sport / level / pos is in the legend (including all
            the subtotals for the sport/level/pos)
        ----------------------------------------------------------------------------*/
        private void EnsureSportLevelPos(string sSport, string sLevel, string sPos)
        {
            Sport sport;
            bool fNewSport = false;

            // make sure we know about this sport and this position
            if (!m_mpSportSport.ContainsKey(sSport))
            {
                sport = new Sport();
                m_mpSportSport.Add(sSport, sport);
                fNewSport = true;
            }

            sport = m_mpSportSport[sSport];

            bool fNewPos, fNewLevel, fNewLevelPos;

            sport.EnsurePos(sLevel, sPos, out fNewLevel, out fNewPos, out fNewLevelPos);

            if (fNewLevelPos)
                m_plsLegend.Add($"{sSport}-{sLevel}-{sPos}");

            if (fNewSport)
                m_plsLegend.Add($"{sSport}-Total");

            if (fNewLevel)
                m_plsLegend.Add($"{sSport}-{sLevel}-Total");

            if (fNewPos)
                m_plsLegend.Add($"{sSport}-{sPos}");
        }

        /* A D D  G A M E */
        /*----------------------------------------------------------------------------
            %%Function: AddGame
            %%Qualified: ArbWeb.CountsData:GameData:Games.AddGame
            %%Contact: rlittle

            Add a game

            This handles ensuring the the sport/level/pos information has been added
            to the legend, so we can build the detail lines later
        ----------------------------------------------------------------------------*/
        public void AddGame(DateTime dttm, string sSite, string sName, string sTeam, string sEmail, string sGameNum, string sHome, string sAway, string sLevel, string sSport, string sPos, string sStatus, bool fCancelled, List<string> plsMisc)
        {
            GameSlot gm = new GameSlot(dttm, sSite, sName, sTeam, sEmail, sGameNum, sHome, sAway, sLevel, sSport, sPos, sStatus, fCancelled, plsMisc);
            AddGame(gm);
        }

        /* A D D  G A M E */
        /*----------------------------------------------------------------------------
            %%Function: AddGame
            %%Qualified: ArbWeb.GameData.GameSlots.AddGame
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public void AddGame(GameSlot gms)
        {
            string sTeamSport = gms.Team + "#-#" + gms.Sport;

            m_plgmsSorted.Add($"{gms.Name}_{gms.Dttm.ToString("yyyyMMdd:HH:mm")}_{m_plgmsSorted.Count}", gms);
            m_plgmsSortedGameNum.Add(
                $"{gms.Dttm.ToString("yyyyMMdd:HH:mm")}_{gms.Site}_{gms.Sport}_{gms.Level}_{gms.GameNum}_{m_plgmsSortedGameNum.Count}", gms);


            if (!m_mpnumgm.ContainsKey(gms.GameNum))
            {
                Game gm;
                m_mpnumgm.Add(gms.GameNum, gm = new Game());
                m_plgmSorted.Add($"{gms.Dttm.ToString("yyyymmdd-HH:MM")}-{gms.GameNum}", gm);
            }

            m_mpnumgm[gms.GameNum].AddGameSlot(gms);

            if (gms.Team != null && gms.Team.Length > 0)
            {
                if (m_mpTeamCount.ContainsKey(gms.Team))
                    m_mpTeamCount[gms.Team]++;
                else
                    m_mpTeamCount.Add(gms.Team, 1);

                if (m_mpTeamCount.ContainsKey(sTeamSport))
                    m_mpTeamCount[sTeamSport]++;
                else
                    m_mpTeamCount.Add(sTeamSport, 1);
            }
            EnsureSportLevelPos(gms.Sport, gms.Level, gms.Pos);
        }

        private string[] SplitTeams(string s)
        {
            string[] rgs = s.Split(';');
            int i;

            for (i = 0; i < rgs.Length; i++)
            {
                rgs[i] = Regex.Replace(rgs[i], "^ *", "");
                rgs[i] = Regex.Replace(rgs[i], " *$", "");
            }

            return rgs;
        }

        private Dictionary<string, DistributeTeamCount> m_mpTeamDtc;

        private int TeamCount(string s)
        {
            if (m_mpTeamCount.ContainsKey(s))
                return m_mpTeamCount[s] + TeamDcCount(s);

            return TeamDcCount(s);
        }

        private int TeamDcCount(string s)
        {
            return 0;
#if NO
					if (m_mpTeamDcCount.ContainsKey(s))
						return m_mpTeamDcCount[s];

					return 0;
#endif
        }

#if NO
				void IncTeamDcCount(string s)
				{
					if (m_mpTeamDcCount.ContainsKey(s))
						m_mpTeamDcCount[s]++;

					m_mpTeamDcCount.Add(s, 1);
				}
#endif

        /* F  T E A M  M A T C H E S  S P O R T */
        /*----------------------------------------------------------------------------
                %%Function: FTeamMatchesSport
                %%Qualified: ArbWeb.CountsData:GameData:Games.FTeamMatchesSport
                %%Contact: rlittle

                Determine if the given team name belongs to a given sport (i.e. look
                for things like "Softball Blast" matching "Softball", etc.
            ----------------------------------------------------------------------------*/
        private bool FTeamMatchesSport(string sTeam, string sSport)
        {
            if (Regex.Match(sSport, ".*Baseball", RegexOptions.IgnoreCase).Success)
            {
                // baseball teams don't have any decoration, or are decorated with baseball
                if (Regex.Match(sTeam, ".*Baseball.*", RegexOptions.IgnoreCase).Success)
                    return true;

                if (Regex.Match(sTeam, ".*Softball.*", RegexOptions.IgnoreCase).Success)
                    return false;

                return true;
            }

            // all other sports should be in the string somewhere
            if (Regex.Match(sTeam, ".*" + sSport + ".*", RegexOptions.IgnoreCase).Success)
                return true;

            return false;
        }

        /* R E D U C E  T E A M S */
        /*----------------------------------------------------------------------------
                %%Function: ReduceTeams
                %%Qualified: ArbWeb.CountsData:GameData:Games.ReduceTeams
                %%Contact: rlittle

                Take "multiple" team allocations and redistribute them to "needier" teams
        ----------------------------------------------------------------------------*/
        public void ReduceTeams()
        {
            m_mpTeamDtc = new Dictionary<string, DistributeTeamCount>();

            // look for team multiples

            // TO DO: Right now, we are allotting games regardless of sport, so baseball teams are getting
            // credit for softball games, etc.  This is fine, unless there's a softball team that needs
            // credit too...
            foreach (string s in new List<string>(m_mpTeamCount.Keys))
            {
                // skip aggregate sport totals
                if (!Regex.Match(s, ".*#-#.*").Success)
                    continue;

                int cDist = m_mpTeamCount[s]; // kvp.Value;

                //						m_srpt.AddMessage(String.Format("{0,-20}{1}", s, m_mpTeamCount[s]), StatusRpt.MSGT.Body);
                if (s.IndexOf(';') == -1)
                    continue;

                // we've got a multiple.  split it up
                DistributeTeamCount dtc = new DistributeTeamCount();
                string[] rgs = CountsData.RexHelper.RgsMatch(s, "(.*)#-#(.*)$");
                string sTeams = rgs[0];
                string sSport = rgs[1];
                string[] rgsTeams = SplitTeams(sTeams);
                bool fIntraSport = false;

                // first, figure out of we're doing an inter or intra sport distribution
                foreach (string sTeam in rgsTeams)
                {
                    if (FTeamMatchesSport(sTeam, sSport))
                    {
                        fIntraSport = true;
                        break;
                    }
                }

                // if we're intra-sport, then we're only going to consider sport totals
                // when we distribute games around...
                foreach (string sTeam in rgsTeams)
                {
                    string sTeamSport = $"{sTeam}#-#{sSport}";

                    if (fIntraSport)
                    {
                        if (FTeamMatchesSport(sTeam, sSport))
                            dtc.AddTeam(sTeam, TeamCount(sTeamSport));
                    }
                    else
                    {
                        // no team matched the sport, so we're going to just distribute
                        // based on total counts for each team, regardless of sport
                        dtc.AddTeam(sTeam, TeamCount(sTeam));
                    }
                }

                dtc.Distribute(cDist);
                // ok, at this point, dtc has the list of distribution changes needed to make this work

                dtc.UpdateTeamTotals(m_mpTeamCount, sSport);

                m_mpTeamDtc.Add(s, dtc);
            }
            // now, adjust the raw team counts

            foreach (GameSlot gm in m_plgmsSorted.Values)
            {
                if (gm.Open)
                    continue;

                if (gm.Team.IndexOf(';') == -1)
                    continue;

                string sSportTeam = $"{gm.Team}#-#{gm.Sport}";

                DistributeTeamCount dtc = m_mpTeamDtc[sSportTeam];

                gm.Team = dtc.STeamNext();
                dtc.DecTeamNext();
            }
        }



        /* G E N  G A M E S  R E P O R T */
        /*----------------------------------------------------------------------------
                %%Function: GenGamesReport
                %%Qualified: ArbWeb.CountsData:GameData:Games.GenGamesReport
                %%Contact: rlittle

                Take the accumulated game data and generate a report of the games
                that Arbiter knows about.  suitable for comparing					
            ----------------------------------------------------------------------------*/
        public void GenGamesReport(string sReport)
        {
            StreamWriter sw = new StreamWriter(sReport, false, Encoding.Default);
            List<string> plsLegend = new List<string>();

            plsLegend.Insert(0, "Game");
            plsLegend.Insert(1, "Date");
            plsLegend.Insert(2, "Time");
            plsLegend.Insert(3, "Site");
            plsLegend.Insert(4, "Level");
            plsLegend.Insert(5, "Home");
            plsLegend.Insert(6, "Away");
            plsLegend.Insert(7, "Sport");

            bool fFirst = true;
            foreach (string s in plsLegend)
            {
                if (!fFirst)
                {
                    sw.Write(",");
                }

                fFirst = false;
                sw.Write(s);
            }
            sw.WriteLine();

            Dictionary<string, bool> mpGame = new Dictionary<string, bool>();

            foreach (GameSlot gm in m_plgmsSorted.Values)
            {
                // only report each game once...
                if (mpGame.ContainsKey(gm.GameNum))
                    continue;

                mpGame.Add(gm.GameNum, true);

                // for each game, report the information, using Legend as the sort order for everything
                sw.WriteLine(gm.SReport(plsLegend));
            }
            sw.Close();
        }

        /* G E N  R E P O R T */
        /*----------------------------------------------------------------------------
                %%Function: GenReport
                %%Qualified: ArbWeb.CountsData:GameData:Games.GenReport
                %%Contact: rlittle

                Take the accumulated game data and generate an analysis report
            ----------------------------------------------------------------------------*/
        public void GenReport(string sReport)
        {
            StreamWriter sw = new StreamWriter(sReport, false, Encoding.Default);
            bool fFirst = true;

            m_plsLegend.Sort();
            m_plsLegend.Insert(0, "UmpireName");
            m_plsLegend.Insert(1, "Team");
            m_plsLegend.Insert(2, "Email");
            m_plsLegend.Insert(3, "Game");
            m_plsLegend.Insert(4, "DateTime");
            m_plsLegend.Insert(5, "Level");
            m_plsLegend.Insert(6, "Home");
            m_plsLegend.Insert(7, "Away");
            m_plsLegend.Insert(8, "Description");
            m_plsLegend.Insert(9, "Cancelled");
            m_plsLegend.Insert(10, "Total");
            m_plsLegend.Add("$$$MISC$$$"); // want this at the end!

            foreach (string s in m_plsLegend)
            {
                if (s == "$$$MISC$$$")
                {
                    foreach (string sMisc in m_plsMiscHeadings)
                    {
                        sw.Write(",\"{0}\"", sMisc);
                    }
                    continue;
                }

                if (!fFirst)
                {
                    sw.Write(",");
                }

                fFirst = false;
                sw.Write(s);
            }
            sw.WriteLine();

            foreach (GameSlot gm in m_plgmsSorted.Values)
            {
                //						if (gm.Open)
                //							continue;

                // for each game, report the information, using Legend as the sort order for everything
                sw.WriteLine(gm.SReport(m_plsLegend));
            }
            sw.Close();
        }


        private void WriteGameRoster(StreamWriter sw, List<GameSlot> plgm, bool fHeader, ArbWeb.Roster rst, Dictionary<string, string> mpSiteRoot)
        {
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
                        sOtherRanks = rste.OtherRanks(gm.Sport, gm.Pos, nBaseRank);
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

        public void GenSiteRosterReport(string sReportFile, ArbWeb.Roster rst, string[] rgsRosterFilter, DateTime dttmStart, DateTime dttmEnd)
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

            // we want these to become (not new sort order).
            //     DateTime/Foo Field Consult/Foo Field, Foo Field/ Consultants
            //     DateTime/Foo Field Game/Foo Field #1/Game 1
            //     DateTime/Bar Field Consult/Bar Field, Bar Field/Consultants
            //     DateTime/Bar Field Game/Bar Field, North Field/Game 1

            // we need to figure out the common root between fields (and we can't rely on punctuation...)
            SortedList<string, GameSlot> mpgames = new SortedList<string, GameSlot>();
            SortedSet<string> plsSites = new SortedSet<string>();

            foreach (GameSlot gm in m_plgmsSortedGameNum.Values)
            {
                if (!plsSites.Contains(gm.Site))
                    plsSites.Add(gm.Site);
            }

            Dictionary<string, string> mpSiteRoot = MapCommonRootsFromList(plsSites);

            // now, build a new game list
            foreach (GameSlot gm in m_plgmsSortedGameNum.Values)
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
                    WriteGameRoster(sw, plgm, fHeader, rst, mpSiteRoot);
                    plgm.Clear();
                    fHeader = false;
                }

                plgm.Add(gm);
                if (gm.SportLevel.ToUpper().Contains("CONSULTANT"))
                    fHeader = true;
            }

            if (plgm.Count > 0)
                WriteGameRoster(sw, plgm, fHeader, rst, mpSiteRoot);

            sw.WriteLine("</table></div></html>");
            sw.Close();
        }

        /* G E N  O P E N  S L O T S  R E P O R T */
        /*----------------------------------------------------------------------------
                %%Function: GenOpenSlotsReport
                %%Qualified: ArbWeb.CountsData:GameData:Games.GenOpenSlotsReport
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        public void GenOpenSlotsReport(string sReport, bool fDetail, bool fFuzzyTimes, bool fDatePivot, string[] rgsSportFilter, string[] rgsSportLevelFilter, SlotAggr sa)
        {
            if (fDetail)
                GenOpenSlotsReportDetail(sReport, rgsSportFilter, rgsSportLevelFilter, sa);
            else
                GenOpenSlotsReport(sReport, rgsSportFilter, rgsSportLevelFilter, fFuzzyTimes, fDatePivot, sa);
        }

        /* G A M E S  F R O M  F I L T E R */
        /*----------------------------------------------------------------------------
            %%Function: GamesFromFilter
            %%Qualified: ArbWeb.GameData.Games.GamesFromFilter
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public ScheduleGames GamesFromFilter(string[] rgsSportFilter, string[] rgsSportLevelFilter, bool fOpenOnly, SlotAggr sa)
        {
            ScheduleGames gms = new ScheduleGames(m_srpt);

            DateTime dttmStart = sa.DttmStart;
            DateTime dttmEnd = sa.DttmEnd;

            SortedList<string, int> plsSports = Utils.PlsUniqueFromRgs(rgsSportFilter);
            SortedList<string, int> plsLevels = Utils.PlsUniqueFromRgs(rgsSportLevelFilter);

            foreach (GameSlot gm in m_plgmsSorted.Values)
            {
                if (!gm.Open && fOpenOnly)
                    continue;

                if (DateTime.Compare(gm.Dttm, dttmStart) < 0 || DateTime.Compare(gm.Dttm, dttmEnd) > 0)
                    continue;

                if (plsSports != null && !(plsSports.ContainsKey(gm.Sport)))
                    continue;

                if (plsLevels != null && !(plsLevels.ContainsKey(gm.SportLevel)))
                    continue;

                gms.AddGame(gm);
            }
            return gms;
        }

        /* R E Q U I R E D  R A N K S */
        /*----------------------------------------------------------------------------
            %%Function: RequiredRanks
            %%Qualified: ArbWeb.GameData.Games.RequiredRanks
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public List<string> RequiredRanks()
        {
            List<string> pls = new List<string>();

            HashSet<string> hs = new HashSet<string>();

            foreach (GameSlot gm in m_plgmsSorted.Values)
            {
                string sPosRank = $"{gm.Sport}, {gm.Pos}";

                if (!hs.Contains(sPosRank))
                {
                    hs.Add(sPosRank);
                    pls.Add(sPosRank);
                }
            }
            return pls;
        }


        /* G E N  O P E N  S L O T S  R E P O R T  D E T A I L */
        /*----------------------------------------------------------------------------
                %%Function: GenOpenSlotsReportDetail
                %%Qualified: ArbWeb.CountsData:GameData:Games.GenOpenSlotsReportDetail
                %%Contact: rlittle

                Generate a report of open slots
            ----------------------------------------------------------------------------*/
        private void GenOpenSlotsReportDetail(string sReport, string[] rgsSportFilter, string[] rgsSportLevelFilter, SlotAggr sa)
        {
            DateTime dttmStart = sa.DttmStart;
            DateTime dttmEnd = sa.DttmEnd;

            StreamWriter sw = new StreamWriter(sReport, false, Encoding.Default);

            string sFormat = "<tr><td>{0}<td>{1}<td>{2}<td>{3}<td>{4}<td>{5}<td>{6}<td>{7}</tr>";

            sw.WriteLine("<html><body><table>");
            sw.WriteLine(String.Format(sFormat, "Game", "Date", "Time", "Field", "Level", "Home", "Away", "Slots"));

            SortedList<string, int> plsSports = Utils.PlsUniqueFromRgs(rgsSportFilter);
            SortedList<string, int> plsLevels = Utils.PlsUniqueFromRgs(rgsSportLevelFilter);

            foreach (GameSlot gm in m_plgmsSorted.Values)
            {
                if (!gm.Open)
                    continue;

                if (DateTime.Compare(gm.Dttm, dttmStart) < 0 || DateTime.Compare(gm.Dttm, dttmEnd) > 0)
                    continue;

                if (plsSports != null && !(plsSports.ContainsKey(gm.Sport)))
                    continue;

                if (plsLevels != null && !(plsLevels.ContainsKey(gm.SportLevel)))
                    continue;

                sw.WriteLine(String.Format(sFormat, gm.GameNum, gm.Dttm.ToString("MM/dd/yy ddd"), gm.Dttm.ToString("hh:mm tt"), gm.Site,
                    $"{gm.Sport}, {gm.Level}", gm.Home, gm.Away, gm.Pos));
            }
            sw.Close();
        }

        /* G E N  O P E N  S L O T S */
        /*----------------------------------------------------------------------------
                %%Function: GenOpenSlots
                %%Qualified: ArbWeb.CountsData:GameData:Games.GenOpenSlots
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        public SlotAggr GenOpenSlots(DateTime dttmStart, DateTime dttmEnd)
        {
            return SlotAggr.Gen(m_plgmsSorted, dttmStart, dttmEnd, null, null, true);
        }

        /* G E N  O P E N  S L O T S  R E P O R T */
        /*----------------------------------------------------------------------------
                %%Function: GenOpenSlotsReport
                %%Qualified: ArbWeb.CountsData:GameData:Games.GenOpenSlotsReport
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private void GenOpenSlotsReport(string sReport, string[] rgsSportsFilter, string[] rgsSportLevelsFilter, bool fFuzzyTimes, bool fDatePivot, SlotAggr sa)
        {
            //					sa.GenReport(sReport, rgsSportsFilter, rgsSportLevelsFilter);
            sa.GenReportBySite(sReport, fFuzzyTimes, fDatePivot, rgsSportsFilter, rgsSportLevelsFilter);
        }

        /* G E T  O P E N  S L O T  S P O R T S */
        /*----------------------------------------------------------------------------
                %%Function: GetOpenSlotSports
                %%Qualified: ArbWeb.CountsData:GameData:Games.GetOpenSlotSports
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        public string[] GetOpenSlotSports(SlotAggr sa)
        {
            string[] rgs = new string[sa.Sports.Length];
            sa.Sports.CopyTo(rgs, 0);
            return rgs;
        }

        public string[] GetSiteRosterSites(SlotAggr sa)
        {
            List<string> plsSports = new List<string>();
            List<string> plsSites = new List<string>();

            foreach (GameSlot gm in m_plgmsSorted.Values)
            {
                if (gm.Dttm < sa.DttmStart || gm.Dttm > sa.DttmEnd)
                    continue;

                if (!plsSports.Contains(gm.SportLevel))
                    plsSports.Add(gm.SportLevel);

                if (!plsSites.Contains(gm.SiteShort))
                    plsSites.Add(gm.SiteShort);
            }
            return plsSites.ToArray();
        }

        /* G E T  O P E N  S L O T  S P O R T  L E V E L S */
        /*----------------------------------------------------------------------------
                %%Function: GetOpenSlotSportLevels
                %%Qualified: ArbWeb.CountsData:GameData:Games.GetOpenSlotSportLevels
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        public string[] GetOpenSlotSportLevels(SlotAggr sa)
        {
            string[] rgs = new string[sa.SportLevels.Length];
            sa.SportLevels.CopyTo(rgs, 0);
            return rgs;
        }

        #region Tests

        private void UnitTest()
        {
            string s = "", s2 = "", s3 = "";

            s = GamesLoader_Arbiter.ReverseName(null, "Mary Van Whatsa Hoozit");
            Debug.Assert(String.Compare(s, "Van Whatsa Hoozit,Mary") == 0);
            Debug.Assert(ArbWeb.Roster.FSplitName(s, out s2, out s3));
            Debug.Assert(s2 == "Mary");
            Debug.Assert(s3 == "Van Whatsa Hoozit");

            s = "Van Whatsa Hoozit, Mary";

            Debug.Assert(ArbWeb.Roster.FSplitName(s, out s2, out s3));
            Debug.Assert(s2 == "Mary");
            Debug.Assert(s3 == "Van Whatsa Hoozit");

            s = "";
            s = GamesLoader_Arbiter.AppendCheck(s, "Foo");
            Debug.Assert(String.Compare(s, "Foo") == 0);
            s = GamesLoader_Arbiter.AppendCheck(s, " Bar ");
            Debug.Assert(String.Compare(s, "Foo Bar") == 0);
            s = GamesLoader_Arbiter.AppendCheck(s, " Baz ");
            Debug.Assert(String.Compare(s, "Foo Bar Baz") == 0);

            // now test Reduction functionality...

            Dictionary<string, int> mpTeamCountSav = m_mpTeamCount;

            m_mpTeamCount = new Dictionary<string, int>();
            m_mpTeamCount.Add("Swans", 5);
            m_mpTeamCount.Add("Swans#-#Redmond Baseball", 5);
            m_mpTeamCount.Add("Crows", 2);
            m_mpTeamCount.Add("Crows#-#Redmond Baseball", 2);
            m_mpTeamCount.Add("Blue Jays", 9);
            m_mpTeamCount.Add("Blue Jays#-#Redmond Baseball", 9);
            m_mpTeamCount.Add("Eagles", 3);
            m_mpTeamCount.Add("Eagles#-#Redmond Baseball", 3);
            m_mpTeamCount.Add("Woodpeckers", 2);
            m_mpTeamCount.Add("Woodpeckers#-#Redmond Baseball", 2);
            m_mpTeamCount.Add("Swans;Eagles", 10);
            m_mpTeamCount.Add("Swans;Eagles#-#Redmond Baseball", 10);
            ReduceTeams();
            Debug.Assert(TeamCount("Swans") == 9);
            Debug.Assert(TeamCount("Swans#-#Redmond Baseball") == 9);
            Debug.Assert(TeamCount("Crows") == 2);
            Debug.Assert(TeamCount("Crows#-#Redmond Baseball") == 2);
            Debug.Assert(TeamCount("Blue Jays") == 9);
            Debug.Assert(TeamCount("Blue Jays#-#Redmond Baseball") == 9);
            Debug.Assert(TeamCount("Eagles") == 9);
            Debug.Assert(TeamCount("Eagles#-#Redmond Baseball") == 9);
            Debug.Assert(TeamCount("Woodpeckers") == 2);
            Debug.Assert(TeamCount("Woodpeckers#-#Redmond Baseball") == 2);

            m_mpTeamCount = new Dictionary<string, int>();
            m_mpTeamCount.Add("Swans", 5);
            m_mpTeamCount.Add("Swans#-#Baseball", 5);
            m_mpTeamCount.Add("Crows", 2);
            m_mpTeamCount.Add("Crows#-#Baseball", 2);
            m_mpTeamCount.Add("Blue Jays", 9);
            m_mpTeamCount.Add("Blue Jays#-#Baseball", 9);
            m_mpTeamCount.Add("Eagles", 3);
            m_mpTeamCount.Add("Eagles#-#Baseball", 3);
            m_mpTeamCount.Add("Woodpeckers", 2);
            m_mpTeamCount.Add("Woodpeckers#-#Baseball", 2);
            m_mpTeamCount.Add("Eagles;Woodpeckers;Swans", 14);
            m_mpTeamCount.Add("Eagles;Woodpeckers;Swans#-#Baseball", 14);
            ReduceTeams();
            Debug.Assert(TeamCount("Swans") == 8);
            Debug.Assert(TeamCount("Swans#-#Baseball") == 8);
            Debug.Assert(TeamCount("Crows") == 2);
            Debug.Assert(TeamCount("Crows#-#Baseball") == 2);
            Debug.Assert(TeamCount("Blue Jays") == 9);
            Debug.Assert(TeamCount("Blue Jays#-#Baseball") == 9);
            Debug.Assert(TeamCount("Eagles") == 8);
            Debug.Assert(TeamCount("Eagles#-#Baseball") == 8);
            Debug.Assert(TeamCount("Woodpeckers") == 8);
            Debug.Assert(TeamCount("Woodpeckers#-#Baseball") == 8);

            m_mpTeamCount = new Dictionary<string, int>();
            m_mpTeamCount.Add("Swans", 5);
            m_mpTeamCount.Add("Swans#-#Baseball", 5);
            m_mpTeamCount.Add("Crows", 2);
            m_mpTeamCount.Add("Crows#-#Baseball", 2);
            m_mpTeamCount.Add("Blue Jays", 9);
            m_mpTeamCount.Add("Blue Jays#-#Baseball", 9);
            m_mpTeamCount.Add("Eagles", 3);
            m_mpTeamCount.Add("Eagles#-#Baseball", 3);
            m_mpTeamCount.Add("Woodpeckers", 2);
            m_mpTeamCount.Add("Woodpeckers#-#Baseball", 2);
            m_mpTeamCount.Add("Eagles;Blue Jays;Swans", 5);
            m_mpTeamCount.Add("Eagles;Blue Jays;Swans#-#Baseball", 5);
            ReduceTeams();

            Debug.Assert(TeamCount("Swans#-#Baseball") == 6);
            Debug.Assert(TeamCount("Swans") == 6);
            Debug.Assert(TeamCount("Crows#-#Baseball") == 2);
            Debug.Assert(TeamCount("Crows") == 2);
            Debug.Assert(TeamCount("Blue Jays#-#Baseball") == 9);
            Debug.Assert(TeamCount("Blue Jays") == 9);
            Debug.Assert(TeamCount("Eagles#-#Baseball") == 7);
            Debug.Assert(TeamCount("Eagles") == 7);
            Debug.Assert(TeamCount("Woodpeckers#-#Baseball") == 2);
            Debug.Assert(TeamCount("Woodpeckers") == 2);


            m_mpTeamCount = new Dictionary<string, int>();
            m_mpTeamCount.Add("Swans", 5);
            m_mpTeamCount.Add("Swans#-#Baseball", 5);
            m_mpTeamCount.Add("Crows", 2);
            m_mpTeamCount.Add("Crows#-#Baseball", 2);
            m_mpTeamCount.Add("Blue Jays", 9);
            m_mpTeamCount.Add("Blue Jays#-#Baseball", 9);
            m_mpTeamCount.Add("Eagles", 3);
            m_mpTeamCount.Add("Eagles#-#Baseball", 3);
            m_mpTeamCount.Add("Woodpeckers", 2);
            m_mpTeamCount.Add("Woodpeckers#-#Baseball", 2);
            m_mpTeamCount.Add("Swans;Eagles", 10);
            m_mpTeamCount.Add("Swans;Eagles#-#Baseball", 10);
            m_mpTeamCount.Add("Eagles;Blue Jays;Swans", 5);
            m_mpTeamCount.Add("Eagles;Blue Jays;Swans#-#Baseball", 5);
            m_mpTeamCount.Add("Eagles;Woodpeckers;Swans", 14);
            m_mpTeamCount.Add("Eagles;Woodpeckers;Swans#-#Baseball", 14);
            ReduceTeams();

            Debug.Assert(TeamCount("Swans") == 12);
            Debug.Assert(TeamCount("Swans#-#Baseball") == 12);
            Debug.Assert(TeamCount("Crows") == 2);
            Debug.Assert(TeamCount("Crows#-#Baseball") == 2);
            Debug.Assert(TeamCount("Blue Jays") == 11);
            Debug.Assert(TeamCount("Blue Jays#-#Baseball") == 11);
            Debug.Assert(TeamCount("Eagles") == 12);
            Debug.Assert(TeamCount("Eagles#-#Baseball") == 12);
            Debug.Assert(TeamCount("Woodpeckers") == 13);
            Debug.Assert(TeamCount("Woodpeckers#-#Baseball") == 13);

            m_mpTeamCount = new Dictionary<string, int>();
            m_mpTeamCount.Add("Swans", 5);
            m_mpTeamCount.Add("Swans#-#Baseball", 5);
            m_mpTeamCount.Add("Softball Eagles", 9);
            m_mpTeamCount.Add("Softball Eagles#-#Softball", 5);
            m_mpTeamCount.Add("Softball Eagles#-#Baseball", 4);
            m_mpTeamCount.Add("Swans;Softball Eagles", 8);
            m_mpTeamCount.Add("Swans;Softball Eagles#-#Baseball", 3);
            m_mpTeamCount.Add("Swans;Softball Eagles#-#Softball", 5);
            ReduceTeams();
            Debug.Assert(TeamCount("Swans") == 8);
            Debug.Assert(TeamCount("Swans#-#Baseball") == 8);
            Debug.Assert(TeamCount("Softball Eagles") == 14);
            Debug.Assert(TeamCount("Softball Eagles#-#Baseball") == 4);
            Debug.Assert(TeamCount("Softball Eagles#-#Softball") == 10);

            m_mpTeamCount = new Dictionary<string, int>();
            m_mpTeamCount.Add("Swans", 5);
            m_mpTeamCount.Add("Swans#-#Baseball", 5);
            m_mpTeamCount.Add("Gulls", 3);
            m_mpTeamCount.Add("Gulls#-#Baseball", 3);
            m_mpTeamCount.Add("Softball Eagles", 9);
            m_mpTeamCount.Add("Softball Eagles#-#Softball", 5);
            m_mpTeamCount.Add("Softball Eagles#-#Baseball", 4);
            m_mpTeamCount.Add("Swans;Softball Eagles", 8);
            m_mpTeamCount.Add("Swans;Softball Eagles#-#Baseball", 3);
            m_mpTeamCount.Add("Swans;Softball Eagles#-#Softball", 5);
            m_mpTeamCount.Add("Swans;Gulls#-#Softball", 10); // softball to get distributed to non-softball teams
            ReduceTeams();
            Debug.Assert(TeamCount("Swans") == 10);
            Debug.Assert(TeamCount("Swans#-#Baseball") == 8);
            Debug.Assert(TeamCount("Swans#-#Softball") == 2);
            Debug.Assert(TeamCount("Gulls") == 11);
            Debug.Assert(TeamCount("Gulls#-#Baseball") == 3);
            Debug.Assert(TeamCount("Gulls#-#Softball") == 8);
            Debug.Assert(TeamCount("Softball Eagles") == 14);
            Debug.Assert(TeamCount("Softball Eagles#-#Baseball") == 4);
            Debug.Assert(TeamCount("Softball Eagles#-#Softball") == 10);
            m_mpTeamCount = mpTeamCountSav;

        }

#endregion

    }
}