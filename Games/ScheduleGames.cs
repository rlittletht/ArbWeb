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
        private Dictionary<string, Sport> m_mpSportSport;
        private List<string> m_plsAnalysisLegend;

        public IEnumerable<GameSlot> SortedGameSlots => m_plgmsSorted.Values;
        public IEnumerable<GameSlot> SortedGameSlotsByGameNumber => m_plgmsSortedGameNum.Values;

        // NOTE:  This isn't just Team -> Count.  This is also Team-Sport -> Count.
        private Dictionary<string, int> m_mpTeamCount;

        public IEnumerable<string> AnalysisLegend => m_plsAnalysisLegend;
        public IEnumerable<string> MiscHeadings => m_plsMiscHeadings;
        
        
        /*----------------------------------------------------------------------------
			%%Function:SetMiscHeadings
			%%Qualified:ArbWeb.Games.ScheduleGames.SetMiscHeadings
        ----------------------------------------------------------------------------*/
        public void SetMiscHeadings(List<string> plsMisc)
        {
            m_plsMiscHeadings = plsMisc;
        }

        /*----------------------------------------------------------------------------
			%%Function:ScheduleGames
			%%Qualified:ArbWeb.Games.ScheduleGames.ScheduleGames
        ----------------------------------------------------------------------------*/
        public ScheduleGames(StatusBox srpt)
        {
            m_plgmsSorted = new SortedList<string, GameSlot>();
            m_plgmsSortedGameNum = new SortedList<string, GameSlot>();
            m_srpt = srpt;
            m_mpSportSport = new Dictionary<string, Sport>();
            m_plsAnalysisLegend = new List<string>();
            m_mpTeamCount = new Dictionary<string, int>();
            m_plgmSorted = new SortedList<string, Game>();
            m_mpnumgm = new Dictionary<string, Game>();
            UnitTest();
        }

        /*----------------------------------------------------------------------------
            %%Function: EnsureSportLevelPos
			%%Qualified:ArbWeb.Games.ScheduleGames.EnsureSportLevelPos

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
                m_plsAnalysisLegend.Add($"{sSport}-{sLevel}-{sPos}");

            if (fNewSport)
                m_plsAnalysisLegend.Add($"{sSport}-Total");

            if (fNewLevel)
                m_plsAnalysisLegend.Add($"{sSport}-{sLevel}-Total");

            if (fNewPos)
                m_plsAnalysisLegend.Add($"{sSport}-{sPos}");
        }

        /*----------------------------------------------------------------------------
			%%Function:AddStaticLegendColumns
			%%Qualified:ArbWeb.Games.ScheduleGames.AddStaticLegendColumns

			Add the static columns that are always in the analysis report			
        ----------------------------------------------------------------------------*/
        public void AddStaticLegendColumns()
        {
	        m_plsAnalysisLegend.Sort();
	        m_plsAnalysisLegend.Insert(0, "UmpireName");
	        m_plsAnalysisLegend.Insert(1, "Team");
	        m_plsAnalysisLegend.Insert(2, "Email");
	        m_plsAnalysisLegend.Insert(3, "Game");
	        m_plsAnalysisLegend.Insert(4, "DateTime");
	        m_plsAnalysisLegend.Insert(5, "Level");
	        m_plsAnalysisLegend.Insert(6, "Home");
	        m_plsAnalysisLegend.Insert(7, "Away");
	        m_plsAnalysisLegend.Insert(8, "Description");
	        m_plsAnalysisLegend.Insert(9, "Cancelled");
	        m_plsAnalysisLegend.Insert(10, "Total");
	        m_plsAnalysisLegend.Add("$$$MISC$$$"); // want this at the end!
        }

        /*----------------------------------------------------------------------------
            %%Function: AddGame
			%%Qualified:ArbWeb.Games.ScheduleGames.AddGame

            Add a game

            This handles ensuring the the sport/level/pos information has been added
            to the legend, so we can build the detail lines later
        ----------------------------------------------------------------------------*/
        public void AddGame(DateTime dttm, string sSite, string sName, string sTeam, string sEmail, string sGameNum, string sHome, string sAway, string sLevel, string sSport, string sPos, string sStatus, bool fCancelled, List<string> plsMisc)
        {
            GameSlot gm = new GameSlot(dttm, sSite, sName, sTeam, sEmail, sGameNum, sHome, sAway, sLevel, sSport, sPos, sStatus, fCancelled, plsMisc);
            AddGame(gm);
        }

        /*----------------------------------------------------------------------------
            %%Function: AddGame
			%%Qualified:ArbWeb.Games.ScheduleGames.AddGame
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

        /*----------------------------------------------------------------------------
			%%Function:SplitTeams
			%%Qualified:ArbWeb.Games.ScheduleGames.SplitTeams
        ----------------------------------------------------------------------------*/
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

        /*----------------------------------------------------------------------------
			%%Function:TeamCount
			%%Qualified:ArbWeb.Games.ScheduleGames.TeamCount
        ----------------------------------------------------------------------------*/
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

        /*----------------------------------------------------------------------------
			%%Function:FTeamMatchesSport
			%%Qualified:ArbWeb.Games.ScheduleGames.FTeamMatchesSport

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

        /*----------------------------------------------------------------------------
            %%Function: ReduceTeams
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

        /*----------------------------------------------------------------------------
			%%Function:GamesFromFilter
			%%Qualified:ArbWeb.Games.ScheduleGames.GamesFromFilter
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

        /*----------------------------------------------------------------------------
			%%Function:RequiredRanks
			%%Qualified:ArbWeb.Games.ScheduleGames.RequiredRanks
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

        /*----------------------------------------------------------------------------
			%%Function:GenOpenSlots
			%%Qualified:ArbWeb.Games.ScheduleGames.GenOpenSlots
        ----------------------------------------------------------------------------*/
        public SlotAggr GenOpenSlots(DateTime dttmStart, DateTime dttmEnd)
        {
	        return SlotAggr.Gen(m_plgmsSorted, dttmStart, dttmEnd, null, null, true);
        }

        /*----------------------------------------------------------------------------
			%%Function:GetOpenSlotSports
			%%Qualified:ArbWeb.Games.ScheduleGames.GetOpenSlotSports
        ----------------------------------------------------------------------------*/
        public string[] GetOpenSlotSports(SlotAggr sa)
        {
            string[] rgs = new string[sa.Sports.Length];
            sa.Sports.CopyTo(rgs, 0);
            return rgs;
        }

        /*----------------------------------------------------------------------------
			%%Function:GetSiteRosterSites
			%%Qualified:ArbWeb.Games.ScheduleGames.GetSiteRosterSites
        ----------------------------------------------------------------------------*/
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

        /*----------------------------------------------------------------------------
			%%Function:GetOpenSlotSportLevels
			%%Qualified:ArbWeb.Games.ScheduleGames.GetOpenSlotSportLevels
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