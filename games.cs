using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using NUnit.Framework;
using StatusBox;

namespace ArbWeb
{
    // ================================================================================
    //  G A M E  D A T A 
    // ================================================================================
    public class GameData
    {
        public class Roster // RST
        {
            private ArbWeb.Roster m_rst;
            private Dictionary<string, Umpire> m_mpNameUmpire;
            private StatusRpt m_srpt;

            public List<string> PlsMiscHeadings { get { return m_rst.PlsMisc; } }
            public RosterEntry RsteLookupEmail(string sEmail)
            {
                return m_rst.RsteLookupEmail(sEmail);
            }
            public string SMiscHeader(int i)
            {
                if (m_rst.PlsMisc != null)
                    return m_rst.PlsMisc[i];
                return "";
            }

            public Roster(StatusRpt srpt)
            {
                m_mpNameUmpire = new Dictionary<string, Umpire>();
                m_rst = new ArbWeb.Roster();
                m_srpt = srpt;
            }

            public bool LoadRoster(string sRoster, int iMiscAffiliation)
            {
                m_rst.ReadRoster(sRoster);
                foreach (RosterEntry rste in m_rst.Plrste)
                    {
                    Umpire ump = new Umpire(rste.First, rste.Last, rste.m_plsMisc[iMiscAffiliation], rste.Email, rste.m_plsMisc);

                    m_mpNameUmpire.Add(ump.Name, ump);
                    }
                return true;
            }

            public Umpire UmpireLookup(string sName)
            {
                if (m_mpNameUmpire.ContainsKey(sName))
                    return m_mpNameUmpire[sName];

                return null;
            }
        } // END ROSTER

        // ================================================================================
        //  G A M E  S L O T
        // ================================================================================
        public class GameSlot // GM
        {
            private DateTime m_dttm;
            private string m_sSite;
            private string m_sName;
            private string m_sHome;
            private string m_sAway;
            private string m_sSport;
            private string m_sPos;
            private string m_sStatus;
            private string m_sTeam;
            private string m_sEmail;
            private string m_sGameNum;
            private string m_sLevel;
            private bool m_fCancelled;
            private List<string> m_plsMisc;

            public GameSlot(DateTime dttm, string sSite, string sName, string sTeam, string sEmail, string sGameNum, string sHome, string sAway, string sLevel, string sSport, string sPos, string sStatus, bool fCancelled, List<string> plsMisc)
            {
                m_dttm = dttm;
                m_sSite = sSite;
                m_sName = sName;
                m_sHome = sHome;
                m_sAway = sAway;
                m_sSport = sSport;
                m_sPos = sPos;
                m_sStatus = sStatus;
                m_sTeam = sTeam;
                m_sEmail = sEmail;
                m_sGameNum = sGameNum;
                m_fCancelled = fCancelled;
                m_sLevel = sLevel;
                m_plsMisc = plsMisc;
            }

            public string Status { get { return m_sStatus; } }
            public string Email { get { return m_sEmail; } }
            public string Name { get { return m_sName; } }
            public List<string> PlsMisc { get { return m_plsMisc; } set { m_plsMisc = value; } }
            public string Team { get { return m_sTeam; } set { m_sTeam = value; } }
            public bool Open { get { return m_sName == null; } }
            public string Home { get { return m_sHome; } }
            public string Away { get { return m_sAway; } }
            public string GameNum { get { return m_sGameNum; } }
            public bool Cancelled { get { return m_fCancelled; } }
            public string Level { get { return m_sLevel; } }
            public string Sport { get { return m_sSport; } }
            public string Pos { get { return m_sPos; } }
            public DateTime Dttm { get { return m_dttm; } }
            public string Site { get { return m_sSite; } }
            public string SportLevel { get { return $"{m_sSport} {m_sLevel}"; } }

            public string SiteShort
            {
                get
                {
                    // get rid of any trailing fields
                    string s = Regex.Replace(m_sSite, " [A-D]$", "");

                    s = Regex.Replace(s, " #[1-9]$", "");
                    s = Regex.Replace(s, " Big$", "");

                    s = Regex.Replace(s, " South$", "");

                    s = Regex.Replace(s, " East$", "");
                    s = Regex.Replace(s, " West$", "");
                    s = Regex.Replace(s, " North$", "");
                    s = Regex.Replace(s, " Varsity Field$", "");
                    s = Regex.Replace(s, " JV Field$", "");

                    s = Regex.Replace(s, " #[1-9][ ]*[69]0'$", "");
                    s = Regex.Replace(s, " #[1-9][ ]*\\([69]0'\\)$", "");
                    return s;
                }
            }

            /* S  R E P O R T */
            /*----------------------------------------------------------------------------
					%%Function: SReport
					%%Qualified: ArbWeb.CountsData:GameData:Game.SReport
					%%Contact: rlittle

					Return a detail string suitable for saving in CSV format.

					Takes the given legend and saves out our collected data according to that
					legend.
				----------------------------------------------------------------------------*/
            public string SReport(List<string> plsLegend)
            {
                Dictionary<string, string> m_mpFieldVal = new Dictionary<string, string>();

                m_mpFieldVal.Add("UmpireName", "\"" + m_sName + "\"");
                m_mpFieldVal.Add("Team", m_sTeam);
                m_mpFieldVal.Add("Email", m_sEmail);
                m_mpFieldVal.Add("Game", m_sGameNum);
//					string sDateTime = m_dttm.ToString("M/d/yyyy ddd h:mm tt");
                string sDateTime = m_dttm.ToString("M/d/yyyy H:mm");

                m_mpFieldVal.Add("DateTime", sDateTime);
                m_mpFieldVal.Add("Date", m_dttm.ToString("M/d/yyyy"));
                m_mpFieldVal.Add("Time", m_dttm.ToString("H:mm"));
                m_mpFieldVal.Add("Level", m_sLevel);
                m_mpFieldVal.Add("Home", m_sHome);
                m_mpFieldVal.Add("Away", m_sAway);
                m_mpFieldVal.Add("Site", m_sSite);
                m_mpFieldVal.Add("Description", $"{m_sPos}: [{m_sGameNum}] {sDateTime}: {m_sHome} vs. {m_sAway} ({m_sSport} {m_sLevel})");
                m_mpFieldVal.Add("Cancelled", m_fCancelled ? "1" : "0");

                m_mpFieldVal.Add("Sport", m_sSport);

                string sSportLevelPos = $"{m_sSport}-{m_sLevel}-{m_sPos}";
                string sSportTotal = $"{m_sSport}-Total";
                string sSportLevelTotal = $"{m_sSport}-{m_sLevel}-Total";
                string sSportPos = $"{m_sSport}-{m_sPos}";
                string sTotal = String.Format("Total");

                m_mpFieldVal.Add(sSportLevelPos, "1");
                m_mpFieldVal.Add(sSportTotal, "1");
                m_mpFieldVal.Add(sSportLevelTotal, "1");
                m_mpFieldVal.Add(sSportPos, "1");
                m_mpFieldVal.Add(sTotal, "1");

                // now that we have a dictionary of values, write it out
                bool fFirst = true;
                string sRet = "";

                foreach (string s in plsLegend)
                    {
                    if (s == "$$$MISC$$$") // expand the Misc values here
                        {
                        // if there's no misc data for this game, no worries -- misc should
                        // be the last entry!
                        if (m_plsMisc == null)
                            continue;

                        foreach (string sMisc in m_plsMisc)
                            {
                            sRet += ",\"" + sMisc + "\"";
                            }
                        continue;
                        }

                    if (fFirst != true)
                        {
                        sRet += ",";
                        }
                    fFirst = false;

                    if (m_mpFieldVal.ContainsKey(s))
                        sRet += m_mpFieldVal[s];
                    else
                        sRet += "0";
                    }
                return sRet;
            }

        }

        public class Game
        {
            private int m_cSlots;
            private int m_cOpen;

            private List<GameSlot> m_plgms;

            public Game()
            {
                m_plgms = new List<GameSlot>();
            }

            public void AddGameSlot(GameSlot gms)
            {
                m_plgms.Add(gms);
                m_cSlots++;
                if (gms.Open)
                    m_cOpen++;
            }

            public int OpenSlots { get { return m_cOpen; } }
            public int TotalSlots { get { return m_cSlots; } }
            public List<GameSlot> Slots { get { return m_plgms; } } 
        }

        // ================================================================================
        //  G A M E S  S L O T S
        // ================================================================================
        public class GameSlots // GMSS
        {
            private const int icolGameLevelDefault = 6;
            private const int icolGameDateTimeDefault = 2;
            private const int icolGameHomeColumnDeltaFromSportLevel = 14 - icolGameLevelDefault;
            private const int icolGameSiteDeltaFromSportLevel = 11 - icolGameLevelDefault;
            private const int icolGameAwayDeltaFromSportLevel = 19 - icolGameLevelDefault;
            private const int icolOfficialDeltaFromSportLevel = 5 - icolGameLevelDefault; // yes this is a negative delta
            private const int icolSlotStatusDeltaFromSportLevel = 17 - icolGameLevelDefault; // THIS IS UNVERIFIED!
            private const int icolGameGameBase = 0;
            
            private int GameGameColumn => icolGameGameBase;
            private int GameSiteColumn => icolGameSiteDeltaFromSportLevel + GameLevelColumn;
            private int GameHomeColumn => icolGameHomeColumnDeltaFromSportLevel + GameLevelColumn;
            private int GameAwayColumn => icolGameAwayDeltaFromSportLevel + GameLevelColumn;
            private int OfficialColumn => icolOfficialDeltaFromSportLevel + GameLevelColumn;
            private int SlotStatusColumn => icolSlotStatusDeltaFromSportLevel + GameLevelColumn;

            private int GameLevelColumn { get; set; }
            private int GameDateTimeColumn { get; set; }

            public GameSlots() {} // just for unit tests

            public SortedList<string, Game> Games { get { return m_plgmSorted; } }

            // ================================================================================
            //  S P O R T 
            // ================================================================================
            public class Sport
            {
                private SortedList<string, string> m_plLevelPos;
                private SortedList<string, string> m_plLevel;
                private SortedList<string, string> m_plPos;

                public Sport()
                {
                    m_plLevelPos = new SortedList<string, string>();
                    m_plLevel = new SortedList<string, string>();
                    m_plPos = new SortedList<string, string>();
                }

                /* E N S U R E  P O S */
                /*----------------------------------------------------------------------------
						%%Function: EnsurePos
						%%Qualified: ArbWeb.CountsData:GameData:Games:Sport.EnsurePos
						%%Contact: rlittle

						Returns true if we needed to add the position
					----------------------------------------------------------------------------*/
                public void EnsurePos(string sLevel, string sPos, out bool fNewLevel, out bool fNewPos, out bool fNewLevelPos)
                {
                    string sKey;
                    string sLevelPos = sLevel + "-" + sPos;

                    sKey = sLevelPos;
                    fNewLevel = fNewPos = fNewLevelPos = false;

                    if (!m_plLevelPos.ContainsKey(sKey))
                        {
                        m_plLevelPos.Add(sKey, sLevelPos);
                        fNewLevelPos = true;

                        if (!m_plLevel.ContainsKey(sLevel))
                            {
                            m_plLevel.Add(sLevel, sLevel);
                            fNewLevel = true;
                            }
                        if (!m_plPos.ContainsKey(sPos))
                            {
                            m_plPos.Add(sPos, sPos);
                            fNewPos = true;
                            }
                        }
                }
            }

            private SortedList<string, Game> m_plgmSorted;
            private Dictionary<string, Game> m_mpnumgm;
 
            private List<string> m_plsMiscHeadings;
            private SortedList<string, GameSlot> m_plgmsSorted;
            private SortedList<string, GameSlot> m_plgmsSortedGameNum;
            private StatusRpt m_srpt;
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
            public GameSlots(StatusRpt srpt)
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
            private void AddGame(DateTime dttm, string sSite, string sName, string sTeam, string sEmail, string sGameNum, string sHome, string sAway, string sLevel, string sSport, string sPos, string sStatus, bool fCancelled, List<string> plsMisc)
            {
                GameSlot gm = new GameSlot(dttm, sSite, sName, sTeam, sEmail, sGameNum, sHome, sAway, sLevel, sSport, sPos, sStatus, fCancelled, plsMisc);
                AddGame(gm);
#if old
                string sTeamSport = sTeam + "#-#" + sSport;

                m_plgmSorted.Add(String.Format("{0}_{1}_{2}", sName, dttm.ToString("yyyyMMdd:HH:mm"), m_plgmSorted.Count), gm);
                m_plgmSortedGameNum.Add(String.Format("{0}_{1}_{2}_{3}_{4}_{5}", dttm.ToString("yyyyMMdd:HH:mm"), sSite, sSport, sLevel, sGameNum, m_plgmSortedGameNum.Count), gm);
                if (sTeam != null && sTeam.Length > 0)
                    {
                    if (m_mpTeamCount.ContainsKey(sTeam))
                        m_mpTeamCount[sTeam]++;
                    else
                        m_mpTeamCount.Add(sTeam, 1);

                    if (m_mpTeamCount.ContainsKey(sTeamSport))
                        m_mpTeamCount[sTeamSport]++;
                    else
                        m_mpTeamCount.Add(sTeamSport, 1);

                    EnsureSportLevelPos(sSport, sLevel, sPos);
                    }
#endif
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

            private class DistributeTeamCount   // DTC
            {
                private class DND
                {
                    private string m_sTeam;
                    private int m_c;
                    private int m_dc;

                    public DND(string sTeam, int c)
                    {
                        m_c = c;
                        m_sTeam = sTeam;
                        m_dc = 0;
                    }

                    public void AddCount(int c)
                    {
                        m_dc += c;
                    }

                    public void SubCount(int c)
                    {
                        m_dc -= c;
                    }

                    public int Count { get { return m_c + m_dc; } }
                    public string Name { get { return m_sTeam; } }
                    public int DCount { get { return m_dc; } set { m_dc = value; } }
                };

                private List<DND> m_pldnd;

                public DistributeTeamCount()
                {
                    m_pldnd = new List<DND>();
                }

                public string STeamNext()
                {
                    foreach (DND dnd in m_pldnd)
                        {
                        if (dnd.DCount == 0)
                            continue;

                        return dnd.Name;
                        }
                    throw new Exception("could not find team->dtc mapping");
                }

                public void DecTeamNext()
                {
                    foreach (DND dnd in m_pldnd)
                        {
                        if (dnd.DCount != 0)
                            {
                            dnd.DCount--;
                            return;
                            }
                        }
                    throw new Exception("could not find team->dtc mapping");
                }

                public void AddTeam(string sTeam, int c)
                {
                    int i;

                    for (i = 0; i < m_pldnd.Count; i++)
                        {
                        DND dnd = m_pldnd[i];

                        if (dnd.Count > c)
                            {
                            m_pldnd.Insert(i, new DND(sTeam, c));
                            break;
                            }
                        }
                    if (i >= m_pldnd.Count)
                        m_pldnd.Add(new DND(sTeam, c));
                }

                /* D I S T R I B U T E */
                /*----------------------------------------------------------------------------
						%%Function: Distribute
						%%Qualified: ArbWeb.CountsData:GameData:Games:DTC.Distribute
						%%Contact: rlittle

					----------------------------------------------------------------------------*/
                public void Distribute(int c)
                {
                    List<DND> pldndUse = m_pldnd;

                    if (m_pldnd.Count == 0)
                        return;

#if DIST_USE_SPORT
    // first, let's see if there are multiple teams for the sport we're trying to distribute...
						foreach(DND dnd in m_pldnd)
							{
							if (FTeamMatchesSport(dnd.Name, sSport))
								pldndUse.Add(dnd);
							}

						if (pldndUse.Count == 1)
							{
							// easy, everythign goes to the one team/sport match
							pldndUse[0].AddCount(c);
							c = 0;
							}
						else if (pldndUse.Count == 0)
							{
							// nobody in the sport.  distribute to everyone regardless of sport
							pldndUse = m_pldnd;
							}
#endif // DIST_USE_SPORT

                    // ok, the idea here is, we always take the lowest team(s) and give
                    // them games until they match the next team, until all teams match
                    int iMac = 0;
                    int cMin = pldndUse[0].Count;
                    int cNext = 0;

                    // m_pldnd[0..iMac] have the same value, and we're trying to get to
                    // cNext

                    while (c > 0)
                        {
                        // find out how many entries *after* iMac match us...

                        while (iMac < pldndUse.Count && cMin == pldndUse[iMac].Count)
                            iMac++;

                        if (iMac >= pldndUse.Count)
                            cNext = Int16.MaxValue;
                        else
                            cNext = pldndUse[iMac].Count;

                        iMac--;

                        int iInner;
                        int cDist = Math.Min((cNext - cMin) * (iMac + 1), c);

                        // we have iMac+1 teams to distribute this to
                        int cEachMin = cDist / (iMac + 1);

                        // each team will get at least cEachMin
                        int cRemain = cDist - cEachMin * (iMac + 1);

                        // and cRemain will get 1 additional to distribute
                        // the remainder
                        for (iInner = 0; iInner <= iMac; iInner++)
                            {
                            pldndUse[iInner].AddCount(cEachMin + (cRemain > 0 ? 1 : 0));
                            cRemain--;
                            }

                        c -= cDist;
                        cMin = cNext;
                        }
                    // ok, distribution done.
#if DIST_USE_SPORT
    // now, update m_pldnd if we weren't working directly with it
						if (m_pldnd != pldndUse)
							{
							int idnd = 0, idndMac = m_pldnd.Count;

							foreach(DND dnd in pldndUse)
								{
								while (idnd < idndMac && String.Compare(dnd.Name, m_pldnd[idnd].Name) != 0)
									idnd++;

								if (idnd >= idndMac)
									throw new Exception("internal error -- couldn't find the dnd that's guaranteed to be there!!");


								m_pldnd[idnd] = dnd;
								}
							}
#endif // DIST_USE_SPORT
                }

                /* U P D A T E  T E A M  T O T A L S */
                /*----------------------------------------------------------------------------
						%%Function: UpdateTeamTotals
						%%Qualified: ArbWeb.CountsData:GameData:Games:DTC.UpdateTeamTotals
						%%Contact: rlittle
						
					----------------------------------------------------------------------------*/
                public void UpdateTeamTotals(Dictionary<string, int> mpTeamCount, string sSport)
                {
                    foreach (DND dnd in m_pldnd)
                        UpdateTeamCount(mpTeamCount, dnd.Name, sSport, dnd.DCount);
                }
            }

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

            /* U P D A T E  T E A M  C O U N T */
            /*----------------------------------------------------------------------------
					%%Function: UpdateTeamCount
					%%Qualified: ArbWeb.CountsData:GameData:Games.UpdateTeamCount
					%%Contact: rlittle
					
			----------------------------------------------------------------------------*/
            public static void UpdateTeamCount(Dictionary<string, int> mpTeamCount, string sTeam, string sSport, int dCount)
            {
                string sTeamSport = $"{sTeam}#-#{sSport}";

                if (!mpTeamCount.ContainsKey(sTeam))
                    mpTeamCount.Add(sTeam, dCount);
                else
                    mpTeamCount[sTeam] += dCount;

                if (!mpTeamCount.ContainsKey(sTeamSport))
                    mpTeamCount.Add(sTeamSport, dCount);
                else
                    mpTeamCount[sTeamSport] += dCount;
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

            public enum ReadState
            {
                ScanForHeader = 1,
                ScanForGame = 2,
                ReadingGame1 = 3,
                ReadingGame2 = 4,
                ReadingOfficials1 = 5,
                ReadingOfficials2 = 6,
                ReadingComments = 7,
            };

            /* A P P E N D  C H E C K */
            /*----------------------------------------------------------------------------
				%%Function: AppendCheck
				%%Qualified: ArbWeb.CountsData:GameData:Games.AppendCheck
				%%Contact: rlittle

				Append s to sAppend -- deals with leading and trailing spaces as well
				as making sure there are spaces separating the arguments
			----------------------------------------------------------------------------*/
            private string AppendCheck(string s, string sAppend)
            {
                sAppend = Regex.Replace(sAppend, "^ *", "");
                sAppend = Regex.Replace(sAppend, " *$", "");
                if (sAppend.Length > 0)
                    {
                    if (s.Length > 1)
                        s = s + " " + sAppend;
                    else
                        s = sAppend;
                    }
                return s;
            }

            private void UnitTest()
            {
                string s = "", s2 = "", s3 = "";

                s = ReverseName(null, "Mary Van Whatsa Hoozit");
                Debug.Assert(String.Compare(s, "Van Whatsa Hoozit,Mary") == 0);
                Debug.Assert(ArbWeb.Roster.FSplitName(s, out s2, out s3));
                Debug.Assert(s2 == "Mary");
                Debug.Assert(s3 == "Van Whatsa Hoozit");

                s = "Van Whatsa Hoozit, Mary";

                Debug.Assert(ArbWeb.Roster.FSplitName(s, out s2, out s3));
                Debug.Assert(s2 == "Mary");
                Debug.Assert(s3 == "Van Whatsa Hoozit");

                s = "";
                s = AppendCheck(s, "Foo");
                Debug.Assert(String.Compare(s, "Foo") == 0);
                s = AppendCheck(s, " Bar ");
                Debug.Assert(String.Compare(s, "Foo Bar") == 0);
                s = AppendCheck(s, " Baz ");
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

            /* R E V E R S E  N A M E  S I M P L E */
            /*----------------------------------------------------------------------------
            	%%Function: ReverseNameSimple
            	%%Qualified: ArbWeb.GameData.GameSlots.ReverseNameSimple
            	%%Contact: rlittle

				Reverse the given "First Last" into "Last,First"

				Handles things like "van Doren, Martin"
            ----------------------------------------------------------------------------*/
            static string ReverseNameSimple(string s)
            {
                string[] rgs;

                rgs = CountsData.RexHelper.RgsMatch(s, "^[ \t]*([^ \t]*) ([^\t]*) *$");
                if (rgs.Length < 2)
                    return s;

                if (rgs[0] == null || rgs[1] == null)
                    return s;
                return $"{rgs[1]},{rgs[0]}";
            }

            [Test]
            [TestCase("Rob Little", "Little,Rob")]
            [TestCase("Martin van Doren", "van Doren,Martin")]
            [TestCase("Byron (Barney) Kinzer", "(Barney) Kinzer,Byron")]
            public static void TestReverseNameSimple(string sIn, string sExpected)
            {
                string sActual = ReverseNameSimple(sIn);

                Assert.AreEqual(sExpected, sActual);
            }

            /* R E V E R S E  N A M E  S I M P L E  P A R E N T H E T I C A L */
            /*----------------------------------------------------------------------------
            	%%Function: ReverseNameParenthetical
            	%%Qualified: ArbWeb.GameData.GameSlots.ReverseNameParenthetical
            	%%Contact: rlittle
            	
                Same as ReverseNameSimple, but allows parenthetical first names

                Byron (Barney) Kinzer => "Kinzer,Byron (Barney)"
            ----------------------------------------------------------------------------*/
            static string ReverseNameParenthetical(string s)
            {
                // first, see if we match a parenthetical
                if (s.IndexOf('(') <= 0)
                    return ReverseNameSimple(s);

                // that was just the quick check. now we need to see if there's a
                // parenthetical match between the first and last name
                string[] rgs;

                rgs = CountsData.RexHelper.RgsMatch(s, "^[ \t]*([^ \t]*)([ \t]*\\([^ \t]*\\))[ \t]*([^\t]*) *$");
                if (rgs.Length < 3)
                    return ReverseNameSimple(s);

                if (rgs[0] == null || rgs[1] == null || rgs[2] == null)
                    return ReverseNameSimple(s);

                return $"{rgs[2]},{rgs[0]}{rgs[1]}";
            }

            [Test]
            [TestCase("Rob Little", "Little,Rob")]
            [TestCase("Martin van Doren", "van Doren,Martin")]
            [TestCase("Byron (Barney) Kinzer", "Kinzer,Byron (Barney)")]
            [TestCase("Byron(Barney) Kinzer", "Kinzer,Byron(Barney)")]
            [TestCase("(Barney)Byron Kinzer", "Kinzer,(Barney)Byron")]
            [TestCase("(Barney) Byron Kinzer", "Byron Kinzer,(Barney)")]
            public static void TestReverseNameParenthetical(string sIn, string sExpected)
            {
                string sActual = ReverseNameParenthetical(sIn);

                Assert.AreEqual(sExpected, sActual);
            }

            /* R E V E R S E  N A M E */
            /*----------------------------------------------------------------------------
				%%Function: ReverseName
				%%Qualified: ArbWeb.CountsData:GameData:Games.ReverseName
				%%Contact: rlittle

                try to figure out the right reversal for the given name. Try each type of
                reversal we know about and see if it matches an entry in the roster. if not, 
                try the next one.  If all else fails, just return the simple reverse
			----------------------------------------------------------------------------*/
            private static string ReverseName(Roster rst, string s)
            {
                string sReverse, sFallback;

                sFallback = sReverse = ReverseNameSimple(s);
                if (rst == null)
                    return sFallback;

                if (rst.UmpireLookup(sReverse) != null)
                    return sReverse;

                sReverse = ReverseNameParenthetical(s);
                if (rst.UmpireLookup(sReverse) != null)
                    return sReverse;

                return sFallback;
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
            [TestCase(new String[] { "foo"}, new string[] {})]
            [TestCase(new String[] { "foo", "bar" }, new string[] { })]
            [TestCase(new String[] { "foo", "foo2", "bar" }, new string[] { "foo|foo", "foo2|foo"})]
            [TestCase(new String[] { "foo", "foo2", "bar2", "bar3", "bar4", "boo"}, new string[] { "foo|foo", "foo2|foo", "bar2|bar", "bar3|bar", "bar4|bar" })]
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
            public GameSlots GamesFromFilter(string[] rgsSportFilter, string[] rgsSportLevelFilter, bool fOpenOnly, SlotAggr sa)
            {
                GameSlots gms = new GameSlots(m_srpt);

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

            /* F  L O A D  G A M E S */
            /*----------------------------------------------------------------------------
				%%Function: FLoadGames
				%%Qualified: ArbWeb.CountsData:GameData:Games.FLoadGames
				%%Contact: rlittle

				Loading the games needs a state machine -- this is a multi line report
			----------------------------------------------------------------------------*/
            public bool FLoadGames(string sGamesReport, Roster rst, bool fIncludeCanceled)
            {
                using (TextReader tr = new StreamReader(sGamesReport))
                {
                    string sLine;
                    string[] rgsFields;
                    ReadState rs = ReadState.ScanForHeader;
                    bool fCanceled = false;
                    bool fOpenSlot = false;
                    bool fIgnore = false;
                    string sGame = "";
                    string sDateTime = "";
                    string sSport = "";
                    string sLevel = "";
                    string sSite = "";
                    string sHome = "";
                    string sAway = "";
                    string sPosLast = "";
                    string sNameLast = "";
                    string sStatusLast = "";

                    Dictionary<string, string> mpNameStatus = new Dictionary<string, string>();
                    Dictionary<string, string> mpNamePos = new Dictionary<string, string>();
                    // m_mpNameSportLevelCount = new Dictionary<string, Dictionary<string, int>>();
                    Umpire ump = null;

                    while ((sLine = tr.ReadLine()) != null)
                    {
                        // first, change "foo, bar" into "foo bar" (get rid of quotes and the comma)
                        sLine = Regex.Replace(sLine, "\"([^\",]*),([^\",]*)\"", "$1$2");

                        if (sLine.Length < GameDateTimeColumn)
                            continue;

                        Regex rex = new Regex(",");
                        rgsFields = rex.Split(sLine);

                        // check for rainouts and cancellations
                        if (FMatchGameCancelled(sLine))
                        {
                            fCanceled = true;
                            // drop us back to reading officials
                            rs = ReadState.ReadingOfficials1;
                            continue;
                        }

                        // look for comments
                        if (FMatchGameComment(sLine))
                        {
                            rs = RsHandleGameComment(rs, sLine);
                            continue;
                        }

                        if (FMatchGameEmpty(sLine))
                            continue;

                        if (FMatchGameTotalLine(rgsFields))
                        {
                            // this is the final "total" line.  the only thing that should follow this is
                            // the final page break
                            // just leave the rs alone for now...
                            continue;
                        }

                        if (rs == ReadState.ScanForHeader)
                        {
                            rs = RsHandleScanForHeader(sLine, rgsFields, rs);
                            continue;
                        }

                        if (rs == ReadState.ReadingComments)
                        {
                            // when reading comments, we can get text in column 1; if the line ends with commas, then this is just
                            // a continuation of the comment (also be careful to look for another comment starting right after ours
                            // ends
                            if (FMatchGameCommentContinuation(sLine))
                            {
                                continue;
                            }

                            rs = ReadState.ReadingOfficials1;
                            // drop back to reading officials
                        }

                        if (rs == ReadState.ReadingGame2)
                            rs = RsHandleReadingGame2(rgsFields, ref sGame, ref sDateTime, ref sLevel, ref sSite,
                                ref sHome, ref sAway, rs);

                        if (rs == ReadState.ReadingOfficials2)
                            rs = RsHandleReadingOfficials2(rst, rgsFields, mpNamePos, mpNameStatus, sNameLast, sPosLast,
                                sStatusLast, rs);

                        if (rs == ReadState.ReadingOfficials1)
                            rs = RsHandleReadingOfficials1(rst, fIncludeCanceled, sLine, rgsFields, mpNamePos,
                                mpNameStatus, fCanceled, sSite, sGame,
                                sHome, sAway, sLevel, sSport, rs, ref sPosLast, ref sNameLast, ref sStatusLast,
                                ref sDateTime, ref fOpenSlot, ref ump);

                        if (FMatchGameArbiterFooter(sLine))
                        {
                            Debug.Assert(
                                rs == ReadState.ReadingComments || rs == ReadState.ScanForHeader ||
                                rs == ReadState.ScanForGame,
                                $"Page break at illegal position: state = {rs}");
                            rs = ReadState.ScanForHeader;
                            continue;
                        }

                        if (rs == ReadState.ScanForGame)
                            rs = RsHandleScanForGame(ref sGame, mpNamePos, mpNameStatus, sLine, ref sDateTime,
                                ref sSport, ref sLevel, ref sSite,
                                ref sHome, ref sAway, ref fCanceled, ref fIgnore, rs);

                        if (rs == ReadState.ReadingGame1)
                            rs = RsHandleReadingGame1(ref sGame, rgsFields, ref sDateTime, ref sSport, ref sSite,
                                ref sHome, ref sAway, rs);
                    }
                }

                return true;
            }

            /* R S  H A N D L E  S C A N  F O R  H E A D E R */
            /*----------------------------------------------------------------------------
			    	%%Function: RsHandleScanForHeader
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleScanForHeader
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private ReadState RsHandleScanForHeader(string sLine, string[] rgsFields, ReadState rs)
            {
                if (Regex.Match(sLine, "Game.*Date.*Sport.*Level").Success == false)
                    return rs;

                // always start looking here, and adjust
                GameDateTimeColumn = icolGameDateTimeDefault;
                GameLevelColumn = icolGameLevelDefault;

                if (!Regex.Match(rgsFields[GameLevelColumn], "Sport.*Level").Success)
                {
                    // check to see if the previous column is Sport/Level, and if so, automagically adjust
                    if (Regex.Match(rgsFields[GameLevelColumn - 1], "Sport.*Level").Success)
	                    GameLevelColumn--;
                    else if (Regex.Match(rgsFields[GameLevelColumn + 1], "Sport.*Level").Success)
	                    GameLevelColumn++;
                }

                if (!Regex.Match(rgsFields[GameDateTimeColumn], "Date.*Time").Success)
                {
	                // check to see if the previous column is Sport/Level, and if so, automagically adjust
	                if (Regex.Match(rgsFields[GameDateTimeColumn - 1], "Date.*Time").Success)
		                GameDateTimeColumn--;
	                else if (Regex.Match(rgsFields[GameDateTimeColumn + 1], "Date.*Time").Success)
		                GameDateTimeColumn++;
                }

                Debug.Assert(Regex.Match(rgsFields[GameDateTimeColumn], "Date.*Time").Success, "Date & time not where expected!!");
                Debug.Assert(Regex.Match(rgsFields[GameLevelColumn], "Sport.*Level").Success, "Sport & level not where expected!!");
                rs = ReadState.ScanForGame;
                return rs;
            }

            /* R S  H A N D L E  G A M E  C O M M E N T */
            /*----------------------------------------------------------------------------
			    	%%Function: RsHandleGameComment
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleGameComment
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private ReadState RsHandleGameComment(ReadState rs, string sLine)
            {
                rs = ReadState.ReadingComments;
//			        m_srpt.AddMessage("Reading comment: " + sLine);
                // skip comments from officials
                return rs;
            }

            /* R S  H A N D L E  R E A D I N G  G A M E  1 */
            /*----------------------------------------------------------------------------
			    	%%Function: RsHandleReadingGame1
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleReadingGame1
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private ReadState RsHandleReadingGame1(ref string sGame, string[] rgsFields, ref string sDateTime, ref string sSport,
                ref string sSite, ref string sHome, ref string sAway, ReadState rs)
            {
// reading the first line of the game.  We should always get the sport and the first part of the team names here
                sGame = AppendCheck(sGame, rgsFields[GameGameColumn]);
                sDateTime = AppendCheck(sDateTime, rgsFields[GameDateTimeColumn]);
                sSport = AppendCheck(sSport, rgsFields[GameLevelColumn]);
                sSite = AppendCheck(sSite, rgsFields[GameSiteColumn]);
                sHome = AppendCheck(sHome, rgsFields[GameHomeColumn]);
                sAway = AppendCheck(sAway, rgsFields[GameAwayColumn]);

                rs = ReadState.ReadingGame2;
                return rs;
            }

            /* R S  H A N D L E  S C A N  F O R  G A M E */
            /*----------------------------------------------------------------------------
			    	%%Function: RsHandleScanForGame
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleScanForGame
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private static ReadState RsHandleScanForGame(ref string sGame, Dictionary<string, string> mpNamePos, Dictionary<string, string> mpNameStatus, string sLine, ref string sDateTime,
                ref string sSport, ref string sLevel, ref string sSite, ref string sHome,
                ref string sAway, ref bool fCanceled, ref bool fIgnore, ReadState rs)
            {
                sGame = "";
                sDateTime = "";
                sSport = "";
                sLevel = "";
                sSite = "";
                sHome = "";
                sAway = "";
                fCanceled = false;
                fIgnore = false;
                mpNamePos.Clear();
                mpNameStatus.Clear();

                if (!(Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Baseball").Success
                      || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Interlock").Success
                      || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Interlock").Success
                      || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Tourn").Success
                      || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Softball").Success
                      || Regex.Match(sLine, ", *[ a-zA-Z0-9-']* *Postseason").Success
                      || Regex.Match(sLine, ", *[ a-zA-Z0-9-']* *All Stars").Success
                      || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Fall Ball").Success
                      || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Administrative").Success
                      || Regex.Match(sLine, ", *50/50").Success
                      || Regex.Match(sLine, ",_Events*").Success
                      || Regex.Match(sLine, ",zEvents*").Success
                      || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Training").Success))
                    Debug.Assert(false, $"failed to find game as expected!: {sLine} ({rs}");
                rs = ReadState.ReadingGame1;
                // fallthrough to ReadingGame1
                return rs;
            }

            /* F  M A T C H  G A M E  A R B I T E R  F O O T E R */
            /*----------------------------------------------------------------------------
			    	%%Function: FMatchGameArbiterFooter
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameArbiterFooter
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private static bool FMatchGameArbiterFooter(string sLine)
            {
                return Regex.Match(sLine, ".*Created by ArbiterSports").Success;
            }

            /* R S  H A N D L E  R E A D I N G  O F F I C I A L S  1 */
            /*----------------------------------------------------------------------------
			    	%%Function: RsHandleReadingOfficials1
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleReadingOfficials1
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private ReadState RsHandleReadingOfficials1(Roster rst, bool fIncludeCanceled, string sLine, string[] rgsFields,
                Dictionary<string, string> mpNamePos, Dictionary<string, string> mpNameStatus, bool fCanceled, string sSite, string sGame,
                string sHome, string sAway, string sLevel, string sSport, ReadState rs,
                ref string sPosLast, ref string sStatusLast, ref string sNameLast, ref string sDateTime, ref bool fOpenSlot, ref Umpire ump)
            {
// Games may have multiple officials, so we have to collect up the officials.
                // we do this in mpNamePos

                // &&&& TODO: omit dates before here...

                if (Regex.Match(sLine, "Attached").Success)
                    return rs;


                if (rgsFields[0].Length < 1)
                    {
                    // look for possible contiuation line; if not there, then we will fall back
                    // to ReadingOfficials1
                    rs = ReadState.ReadingOfficials2;
                    sPosLast = rgsFields[1];
                    sNameLast = rgsFields[OfficialColumn];
                    sStatusLast = rgsFields[SlotStatusColumn];

                    if (Regex.Match(rgsFields[OfficialColumn], "_____").Success)
                        {
                        fOpenSlot = true;
                        mpNamePos.Add($"!!OPEN{mpNamePos.Count}", rgsFields[1]);
                        mpNameStatus.Add($"!!OPEN{mpNameStatus.Count}", rgsFields[SlotStatusColumn]);
                        return rs;
                        }
                    else
                        {
                        string sName = ReverseName(rst, rgsFields[OfficialColumn]);
                        mpNamePos.Add(sName, rgsFields[1]);
                        mpNameStatus.Add(sName, rgsFields[SlotStatusColumn]);
                        return rs;
                        }
                    }
                // otherwise we're done!!
                //						m_srpt.AddMessage("recording results...");
                if (!fCanceled || fIncludeCanceled)
                    {
                    // record our game


                    // we've got all the info for one particular game and its officials.

                    // walk through the officials that we have
                    foreach (string sName in mpNamePos.Keys)
                        {
                        string sPos = mpNamePos[sName];
                        string sStatus = mpNameStatus[sName];
                        string sEmail;
                        string sTeam;
                        string sNameUse = sName;
                        List<string> plsMisc = null;

                        if (Regex.Match(sName, "!!OPEN.*").Success)
                            {
                            sNameUse = null;
                            sEmail = "";
                            sTeam = "";
                            sStatus = "";
                            }
                        else
                            {
                            ump = rst.UmpireLookup(sName);

                            if (ump == null)
                                {
                                if (sName != "")
                                    m_srpt.AddMessage(
	                                    $"Cannot find info for Umpire: {sName}",
                                                      StatusRpt.MSGT.Error);
                                sEmail = "";
                                sTeam = "";
                                }
                            else
                                {
                                sEmail = ump.Contact;
                                sTeam = ump.Misc;
                                plsMisc = ump.PlsMisc;
                                }
                            }
                        if (sPos != "Training")
                            {
                            if (sDateTime.EndsWith("TBA"))
                                sDateTime = sDateTime.Substring(0, sDateTime.Length - 4) + "00:00";
                            AddGame(DateTime.Parse(sDateTime), sSite, sNameUse, sTeam, sEmail, sGame, sHome, sAway, sLevel, sSport,
                                    sPos, sStatus, fCanceled, plsMisc);
                            }
                        }
                    }
                rs = ReadState.ScanForGame;
                return rs;
            }

            private static bool FEmptyField(string s)
            {
                return String.IsNullOrWhiteSpace(s);
            }
            /* R S  H A N D L E  R E A D I N G  G A M E  2 */
            /*----------------------------------------------------------------------------
			    	%%Function: RsHandleReadingOfficials2
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleReadingOfficials2
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private static ReadState RsHandleReadingOfficials2(Roster rst, string[] rgsFields, Dictionary<string, string> mpNamePos, Dictionary<string, string> mpNameStatus, string sNameLast,
                string sPosLast, string sStatusLast, ReadState rs)
            {
// we are reading the subsequent game lines.  these are not guaranteed to be there (it depends on field
                // overflows
                if (FEmptyField(rgsFields[1]) && FEmptyField(rgsFields[0]) && !FEmptyField(rgsFields[3]))
                    {
                    // nothing in that column means we have a continuation.  now lets concatenate all our stuff
                    mpNamePos.Remove(ReverseName(rst, sNameLast));
                    mpNameStatus.Remove(ReverseName(rst, sNameLast));
                    string sName = $"{sNameLast} {rgsFields[3]}";
                    sName = ReverseName(rst, sName);
                    mpNamePos.Add(sName, sPosLast);
                    mpNameStatus.Add(sName, sStatusLast);
                    return rs;
                    }
                rs = ReadState.ReadingOfficials1;
                // fallthrough to reading officials
                return rs;
            }

            /* R S  H A N D L E  R E A D I N G  G A M E  2 */
            /*----------------------------------------------------------------------------
			    	%%Function: RsHandleReadingOfficials2
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleReadingOfficials2
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private ReadState RsHandleReadingGame2(string[] rgsFields, ref string sGame, ref string sDateTime, ref string sLevel,
                ref string sSite, ref string sHome, ref string sAway, ReadState rs)
            {
// we are reading the subsequent game lines.  these are not guaranteed to be there (it depends on field
                // overflows
                if (FEmptyField(rgsFields[1]) && FEmptyField(rgsFields[GameGameColumn]))
                    {
                    // nothing in that column means we have a continuation.  now lets concatenate all our stuff
                    sGame = AppendCheck(sGame, rgsFields[GameGameColumn]);
                    sDateTime = AppendCheck(sDateTime, rgsFields[GameDateTimeColumn]);
                    sLevel = AppendCheck(sLevel, rgsFields[GameLevelColumn]);
                    sSite = AppendCheck(sSite, rgsFields[GameSiteColumn]);
                    sHome = AppendCheck(sHome, rgsFields[GameHomeColumn]);
                    sAway = AppendCheck(sAway, rgsFields[GameAwayColumn]);
                    return rs;
                    }
                rs = ReadState.ReadingOfficials1;
                // fallthrough to reading officials
                return rs;
            }

            /* F  M A T C H  G A M E  C O M M E N T  C O N T I N U A T I O N */
            /*----------------------------------------------------------------------------
			    	%%Function: FMatchGameCommentContinuation
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameCommentContinuation
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private static bool FMatchGameCommentContinuation(string sLine)
            {
                return Regex.Match(sLine, ",,,,,,,,,,,,,,,,,$").Success
                       && !Regex.Match(sLine, "^\\*\\*\\*").Success;
            }

            /* F  M A T C H  G A M E  T O T A L  L I N E */
            /*----------------------------------------------------------------------------
			    	%%Function: FMatchGameTotalLine
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameTotalLine
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private static bool FMatchGameTotalLine(string[] rgsFields)
            {
                return Regex.Match(rgsFields[13], "Total:").Success
                       || Regex.Match(rgsFields[14], "Total:").Success
                       || Regex.Match(rgsFields[15], "Total:").Success;
            }

            /* F  M A T C H  G A M E  E M P T Y */
            /*----------------------------------------------------------------------------
			    	%%Function: FMatchGameEmpty
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameEmpty
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private static bool FMatchGameEmpty(string sLine)
            {
                return Regex.Match(sLine, "^,,,,,,,,,,,,,,,,,,").Success
                       || Regex.Match(sLine, "^,,,,,,,,,,,,,,,,,").Success;
            }

            /* F  M A T C H  G A M E  C O M M E N T */
            /*----------------------------------------------------------------------------
			    	%%Function: FMatchGameComment
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameComment
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private static bool FMatchGameComment(string sLine)
            {
                return Regex.Match(sLine, "^\"*\\[.*by.*\\]").Success
                       || Regex.Match(sLine, "^[ \t]*\\[.*/.*/.*by.*\\]").Success
                       || Regex.Match(sLine, "^Comments: ").Success;
            }

            /* F  M A T C H  G A M E  C A N C E L L E D */
            /*----------------------------------------------------------------------------
			    	%%Function: FMatchGameCancelled
			    	%%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameCancelled
			    	%%Contact: rlittle
			    	
			    ----------------------------------------------------------------------------*/
            private static bool FMatchGameCancelled(string sLine)
            {
                return Regex.Match(sLine, ".*\\*\\*\\*.*CANCEL*ED").Success
                       || Regex.Match(sLine, ".*\\*\\*\\*.*FORFEITED").Success
                       || Regex.Match(sLine, ".*\\*\\*\\*.*RAINED OUT").Success
                       || Regex.Match(sLine, ".*\\*\\*\\*.*SUSPEND*ED").Success;
            }
        }

        private Roster m_rst;
        private GameSlots m_gms;
        private StatusRpt m_srpt;

        public string SMiscHeader(int i)
        {
            return m_rst.SMiscHeader(i);
        }

        /* L O A D  R O S T E R */
        /*----------------------------------------------------------------------------
				%%Function: LoadRoster
				%%Qualified: GenCount.CountsData:GameData.LoadRoster
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public bool FLoadRoster(string sRoster, int iMiscAffiliation)
        {
            m_rst = new Roster(m_srpt);
            bool f = m_rst.LoadRoster(sRoster, iMiscAffiliation);

            if (f)
                {
                m_gms.SetMiscHeadings(m_rst.PlsMiscHeadings);
                }
            return f;
        }

        /* F  L O A D  G A M E S */
        /*----------------------------------------------------------------------------
				%%Function: FLoadGames
				%%Qualified: ArbWeb.CountsData:GameData.FLoadGames
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public bool FLoadGames(string sGames, bool fIncludeCanceled)
        {
            return m_gms.FLoadGames(sGames, m_rst, fIncludeCanceled);
        }

        /* G E N  R E P O R T */
        /*----------------------------------------------------------------------------
				%%Function: GenReport
				%%Qualified: ArbWeb.CountsData:GameData.GenReport
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public void GenAnalysis(string sReport)
        {
            m_gms.ReduceTeams();
            m_gms.GenReport(sReport);
        }

        public void GenGamesReport(string sReport)
        {
            m_gms.GenGamesReport(sReport);
        }

        /* G E N  O P E N  S L O T S */
        /*----------------------------------------------------------------------------
				%%Function: GenOpenSlots
				%%Qualified: ArbWeb.CountsData:GameData.GenOpenSlots
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public SlotAggr GenOpenSlots(DateTime dttmStart, DateTime dttmEnd)
        {
            return m_gms.GenOpenSlots(dttmStart, dttmEnd);
        }

        public void GenSiteRosterReport(string sReportFile, ArbWeb.Roster rst, string[] rgsRosterFilter, DateTime dttmStart, DateTime dttmEnd)
        {
            m_gms.GenSiteRosterReport(sReportFile, rst, rgsRosterFilter, dttmStart, dttmEnd);
        }
        /* G E N  O P E N  S L O T S */
        /*----------------------------------------------------------------------------
				%%Function: GenOpenSlots
				%%Qualified: ArbWeb.CountsData:GameData.GenOpenSlots
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public void GenOpenSlotsReport(string sReport, bool fDetail, bool fFuzzyTimes, bool fDatePivot, string[] rgsSportFilter, string[] rgsSportLevelFilter, SlotAggr sa)
        {
            m_gms.GenOpenSlotsReport(sReport, fDetail, fFuzzyTimes, fDatePivot, rgsSportFilter, rgsSportLevelFilter, sa);
        }

        /* G A M E S  F R O M  F I L T E R */
        /*----------------------------------------------------------------------------
        	%%Function: GamesFromFilter
        	%%Qualified: ArbWeb.GameData.GamesFromFilter
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public GameSlots GamesFromFilter(string[] rgsSportFilter, string[] rgsSportLevelFilter, bool fOpenOnly, SlotAggr sa)
        {
            return m_gms.GamesFromFilter(rgsSportFilter, rgsSportLevelFilter, fOpenOnly, sa);
        }

        /* G E T  O P E N  S L O T  S P O R T S */
        /*----------------------------------------------------------------------------
				%%Function: GetOpenSlotSports
				%%Qualified: ArbWeb.CountsData:GameData.GetOpenSlotSports
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public string[] GetOpenSlotSports(SlotAggr sa)
        {
            return m_gms.GetOpenSlotSports(sa);
        }

        public string[] GetSiteRosterSites(SlotAggr sa)
        {
            return m_gms.GetSiteRosterSites(sa);
        }

        /* G E T  O P E N  S L O T  S P O R T  L E V E L S */
        /*----------------------------------------------------------------------------
				%%Function: GetOpenSlotSportLevels
				%%Qualified: ArbWeb.CountsData:GameData.GetOpenSlotSportLevels
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public string[] GetOpenSlotSportLevels(SlotAggr sa)
        {
            return m_gms.GetOpenSlotSportLevels(sa);
        }

        /* G E N  G A M E  S T A T S */
        /*----------------------------------------------------------------------------
				%%Function: GameData
				%%Qualified: GenCount.CountsData:GameData.GameData
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public GameData(StatusRpt srpt)
        {
            //  m_sRoster = null;
            m_srpt = srpt;
            m_rst = new Roster(srpt);
            m_gms = new GameSlots(srpt);
        }
    } // END  GameData

}