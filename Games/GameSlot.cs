using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ArbWeb.Games
{
    // ================================================================================
    //  G A M E  S L O T
    //
    // The GameSlot knows everything about that slot, including the sport, level,
    // the official assigned to the game, etc
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

        /*----------------------------------------------------------------------------
            %%Function: MakeCsvLine
            %%Qualified: ArbWeb.CountsData:GameData:Game.MakeCsvLine
            %%Contact: rlittle

            Return a detail string suitable for saving in CSV format.

            Takes the given legend and saves out our collected data according to that
            legend.
        ----------------------------------------------------------------------------*/
        public string MakeCsvLine(IEnumerable<string> legend)
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

            foreach (string s in legend)
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
}