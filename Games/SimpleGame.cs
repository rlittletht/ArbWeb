using System;
using System.Collections.Generic;
using System.Net.NetworkInformation;

namespace ArbWeb.Games
{
    // these are just simple games - no slot information
    public class SimpleGame
    {
        public string Site { get; set; }
        public string Level { get; set; }
        public string Home { get; set; }
        public string Away { get; set; }
        public string Sport { get; set; }
        public string Number { get; set; }
        public string Status { get; set; }

        public DateTime StartDateTime { get; set; }

        public override int GetHashCode()
        {
            return SortKey.GetHashCode();
        }

        public string SortKey => $"{StartDateTime:u}-{Site?.ToUpper()}-{Sport?.ToUpper()}-{Level?.ToUpper()}-{Home?.ToUpper()}-{Status?.ToUpper()}";

        public SimpleGame()
        {
        }

        /*----------------------------------------------------------------------------
            %%Function: SimpleGame
            %%Qualified: ArbWeb.Games.SimpleGame.SimpleGame
        ----------------------------------------------------------------------------*/
        public SimpleGame(GameSlot game)
        {
            StartDateTime = game.Dttm;
            Site = game.Site;
            Level = game.Level;
            Home = game.Home;
            Away = game.Away;
            Sport = game.Sport;
            Number = game.GameNum;
            Status = game.Cancelled ? "Cancelled" : "";
        }

        /*----------------------------------------------------------------------------
            %%Function: SimpleGame
            %%Qualified: ArbWeb.Games.SimpleGame.SimpleGame
        ----------------------------------------------------------------------------*/
        public SimpleGame(DateTime startDateTime, string site, string level, string home, string away, string number, string status, string sport)
        {
            StartDateTime = startDateTime;
            Site = site;
            Level = level;
            Home = home;
            Away = away;
            Number = number;
            Sport = sport;
            Status = status;
        }

        /*----------------------------------------------------------------------------
            %%Function: IsEqual
            %%Qualified: ArbWeb.Games.SimpleGame.IsEqual
        ----------------------------------------------------------------------------*/
        public static bool AreEqual(SimpleGame left, SimpleGame right, ScheduleMaps maps)
        {
            if (left.StartDateTime != right.StartDateTime)
                return false;

            if (!(left.Level.ToUpper().Contains("INTERMEDIATE") && right.Level.ToUpper().Contains("INTERMEDIATE")))
            {
                if (String.Compare(left.Level, right.Level, true) != 0)
                    return false;
            }

            if (FuzzyMatcher.IsGameFuzzySportMatch(left, right) == 0)
                return false;

            string home = left.Home.ToUpper();
            string away = left.Away.ToUpper();
            string site = left.Site.ToUpper();

            if (maps != null)
            {
                home = maps.TeamsMap.ContainsKey(home) ? maps.TeamsMap[home] : left.Home;

                away = maps.TeamsMap.ContainsKey(away) ? maps.TeamsMap[away] : left.Away;

                site = maps.SitesMap.ContainsKey(site) ? maps.SitesMap[site] : left.Site;
            }

            if (String.Compare(home, right.Home, true) != 0)
                return false;

            if (String.Compare(away, right.Away, true) != 0)
                return false;

            if (String.Compare(site, right.Site, true) != 0)
                return false;

            if (String.Compare(left.Status, right.Status, true) != 0)
            {
                if (String.IsNullOrEmpty(right.Status))
                {
                    if (!left.Status.ToUpper().Contains("RESCHED"))
                        return false;
                }
            }


            return true;
        }


        /*----------------------------------------------------------------------------
            %%Function: IsEqual
            %%Qualified: ArbWeb.Games.SimpleGame.IsEqual
        ----------------------------------------------------------------------------*/
        public bool IsEqual(SimpleGame right, ScheduleMaps maps)
        {
            return SimpleGame.AreEqual(this, right, maps);
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

            string sDateTime = StartDateTime.ToString("M/d/yyyy H:mm");

            m_mpFieldVal.Add("DateTime", sDateTime);
            m_mpFieldVal.Add("Date", StartDateTime.ToString("M/d/yyyy"));
            m_mpFieldVal.Add("Time", StartDateTime.ToString("H:mm"));
            m_mpFieldVal.Add("Level", Level);
            m_mpFieldVal.Add("Home", Home);
            m_mpFieldVal.Add("Away", Away);
            m_mpFieldVal.Add("Site", Site);
            m_mpFieldVal.Add("Sport", Sport);
            m_mpFieldVal.Add("Game", Number);
            m_mpFieldVal.Add("Status", Status);

            // now that we have a dictionary of values, write it out
            bool fFirst = true;
            string sRet = "";

            foreach (string s in legend)
            {
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
