using System;
using System.Collections.Generic;

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
		
		public DateTime StartDateTime { get; set; }

		public string SortKey => $"{StartDateTime:u}-{Site}-{Sport}-{Level}-{Home}";
		
		public SimpleGame()
		{
		}

		public SimpleGame(GameSlot game)
		{
			StartDateTime = game.Dttm;
			Site = game.Site;
			Level = game.Level;
			Home = game.Home;
			Away = game.Away;
			Sport = game.Sport;
			Number = game.GameNum;
		}

		public SimpleGame(DateTime startDateTime, string site, string level, string home, string away, string number)
		{
			StartDateTime = startDateTime;
			Site = site;
			Level = level;
			Home = home;
			Away = away;
			Number = number;
			Sport = null;
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