using System;
using System.Collections.Generic;
using System.Security.Policy;
using System.Security.RightsManagement;

namespace ArbWeb.Games
{
	// SimpleGame + diff type
	public class SimpleDiffGame : SimpleGame
	{
		public enum DiffOp
		{
			None,
			Delete,
			Insert
		}

		public new string SortKey => $"{StartDateTime:u}-{OpToString()}-{Site}-{Sport}-{Level}-{Home}-{Away}-{Status}";

		private DiffOp Op { get; set; }

		/*----------------------------------------------------------------------------
			%%Function: OpToString
			%%Qualified: ArbWeb.Games.SimpleDiffGame.OpToString
		----------------------------------------------------------------------------*/
		public string OpToString()
		{
			return (Op == DiffOp.Delete ? "<" : (Op == DiffOp.Insert ? ">" : ""));
		}
		
		/*----------------------------------------------------------------------------
			%%Function: SimpleDiffGame
			%%Qualified: ArbWeb.Games.SimpleDiffGame.SimpleDiffGame
		----------------------------------------------------------------------------*/
		public SimpleDiffGame(SimpleGame game, DiffOp op)
		{
			Number = game.Number;
			Home = game.Home;
			Away = game.Away;
			StartDateTime = game.StartDateTime;
			Site = game.Site;
			Level = game.Level;
			Sport = game.Sport;
			Status = game.Status;
			Op = op;
		}

		public new string MakeCsvLine(IEnumerable<string> legend)
		{
			Dictionary<string, string> m_mpFieldVal = new Dictionary<string, string>();

			string sDateTime = StartDateTime.ToString("M/d/yyyy H:mm");

			m_mpFieldVal.Add("Diff", OpToString());
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