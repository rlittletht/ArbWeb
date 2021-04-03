using System.Collections.Generic;
using System.IO;
using System.Text;
using ArbWeb.Games;

namespace ArbWeb.Reports
{
	public class SimpleGameReport
	{
		public static void GenSimpleGamesReport(SimpleSchedule schedule, string sOutputFile)
		{
			using (StreamWriter sw = new StreamWriter(sOutputFile, false, Encoding.Default))
			{
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

				foreach (SimpleGame game in schedule.Games)
				{
					// for each game, report the information, using Legend as the sort order for everything
					sw.WriteLine(game.MakeCsvLine(plsLegend));
				}

				sw.Close();
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
		public static void GenGamesReport(ScheduleGames games, string sOutputFile)
		{
			SimpleSchedule schedule = SimpleSchedule.BuildFromScheduleGames(games);
			GenSimpleGamesReport(schedule, sOutputFile);
		}
    }
}