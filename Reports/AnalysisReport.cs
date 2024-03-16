using System.IO;
using System.Text;
using ArbWeb.Games;

namespace ArbWeb.Reports
{
    public class AnalysisReport
    {
        /* G E N  R E P O R T */
        /*----------------------------------------------------------------------------
            %%Function: GenReport
            %%Qualified: ArbWeb.CountsData:GameData:Games.GenReport
            %%Contact: rlittle

            Take the accumulated game data and generate an analysis report
        ----------------------------------------------------------------------------*/
        public static void GenReport(ScheduleGames games, string sReport)
        {
            StreamWriter sw = new StreamWriter(sReport, false, Encoding.Default);
            bool fFirst = true;

            games.ReduceTeams();
            games.AddStaticLegendColumns();

            foreach (string s in games.AnalysisLegend)
            {
                if (s == "$$$MISC$$$")
                {
                    foreach (string sMisc in games.MiscHeadings)
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

            foreach (GameSlot gm in games.SortedGameSlots)
            {
                //						if (gm.Open)
                //							continue;

                // for each game, report the information, using Legend as the sort order for everything
                sw.WriteLine(gm.MakeCsvLine(games.AnalysisLegend));
            }

            sw.Close();
        }
    }
}
