using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TCore.CSV;
using TCore.StatusBox;
using TCore.UI;

namespace ArbWeb.Games
{
    public class ScheduleStats
    {
        private StatusBox m_statusBox;

        public ScheduleStats(StatusBox srpt)
        {
            m_statusBox = srpt;
        }

        public void GenerateScheduleStats(Form parentForm)
        {
//            if (!InputBox.ShowBrowseBox("Roster", "", out string roster, "Rosters|*roster*.csv", 480, true, parentForm))
//                return;

            if (!InputBox.ShowBrowseBox("Games CSV file", "", out string gamesFile, "Games Report Files|*games*.csv", 480, true, parentForm))
                return;

            if (!InputBox.ShowBrowseBox("Output text file", "", out string outfile, "Text report|*.csv", 480, false, parentForm))
                return;

            Schedule sched = new Schedule(m_statusBox);
//            m_statusBox.AddMessage("Loading roster...", MSGT.Header, false);

//            sched.FLoadRoster(roster, 0);

//            m_statusBox.PopLevel();
            m_statusBox.AddMessage("Loading games...", MSGT.Header, false);
            sched.FLoadGames(gamesFile, true);
            m_statusBox.PopLevel();

            // at this point we have the schedule...

            // we want to build a map of Sport+Level to Site Short Names
            ScheduleGames games = sched.Games;
            Dictionary<string, Dictionary<string, int>> mpSportLevelsSites = new Dictionary<string, Dictionary<string, int>>();

            foreach (Game game in games.Games.Values)
            {
                if (game.Slots.Count == 0)
                    continue;

                string sportLevel = $"{game.Slots[0].SportLevel}";
                string siteShort = $"{game.Slots[0].SiteShort}";

                if (!mpSportLevelsSites.ContainsKey(sportLevel))
                {
                    mpSportLevelsSites.Add(sportLevel, new Dictionary<string, int>());
                }

                if (!mpSportLevelsSites[sportLevel].ContainsKey(siteShort))
                    mpSportLevelsSites[sportLevel].Add(siteShort, 0);

                mpSportLevelsSites[sportLevel][siteShort]++;
            }

            using (StreamWriter sw = new StreamWriter(outfile, false, Encoding.Default))
            {
                sw.WriteLine(Csv.CsvFromRgs(new string[] { "SportLevel", "SiteShort", "Count" }));

                foreach (string sportLevel in mpSportLevelsSites.Keys)
                {
                    foreach (string siteShort in mpSportLevelsSites[sportLevel].Keys)
                    {
                        sw.WriteLine(Csv.CsvFromRgs(new string[] { sportLevel, siteShort, mpSportLevelsSites[sportLevel][siteShort].ToString() }));
                    }
                }

                sw.Flush();
                sw.Close();
            }
        }
    }
}
