using System;
using System.Collections.Generic;

namespace ArbWeb.Games
{
    public class DistributeTeamCount // DTC
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

            public int Count
            {
                get { return m_c + m_dc; }
            }

            public string Name
            {
                get { return m_sTeam; }
            }

            public int DCount
            {
                get { return m_dc; }
                set { m_dc = value; }
            }
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

        /* U P D A T E  T E A M  C O U N T */
        /*----------------------------------------------------------------------------
                %%Function: UpdateTeamCount
                %%Qualified: ArbWeb.CountsData:GameData:Games.UpdateTeamCount
                %%Contact: rlittle

        ----------------------------------------------------------------------------*/

        private static void UpdateTeamCount(Dictionary<string, int> mpTeamCount, string sTeam, string sSport, int dCount)
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
    }
}
