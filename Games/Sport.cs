using System.Collections.Generic;

namespace ArbWeb.Games
{
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
}
