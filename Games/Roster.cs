using System.Collections.Generic;
using TCore.StatusBox;

namespace ArbWeb.Games
{
    public class Roster // RST
    {
        private ArbWeb.Roster m_rst;
        private Dictionary<string, Umpire> m_mpNameUmpire;
        private StatusBox m_srpt;

        public List<string> PlsMiscHeadings
        {
            get { return m_rst.PlsMisc; }
        }

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

        public Roster(StatusBox srpt)
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
}
