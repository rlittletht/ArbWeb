using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ArbWeb
{
	// ================================================================================
	//  U M P I R E 
	// ================================================================================
    public class Umpire // UMP
    {
        private string m_sFirst;
        private string m_sLast;
        private string m_sContact;
        private string m_sMisc;
        private List<string> m_plsMisc;

        /* U M P I R E */
        /*----------------------------------------------------------------------------
			%%Function: Umpire
			%%Qualified: ArbWeb.Umpire.Umpire
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
        public Umpire(string sFirst, string sLast, string sAffiliation, string sContact, List<string> plsMisc)
        {
            m_sFirst = sFirst;
            m_sLast = sLast;
            m_sContact = sContact;
            m_sMisc = Regex.Replace(sAffiliation, "[ ]*20[0-9][0-9]$", "");
            m_plsMisc = plsMisc;
        }

        public string FirstName { get { return m_sFirst; } }
        public string LastName { get { return m_sLast; } }
        public string Contact { get { return m_sContact; } }
        public string Misc { get { return m_sMisc; } }
        public string Name { get { return $"{m_sLast},{m_sFirst}"; } }
        public List<string> PlsMisc { get { return m_plsMisc; } }

    } // END UMPIRE
}