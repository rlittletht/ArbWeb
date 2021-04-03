
using System.Collections.Generic;


namespace ArbWeb
{
    internal class Utils
    {
        /* P L S  U N I Q U E  F R O M  R G S */
        /*----------------------------------------------------------------------------
			%%Function: PlsUniqueFromRgs
			%%Qualified: ArbWeb.CountsData:GameData:Games.PlsUniqueFromRgs

		----------------------------------------------------------------------------*/
        public static SortedList<string, int> PlsUniqueFromRgs(string[] rgs)
        {
            if (rgs == null)
                return null;

            SortedList<string, int> pls = new SortedList<string, int>();
            foreach (string s in rgs)
                {
                if (!pls.ContainsKey(s))
                    pls.Add(s, 0);
                }
            return pls;
        }
    }
}