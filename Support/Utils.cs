﻿using System;
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

        public static void ThrowIfNot(bool f, string s)
        {
            if (!f)
                throw new Exception(s);
        }

        public static void ThrowIfNot(bool f)
        {
            if (!f)
                throw new Exception("Unknown failure");
        }
    }
}
