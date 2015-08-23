using System;
using System.Collections.Generic;
using System.IO;

namespace ArbWeb
{
    // ================================================================================
    // S L O T  C O U N T 
    // ================================================================================
    internal class SlotCount // SC
    {
        private DateTime m_dttmSlot;
        private Dictionary<string, int> m_mpSportCount;
        private Dictionary<string, int> m_mpSportLevelCount;
        private Dictionary<string, int> m_mpSiteCount;

        /* S L O T  C O U N T */
        /*----------------------------------------------------------------------------
        	%%Function: SlotCount
        	%%Qualified: ArbWeb.SlotCount.SlotCount
        	%%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public SlotCount(GameData.GameSlot gm)
        {
            m_dttmSlot = gm.Dttm;
            m_mpSportCount = new Dictionary<string, int>();
            m_mpSportCount.Add(gm.Sport, 1);
            m_mpSportLevelCount = new Dictionary<string, int>();
            m_mpSportLevelCount.Add(gm.SportLevel, 1);
            m_mpSiteCount = new Dictionary<string, int>();
            m_mpSiteCount.Add(gm.SiteShort, 1);
        }

        public SlotCount()
        {
        }

        /* A D D  S L O T */
        /*----------------------------------------------------------------------------
        	%%Function: AddSlot
        	%%Qualified: ArbWeb.SlotCount.AddSlot
        	%%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public void AddSlot(GameData.GameSlot gm)
        {
            if (m_mpSportCount.ContainsKey(gm.Sport))
                m_mpSportCount[gm.Sport]++;
            else
                m_mpSportCount.Add(gm.Sport, 1);

            if (m_mpSportLevelCount.ContainsKey(gm.SportLevel))
                m_mpSportLevelCount[gm.SportLevel]++;
            else
                m_mpSportLevelCount.Add(gm.SportLevel, 1);

            string sSite = gm.SiteShort;

            if (m_mpSiteCount.ContainsKey(sSite))
                m_mpSiteCount[sSite]++;
            else
                m_mpSiteCount.Add(sSite, 1);
        }

        public string[] Sports
        {
            get
            {
                string[] rgs = new string[m_mpSportCount.Count];
                m_mpSportCount.Keys.CopyTo(rgs, 0);
                return rgs;
            }
        }

        public string[] SportLevels
        {
            get
            {
                string[] rgs = new string[m_mpSportLevelCount.Count];
                m_mpSportLevelCount.Keys.CopyTo(rgs, 0);
                return rgs;
            }
        }

        public string[] Sites
        {
            get
            {
                string[] rgs = new string[m_mpSiteCount.Count];
                m_mpSiteCount.Keys.CopyTo(rgs, 0);
                return rgs;
            }
        }

        public DateTime Dttm { get { return m_dttmSlot; } }

        /* M E R G E */
        /*----------------------------------------------------------------------------
        	%%Function: Merge
        	%%Qualified: ArbWeb.SlotCount.Merge
        	%%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public SlotCount Merge(SlotCount sc)
        {
            SlotCount scNew = new SlotCount();

            scNew.m_dttmSlot = m_dttmSlot;
            scNew.m_mpSportCount = new Dictionary<string, int>();
            scNew.m_mpSportLevelCount = new Dictionary<string, int>();
            scNew.m_mpSiteCount = new Dictionary<string, int>();

            foreach (string sSport in m_mpSportCount.Keys)
                {
                int c = m_mpSportCount[sSport];

                if (sc.m_mpSportCount.ContainsKey(sSport))
                    c += sc.m_mpSportCount[sSport];
                scNew.m_mpSportCount.Add(sSport, c);
                }

            foreach (string sSport in sc.m_mpSportCount.Keys)
                {
                int c = sc.m_mpSportCount[sSport];

                if (!m_mpSportCount.ContainsKey(sSport))
                    scNew.m_mpSportCount.Add(sSport, c);
                }

            foreach (string sSport in m_mpSportLevelCount.Keys)
                {
                int c = m_mpSportLevelCount[sSport];

                if (sc.m_mpSportLevelCount.ContainsKey(sSport))
                    c += sc.m_mpSportLevelCount[sSport];
                scNew.m_mpSportLevelCount.Add(sSport, c);
                }

            foreach (string sSport in sc.m_mpSportLevelCount.Keys)
                {
                int c = sc.m_mpSportLevelCount[sSport];

                if (!m_mpSportLevelCount.ContainsKey(sSport))
                    scNew.m_mpSportLevelCount.Add(sSport, c);
                }

            foreach (string sSport in m_mpSiteCount.Keys)
                {
                int c = m_mpSiteCount[sSport];

                if (sc.m_mpSiteCount.ContainsKey(sSport))
                    c += sc.m_mpSiteCount[sSport];
                scNew.m_mpSiteCount.Add(sSport, c);
                }

            foreach (string sSport in sc.m_mpSiteCount.Keys)
                {
                int c = sc.m_mpSiteCount[sSport];

                if (!m_mpSiteCount.ContainsKey(sSport))
                    scNew.m_mpSiteCount.Add(sSport, c);
                }

            return scNew;
        }


        public bool FMatchFuzzyTime(SlotCount sc)
        {
            if (sc.Dttm.Date == m_dttmSlot.Date)
                {
                if (sc.Dttm.Hour < 12 && m_dttmSlot.Hour < 12)
                    return true;

                if ((sc.Dttm.Hour >= 12 && sc.Dttm.Hour < 16)
                    && (m_dttmSlot.Hour >= 12 && m_dttmSlot.Hour < 16))
                    return true;

                if (sc.Dttm.Hour >= 16 && m_dttmSlot.Hour >= 16)
                    return true;
                }
            return false;
        }


        public int OpenCount(string sSport)
        {
            if (m_mpSiteCount.ContainsKey(sSport))
                return m_mpSiteCount[sSport];

            if (!m_mpSportCount.ContainsKey(sSport)
                && !m_mpSportLevelCount.ContainsKey(sSport))
                {
                return 0;
                }

            if (m_mpSportLevelCount.ContainsKey(sSport))
                return m_mpSportLevelCount[sSport];

            return m_mpSportCount[sSport];
        }

    }

    // ================================================================================
    // S L O T  A G G R 
    // ================================================================================
    public class SlotAggr
    {
        private SortedList<DateTime, SlotCount> m_mpSlotSc;
        private SortedList<string, GameData.GameSlot> m_plgm;
        private List<string> m_plsSportLevels;
        private List<string> m_plsSports;
        private List<string> m_plsSites;

        private DateTime m_dttmStart;
        private DateTime m_dttmEnd;


        public DateTime DttmStart { get { return m_dttmStart; } }
        public DateTime DttmEnd { get { return m_dttmEnd; } }

        /* S L O T  A G G R */
        /*----------------------------------------------------------------------------
        	%%Function: SlotAggr
        	%%Qualified: ArbWeb.SlotAggr.SlotAggr
        	%%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public SlotAggr()
        {
        }

        public string[] Sports
        {
            get
            {
                string[] rgs = new string[m_plsSports.Count];
                m_plsSports.CopyTo(rgs, 0);
                return rgs;
            }
        }

        public string[] SportLevels
        {
            get
            {
                string[] rgs = new string[m_plsSportLevels.Count];
                m_plsSportLevels.CopyTo(rgs, 0);
                return rgs;
            }
        }

        public string[] Sites
        {
            get
            {
                string[] rgs = new string[m_plsSites.Count];
                m_plsSites.CopyTo(rgs, 0);
                return rgs;
            }
        }

        /* G E N */
        /*----------------------------------------------------------------------------
        	%%Function: Gen
        	%%Qualified: ArbWeb.SlotAggr.Gen
        	%%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public static SlotAggr Gen(SortedList<string, GameData.GameSlot> plgm, DateTime dttmStart, DateTime dttmEnd, string[] rgsSportFilter, string[] rgsSportLevelFilter, bool fOnlyOpen)
        {
            SlotAggr os = new SlotAggr();

            os.m_plgm = plgm;
            os.m_dttmStart = dttmStart;
            os.m_dttmEnd = dttmEnd;
            os.m_mpSlotSc = new SortedList<DateTime, SlotCount>();
            os.m_plsSports = new List<string>();
            os.m_plsSportLevels = new List<string>();
            os.m_plsSites = new List<string>();

            foreach (GameData.GameSlot gm in plgm.Values)
                {
                if (!gm.Open && fOnlyOpen)
                    continue;

                if (DateTime.Compare(gm.Dttm, dttmStart) < 0 || DateTime.Compare(gm.Dttm, dttmEnd) > 0)
                    continue;

                if (rgsSportFilter != null)
                    {
                    bool fMatch = false;

                    foreach (string s in rgsSportFilter)
                        if (String.Compare(gm.Sport, s, true) == 0)
                            fMatch = true;

                    if (fMatch == false)
                        continue;
                    }


                if (rgsSportLevelFilter != null)
                    {
                    bool fMatch = false;

                    foreach (string s in rgsSportLevelFilter)
                        if (String.Compare(gm.SportLevel, s, true) == 0)
                            fMatch = true;

                    if (fMatch == false)
                        continue;
                    }

                if (os.m_mpSlotSc.ContainsKey(gm.Dttm))
                    {
                    os.m_mpSlotSc[gm.Dttm].AddSlot(gm);
                    }
                else
                    {
                    os.m_mpSlotSc.Add(gm.Dttm, new SlotCount(gm));
                    }
                }

            foreach (SlotCount sc in os.m_mpSlotSc.Values)
                {
                foreach (string sSport in sc.Sports)
                    {
                    if (!os.m_plsSports.Contains(sSport))
                        os.m_plsSports.Add(sSport);
                    }
                foreach (string sSportLevel in sc.SportLevels)
                    {
                    if (!os.m_plsSportLevels.Contains(sSportLevel))
                        os.m_plsSportLevels.Add(sSportLevel);
                    }
                foreach (string sSite in sc.Sites)
                    {
                    if (!os.m_plsSites.Contains(sSite))
                        os.m_plsSites.Add(sSite);
                    }
                }
            return os;
        }

        /* G E N  R E P O R T */
        /*----------------------------------------------------------------------------
        	%%Function: GenReport
        	%%Qualified: ArbWeb.SlotAggr.GenReport
        	%%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public void GenReport(string sReport, string[] rgsSportFilter, string[] rgsSportLevelFilter)
        {
            StreamWriter sw = new StreamWriter(sReport, false, System.Text.Encoding.Default);
            List<string> plsUse;
            List<string> plsCategory;

            SortedList<string, int> mpFilter = null;

            if (rgsSportFilter != null)
                {
                if (rgsSportLevelFilter != null)
                    {
                    plsUse = m_plsSportLevels;
                    plsCategory = m_plsSports;

                    mpFilter = Utils.PlsUniqueFromRgs(rgsSportLevelFilter);
                    }
                else
                    {
                    plsUse = m_plsSports;
                    plsCategory = null;
                    mpFilter = Utils.PlsUniqueFromRgs(rgsSportFilter);
                    }
                }
            else
                {
                plsUse = m_plsSportLevels;
                plsCategory = m_plsSports;
                }

            string sFormat = "<tr><td>{0}<td>{1}"; //<td>{2}</tr>

            Dictionary<string, int> mpSportCount = new Dictionary<string, int>();

            int cCols = 0;
            foreach (string sSport in plsUse)
                {
                if (mpFilter != null && !mpFilter.ContainsKey(sSport))
                    continue;

                if (plsCategory != null)
                    {
                    string sCat = null;

                    // let's figure out which category we belong in
                    foreach (string sCatT in plsCategory)
                        {
                        if (sSport.StartsWith(sCatT))
                            {
                            sCat = sCatT;
                            break;
                            }
                        }
                    if (mpSportCount.ContainsKey(sCat))
                        mpSportCount[sCat]++;
                    else
                        mpSportCount.Add(sCat, 1);
                    }
                cCols++;
                }
            // write the legend
            sw.WriteLine("<html><body><table>");

            if (plsCategory != null)
                {
                sw.Write(String.Format(sFormat, "", ""));

                foreach (string sCatT in plsCategory)
                    {
                    if (mpSportCount.ContainsKey(sCatT))
                        {
                        sw.WriteLine("<td colspan={0}>{1}", mpSportCount[sCatT], sCatT);
                        }
                    }
                sw.Write("<td></tr>");
                }

            sw.Write(String.Format(sFormat, "Date", "Slot"));

            // at this point, make sure we have a plsCategory with the categories or "" as a single (match all) category
            if (plsCategory == null)
                {
                plsCategory = new List<string>();
                plsCategory.Add("");
                }

            foreach (string sCatT in plsCategory)
                {
                // match all the items we have that begin with sCatT
                foreach (string sSport in plsUse)
                    {
                    if (mpFilter != null && !mpFilter.ContainsKey(sSport))
                        continue;

                    if (sSport.StartsWith(sCatT))
                        {
                        string s = sSport.Remove(0, sCatT.Length + (sCatT == "" ? 0 : 1));
                        sw.Write("<td>{0}", s);
                        }
                    }
                }

            sw.WriteLine("<td>Total</tr>");

            // now, put out the actual values...

            foreach (SlotCount sc in m_mpSlotSc.Values)
                {
                int cTotal = 0;
                string sWrite = String.Format(sFormat, sc.Dttm.ToString("MM/dd/yy ddd"), sc.Dttm.ToString("hh:mm tt"));

                foreach (string sCatT in plsCategory)
                    {
                    // match all the items we have that begin with sCatT
                    foreach (string sSport in plsUse)
                        {
                        if (mpFilter != null && !mpFilter.ContainsKey(sSport))
                            continue;

                        if (sSport.StartsWith(sCatT))
                            {
                            int c = sc.OpenCount(sSport);

                            if (c > 0)
                                sWrite += String.Format("<td>{0} umpires", sc.OpenCount(sSport));
                            else
                                sWrite += "<td>";

                            cTotal += sc.OpenCount(sSport);
                            }
                        }
                    }

                if (cTotal > 0)
                    {
                    sw.Write(sWrite);
                    sw.WriteLine("<td>{0} umpires</tr>", cTotal);
                    }
                }
            sw.Close();
        }

        // ================================================================================
        // H T M L  T A B L E  R E P O R T 
        // ================================================================================
        public class HtmlTableReport
        {
            private List<string> m_rgsCols = new List<string>();
            private SortedList<string, string> m_rgsAxis1 = new SortedList<string, string>();
            private SortedList<string, string> m_rgsAxis2 = new SortedList<string, string>();
            private Dictionary<string, string> m_mpAxisValues = new Dictionary<string, string>();
            private string m_sAxis1Title;
            private string m_sAxis2Title;

            public HtmlTableReport()
            {
            }

            public string Axis1Title { get { return m_sAxis1Title; } set { m_sAxis1Title = value; } }
            public string Axis2Title { get { return m_sAxis2Title; } set { m_sAxis2Title = value; } }

            /* A D D  V A L U E */
            /*----------------------------------------------------------------------------
            	%%Function: AddValue
            	%%Qualified: ArbWeb.SlotAggr:HtmlTableReport.AddValue
            	%%Contact: rlittle
            	
            ----------------------------------------------------------------------------*/
            public void AddValue(string sAxis1, string sAxis1Key, string sAxis2, string sAxis2Key, string sValue)
            {
                if (!m_rgsAxis1.ContainsKey(sAxis1Key))
                    m_rgsAxis1.Add(sAxis1Key, sAxis1);

                if (!m_rgsAxis2.ContainsKey(sAxis2Key))
                    m_rgsAxis2.Add(sAxis2Key, sAxis2);

                string sKey = String.Format("({0},{1})", sAxis1, sAxis2);

                if (m_mpAxisValues.ContainsKey(sKey))
                    throw new Exception("duplicate axis value mapping " + sKey);

                m_mpAxisValues.Add(sKey, sValue);
            }

            /* G E N  R E P O R T */
            /*----------------------------------------------------------------------------
            	%%Function: GenReport
            	%%Qualified: ArbWeb.SlotAggr:HtmlTableReport.GenReport
            	%%Contact: rlittle
            	
            ----------------------------------------------------------------------------*/
            public void GenReport(StreamWriter sw, bool fAxis1Pivot)
            {
                string sFirstRowStyle = "";

                sw.WriteLine("<html><body><table style='mso-yfti-tbllook: 1504;border-collapse: collapse;padding-left: 4pt;padding-right: 4pt' cellspacing=0><thead><tr " + sFirstRowStyle + ">");

                SortedList<string, string> rgsCols = fAxis1Pivot ? m_rgsAxis1 : m_rgsAxis2;
                SortedList<string, string> rgsRows = fAxis1Pivot ? m_rgsAxis2 : m_rgsAxis1;

                // the axis we pivot on is the axis that gives column titles
                int[] rgcColspan = new int[rgsCols.Keys.Count];
                string[] rgsFirstLine = new string[rgsCols.Keys.Count];
                bool fNeed2Lines = false;

                for (int i = 0, iMac = rgsCols.Keys.Count; i < iMac; i++)
                    {
                    string s = rgsCols[rgsCols.Keys[i]];
                    int cchFirst = s.IndexOf('\r');

                    rgcColspan[i] = 1;
                    if (cchFirst >= 0)
                        rgsFirstLine[i] = s.Substring(0, cchFirst);
                    else
                        rgsFirstLine[i] = "";

                    if (cchFirst > 0)
                        {
                        fNeed2Lines = true;
                        int iT = i + 1;

                        while (iT < iMac)
                            {
                            string sT = rgsCols[rgsCols.Keys[iT]];
                            if (cchFirst > sT.Length)
                                break;
                            if (String.Compare(s.Substring(0, cchFirst), sT.Substring(0, cchFirst)) != 0)
                                break;

                            rgcColspan[i]++;
                            rgsFirstLine[iT] = rgsFirstLine[i];
                            rgcColspan[iT] = 0;
                            iT++;
                            }
                        i = iT - 1;
                        }
                    }

                string sStyle = String.Format("style='{0}'", sFirstRowStyle);
                if (fNeed2Lines)
                    {
                    string sRow1Style = "border-bottom: none;";
                    string sRow2Style = "border-top: none;";

                    sStyle = String.Format("style='{0}{1}'", sFirstRowStyle, sRow1Style);

                    sw.WriteLine(String.Format("<td align=center valign=middle rowspan=2 " + sStyle + ">{0}", !fAxis1Pivot ? m_sAxis1Title : m_sAxis2Title));

                    for (int i = 0, iMac = rgsCols.Keys.Count; i < iMac; i++)
                        {
                        if (rgcColspan[i] == 0)
                            continue;

                        sw.WriteLine(String.Format("<td align=center valign=middle colspan={0} " + sStyle + ">{1}", rgcColspan[i], rgsFirstLine[i]));
                        }

                    sStyle = String.Format("style='{0}{1}'", sFirstRowStyle, sRow2Style);
                    sw.WriteLine("<tr " + sStyle + ">");
                    for (int i = 0, iMac = rgsCols.Keys.Count; i < iMac; i++)
                        {
                        string s = rgsCols[rgsCols.Keys[i]].Substring(rgsFirstLine[i].Length + 1);
                        sw.WriteLine(String.Format("<td align=center valign=middle " + sStyle + ">{0}", s));
                        }
                    }
                else
                    {
                    sw.WriteLine(String.Format("<td align=center valign=middle " + sStyle + ">{0}", !fAxis1Pivot ? m_sAxis1Title : m_sAxis2Title));
                    foreach (string sT in rgsCols.Keys)
                        {
                        string s = rgsCols[sT];
                        sw.WriteLine(String.Format("<td align=center valign=middle " + sStyle + ">{0}", s));
                        }
                    }

                sw.WriteLine("</thead>");
                // now, each row is from the other axis
                foreach (string sRowT in rgsRows.Keys)
                    {
                    string sRow = rgsRows[sRowT];
                    sw.WriteLine(String.Format("<tr><td>{0}", sRow));

                    foreach (string sColT in rgsCols.Keys)
                        {
                        string sCol = rgsCols[sColT];
                        string sKey;

                        if (fAxis1Pivot)
                            sKey = String.Format("({0},{1})", sCol, sRow);
                        else
                            sKey = String.Format("({0},{1})", sRow, sCol);

                        if (m_mpAxisValues.ContainsKey(sKey))
                            sw.WriteLine(String.Format("<td>{0}", m_mpAxisValues[sKey]));
                        else
                            sw.WriteLine("<td>");
                        }
                    }
                sw.WriteLine("</table>");
            }


        }

        /* G E N  R E P O R T  B Y  S I T E */
        /*----------------------------------------------------------------------------
        	%%Function: GenReportBySite
        	%%Qualified: ArbWeb.SlotAggr.GenReportBySite
        	%%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public void GenReportBySite(string sReport, bool fFuzzyTimes, bool fDatePivot, string[] rgsSportFilter, string[] rgsSportLevelFilter)
        {
            SlotAggr os = SlotAggr.Gen(m_plgm, m_dttmStart, m_dttmEnd, rgsSportFilter, rgsSportLevelFilter, true);

            StreamWriter sw = new StreamWriter(sReport, false, System.Text.Encoding.Default);
            List<string> plsUse;
            SortedList<string, int> mpFilter = null;
            HtmlTableReport htr = new HtmlTableReport();
            Dictionary<string, int> mpSiteCount = new Dictionary<string, int>();
            int cTotalTotal = 0;

            htr.Axis1Title = "Date";
            htr.Axis2Title = "Site";

            plsUse = new List<string>();

            foreach (string sSite in os.m_plsSites)
                {
                if (mpFilter != null && !mpFilter.ContainsKey(sSite))
                    continue;

                plsUse.Add(sSite);
                }

            // now, put out the actual values...

            for (int isc = 0, iscMac = os.m_mpSlotSc.Keys.Count; isc < iscMac; isc++)
                {
                SlotCount sc = os.m_mpSlotSc[os.m_mpSlotSc.Keys[isc]];

                // merge adject SlotCounts for fuzzytime
                if (fFuzzyTimes)
                    {
                    while (isc + 1 < iscMac)
                        {
                        SlotCount scNext = os.m_mpSlotSc[os.m_mpSlotSc.Keys[isc + 1]];

                        if (!scNext.FMatchFuzzyTime(sc))
                            break;

                        sc = sc.Merge(scNext);
                        isc++;
                        }
                    }

                int cTotal = 0;

                string sWrite;
                string sKey;
                string sFormat = "{0}\r{1}";


                if (fFuzzyTimes)
                    {
                    string sTime;
                    string sTimeKey;

                    if (sc.Dttm.Hour < 12)
                        {
                        sTime = "Morning";
                        sTimeKey = "0";
                        }
                    else if (sc.Dttm.Hour >= 16)
                        {
                        sTime = "Evening";
                        sTimeKey = "2";
                        }
                    else
                        {
                        sTime = "Afternoon";
                        sTimeKey = "1";
                        }

                    sWrite = String.Format(sFormat, sc.Dttm.ToString("ddd M/d"), sTime);
                    sKey = String.Format(sFormat, sc.Dttm.ToString("MN/dd ddd"), sTimeKey);
                    }
                else
                    {
                    sWrite = String.Format(sFormat, sc.Dttm.ToString("M/d ddd"), sc.Dttm.ToString("h:mm tt"));
                    sKey = String.Format(sFormat, sc.Dttm.ToString("MM/dd ddd"), sc.Dttm.ToString("HH:mm"));
                    }

                // match all the items we have that begin with sCatT
                foreach (string sSite in plsUse)
                    {
                    int c = sc.OpenCount(sSite);

                    if (!mpSiteCount.ContainsKey(sSite))
                        mpSiteCount.Add(sSite, 0);

                    mpSiteCount[sSite] += c;

                    if (c > 0)
                        htr.AddValue(sWrite, sKey, sSite, sSite, String.Format("{0} umps", c));

                    cTotal += sc.OpenCount(sSite);
                    }

                if (cTotal > 0)
                    {
                    htr.AddValue(sWrite, sKey, "Total", "zzzTotal", String.Format("{0} umps", cTotal));
                    cTotalTotal += cTotal;
                    }
                }
            foreach (string sSite in mpSiteCount.Keys)
                {
                htr.AddValue("\rTotal", "zzz\rzzzTotal", sSite, sSite, String.Format("{0} umps", mpSiteCount[sSite]));
                }
            htr.AddValue("\rTotal", "zzz\rzzzTotal", "Total", "zzzTotal", String.Format("{0} umps", cTotalTotal));
            htr.GenReport(sw, fDatePivot);
            sw.Close();
        }
    }
}