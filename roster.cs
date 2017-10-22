using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using NUnit.Framework;

namespace ArbWeb
{

    // ================================================================================
    //  R O S T E R     E N T R Y 
    // ================================================================================
    public class RosterEntry // RSTE
    {
        public string m_sEmail;
        public string m_sFirst;
        public string m_sLast;
        public string m_sAddress1;
        public string m_sAddress2;
        public string m_sCity;
        public string m_sState;
        public string m_sZip;
        public string m_sOfficialNumber;
        public string m_sDateOfBirth;
        public string m_sDateJoined;
        public string m_sLastSignin;
        public string m_sGamesPerDay;
        public string m_sGamesPerWeek;
        public string m_sTotalGames;
        public string m_sWaitMinutes;
        public string m_sPhone1;
        public string m_sPhone2;
        public string m_sPhone3;

        public bool IsUploadableQuickroster {  get { return m_rstt == Roster.RSTT.QuickFull2 || m_rstt == Roster.RSTT.QuickFull; } }

        public int m_cRankings;
        public Dictionary<string, int> m_mpRanking;

        public bool m_fReady;
        public bool m_fActive;

        public List<string> m_plsMisc;

        public bool m_fMarked;

        public string Phone1 { get { return m_sPhone1; } }

        public string CellPhone
        {
            get
            {
                if (m_sPhone1 != null && m_sPhone1.Contains("C:"))
                    return m_sPhone1;
                else if (m_sPhone2 != null && m_sPhone2.Contains("C:"))
                    return m_sPhone2;
                else if (m_sPhone3 != null && m_sPhone3.Contains("C:"))
                    return m_sPhone3;
                else
                    return m_sPhone1;
            }
        }

        public bool FHasPhoneNumber(int iPhone)
        {
            return !String.IsNullOrEmpty(NumberRawForPhone(iPhone));
        }

        static void ExtractNumberParts(string sNumberRaw, out string sNumber, out string sType)
        {
            sNumber = null;
            sType = null;

            if (String.IsNullOrEmpty(sNumberRaw))
                {
                return;
                }

            if (sNumberRaw.StartsWith("C:"))
                {
                sType = "Cellular";
                }
            else if (sNumberRaw.StartsWith("W:"))
                {
                sType = "Work";
                }
            else if (sNumberRaw.StartsWith("H"))
                {
                sType = "Home";
                }
            else
                {
                throw new Exception("Unknown phone type");
                }

            sNumber = sNumberRaw.Substring(2);
        }

        static string SConstructPhoneNumberFromParts(string sPhoneNumber, string sTypeDefault)
        {
            if (String.IsNullOrEmpty(sPhoneNumber))
                return sPhoneNumber;

            if (sPhoneNumber.StartsWith("C:") || sPhoneNumber.StartsWith("W:") || sPhoneNumber.StartsWith("H:"))
                return sPhoneNumber;

            return String.Format("{0}:{1}", sTypeDefault.Substring(0, 1), sPhoneNumber);
        }


        [TestCase("C:425-555-1212", "425-555-1212", "Cellular")]
        [TestCase("H:425-555-1212", "425-555-1212", "Home")]
        [TestCase("W:425-555-1212", "425-555-1212", "Work")]
        [TestCase(null, null, null)]
        [Test]
        public static void TestExtractNumberParts(string sNumberRaw, string sNumberExpected, string sTypeExpected)
        {
            string sNumberActual, sTypeActual;

            ExtractNumberParts(sNumberRaw, out sNumberActual, out sTypeActual);
            Assert.AreEqual(sNumberExpected, sNumberActual);
            Assert.AreEqual(sTypeExpected, sTypeActual);
        }
        string NumberRawForPhone(int iPhone)
        {
            string sNumberRaw = null;

            switch (iPhone)
                {
                case 1:
                    sNumberRaw = m_sPhone1;
                    break;
                case 2:
                    sNumberRaw = m_sPhone2;
                    break;
                case 3:
                    sNumberRaw = m_sPhone3;
                    break;
                }
            return sNumberRaw;
        }
        public void GetPhoneNumber(int iPhone, out string sNumber, out string sType)
        {
            string sNumberRaw = NumberRawForPhone(iPhone);
            ExtractNumberParts(sNumberRaw, out sNumber, out sType);
        }

        public void SetPhoneNumber(int iPhone, string sPhoneNumber, string sTypeDefault)
        {
            string sNumberRaw = SConstructPhoneNumberFromParts(sPhoneNumber, sTypeDefault);

            switch (iPhone)
                {
                case 1:
                    m_sPhone1 = sNumberRaw;
                    break;
                case 2:
                    m_sPhone2 = sNumberRaw;
                    break;
                case 3:
                    m_sPhone3 = sNumberRaw;
                    break;
                }
        }

        public void SetNextPhoneNumber(string sPhoneNumber, string sTypeDefault)
        {
            int iPhone = 1;

            if (String.IsNullOrEmpty(m_sPhone1))
                iPhone = 1;
            else if (String.IsNullOrEmpty(m_sPhone2))
                iPhone = 2;
            else if (String.IsNullOrEmpty(m_sPhone3))
                iPhone = 3;
            else
                return;

            SetPhoneNumber(iPhone, sPhoneNumber, sTypeDefault);
        }

        public Roster.RSTT m_rstt;

        public bool FMatchAnyMisc(string sRegexFilter)
        {
            foreach (string s2 in m_plsMisc)
                {
                if (Regex.Match(s2, sRegexFilter, System.Text.RegularExpressions.RegexOptions.IgnoreCase).Success)
                    {
                    return true;
                    }
                }
            return false;
        }

        public bool FEqualsMisc(RosterEntry rste)
        {
            if (m_plsMisc?.Count != rste?.m_plsMisc?.Count)
                return false;

            int i;

            for (i = 0; i < m_plsMisc.Count; i++)
                {
                if (String.Compare(m_plsMisc[i], rste.m_plsMisc[i]) != 0)
                    return false;
                }

            return true;
        }

        /* F  E Q U A L S */
        /*----------------------------------------------------------------------------
        	%%Function: FEquals
        	%%Qualified: ArbWeb.RosterEntry.FEquals
        	%%Contact: rlittle
        	
            Compare the settable server fields to see if they are equivalent
        ----------------------------------------------------------------------------*/
        public bool FEquals(RosterEntry rste)
        {
            if (String.Compare(m_sFirst, rste.m_sFirst) != 0)
                return false;
            if (String.Compare(m_sLast, rste.m_sLast) != 0)
                return false;
            if (String.Compare(m_sAddress1, rste.m_sAddress1) != 0)
                return false;
            if (String.Compare(m_sAddress2, rste.m_sAddress2) != 0)
                return false;
            if (String.Compare(m_sCity, rste.m_sCity) != 0)
                return false;
            if (String.Compare(m_sState, rste.m_sState) != 0)
                return false;
            if (String.Compare(m_sZip, rste.m_sZip) != 0)
                return false;
            if (String.Compare(m_sOfficialNumber, rste.m_sOfficialNumber) != 0)
                return false;
            if (String.Compare(m_sDateOfBirth, rste.m_sDateOfBirth) != 0)
                return false;
//            if (String.Compare(m_sDateJoined, rste.m_sDateJoined) != 0)
//                return false;
            if (String.Compare(m_sGamesPerDay, rste.m_sGamesPerDay) != 0)
                return false;
            if (String.Compare(m_sGamesPerWeek, rste.m_sGamesPerWeek) != 0)
                return false;
            if (String.Compare(m_sTotalGames, rste.m_sTotalGames) != 0)
                return false;
            if (String.Compare(m_sWaitMinutes, rste.m_sWaitMinutes) != 0)
                return false;
            if (String.Compare(m_sPhone1, rste.m_sPhone1) != 0)
                return false;
            if (String.Compare(m_sPhone2, rste.m_sPhone2) != 0)
                return false;
            if (String.Compare(m_sPhone3, rste.m_sPhone3) != 0)
                return false;

            return true;
        }

        public string OtherRanks(string sSport, string sPos, int nBase)
        {
            string sOther = "";
            foreach (string sKey in m_mpRanking.Keys)
                {
                if (sKey.StartsWith(sSport) && !sKey.EndsWith(sPos))
                    {
                    if (m_mpRanking[sKey] == nBase || m_mpRanking[sKey] == 1)
                        continue;

                    if (sOther.Length > 0)
                        sOther += ", ";
                    sOther = String.Format("{0}{1} ({2})", sOther, sKey.Substring(sSport.Length + 2), m_mpRanking[sKey]);
                    }
                }
            return sOther;
        }

        /* R  S  T  E */
        /*----------------------------------------------------------------------------
			%%Function: RSTE
			%%Qualified: ArbWeb.RSTE.RSTE
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public RosterEntry()
        {
            m_plsMisc = new List<string>();
            m_mpRanking = new Dictionary<string, int>();
            m_rstt = Roster.RSTT.Full;
        }

        public enum QuickShortColumns
        {
            FirstName = 0,
            LastName = 1, 
            Address1 = 2,
            Address2 = 3,
            City = 4,
            State = 5,
            PostalCode = 6,
            HomePhone = 7,
            WorkPhone = 8,
            CellPhone = 9,
            Email = 10,
            OfficialNumber = 11,
            DateJoined = 12,
            BuiltInMac = 13
        }

        static bool FVerifyQuickShortColumns(string[] rgs)
        {
            if (String.Compare(rgs[(int)QuickShortColumns.FirstName], "FirstName", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.LastName], "LastName", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.Address1], "Address1", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.Address2], "Address2", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.City], "City", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.State], "State", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.PostalCode], "PostalCode", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.HomePhone], "HomePhone", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.WorkPhone], "WorkPhone", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.CellPhone], "CellPhone", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.Email], "Email", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.OfficialNumber], "OfficalNumber", false) != 0  // take into account bad heading from ArbiterSports.
                && String.Compare(rgs[(int)QuickShortColumns.OfficialNumber], "OfficialNumber", false) != 0)
                return false;

            if (String.Compare(rgs[(int)QuickShortColumns.DateJoined], "DateJoined", false) != 0)
                return false;

            return true;
        }
        public static bool FVerifyHeaderColumns(string[] rgs, Roster.RSTT rstt)
        {
            if (rstt == Roster.RSTT.QuickShort)
                return FVerifyQuickShortColumns(rgs);

            return true;
        }

        /* R  S  T  E */
        /*----------------------------------------------------------------------------
			%%Function: RSTE
			%%Qualified: ArbWeb.RSTE.RSTE
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public RosterEntry(string[] rgs, List<string> plsMiscHeadings, List<string> plsRankings, Roster.RSTT rstt)
        {
            m_rstt = rstt;

            int i;

            m_fMarked = false;

            if (m_rstt == Roster.RSTT.QuickShort)
                {
                if (rgs.Length < 13)
                    throw new Exception("input line too short -- not enough fields");

                m_sFirst = rgs[(int)QuickShortColumns.FirstName];
                m_sLast = rgs[(int)QuickShortColumns.LastName];
                m_sEmail = rgs[(int)QuickShortColumns.Email];
                m_sAddress1 = rgs[(int)QuickShortColumns.Address1];
                m_sAddress2 = rgs[(int)QuickShortColumns.Address2];
                m_sCity = rgs[(int)QuickShortColumns.City];
                m_sState = rgs[(int)QuickShortColumns.State];
                m_sZip = rgs[(int)QuickShortColumns.PostalCode];

                SetNextPhoneNumber(rgs[(int)QuickShortColumns.HomePhone], "H");
                SetNextPhoneNumber(rgs[(int)QuickShortColumns.WorkPhone], "W");
                SetNextPhoneNumber(rgs[(int)QuickShortColumns.CellPhone], "C");

                m_sOfficialNumber = rgs[(int)QuickShortColumns.OfficialNumber];
                m_sDateJoined = rgs[(int)QuickShortColumns.DateJoined];
                i = (int)QuickShortColumns.BuiltInMac;

                m_mpRanking = new Dictionary<string, int>();
                }
            else
                {
                if (m_rstt == Roster.RSTT.QuickFull)
                    {
                    if (rgs.Length < 17)
                        throw new Exception("input line too short -- not enough fields");
                    }
                else if (m_rstt == Roster.RSTT.QuickFull2)
                    {
                    if (rgs.Length < 19)
                        throw new Exception("input line too short -- not enough fields");
                    }
                else
                    {
                    if (rgs.Length < 17)
                        throw new Exception("input line too short -- not enough fields");
                    }

                m_sFirst = rgs[0];
                m_sLast = rgs[1];
                m_sEmail = rgs[2];
                m_sAddress1 = rgs[3];
                m_sAddress2 = rgs[4];
                m_sCity = rgs[5];
                m_sState = rgs[6];
                m_sZip = rgs[7];
                SetNextPhoneNumber(rgs[8], "H");
                SetNextPhoneNumber(rgs[9], "W");
                SetNextPhoneNumber(rgs[10], "C");

                m_sOfficialNumber = rgs[11];
                if (m_rstt != Roster.RSTT.QuickFull && m_rstt != Roster.RSTT.QuickFull2)
                    {
                    throw (new Exception(
                        "This is probably rstt.Full - this was never updated to account for the phone fields we now write out."));
#if deadcode
                    m_sDateOfBirth = rgs[12];
                    m_sDateJoined = rgs[13];
                    m_sGamesPerDay = rgs[14];
                    m_sGamesPerWeek = rgs[15];
                    m_sTotalGames = rgs[16];
                    m_sWaitMinutes = rgs[17];

                    m_fReady = rgs[18] == "1" ? true : false;
                    m_fActive = rgs[19] == "1" ? true : false;
                    i = 20;
#endif
                    }
                else
                    {
                    // read in the misc fields
                    i = 12;
                    m_plsMisc = new List<string>();

                    foreach (string s in plsMiscHeadings)
                        {
                        m_plsMisc.Add(rgs[i]);
                        i++;
                        }

                    // now, read in the DateJoined
                    m_sDateJoined = rgs[i++];
                    if (m_rstt == Roster.RSTT.QuickFull2)
                        m_sLastSignin = rgs[i++];

                    // at this point, i points to the rankings...
                    }

                m_mpRanking = new Dictionary<string, int>();

                if (plsRankings != null)
                    {
                    int iRank = 0;
                    for (; iRank < plsRankings.Count; iRank++)
                        {
                        int nRank;
                        try
                            {
                            if (rgs[i + iRank] == "")
                                nRank = 0;
                            else
                                nRank = Int32.Parse(rgs[i + iRank]);
                            }
                        catch
                            {
                            nRank = 0;
                            }
                        if (nRank > 0)
                            m_mpRanking.Add(plsRankings[iRank], nRank);
                        }
                    i += iRank;
                    }
                }

            if (m_rstt != Roster.RSTT.QuickFull && m_rstt != Roster.RSTT.QuickFull2)
                {
                m_plsMisc = new List<string>();

                // and the rest are misc fields
                for (; i < rgs.Length; i++)
                    m_plsMisc.Add(rgs[i]);
                }
        }

        public string Name { get { return String.Format("{0} {1}", m_sFirst, m_sLast); } }
        public string First { get { return m_sFirst; } }
        public string Last { get { return m_sLast; } }
        public string Email { get { return m_sEmail; } }
        public bool Marked { get { return m_fMarked; } set { m_fMarked = value; } }
        public List<string> Misc {  get { return m_plsMisc; } }

        public int Rank(string s)
        {
            if (m_mpRanking == null || !m_mpRanking.ContainsKey(s))
                return -1;

            return m_mpRanking[s];
        }

        /* F  R A N K E D */
        /*----------------------------------------------------------------------------
			%%Function: FRanked
			%%Qualified: ArbWeb.RSTE.FRanked
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public bool FRanked(string s)
        {
            if (m_mpRanking == null || !m_mpRanking.ContainsKey(s))
                return false;

            return true;
        }

        public bool FRankedReal(string s)
        {
            if (!FRanked(s))
                return false;

            return m_mpRanking[s] > 0;
        }
        /* S E T  E M A I L */
        /*----------------------------------------------------------------------------
			%%Function: SetEmail
			%%Qualified: ArbWeb.RSTE.SetEmail
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public void SetEmail(string s)
        {
            m_sEmail = s.Substring(s.IndexOf(":") + 1);
        }

        /* P L S  M I S C  F R O M  H E A D I N G  L I N E */
        /*----------------------------------------------------------------------------
			%%Function: PlsMiscFromHeadingLine
			%%Qualified: ArbWeb.RSTE.PlsMiscFromHeadingLine
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public static List<string> PlsMiscFromHeadingLine(string[] rgs)
        {
            List<string> pls = new List<string>();
            if (rgs[12] == "DateJoined")
                {
                // this is a simple quickroster (straight from Arbiter)

                // everything past the 12th column is a misc field
                for (int i = 13; i < rgs.Length; i++)
                    {
                    pls.Add(rgs[i]);
                    }
                }
            else if (rgs[10] != "DateJoined" && rgs[9] != "DateJoined")
                {
                // this is a roster that we've written out, but it was
                // based on a quick roster, so its still missing some
                // information (so we'll continue to consider it a quick
                // roster)

                // misc fields are between OfficialNumber and DateJoined

                bool fCollecting = false;

                foreach (string s in rgs)
                    {
                    if (fCollecting)
                        {
                        if (s == "DateJoined")
                            break;

                        pls.Add(s);
                        }
                    else
                        {
                        if (s == "OfficialNumber")
                            fCollecting = true;
                        }
                    }
                if (!fCollecting)
                    throw (new Exception("never found DateJoined in a quickroster descendent file"));
                }
            else
                throw (new Exception("shouldn't try to read misc headings from a non-quick roster"));

            return pls;
        }

        /* P L S  R A N K I N G S  F R O M  H E A D I N G  L I N E */
        /*----------------------------------------------------------------------------
			%%Function: PlsRankingsFromHeadingLine
			%%Qualified: ArbWeb.RSTE.PlsRankingsFromHeadingLine
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public static List<string> PlsRankingsFromHeadingLine(string[] rgs, out Roster.RSTT rstt)
        {
            // first, check to see if this is a quick roster
            if (rgs[2] == "Address1")
                {
                rstt = Roster.RSTT.QuickShort;

                // there will be no ranks, but misc fields will be labeled
                return null;
                }

            // ok, that was the obvious quick roster.  now we might have a roster we have written
            // out that was based on a quick roster (it will be missing some fields, but it WILL 
            // have rankings...)

            int iRank;

            if (rgs[10] != "DateJoined" && rgs[9] != "DateJoined")
                {
                // this is based on a quick roster...read the misc and the rankings but remember that
                // its a quick roster

                rstt = Roster.RSTT.QuickFull;

                // rankings start right after the DateJoined field (or LastSignin for QuickFull2)
                iRank = 0;
                while (iRank < rgs.Length)
                    {
                    if (rgs[iRank] == "DateJoined")
                        {
                        if (iRank + 1 < rgs.Length && rgs[iRank + 1] == "LastSignin")
                            {
                            rstt = Roster.RSTT.QuickFull2;
                            iRank++;
                            }
                        break;
                        }
                    iRank++;
                    }
                iRank++;
                if (iRank > rgs.Length)
                    throw (new Exception("bad format in heading line -- found no DateJoined in a quickroster, or date joined beyond the end of the array"));
                }
            else
                {
                throw (new Exception("No idea if this code works anymore..."));
                // ranks start at column 17 and go until we see "Misc..." (or run out of fields)

#if deadcode
                rstt = Roster.RSTT.Full;
                iRank = 17;
#endif
            }

            List<string> pls = new List<string>();

            while (iRank < rgs.Length)
                {
                if (rgs[iRank].Contains("Misc"))
                    break;
                pls.Add(rgs[iRank]);
                iRank++;
                }

            if (pls.Count == 0)
                return null;

            return pls;
        }

        /* W R I T E  H E A D E R  T O  F I L E */
        /*----------------------------------------------------------------------------
			%%Function: WriteHeaderToFile
			%%Qualified: ArbWeb.RSTE.WriteHeaderToFile
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public static void WriteHeaderToFile(string sFile, List<string> plsMiscHeadings, List<string> plsRankings, Roster.RSTT rstt)
        {
            StreamWriter sw = new StreamWriter(sFile, false, System.Text.Encoding.Default);

            if (sw == null)
                throw new Exception("could not create file to write header");

            if (rstt != Roster.RSTT.Full)
                {
                // be careful here -- we have to have a predictable field BEFORE the misc fields, 
                // as well as AFTER the misc fields.  This way, we can tell the difference between
                // misc fields and rankings
                sw.Write("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\"",
                         "First", "Last", "Email", "Address1", "Address2", "City", "State", "Zip", "Phone1", "Phone2", "Phone3", "OfficialNumber");

                foreach (string s in plsMiscHeadings)
                    {
                    sw.Write(",\"{0}\"", s);
                    }

                sw.Write(",\"DateJoined\"");
                sw.Write(",\"LastSignin\"");
                }
            else
                {
                sw.Write("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\",\"{12}\",\"{13}\",\"{14}\",\"{15}\",\"{16}\",\"{17}\",\"{18}\",\"{19}\",\"{20}\"",
                         "First", "Last", "Email", "Address1", "Address2", "City", "State", "Zip", "Phone1", "Phone2", "Phone3", "OfficialNumber", "DateOfBirth", "DateJoined", "LastSignin",
                         "GamesPerDay", "GamesPerWeek", "TotalGames", "WaitMinutes", "Ready", "Active");
                }


            if (plsRankings != null)
                {
                foreach (string s in plsRankings)
                    sw.Write(",\"{0}\"", s);
                }

            if (rstt == Roster.RSTT.Full)
                sw.Write(",\"{0}\"", "Misc...");

            sw.WriteLine();
            sw.Close();
        }

        /* A P P E N D  T O  F I L E */
        /*----------------------------------------------------------------------------
			%%Function: AppendToFile
			%%Qualified: ArbWeb.RSTE.AppendToFile
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public void AppendToFile(string sFile, List<string> plsRankings)
        {
            StreamWriter sw = new StreamWriter(sFile, true, System.Text.Encoding.Default);

            if (sw == null)
                throw new Exception("could not append to file");

            if (m_rstt != Roster.RSTT.Full)
                {
                sw.Write("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\"",
                         m_sFirst, m_sLast, m_sEmail, m_sAddress1, m_sAddress2, m_sCity, m_sState, m_sZip, m_sPhone1, m_sPhone2, m_sPhone3, m_sOfficialNumber);
                }
            else
                {
                sw.Write("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\",\"{12}\",\"{13}\",\"{14}\",\"{15}\",\"{16}\",\"{17}\",\"{18}\",\"{19}\",\"{20}\"",
                         m_sFirst, m_sLast, m_sEmail, m_sAddress1, m_sAddress2, m_sCity, m_sState, m_sZip, m_sPhone1, m_sPhone2, m_sPhone3, m_sOfficialNumber, m_sDateOfBirth, m_sDateJoined, m_sLastSignin,
                         m_sGamesPerDay, m_sGamesPerWeek, m_sTotalGames, m_sWaitMinutes, m_fReady ? "1" : "0", m_fActive ? "1" : "0");
                }

            if (m_rstt != Roster.RSTT.Full)
                {
                foreach (string s in m_plsMisc)
                    sw.Write(",\"{0}\"", s);

                sw.Write(",\"{0}\"", m_sDateJoined);
                sw.Write(",\"{0}\"", m_sLastSignin);
                }

            if (plsRankings != null)
                {
                foreach (string s in plsRankings)
                    {
                    if (m_mpRanking.ContainsKey(s))
                        sw.Write(",\"{0}\"", m_mpRanking[s]);
                    else
                        sw.Write(",\"\"");
                    }
                }

            if (m_rstt == Roster.RSTT.Full)
                {
                foreach (string s in m_plsMisc)
                    sw.Write(",\"{0}\"", s);
                }

            sw.WriteLine();
            sw.Close();
        }

//			public RSTE(string sEmail, List<string> plsMisc)
//			{
//			    m_sEmail = sEmail;
//			    m_plsMisc = new List<string>(plsMisc);
//			}

    }

    // ================================================================================
    //  R O S T E R 
    // ================================================================================
    public class Roster // RST
    {
        private List<RosterEntry> m_plrste;
        private List<string> m_plsRankings;
        private List<string> m_plsMisc;

        private RSTT m_rstt;

        public enum RSTT // RosterType
        {
            Full,
            QuickShort,
//    		QuickShort2,	// this includes LastSignin
            QuickFull,
            QuickFull2 // this includes LastSignin
        };

        public Roster()
        {
            m_plrste = new List<RosterEntry>();
            m_plsRankings = new List<string>();
        }

        public bool IsQuick { get { return m_rstt != RSTT.Full; } }
        public bool IsUploadableQuickroster {  get { return m_rstt == RSTT.QuickFull || m_rstt == RSTT.QuickFull2; } }
        public RSTT Rstt { get { return m_rstt; } }

        public bool HasRankings { get { return m_plsRankings.Count > 0; } }

        public List<string> PlsRankings { get { return m_plsRankings; } set { m_plsRankings = value; } }

        public List<string> PlsMisc { get => m_plsMisc; set => m_plsMisc = value; }

        /* R E A D  R O S T E R */
        /*----------------------------------------------------------------------------
			%%Function: ReadRoster
			%%Qualified: ArbWeb.RST.ReadRoster
			%%Contact: rlittle

			Handle both a full roster and a quick roster
		----------------------------------------------------------------------------*/
        public void ReadRoster(string sFile)
        {

            TextReader tr = new StreamReader(sFile);
            string sLine;
            string[] rgs;
            m_plrste = new List<RosterEntry>();
            bool fFirst = true;
            int cMiscMax = 0;

            m_plsRankings = null;
            while ((sLine = tr.ReadLine()) != null)
                {
                if (String.IsNullOrEmpty(sLine) || String.IsNullOrWhiteSpace(sLine))
                    continue;

                rgs = Csv.LineToArray(sLine);

                if (fFirst && sLine.Contains("First") && sLine.Contains("Email") && sLine.Contains("Address2"))
                    {
                    // grab the list of rankings from the first line
                    m_plsRankings = RosterEntry.PlsRankingsFromHeadingLine(rgs, out m_rstt);
                    if (m_rstt != RSTT.Full)
                        m_plsMisc = RosterEntry.PlsMiscFromHeadingLine(rgs);

                    fFirst = false;
                    if (!RosterEntry.FVerifyHeaderColumns(rgs, m_rstt))
                        throw new Exception("Column headers broken");

                    continue; // skip heading line
                    }

                fFirst = false;
                RosterEntry rste = new RosterEntry(rgs, m_plsMisc, m_plsRankings, m_rstt);

                if (m_rstt == RSTT.Full)
                    {
                    cMiscMax = Math.Max(cMiscMax, rste.m_plsMisc.Count);
                    }

                m_plrste.Add(rste);
                }

            if (m_rstt == RSTT.Full)
                {
                m_plsMisc = new List<string>();

                for (int i = 0; i < cMiscMax; i++)
                    {
                    m_plsMisc.Add(String.Format("Misc{0}", i));
                    }
                }
        }

        /* W R I T E  R O S T E R */
        /*----------------------------------------------------------------------------
			%%Function: WriteRoster
			%%Qualified: ArbWeb.RST.WriteRoster
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public void WriteRoster(string sFile)
        {
            RosterEntry.WriteHeaderToFile(sFile, m_plsMisc, m_plsRankings, m_rstt);

            foreach (RosterEntry rste in m_plrste)
                rste.AppendToFile(sFile, m_plsRankings);
        }

        /* P L R S T E  U N M A R K E D */
        /*----------------------------------------------------------------------------
			%%Function: PlrsteUnmarked
			%%Qualified: ArbWeb.RST.PlrsteUnmarked
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public List<RosterEntry> PlrsteUnmarked()
        {
            List<RosterEntry> plrste = new List<RosterEntry>();

            foreach (RosterEntry rste in m_plrste)
                {
                if (!rste.Marked && !String.IsNullOrEmpty(rste.Email))
                    plrste.Add(rste);
                }
            return plrste;
        }

        public List<RosterEntry> Plrste { get { return m_plrste; } }

        /* F  S P L I T  N A M E */
        /*----------------------------------------------------------------------------
			%%Function: FSplitName
			%%Qualified: ArbWeb.RST.FSplitName
			%%Contact: rlittle

			split name 
		----------------------------------------------------------------------------*/
        public static bool FSplitName(string sReversed, out string sFirst, out string sLast)
        {
            string[] rgs;
            sFirst = sLast = null;

            // rgs = CountsData.RexHelper.RgsMatch(sReversed, "^[ \t]*([^\t]*)[ \t]*,[ \t]*([^ \t]*) *$");
            rgs = CountsData.RexHelper.RgsMatch(sReversed, "^[ \t]*([^\t]*)[ \t]*,[ \t]*([^\t,]*) *$");

            if (rgs[0] == null || rgs[1] == null)
                return false;

            sFirst = rgs[1];
            sLast = rgs[0];

            return true;
        }

        /* R S T E  L O O K U P  R E V E R S E D  N A M E */
        /*----------------------------------------------------------------------------
			%%Function: RsteLookupReversedName
			%%Qualified: ArbWeb.RST.RsteLookupReversedName
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public RosterEntry RsteLookupReversedName(string sReversed)
        {
            string sFirst;
            string sLast;

            if (!Roster.FSplitName(sReversed, out sFirst, out sLast))
                return null;

            foreach (RosterEntry rste in m_plrste)
                {
                if (String.Compare(rste.m_sFirst, sFirst, true) == 0
                    && String.Compare(rste.m_sLast, sLast, true) == 0)
                    {
                    return rste;
                    }
                }

            return null;
        }

        /* A D D */
        /*----------------------------------------------------------------------------
			%%Function: Add
			%%Qualified: ArbWeb.RST.Add
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public void Add(RosterEntry rste)
        {
            m_plrste.Add(rste);
        }

        public Roster FilterByRanks(List<string> plsRequiredRanks)
        {
            Roster rst = new Roster();

            foreach (RosterEntry rste in m_plrste)
                {
                // each entry must be rated for at least one of the required ranks
                foreach (string s in plsRequiredRanks)
                    {
                    if (rste.FRankedReal(s))
                        {
                        rst.Add(rste);
                        break;
                        }
                    }
                }

            return rst;
        }

        /* F  A D D  R A N K I N G */
        /*----------------------------------------------------------------------------
			%%Function: FAddRanking
			%%Qualified: ArbWeb.RST.FAddRanking
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public bool FAddRanking(string sName, string sPosition, int nRank)
        {
            RosterEntry rste = RsteLookupReversedName(sName);

            if (rste == null)
                return false;

            if (rste.m_mpRanking.ContainsKey(sPosition))
                rste.m_mpRanking[sPosition] = nRank;
            else
                rste.m_mpRanking.Add(sPosition, nRank);

            // now, see if its in our list
            if (!m_plsRankings.Contains(sPosition))
                m_plsRankings.Add(sPosition);

            return true;
        }


        /* R S T E  L O O K U P  E M A I L */
        /*----------------------------------------------------------------------------
			%%Function: RsteLookupEmail
			%%Qualified: ArbWeb.RST.RsteLookupEmail
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public RosterEntry RsteLookupEmail(string sEmail)
        {
            if (sEmail == null)
                return null;

            if (sEmail.IndexOf(":") >= 0)
                sEmail = sEmail.Substring(sEmail.IndexOf(":") + 1);

            if (sEmail.Length == 0)
                return null;

            foreach (RosterEntry rste in m_plrste)
                {
                if (string.Compare(rste.m_sEmail, sEmail, true /*ignoreCase*/) == 0)
                    return rste;
                }
            return null;
        }

        /* P L S  L O O K U P  E M A I L */
        /*----------------------------------------------------------------------------
			%%Function: PlsMiscLookupEmail
			%%Qualified: ArbWeb.RST.PlsMiscLookupEmail
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public List<string> PlsMiscLookupEmail(string sEmail)
        {
            RosterEntry rste = RsteLookupEmail(sEmail);
            if (rste == null)
                return null;

            return rste.m_plsMisc;
        }

        /* S  B U I L D  A D D R E S S  L I N E */
        /*----------------------------------------------------------------------------
			%%Function: SBuildAddressLine
			%%Qualified: ArbWeb.RST.SBuildAddressLine
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        public string SBuildAddressLine(string sRegexFilter)
        {
            string s = "";
            bool fFirst = true;

            foreach (RosterEntry rste in m_plrste)
                {
                if (sRegexFilter != null)
                    {
                    if (rste.FMatchAnyMisc(sRegexFilter))
                        continue;
                    }
                if (!fFirst)
                    s += ";";
                fFirst = false;
                s += rste.m_sEmail;
                }
            return s;
        }
    }
}