using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;

using Microsoft.Win32;
using System.Collections;
using System.Net;
using AxSHDocVw;
using mshtml;
using Microsoft.Office;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArbWeb
	{
//          _  _ _  _ ___  _ ____ ____ 
//          |  | |\/| |__] | |__/ |___ 
//          |__| |  | |    | |  \ |___ 
//
	public class Umpire	// UMP
		{
		string m_sFirst;
		string m_sLast;
		string m_sContact;
		string m_sMisc;
		List<string> m_plsMisc;

		/* U M P I R E */
		/*----------------------------------------------------------------------------
			%%Function: Umpire
			%%Qualified: GenCount.GenCounts:GenGameStats:Umpire.Umpire
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

		public string FirstName { get { return m_sFirst;}}
		public string LastName { get { return m_sLast;}}
		public string Contact { get { return m_sContact;}}
		public string Misc { get { return m_sMisc;}}
		public string Name { get { return String.Format("{0},{1}", m_sLast, m_sFirst);}}
		public List<string> PlsMisc { get { return m_plsMisc;}}

		} // END UMPIRE

	class Csv
	{
		public static string[] LineToArray(string line)
		{
			String pattern = ",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))";
			Regex r = new Regex(pattern);

			string[] rgs = r.Split(line);

			for (int i = 0; i < rgs.Length; i++)
				{
				if (rgs[i].Length > 0 && rgs[i][0] == '"')
					rgs[i] = rgs[i].Substring(1, rgs[i].Length - 2);
				}
			return rgs;
	   }
	};

//          ____ ____ ____ ___ ____ ____ 
//          |__/ |  | [__   |  |___ |__/ 
//          |  \ |__| ___]  |  |___ |  \ 
// 
	public class RSTE // Roster Entry
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

		public int m_cRankings;
		public Dictionary<string, int> m_mpRanking;

		public bool m_fReady;
		public bool m_fActive;

		public List<string> m_plsMisc;

		public bool m_fMarked;

		public RST.RSTT m_rstt;

		public RSTE()
		{
			m_plsMisc = new List<string>();
			m_mpRanking = new Dictionary<string, int>();
			m_rstt = RST.RSTT.Full;
		}

		public RSTE(string[] rgs, List<string> plsMiscHeadings, List<string> plsRankings, RST.RSTT rstt)
		{
			m_rstt = rstt;

			int i;

			m_fMarked = false;

			if (m_rstt == RST.RSTT.QuickShort || m_rstt == RST.RSTT.QuickShort2)
				{
				if (rgs.Length < 13)
					throw new Exception("input line too short -- not enough fields");

				m_sFirst = rgs[0];
				m_sLast = rgs[1];
				m_sEmail = rgs[10];
				m_sAddress1 = rgs[2];
				m_sAddress2 = rgs[3];
				m_sCity = rgs[4];
				m_sState = rgs[5];
				m_sZip = rgs[6];
				m_sOfficialNumber = rgs[11];
				m_sDateJoined = rgs[12];
    			i = 13;

				if (m_rstt == RST.RSTT.QuickShort2)
					{
					m_sLastSignin = rgs[13];
    				i++;
					}
				m_mpRanking = new Dictionary<string, int>();
				}
			else
				{
				if (m_rstt == RST.RSTT.QuickFull)
					{
					if (rgs.Length < 14)
						throw new Exception("input line too short -- not enough fields");
					}
    			else if (m_rstt == RST.RSTT.QuickFull2)
					{
					if (rgs.Length < 15)
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
				m_sOfficialNumber = rgs[8];
				if (m_rstt != RST.RSTT.QuickFull && m_rstt != RST.RSTT.QuickFull2)
					{
					m_sDateOfBirth = rgs[9];
					m_sDateJoined = rgs[10];
					m_sGamesPerDay = rgs[11];
					m_sGamesPerWeek = rgs[12];
					m_sTotalGames = rgs[13];
					m_sWaitMinutes = rgs[14];
	
					m_fReady = rgs[15] == "1" ? true : false;
					m_fActive = rgs[16] == "1" ? true : false;
					i = 17;
					}
				else
					{
					// read in the misc fields
					i = 9;
					m_plsMisc = new List<string>();

					foreach(string s in plsMiscHeadings)
						{
						m_plsMisc.Add(rgs[i]);
						i++;
						}

					// now, read in the DateJoined
					m_sDateJoined = rgs[i++];
    				if (m_rstt == RST.RSTT.QuickFull2)
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
						try {
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

			if (m_rstt != RST.RSTT.QuickFull && m_rstt != RST.RSTT.QuickFull2)
				{
				m_plsMisc = new List<string>();

				// and the rest are misc fields
				for (; i < rgs.Length; i++)
					m_plsMisc.Add(rgs[i]);
				}
		}

		public string Name { get { return String.Format("{0} {1}", m_sFirst, m_sLast); } }

		public bool Marked { get { return m_fMarked; } set { m_fMarked = value; } }

		public int Rank(string s)
		{
			if (m_mpRanking == null || !m_mpRanking.ContainsKey(s))
				return -1;

			return m_mpRanking[s];
		}

		public bool FRanked(string s)
		{
			if (m_mpRanking == null || !m_mpRanking.ContainsKey(s))
				return false;

			return true;
		}

		public void SetEmail(string s)
		{
			m_sEmail = s.Substring(s.IndexOf(":") + 1);
		}

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

				foreach(string s in rgs)
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
					throw(new Exception("never found DateJoined in a quickroster descendent file"));
				}
			else
				throw(new Exception("shouldn't try to read misc headings from a non-quick roster"));

			return pls;
		}

		public static List<string> PlsRankingsFromHeadingLine(string[] rgs, out RST.RSTT rstt)
		{
			// first, check to see if this is a quick roster
			if (rgs[2] == "Address1")
				{
				rstt = RST.RSTT.QuickShort;
				
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

                rstt = RST.RSTT.QuickFull;

				// rankings start right after the DateJoined field (or LastSignin for QuickFull2)
				iRank = 0;
				while (iRank < rgs.Length)
					{
					if (rgs[iRank] == "DateJoined")
					    {
					    if (iRank + 1 < rgs.Length && rgs[iRank + 1] == "LastSignin")
					        {
					        rstt = RST.RSTT.QuickFull2;
					        iRank++;
					        }
					    break;
					    }
					iRank++;
					}
				iRank++;
				if (iRank >= rgs.Length)
					throw(new Exception("bad format in heading line -- found no DateJoined in a quickroster"));
				}
			else
				{
				// ranks start at column 17 and go until we see "Misc..." (or run out of fields)

                rstt = RST.RSTT.Full;
				iRank = 17;
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

		public static void WriteHeaderToFile(string sFile, List<string> plsMiscHeadings, List<string> plsRankings, RST.RSTT rstt)
		{
			StreamWriter sw = new StreamWriter(sFile, false, System.Text.Encoding.Default);

			if (sw == null)
				throw new Exception("could not create file to write header");

		    if (rstt != RST.RSTT.Full)
				{
				// be careful here -- we have to have a predictable field BEFORE the misc fields, 
				// as well as AFTER the misc fields.  This way, we can tell the difference between
				// misc fields and rankings
				sw.Write("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\"", 
						 "First", "Last", "Email", "Address1", "Address2", "City", "State", "Zip", "OfficialNumber");

				foreach(string s in plsMiscHeadings)
					{
					sw.Write(",\"{0}\"", s);
					}

				sw.Write(",\"DateJoined\"");
				sw.Write(",\"LastSignin\"");
				}
			else
				{
				sw.Write("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\",\"{12}\",\"{13}\",\"{14}\",\"{15}\",\"{16}\",\"{17}\"", 
						 "First", "Last", "Email", "Address1", "Address2", "City", "State", "Zip", "OfficialNumber", "DateOfBirth", "DateJoined", "LastSignin",
						 "GamesPerDay", "GamesPerWeek", "TotalGames", "WaitMinutes", "Ready", "Active");
				}


			if (plsRankings != null)
				{
				foreach(string s in plsRankings)
					sw.Write(",\"{0}\"", s);
				}

            if (rstt == RST.RSTT.Full)
				sw.Write(",\"{0}\"", "Misc...");

			sw.WriteLine();
			sw.Close();
		}

		public void AppendToFile(string sFile, List<string> plsRankings)
		{
			StreamWriter sw = new StreamWriter(sFile, true, System.Text.Encoding.Default);

			if (sw == null)
				throw new Exception("could not append to file");

			if (m_rstt != RST.RSTT.Full)
				{
				sw.Write("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\"", 
						 m_sFirst, m_sLast, m_sEmail, m_sAddress1, m_sAddress2, m_sCity, m_sState, m_sZip, m_sOfficialNumber);
				}
			else
				{
				sw.Write("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",\"{10}\",\"{11}\",\"{12}\",\"{13}\",\"{14}\",\"{15}\",\"{16}\",\"{17}\"", 
						 m_sFirst, m_sLast, m_sEmail, m_sAddress1, m_sAddress2, m_sCity, m_sState, m_sZip, m_sOfficialNumber, m_sDateOfBirth, m_sDateJoined, m_sLastSignin,
						 m_sGamesPerDay, m_sGamesPerWeek, m_sTotalGames, m_sWaitMinutes, m_fReady ? "1" : "0", m_fActive ? "1" : "0");
				}

			if (m_rstt != RST.RSTT.Full)
				{
				foreach(string s in m_plsMisc)
					sw.Write(",\"{0}\"", s);

				sw.Write(",\"{0}\"", m_sDateJoined);
				sw.Write(",\"{0}\"", m_sLastSignin);
				}

			if (plsRankings != null)
				{
				foreach(string s in plsRankings)
					{
					if (m_mpRanking.ContainsKey(s))
						sw.Write(",\"{0}\"", m_mpRanking[s]);
					else
						sw.Write(",\"\"");
					}
				}

            if (m_rstt == RST.RSTT.Full)
				{
				foreach(string s in m_plsMisc)
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

	public class RST 
	{
		List<RSTE> m_plrste;
		List<string> m_plsRankings;
		List<string> m_plsMisc;

		RSTT m_rstt;

		public enum RSTT // RosterType
		{
			Full,
			QuickShort,
    		QuickShort2,	// this includes LastSignin
			QuickFull,
    		QuickFull2		// this includes LastSignin
		};

		public RST()
		{
			m_plrste = new List<RSTE>();
			m_plsRankings = new List<string>();
		}

		public bool IsQuick { get { return m_rstt != RSTT.Full; } }
		public RSTT Rstt { get { return m_rstt; } }

		public bool HasRankings { get { return m_plsRankings.Count > 0; } }

		public List<string> PlsRankings { get { return m_plsRankings; } set { m_plsRankings = value; } }

		public List<string> PlsMisc { get { return m_plsMisc; } }

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
			string []rgs;
			m_plrste = new List<RSTE>();
			bool fFirst = true;
			int cMiscMax = 0;

			m_plsRankings = null;
			while ((sLine = tr.ReadLine()) != null)
				{
				rgs = Csv.LineToArray(sLine);

				if (fFirst && sLine.Contains("First") && sLine.Contains("Email") && sLine.Contains("Address2"))
					{
					// grab the list of rankings from the first line
					m_plsRankings = RSTE.PlsRankingsFromHeadingLine(rgs, out m_rstt);
					if (m_rstt != RSTT.Full)
						m_plsMisc = RSTE.PlsMiscFromHeadingLine(rgs);

					fFirst = false;
					continue;	// skip heading line
					}

				fFirst = false;
				RSTE rste = new RSTE(rgs, m_plsMisc, m_plsRankings, m_rstt);

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

		public void WriteRoster(string sFile)
		{
			RSTE.WriteHeaderToFile(sFile, m_plsMisc, m_plsRankings, m_rstt);

			foreach (RSTE rste in m_plrste)
				rste.AppendToFile(sFile, m_plsRankings);
		}

		public List<RSTE> PlrsteUnmarked()
		{
			List<RSTE> plrste = new List<RSTE>();

			foreach (RSTE rste in m_plrste)
				{
				if (!rste.Marked)
					plrste.Add(rste);
				}
			return plrste;
		}

		public List<RSTE> Plrste { get { return m_plrste; } }

		/* F  S P L I T  N A M E */
		/*----------------------------------------------------------------------------
			%%Function: FSplitName
			%%Qualified: ArbWeb.RST.FSplitName
			%%Contact: rlittle

			split name 
		----------------------------------------------------------------------------*/
		static public bool FSplitName(string sReversed, out string sFirst, out string sLast)
		{
			string[] rgs;
            sFirst = sLast = null;

            // rgs = GenCounts.RexHelper.RgsMatch(sReversed, "^[ \t]*([^\t]*)[ \t]*,[ \t]*([^ \t]*) *$");
            rgs = GenCounts.RexHelper.RgsMatch(sReversed, "^[ \t]*([^\t]*)[ \t]*,[ \t]*([^\t,]*) *$");

			if (rgs[0] == null || rgs[1] == null)
				return false;

			sFirst = rgs[1];
			sLast = rgs[0];

			return true;
		}

		public RSTE RsteLookupReversedName(string sReversed)
		{
			string sFirst;
			string sLast;

			if (!RST.FSplitName(sReversed, out sFirst, out sLast))
				return null;

			foreach (RSTE rste in m_plrste)
				{
				if (String.Compare(rste.m_sFirst, sFirst, true) == 0
					&& String.Compare(rste.m_sLast, sLast, true) == 0)
					{
					return rste;
					}
				}

			return null;
		}

		public void Add(RSTE rste)
		{
			m_plrste.Add(rste);
		}

		public bool FAddRanking(string sName, string sPosition, int nRank)
		{
			RSTE rste = RsteLookupReversedName(sName);

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


		public RSTE RsteLookupEmail(string sEmail)
		{
    		if (sEmail == null)
    			return null;

			if (sEmail.IndexOf(":") >= 0)
				sEmail = sEmail.Substring(sEmail.IndexOf(":") + 1);

			if (sEmail.Length == 0)
				return null;

			foreach (RSTE rste in m_plrste)
				{
				if (string.Compare(rste.m_sEmail, sEmail, true/*ignoreCase*/) == 0)
					return rste;
				}
			return null;
		}

		public List<string> PlsLookupEmail(string sEmail)
		{
			RSTE rste = RsteLookupEmail(sEmail);
			if (rste == null)
				return null;

			return rste.m_plsMisc;
		}

		public string SBuildAddressLine(string sRegexFilter)
		{
			string s = "";
			bool fFirst = true;

			foreach (RSTE rste in m_plrste)
				{
				if (sRegexFilter != null)
					{
					bool fMatched = false;

					foreach (string s2 in rste.m_plsMisc)
						{
						if (Regex.Match(s2, sRegexFilter, System.Text.RegularExpressions.RegexOptions.IgnoreCase).Success)
							{
							fMatched = true;
							break;
							}
						}
					if (!fMatched)
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

	public partial class GenCounts
		{
		public static class RexHelper
			{
			public static string[] RgsMatch(string sInput, string sRex)
			{
				Regex rex = new Regex(sRex);
				Match m = rex.Match(sInput);
				int c = m.Groups.Count - 1;

				string[] rgs = new string[m.Groups.Count - 1];
				int i;
				for (i = 1; i <= c; i++)
					rgs[i - 1] = m.Groups[i].Value;

				return rgs;
			}
			}

//      ____ ____ _  _    ____ ____ _  _ ____    ____ ___ ____ ___ ____ 
//      | __ |___ |\ |    | __ |__| |\/| |___    [__   |  |__|  |  [__  
//      |__] |___ | \|    |__] |  | |  | |___    ___]  |  |  |  |  ___] 
// 
		public class GenGameStats
			{
			public class Roster	// RST
				{
				RST m_rst;
				Dictionary<string, Umpire> m_mpNameUmpire;
				StatusBox.StatusRpt m_srpt;

				public List<string> PlsMiscHeadings { get { return m_rst.PlsMisc; } }

				public string SMiscHeader(int i)
				{
					if (m_rst.PlsMisc != null)
						return m_rst.PlsMisc[i];
					return "";
				}

				public Roster(StatusBox.StatusRpt srpt)
				{
					m_mpNameUmpire = new Dictionary<string, Umpire>();
					m_rst = new RST();
					m_srpt = srpt;
				}

				public bool LoadRoster(string sRoster, int iMiscAffiliation)
				{
					m_rst.ReadRoster(sRoster);
					foreach (RSTE rste in m_rst.Plrste)
						{
						Umpire ump = new Umpire(rste.m_sFirst, rste.m_sLast, rste.m_plsMisc[iMiscAffiliation], rste.m_sEmail, rste.m_plsMisc);

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

//          ____ ____ _  _ ____ 
//          | __ |__| |\/| |___ 
//          |__] |  | |  | |___ 

			public class Game // GM
			{
				DateTime m_dttm;
				string m_sSite;
				string m_sName;
				string m_sHome;
				string m_sAway;
				string m_sSport;
				string m_sPos;
				string m_sTeam;
				string m_sEmail;
				string m_sGameNum;
				string m_sLevel;
				bool m_fCancelled;
				List<string> m_plsMisc;

				public Game(DateTime dttm, string sSite, string sName, string sTeam, string sEmail, string sGameNum, string sHome, string sAway, string sLevel, string sSport, string sPos, bool fCancelled, List<string> plsMisc)
				{
					m_dttm = dttm;
					m_sSite = sSite;
					m_sName = sName;
					m_sHome = sHome;
					m_sAway = sAway;
					m_sSport = sSport;
					m_sPos = sPos;
					m_sTeam = sTeam;
					m_sEmail = sEmail;
					m_sGameNum = sGameNum;
					m_fCancelled = fCancelled;
					m_sLevel = sLevel;
					m_plsMisc = plsMisc;
				}

				public List<string> PlsMisc
				{
					get { return m_plsMisc; }
					set { m_plsMisc = value; }
				}

				public string Team
				{
					get { return m_sTeam; }
					set { m_sTeam = value; }
				}

				public bool Open
				{
					get { return m_sName == null; }
				}

				public string Home
				{
					get { return m_sHome; }
				}

				public string Away
				{
					get { return m_sAway; }
				}

				public string GameNum
				{
					get { return m_sGameNum; }
				}

				public bool Cancelled
				{
					get { return m_fCancelled; } 
				}

				public string Level
				{
					get { return m_sLevel; }
				}

				public string Sport
				{
					get { return m_sSport; }
				}

				public string Pos
				{
					get { return m_sPos; } 
				}

				public DateTime Dttm
				{
					get { return m_dttm; }
				}

				public string Site
				{
					get { return m_sSite; }
				}

				public string SportLevel
				{
					get { return String.Format("{0} {1}", m_sSport, m_sLevel); }
				}

				public string SiteShort
				{
					get
					{
						// get rid of any trailing fields
						string s = Regex.Replace(m_sSite, " [A-D]$", "");

						s = Regex.Replace(s, " #[1-9]$", "");

						s = Regex.Replace(s, " South$", "");

						s = Regex.Replace(s, " East$", "");
						s = Regex.Replace(s, " West$", "");
						s = Regex.Replace(s, " North$", "");
						s = Regex.Replace(s, " Varsity Field$", "");
						s = Regex.Replace(s, " JV Field$", "");

						s = Regex.Replace(s, " #[1-9][ ]*[69]0'$", "");
						s = Regex.Replace(s, " #[1-9][ ]*\\([69]0'\\)$", "");
						return s;
					}
				}


				/* S  R E P O R T */
				/*----------------------------------------------------------------------------
					%%Function: SReport
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Game.SReport
					%%Contact: rlittle

					Return a detail string suitable for saving in CSV format.

					Takes the given legend and saves out our collected data according to that
					legend.
				----------------------------------------------------------------------------*/
				public string SReport(List<string> plsLegend)
				{
					Dictionary<string, string> m_mpFieldVal = new Dictionary<string, string>();

					m_mpFieldVal.Add("UmpireName", "\"" + m_sName + "\"");
					m_mpFieldVal.Add("Team", m_sTeam);
					m_mpFieldVal.Add("Email", m_sEmail);
					m_mpFieldVal.Add("Game", m_sGameNum);
//					string sDateTime = m_dttm.ToString("M/d/yyyy ddd h:mm tt");
					string sDateTime = m_dttm.ToString("M/d/yyyy H:mm");

					m_mpFieldVal.Add("DateTime", sDateTime);
					m_mpFieldVal.Add("Date", m_dttm.ToString("M/d/yyyy"));
					m_mpFieldVal.Add("Time", m_dttm.ToString("H:mm"));
					m_mpFieldVal.Add("Level", m_sLevel);
					m_mpFieldVal.Add("Home", m_sHome);
					m_mpFieldVal.Add("Away", m_sAway);
					m_mpFieldVal.Add("Site", m_sSite);
					m_mpFieldVal.Add("Description", String.Format("{0}: [{1}] {2}: {3} vs. {4} ({5} {6})", m_sPos, m_sGameNum, sDateTime, m_sHome, m_sAway, m_sSport, m_sLevel));
					m_mpFieldVal.Add("Cancelled", m_fCancelled ? "1" : "0");

					m_mpFieldVal.Add("Sport", m_sSport);

					string sSportLevelPos = String.Format("{0}-{1}-{2}", m_sSport, m_sLevel, m_sPos);
					string sSportTotal = String.Format("{0}-Total", m_sSport);
					string sSportLevelTotal = String.Format("{0}-{1}-Total", m_sSport, m_sLevel);
					string sSportPos = String.Format("{0}-{1}", m_sSport, m_sPos);
					string sTotal = String.Format("Total");

					m_mpFieldVal.Add(sSportLevelPos, "1");
					m_mpFieldVal.Add(sSportTotal, "1");
					m_mpFieldVal.Add(sSportLevelTotal, "1");
					m_mpFieldVal.Add(sSportPos, "1");
					m_mpFieldVal.Add(sTotal, "1");

					// now that we have a dictionary of values, write it out
					bool fFirst = true;
					string sRet = "";

					foreach (string s in plsLegend)
						{
						if (s == "$$$MISC$$$") // expand the Misc values here
							{
							// if there's no misc data for this game, no worries -- misc should
							// be the last entry!
							if (m_plsMisc == null)
								continue;

							foreach (string sMisc in m_plsMisc)
								{
								sRet += ",\"" + sMisc + "\"";
								}
							continue;
							}

						if (fFirst != true)
							{
							sRet += ",";
							}
						fFirst = false;

						if (m_mpFieldVal.ContainsKey(s))
							sRet += m_mpFieldVal[s];
						else
							sRet += "0";
						}
					return sRet;
				}

			}

//          ____ ____ _  _ ____ ____ 
//          | __ |__| |\/| |___ [__  
//          |__] |  | |  | |___ ___] 

			public class Games // GMC
		    {
			    private const int icolGameAway = 18;
			    private const int icolGameHome = 13;
			    private const int icolGameSite = 10;
			    private const int icolGameGame = 0;
                private const int icolOfficial = 4;

			    public class Sport
				{
					SortedList<string, string> m_plLevelPos;
					SortedList<string, string> m_plLevel;
					SortedList<string, string> m_plPos;

					public Sport()
					{
						m_plLevelPos = new SortedList<string, string>();
						m_plLevel = new SortedList<string, string>();
						m_plPos = new SortedList<string, string>();
					}

					/* F  E N S U R E  P O S */
					/*----------------------------------------------------------------------------
						%%Function: FEnsurePos
						%%Qualified: ArbWeb.GenCounts:GenGameStats:Games:Sport.FEnsurePos
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

				List<string> m_plsMiscHeadings;
				SortedList<string, Game> m_plgm;
				StatusBox.StatusRpt m_srpt;
				Dictionary<string, Dictionary<string, int>> m_mpNameSportLevelCount;
				Dictionary<string, Sport> m_mpSportSport;
				List<string> m_plsLegend;

				// NOTE:  This isn't just Team -> Count.  This is also Team-Sport -> Count.
				Dictionary<string, int> m_mpTeamCount;

				public void SetMiscHeadings(List<string> plsMisc)
				{
					m_plsMiscHeadings = plsMisc;
				}

                public Games(StatusBox.StatusRpt srpt)
				{
					m_plgm = new SortedList<string, Game>();
					m_srpt = srpt;
					m_mpSportSport = new Dictionary<string, Sport>();
					m_plsLegend = new List<string>();
					m_mpTeamCount = new Dictionary<string, int>();

					UnitTest();
				}

				/* E N S U R E  S P O R T  L E V E L  P O S */
				/*----------------------------------------------------------------------------
					%%Function: EnsureSportLevelPos
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.EnsureSportLevelPos
					%%Contact: rlittle

					Make sure that the sport / level / pos is in the legend (including all
					the subtotals for the sport/level/pos)
				----------------------------------------------------------------------------*/
				void EnsureSportLevelPos(string sSport, string sLevel, string sPos)
				{
				    Sport sport;
				    bool fNewSport = false;
				    
					// make sure we know about this sport and this position
					if (!m_mpSportSport.ContainsKey(sSport))
						{
						sport = new Sport();
						m_mpSportSport.Add(sSport, sport);
						fNewSport = true;
						}

					sport = m_mpSportSport[sSport];

					bool fNewPos, fNewLevel, fNewLevelPos;

                    sport.EnsurePos(sLevel, sPos, out fNewLevel, out fNewPos, out fNewLevelPos);

					if (fNewLevelPos)
						m_plsLegend.Add(String.Format("{0}-{1}-{2}", sSport, sLevel, sPos));

					if (fNewSport)
						m_plsLegend.Add(String.Format("{0}-Total", sSport));

					if (fNewLevel)
						m_plsLegend.Add(String.Format("{0}-{1}-Total", sSport, sLevel));

					if (fNewPos)
						m_plsLegend.Add(String.Format("{0}-{1}", sSport, sPos));
				}

				/* A D D  G A M E */
				/*----------------------------------------------------------------------------
					%%Function: AddGame
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.AddGame
					%%Contact: rlittle

					Add a game

					This handles ensuring the the sport/level/pos information has been added
					to the legend, so we can build the detail lines later
				----------------------------------------------------------------------------*/
				void AddGame(DateTime dttm, string sSite, string sName, string sTeam, string sEmail, string sGameNum, string sHome, string sAway, string sLevel, string sSport, string sPos, bool fCancelled, List<string> plsMisc)
				{
					Game gm = new Game(dttm, sSite, sName, sTeam, sEmail, sGameNum, sHome, sAway, sLevel, sSport, sPos, fCancelled, plsMisc);
					string sTeamSport = sTeam + "#-#" + sSport;

					m_plgm.Add(String.Format("{0}_{1}_{2}", sName, dttm.ToString("yyyyMMdd:HH:mm"), m_plgm.Count), gm);

					if (sTeam != null && sTeam.Length > 0)
						{
						if (m_mpTeamCount.ContainsKey(sTeam))
							m_mpTeamCount[sTeam]++;
						else
							m_mpTeamCount.Add(sTeam, 1);

						if (m_mpTeamCount.ContainsKey(sTeamSport))
							m_mpTeamCount[sTeamSport]++;
						else
							m_mpTeamCount.Add(sTeamSport, 1);

						EnsureSportLevelPos(sSport, sLevel, sPos);
						}

				}

				string[] SplitTeams(string s)
				{
					string[] rgs = s.Split(';');
					int i;

					for (i  = 0; i < rgs.Length; i++)
						{
						rgs[i] = Regex.Replace(rgs[i], "^ *", "");
						rgs[i] = Regex.Replace(rgs[i], " *$", "");
						}

					return rgs;
				}

				Dictionary<string, DTC> m_mpTeamDtc;

				int TeamCount(string s)
				{
					if (m_mpTeamCount.ContainsKey(s))
						return m_mpTeamCount[s] + TeamDcCount(s);

					return TeamDcCount(s);
				}

				int TeamDcCount(string s)
				{
				    return 0;
#if NO
					if (m_mpTeamDcCount.ContainsKey(s))
						return m_mpTeamDcCount[s];

					return 0;
#endif
				}

#if NO
				void IncTeamDcCount(string s)
				{
					if (m_mpTeamDcCount.ContainsKey(s))
						m_mpTeamDcCount[s]++;

					m_mpTeamDcCount.Add(s, 1);
				}
#endif
				class DTC
				{
					class DND
					{
						string m_sTeam;
						int m_c;
						int m_dc;

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

						public int Count { get { return m_c + m_dc; } }
						public string Name { get { return m_sTeam; } }
						public int DCount { get { return m_dc; } set { m_dc = value; } }
					};

					List<DND> m_pldnd;

					public DTC()
					{
						m_pldnd = new List<DND>();
					}

					public string STeamNext()
					{
						foreach(DND dnd in m_pldnd)
							{
							if (dnd.DCount == 0)
								continue;

							return dnd.Name;
							}
						throw new Exception("could not find team->dtc mapping");
					}

					public void DecTeamNext()
					{
						foreach(DND dnd in m_pldnd)
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
						%%Qualified: ArbWeb.GenCounts:GenGameStats:Games:DTC.Distribute
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
							int cDist = Math.Min((cNext - cMin)*(iMac + 1), c);

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

					public void UpdateTeamTotals(Dictionary<string, int> mpTeamCount, string sSport)
					{
						foreach (DND dnd in m_pldnd)
							Games.UpdateTeamCount(mpTeamCount, dnd.Name, sSport, dnd.DCount);
					}
				}

				/* F  T E A M  M A T C H E S  S P O R T */
				/*----------------------------------------------------------------------------
					%%Function: FTeamMatchesSport
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games:DTC.FTeamMatchesSport
					%%Contact: rlittle

					Determine if the given team name belongs to a given sport (i.e. look
					for things like "Softball Blast" matching "Softball", etc.
				----------------------------------------------------------------------------*/
				bool FTeamMatchesSport(string sTeam, string sSport)
				{
					if (Regex.Match(sSport, ".*Baseball", RegexOptions.IgnoreCase).Success)
						{
						// baseball teams don't have any decoration, or are decorated with baseball
						if (Regex.Match(sTeam, ".*Baseball.*", RegexOptions.IgnoreCase).Success)
							return true;

						if (Regex.Match(sTeam, ".*Softball.*", RegexOptions.IgnoreCase).Success)
							return false;
						
						return true;
						}

					// all other sports should be in the string somewhere
					if (Regex.Match(sTeam, ".*" + sSport + ".*", RegexOptions.IgnoreCase).Success)
						return true;

					return false;
				}

				static public void UpdateTeamCount(Dictionary<string, int> mpTeamCount, string sTeam, string sSport, int dCount)
				{
					string sTeamSport = String.Format("{0}#-#{1}", sTeam, sSport);

					if (!mpTeamCount.ContainsKey(sTeam))
						mpTeamCount.Add(sTeam, dCount);
					else
						mpTeamCount[sTeam] += dCount;

					if (!mpTeamCount.ContainsKey(sTeamSport))
						mpTeamCount.Add(sTeamSport, dCount);
					else
						mpTeamCount[sTeamSport] += dCount;
				}

				/* R E D U C E  T E A M S */
				/*----------------------------------------------------------------------------
					%%Function: ReduceTeams
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.ReduceTeams
					%%Contact: rlittle

					Take "multiple" team allocations and redistribute them to "needier" teams
				----------------------------------------------------------------------------*/
				public void ReduceTeams()
				{
					m_mpTeamDtc = new Dictionary<string, DTC>();

					// look for team multiples

					// TODO: Right now, we are allotting games regardless of sport, so baseball teams are getting
					// credit for softball games, etc.  This is fine, unless there's a softball team that needs
					// credit too...
					foreach(string s in new List<string>(m_mpTeamCount.Keys))
						{
						// skip aggregate sport totals
						if (!Regex.Match(s, ".*#-#.*").Success)
							continue;
						
						int cDist = m_mpTeamCount[s]; // kvp.Value;

//						m_srpt.AddMessage(String.Format("{0,-20}{1}", s, m_mpTeamCount[s]), StatusRpt.MSGT.Body);
						if (s.IndexOf(';') == -1)
							continue;

						// we've got a multiple.  split it up
						DTC dtc = new DTC();
						string []rgs = RexHelper.RgsMatch(s, "(.*)#-#(.*)$");
						string sTeams = rgs[0];
						string sSport = rgs[1];
						string[] rgsTeams = SplitTeams(sTeams);
						bool fIntraSport = false;

						// first, figure out of we're doing an inter or intra sport distribution
						foreach(string sTeam in rgsTeams)
							{
							if (FTeamMatchesSport(sTeam, sSport))
								{
								fIntraSport = true;
								break;
								}
							}

						// if we're intra-sport, then we're only going to consider sport totals
						// when we distribute games around...
						foreach(string sTeam in rgsTeams)
							{
							string sTeamSport = String.Format("{0}#-#{1}", sTeam, sSport);

							if (fIntraSport)
								{
								if (FTeamMatchesSport(sTeam, sSport))
									dtc.AddTeam(sTeam, TeamCount(sTeamSport));
								}
							else
								{
								// no team matched the sport, so we're going to just distribute
								// based on total counts for each team, regardless of sport
								dtc.AddTeam(sTeam, TeamCount(sTeam));
								}
							}

						dtc.Distribute(cDist);
						// ok, at this point, dtc has the list of distribution changes needed to make this work

						dtc.UpdateTeamTotals(m_mpTeamCount, sSport);

						m_mpTeamDtc.Add(s, dtc);
						}
					// now, adjust the raw team counts

					foreach(Game gm in m_plgm.Values)
						{
						if (gm.Open)
							continue;

						if (gm.Team.IndexOf(';') == -1)
							continue;

						string sSportTeam = String.Format("{0}#-#{1}", gm.Team, gm.Sport);

						DTC dtc = m_mpTeamDtc[sSportTeam];

						gm.Team = dtc.STeamNext();
						dtc.DecTeamNext();
						}
				}

				public enum ReadState
				{
					ScanForHeader = 1,
					ScanForGame = 2,
					ReadingGame1 = 3,
					ReadingGame2 = 4,
					ReadingOfficials1 = 5,
					ReadingOfficials2 = 6,
					ReadingComments = 7,
				};

				/* A P P E N D  C H E C K */
				/*----------------------------------------------------------------------------
					%%Function: AppendCheck
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.AppendCheck
					%%Contact: rlittle

					Append s to sAppend -- deals with leading and trailing spaces as well
					as making sure there are spaces separating the arguments
				----------------------------------------------------------------------------*/
				string AppendCheck(string s, string sAppend)
				{
					sAppend = Regex.Replace(sAppend, "^ *", "");
					sAppend = Regex.Replace(sAppend, " *$", "");
					if (sAppend.Length > 0)
						{
						if (s.Length > 1)
							s = s + " " + sAppend;
						else
							s = sAppend;
						}
					return s;
				}

				void UnitTest()
				{
					string s = "", s2 = "", s3 = "";

					s = GenCounts.GenGameStats.Games.ReverseName("Mary Van Whatsa Hoozit");
					Debug.Assert(String.Compare(s, "Van Whatsa Hoozit,Mary") == 0);
					Debug.Assert(RST.FSplitName(s, out s2, out s3));
					Debug.Assert(s2 == "Mary");
					Debug.Assert(s3 == "Van Whatsa Hoozit");

					s =  "Van Whatsa Hoozit, Mary";

					Debug.Assert(RST.FSplitName(s, out s2, out s3));
					Debug.Assert(s2 == "Mary");
					Debug.Assert(s3 == "Van Whatsa Hoozit");

					s = "";
					s = AppendCheck(s, "Foo");
					Debug.Assert(String.Compare(s, "Foo") == 0);
					s = AppendCheck(s, " Bar ");
					Debug.Assert(String.Compare(s, "Foo Bar") == 0);
					s = AppendCheck(s, " Baz ");
					Debug.Assert(String.Compare(s, "Foo Bar Baz") == 0);

					// now test Reduction functionality...

					Dictionary<string, int> mpTeamCountSav = m_mpTeamCount;

					m_mpTeamCount = new Dictionary<string, int>();
						m_mpTeamCount.Add("Swans", 5); 
						m_mpTeamCount.Add("Swans#-#Redmond Baseball", 5); 
						m_mpTeamCount.Add("Crows", 2);
						m_mpTeamCount.Add("Crows#-#Redmond Baseball", 2);
						m_mpTeamCount.Add("Blue Jays", 9);
						m_mpTeamCount.Add("Blue Jays#-#Redmond Baseball", 9);
						m_mpTeamCount.Add("Eagles", 3);
						m_mpTeamCount.Add("Eagles#-#Redmond Baseball", 3);
						m_mpTeamCount.Add("Woodpeckers", 2);
						m_mpTeamCount.Add("Woodpeckers#-#Redmond Baseball", 2);
						m_mpTeamCount.Add("Swans;Eagles", 10);
						m_mpTeamCount.Add("Swans;Eagles#-#Redmond Baseball", 10);
					ReduceTeams();
					Debug.Assert(TeamCount("Swans") == 9);
                    Debug.Assert(TeamCount("Swans#-#Redmond Baseball") == 9);
                    Debug.Assert(TeamCount("Crows") == 2);
                    Debug.Assert(TeamCount("Crows#-#Redmond Baseball") == 2);
					Debug.Assert(TeamCount("Blue Jays") == 9);
                    Debug.Assert(TeamCount("Blue Jays#-#Redmond Baseball") == 9);
					Debug.Assert(TeamCount("Eagles") == 9);
                    Debug.Assert(TeamCount("Eagles#-#Redmond Baseball") == 9);
					Debug.Assert(TeamCount("Woodpeckers") == 2);
                    Debug.Assert(TeamCount("Woodpeckers#-#Redmond Baseball") == 2);

					m_mpTeamCount = new Dictionary<string, int>();
						m_mpTeamCount.Add("Swans", 5); 
						m_mpTeamCount.Add("Swans#-#Baseball", 5); 
						m_mpTeamCount.Add("Crows", 2);
						m_mpTeamCount.Add("Crows#-#Baseball", 2);
						m_mpTeamCount.Add("Blue Jays", 9);
						m_mpTeamCount.Add("Blue Jays#-#Baseball", 9);
						m_mpTeamCount.Add("Eagles", 3);
						m_mpTeamCount.Add("Eagles#-#Baseball", 3);
						m_mpTeamCount.Add("Woodpeckers", 2);
						m_mpTeamCount.Add("Woodpeckers#-#Baseball", 2);
						m_mpTeamCount.Add("Eagles;Woodpeckers;Swans", 14);
						m_mpTeamCount.Add("Eagles;Woodpeckers;Swans#-#Baseball", 14);
					ReduceTeams();
                    Debug.Assert(TeamCount("Swans") == 8);
                    Debug.Assert(TeamCount("Swans#-#Baseball") == 8);
                    Debug.Assert(TeamCount("Crows") == 2);
                    Debug.Assert(TeamCount("Crows#-#Baseball") == 2);
                    Debug.Assert(TeamCount("Blue Jays") == 9);
                    Debug.Assert(TeamCount("Blue Jays#-#Baseball") == 9);
                    Debug.Assert(TeamCount("Eagles") == 8);
                    Debug.Assert(TeamCount("Eagles#-#Baseball") == 8);
                    Debug.Assert(TeamCount("Woodpeckers") == 8);
                    Debug.Assert(TeamCount("Woodpeckers#-#Baseball") == 8);

					m_mpTeamCount = new Dictionary<string, int>();
						m_mpTeamCount.Add("Swans", 5); 
						m_mpTeamCount.Add("Swans#-#Baseball", 5); 
						m_mpTeamCount.Add("Crows", 2);
						m_mpTeamCount.Add("Crows#-#Baseball", 2);
						m_mpTeamCount.Add("Blue Jays", 9);
						m_mpTeamCount.Add("Blue Jays#-#Baseball", 9);
						m_mpTeamCount.Add("Eagles", 3);
						m_mpTeamCount.Add("Eagles#-#Baseball", 3);
						m_mpTeamCount.Add("Woodpeckers", 2);
						m_mpTeamCount.Add("Woodpeckers#-#Baseball", 2);
						m_mpTeamCount.Add("Eagles;Blue Jays;Swans", 5);
						m_mpTeamCount.Add("Eagles;Blue Jays;Swans#-#Baseball", 5);
					ReduceTeams();

                    Debug.Assert(TeamCount("Swans#-#Baseball") == 6);
					Debug.Assert(TeamCount("Swans") == 6);
                    Debug.Assert(TeamCount("Crows#-#Baseball") == 2);
                    Debug.Assert(TeamCount("Crows") == 2);
                    Debug.Assert(TeamCount("Blue Jays#-#Baseball") == 9);
                    Debug.Assert(TeamCount("Blue Jays") == 9);
                    Debug.Assert(TeamCount("Eagles#-#Baseball") == 7);
                    Debug.Assert(TeamCount("Eagles") == 7);
                    Debug.Assert(TeamCount("Woodpeckers#-#Baseball") == 2);
                    Debug.Assert(TeamCount("Woodpeckers") == 2);


					m_mpTeamCount = new Dictionary<string, int>();
						m_mpTeamCount.Add("Swans", 5);
						m_mpTeamCount.Add("Swans#-#Baseball", 5);
						m_mpTeamCount.Add("Crows", 2);
						m_mpTeamCount.Add("Crows#-#Baseball", 2);
						m_mpTeamCount.Add("Blue Jays", 9);
						m_mpTeamCount.Add("Blue Jays#-#Baseball", 9);
						m_mpTeamCount.Add("Eagles", 3);
						m_mpTeamCount.Add("Eagles#-#Baseball", 3);
						m_mpTeamCount.Add("Woodpeckers", 2);
						m_mpTeamCount.Add("Woodpeckers#-#Baseball", 2);
						m_mpTeamCount.Add("Swans;Eagles", 10);
						m_mpTeamCount.Add("Swans;Eagles#-#Baseball", 10);
						m_mpTeamCount.Add("Eagles;Blue Jays;Swans", 5);
						m_mpTeamCount.Add("Eagles;Blue Jays;Swans#-#Baseball", 5);
						m_mpTeamCount.Add("Eagles;Woodpeckers;Swans", 14);
						m_mpTeamCount.Add("Eagles;Woodpeckers;Swans#-#Baseball", 14);
					ReduceTeams();

                    Debug.Assert(TeamCount("Swans") == 12);
                    Debug.Assert(TeamCount("Swans#-#Baseball") == 12);
                    Debug.Assert(TeamCount("Crows") == 2);
                    Debug.Assert(TeamCount("Crows#-#Baseball") == 2);
                    Debug.Assert(TeamCount("Blue Jays") == 11);
                    Debug.Assert(TeamCount("Blue Jays#-#Baseball") == 11);
                    Debug.Assert(TeamCount("Eagles") == 12);
                    Debug.Assert(TeamCount("Eagles#-#Baseball") == 12);
                    Debug.Assert(TeamCount("Woodpeckers") == 13);
                    Debug.Assert(TeamCount("Woodpeckers#-#Baseball") == 13);

					m_mpTeamCount = new Dictionary<string, int>();
						m_mpTeamCount.Add("Swans", 5);
						m_mpTeamCount.Add("Swans#-#Baseball", 5);
						m_mpTeamCount.Add("Softball Eagles", 9);
						m_mpTeamCount.Add("Softball Eagles#-#Softball", 5);
						m_mpTeamCount.Add("Softball Eagles#-#Baseball", 4);
						m_mpTeamCount.Add("Swans;Softball Eagles", 8);
						m_mpTeamCount.Add("Swans;Softball Eagles#-#Baseball", 3);
						m_mpTeamCount.Add("Swans;Softball Eagles#-#Softball", 5);
					ReduceTeams();
					Debug.Assert(TeamCount("Swans") == 8);
					Debug.Assert(TeamCount("Swans#-#Baseball") == 8);
					Debug.Assert(TeamCount("Softball Eagles") == 14);
					Debug.Assert(TeamCount("Softball Eagles#-#Baseball") == 4);
					Debug.Assert(TeamCount("Softball Eagles#-#Softball") == 10);

					m_mpTeamCount = new Dictionary<string, int>();
						m_mpTeamCount.Add("Swans", 5);
						m_mpTeamCount.Add("Swans#-#Baseball", 5);
						m_mpTeamCount.Add("Gulls", 3);
						m_mpTeamCount.Add("Gulls#-#Baseball", 3);
						m_mpTeamCount.Add("Softball Eagles", 9);
						m_mpTeamCount.Add("Softball Eagles#-#Softball", 5);
						m_mpTeamCount.Add("Softball Eagles#-#Baseball", 4);
						m_mpTeamCount.Add("Swans;Softball Eagles", 8);
						m_mpTeamCount.Add("Swans;Softball Eagles#-#Baseball", 3);
						m_mpTeamCount.Add("Swans;Softball Eagles#-#Softball", 5);
						m_mpTeamCount.Add("Swans;Gulls#-#Softball", 10);	// softball to get distributed to non-softball teams
					ReduceTeams();
					Debug.Assert(TeamCount("Swans") == 10);
					Debug.Assert(TeamCount("Swans#-#Baseball") == 8);
					Debug.Assert(TeamCount("Swans#-#Softball") == 2);
					Debug.Assert(TeamCount("Gulls") == 11);
					Debug.Assert(TeamCount("Gulls#-#Baseball") == 3);
					Debug.Assert(TeamCount("Gulls#-#Softball") == 8);
					Debug.Assert(TeamCount("Softball Eagles") == 14);
					Debug.Assert(TeamCount("Softball Eagles#-#Baseball") == 4);
					Debug.Assert(TeamCount("Softball Eagles#-#Softball") == 10);
					m_mpTeamCount = mpTeamCountSav;

				}

					
				/* R E V E R S E  N A M E */
				/*----------------------------------------------------------------------------
					%%Function: ReverseName
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.ReverseName
					%%Contact: rlittle

					Reverse the given "First Last" into "Last,First"

					Handles things like "van Doren, Martin"
				----------------------------------------------------------------------------*/
				static string ReverseName(string s)
				{
					string[] rgs;

					rgs = RexHelper.RgsMatch(s, "^[ \t]*([^ \t]*) ([^\t]*) *$");
					if (rgs.Length < 2)
						return s;

					if (rgs[0] == null || rgs[1] == null)
					    return s;
					return String.Format("{0},{1}", rgs[1], rgs[0]);
				}


				/* G E N  G A M E S  R E P O R T */
				/*----------------------------------------------------------------------------
					%%Function: GenGamesReport
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.GenGamesReport
					%%Contact: rlittle

					Take the accumulated game data and generate a report of the games
					that Arbiter knows about.  suitable for comparing					
				----------------------------------------------------------------------------*/
				public void GenGamesReport(string sReport)
				{
					StreamWriter sw = new StreamWriter(sReport, false, System.Text.Encoding.Default);
					List <string> plsLegend = new List<string>();

					plsLegend.Insert(0, "Game");
					plsLegend.Insert(1, "Date");
					plsLegend.Insert(2, "Time");
                    plsLegend.Insert(3, "Site");
                    plsLegend.Insert(4, "Level");
					plsLegend.Insert(5, "Home");
					plsLegend.Insert(6, "Away");
					plsLegend.Insert(7, "Sport");

                    bool fFirst = true;
					foreach(string s in plsLegend)
						{
						if (!fFirst)
							{
							sw.Write(",");
							}

						fFirst = false;
						sw.Write(s);
						}
					sw.WriteLine();

					Dictionary<string, bool> mpGame = new Dictionary<string, bool>();

					foreach(Game gm in m_plgm.Values)
						{
						// only report each game once...
						if (mpGame.ContainsKey(gm.GameNum))
							continue;

						mpGame.Add(gm.GameNum, true);
						    
						// for each game, report the information, using Legend as the sort order for everything
						sw.WriteLine(gm.SReport(plsLegend));
						}
					sw.Close();
				}

				/* G E N  R E P O R T */
				/*----------------------------------------------------------------------------
					%%Function: GenReport
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.GenReport
					%%Contact: rlittle

					Take the accumulated game data and generate an analysis report
				----------------------------------------------------------------------------*/
				public void GenReport(string sReport)
				{
					StreamWriter sw = new StreamWriter(sReport, false, System.Text.Encoding.Default);
					sw.WriteLine("Analysis (with detail) of Umpire Assignments");
					sw.WriteLine("--------------------------------------------");

					bool fFirst = true;

					m_plsLegend.Sort();
					m_plsLegend.Insert(0, "UmpireName");
					m_plsLegend.Insert(1, "Team");
					m_plsLegend.Insert(2, "Email");
					m_plsLegend.Insert(3, "Game");
					m_plsLegend.Insert(4, "DateTime");
					m_plsLegend.Insert(5, "Level");
					m_plsLegend.Insert(6, "Home");
					m_plsLegend.Insert(7, "Away");
					m_plsLegend.Insert(8, "Description");
					m_plsLegend.Insert(9, "Cancelled");
					m_plsLegend.Insert(10, "Total");
					m_plsLegend.Add("$$$MISC$$$");	// want this at the end!

					foreach(string s in m_plsLegend)
						{
						if (s == "$$$MISC$$$")
							{
							foreach (string sMisc in m_plsMiscHeadings)
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

					foreach(Game gm in m_plgm.Values)
						{
//						if (gm.Open)
//							continue;

						// for each game, report the information, using Legend as the sort order for everything
						sw.WriteLine(gm.SReport(m_plsLegend));
						}
					sw.Close();
				}


				class SlotCount // SC
				{
					DateTime m_dttmSlot;
					Dictionary<string, int> m_mpSportCount;
					Dictionary<string, int> m_mpSportLevelCount;
					Dictionary<string, int> m_mpSiteCount;

					public SlotCount(Game gm)
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

					public void AddSlot(Game gm)
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

					public SlotCount Merge(SlotCount sc)
					{
						SlotCount scNew = new SlotCount();

						scNew.m_dttmSlot = m_dttmSlot;
						scNew.m_mpSportCount = new Dictionary<string, int>();
						scNew.m_mpSportLevelCount = new Dictionary<string, int>();
						scNew.m_mpSiteCount = new Dictionary<string, int>();

						foreach(string sSport in m_mpSportCount.Keys)
							{
							int c = m_mpSportCount[sSport];

							if (sc.m_mpSportCount.ContainsKey(sSport))
								c += sc.m_mpSportCount[sSport];
							scNew.m_mpSportCount.Add(sSport, c);
							}

						foreach(string sSport in sc.m_mpSportCount.Keys)
							{
							int c = sc.m_mpSportCount[sSport];

							if (!m_mpSportCount.ContainsKey(sSport))
								scNew.m_mpSportCount.Add(sSport, c);
							}

						foreach(string sSport in m_mpSportLevelCount.Keys)
							{
							int c = m_mpSportLevelCount[sSport];

							if (sc.m_mpSportLevelCount.ContainsKey(sSport))
								c += sc.m_mpSportLevelCount[sSport];
							scNew.m_mpSportLevelCount.Add(sSport, c);
							}

						foreach(string sSport in sc.m_mpSportLevelCount.Keys)
							{
							int c = sc.m_mpSportLevelCount[sSport];

							if (!m_mpSportLevelCount.ContainsKey(sSport))
								scNew.m_mpSportLevelCount.Add(sSport, c);
							}

						foreach(string sSport in m_mpSiteCount.Keys)
							{
							int c = m_mpSiteCount[sSport];

							if (sc.m_mpSiteCount.ContainsKey(sSport))
								c += sc.m_mpSiteCount[sSport];
							scNew.m_mpSiteCount.Add(sSport, c);
							}

						foreach(string sSport in sc.m_mpSiteCount.Keys)
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

				OpenSlots m_os;
			    private int icolGameLevel;
			    private int icolGameDateTime;

			    class OpenSlots
				{
					SortedList<DateTime, SlotCount> m_mpSlotSc;
					SortedList<string, Game> m_plgm;
					List<string> m_plsSportLevels;
					List<string> m_plsSports;
					List<string> m_plsSites;

					DateTime m_dttmStart;
					DateTime m_dttmEnd;


					public DateTime DttmStart { get { return m_dttmStart; } }
					public DateTime DttmEnd   { get { return m_dttmEnd; } }

					/* O P E N  S L O T S */
					/*----------------------------------------------------------------------------
						%%Function: OpenSlots
						%%Qualified: ArbWeb.GenCounts:GenGameStats:Games:OpenSlots.OpenSlots
						%%Contact: rlittle

					----------------------------------------------------------------------------*/
					public OpenSlots()
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
							string [] rgs = new string[m_plsSites.Count];
							m_plsSites.CopyTo(rgs, 0);
							return rgs;
							}
					}

					/* G E N */
					/*----------------------------------------------------------------------------
						%%Function: Gen
						%%Qualified: ArbWeb.GenCounts:GenGameStats:Games:OpenSlots.Gen
						%%Contact: rlittle

					----------------------------------------------------------------------------*/
					public static OpenSlots Gen(SortedList<string, Game> plgm, DateTime dttmStart, DateTime dttmEnd, string[] rgsSportFilter, string[] rgsSportLevelFilter)
					{
						OpenSlots os = new OpenSlots();

						os.m_plgm = plgm;
						os.m_dttmStart = dttmStart;
						os.m_dttmEnd = dttmEnd;
						os.m_mpSlotSc = new SortedList<DateTime, SlotCount>();
						os.m_plsSports = new List<string>();
                        os.m_plsSportLevels = new List<string>();
						os.m_plsSites = new List<string>();

						foreach(Game gm in plgm.Values)
							{
							if (!gm.Open)
								continue;

							if (DateTime.Compare(gm.Dttm, dttmStart) < 0 || DateTime.Compare(gm.Dttm, dttmEnd) > 0)
								continue;

							if (rgsSportFilter != null)
								{
								bool fMatch = false;

								foreach(string s in rgsSportFilter)
									if (String.Compare(gm.Sport, s, true) == 0)
										fMatch = true;

								if (fMatch == false)
									continue;
								}


							if (rgsSportLevelFilter != null)
								{
								bool fMatch = false;

								foreach(string s in rgsSportLevelFilter)
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
						%%Qualified: ArbWeb.GenCounts:GenGameStats:Games:OpenSlots.GenReport
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

								mpFilter = PlsUniqueFromRgs(rgsSportLevelFilter);
								}
							else
								{
								plsUse = m_plsSports;
								plsCategory = null;
								mpFilter = PlsUniqueFromRgs(rgsSportFilter);
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
						foreach(string sSport in plsUse)
							{
							if (mpFilter != null && !mpFilter.ContainsKey(sSport))
								continue;

							if (plsCategory != null)
								{
								string sCat = null;

								// let's figure out which category we belong in
								foreach(string sCatT in plsCategory)
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

							foreach(string sCatT in plsCategory)
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
							foreach(string sSport in plsUse)
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
								foreach(string sSport in plsUse)
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

					public class HtmlTableReport
					{
						List<string> m_rgsCols = new List<string>();
						SortedList<string, string> m_rgsAxis1 = new SortedList<string, string>();
						SortedList<string, string> m_rgsAxis2 = new SortedList<string, string>();
						Dictionary<string, string> m_mpAxisValues = new Dictionary<string, string>();
						string m_sAxis1Title;
						string m_sAxis2Title;

						public HtmlTableReport() {}

						public string Axis1Title { get { return m_sAxis1Title; } set { m_sAxis1Title = value; } }
						public string Axis2Title { get { return m_sAxis2Title; } set { m_sAxis2Title = value; } }

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
						%%Qualified: ArbWeb.GenCounts:GenGameStats:Games:OpenSlots.GenReportBySite
						%%Contact: rlittle

					----------------------------------------------------------------------------*/
					public void GenReportBySite(string sReport, bool fFuzzyTimes, bool fDatePivot, string[] rgsSportFilter, string[] rgsSportLevelFilter)
					{
						OpenSlots os = OpenSlots.Gen(m_plgm, m_dttmStart, m_dttmEnd, rgsSportFilter, rgsSportLevelFilter);

						StreamWriter sw = new StreamWriter(sReport, false, System.Text.Encoding.Default);
						List<string> plsUse;
						SortedList<string, int> mpFilter = null;
						HtmlTableReport htr = new HtmlTableReport();
						Dictionary<string, int> mpSiteCount = new Dictionary<string, int>();
						int cTotalTotal = 0;

						htr.Axis1Title = "Date";
						htr.Axis2Title = "Site";

						plsUse = new List<string>();

						foreach(string sSite in os.m_plsSites)
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
							foreach(string sSite in plsUse)
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


				/* G E N  O P E N  S L O T S  R E P O R T */
				/*----------------------------------------------------------------------------
					%%Function: GenOpenSlotsReport
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.GenOpenSlotsReport
					%%Contact: rlittle

				----------------------------------------------------------------------------*/
				public void GenOpenSlotsReport(string sReport, bool fDetail, bool fFuzzyTimes, bool fDatePivot, string []rgsSportFilter, string []rgsSportLevelFilter)
				{
					if (fDetail)
						GenOpenSlotsReportDetail(sReport, rgsSportFilter, rgsSportLevelFilter);
					else
						GenOpenSlotsReport(sReport, rgsSportFilter, rgsSportLevelFilter, fFuzzyTimes, fDatePivot);
				}

				/* P L S  U N I Q U E  F R O M  R G S */
				/*----------------------------------------------------------------------------
					%%Function: PlsUniqueFromRgs
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.PlsUniqueFromRgs
					%%Contact: rlittle

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

				/* G E N  O P E N  S L O T S  R E P O R T  D E T A I L */
				/*----------------------------------------------------------------------------
					%%Function: GenOpenSlotsReportDetail
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.GenOpenSlotsReportDetail
					%%Contact: rlittle

					Generate a report of open slots
				----------------------------------------------------------------------------*/
				private void GenOpenSlotsReportDetail(string sReport, string []rgsSportFilter, string []rgsSportLevelFilter)
				{
					DateTime dttmStart = m_os.DttmStart;
					DateTime dttmEnd = m_os.DttmEnd;

					StreamWriter sw = new StreamWriter(sReport, false, System.Text.Encoding.Default);

					string sFormat = "<tr><td>{0}<td>{1}<td>{2}<td>{3}<td>{4}<td>{5}<td>{6}<td>{7}</tr>";

					sw.WriteLine("<html><body><table>");
					sw.WriteLine(String.Format(sFormat, "Game", "Date", "Time", "Field", "Level", "Home", "Away", "Slots"));

					SortedList<string, int> plsSports = PlsUniqueFromRgs(rgsSportFilter);
					SortedList<string, int> plsLevels = PlsUniqueFromRgs(rgsSportLevelFilter);

					foreach(Game gm in m_plgm.Values)
						{
						if (!gm.Open)
							continue;

						if (DateTime.Compare(gm.Dttm, dttmStart) < 0 || DateTime.Compare(gm.Dttm, dttmEnd) > 0)
							continue;

						if (plsSports != null && !(plsSports.ContainsKey(gm.Sport)))
							continue;

						if (plsLevels != null && !(plsLevels.ContainsKey(gm.SportLevel)))
							continue;

						sw.WriteLine(String.Format(sFormat, gm.GameNum, gm.Dttm.ToString("MM/dd/yy ddd"), gm.Dttm.ToString("hh:mm tt"), gm.Site, String.Format("{0}, {1}", gm.Sport, gm.Level), gm.Home, gm.Away, gm.Pos));
						}
					sw.Close();
				}

				/* G E N  O P E N  S L O T S */
				/*----------------------------------------------------------------------------
					%%Function: GenOpenSlots
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.GenOpenSlots
					%%Contact: rlittle

				----------------------------------------------------------------------------*/
				public void GenOpenSlots(DateTime dttmStart, DateTime dttmEnd)
				{
					m_os = OpenSlots.Gen(m_plgm, dttmStart, dttmEnd, null, null);
				}

				/* G E N  O P E N  S L O T S  R E P O R T */
				/*----------------------------------------------------------------------------
					%%Function: GenOpenSlotsReport
					%%Qualified: ArbWeb.GenCounts:GenGameStats:Games.GenOpenSlotsReport
					%%Contact: rlittle

				----------------------------------------------------------------------------*/
				private void GenOpenSlotsReport(string sReport, string[] rgsSportsFilter, string[] rgsSportLevelsFilter, bool fFuzzyTimes, bool fDatePivot)
				{
//					m_os.GenReport(sReport, rgsSportsFilter, rgsSportLevelsFilter);
					m_os.GenReportBySite(sReport, fFuzzyTimes, fDatePivot, rgsSportsFilter, rgsSportLevelsFilter);

				}

				public string[] GetOpenSlotSports()
				{
					string[] rgs = new string[m_os.Sports.Length];
					m_os.Sports.CopyTo(rgs, 0);
					return rgs;
				}

				public string[] GetOpenSlotSportLevels()
				{
					string[] rgs = new string[m_os.SportLevels.Length];
					m_os.SportLevels.CopyTo(rgs, 0);
					return rgs;
				}

				/* F  L O A D  G A M E S */
				/*----------------------------------------------------------------------------
					%%Function: FLoadGames
					%%Qualified: GenCount.GenCounts:GenGameStats:Games.FLoadGames
					%%Contact: rlittle

					Loading the games needs a state machine -- this is a multi line report
				----------------------------------------------------------------------------*/
				public bool FLoadGames(string sGamesReport, Roster rst, bool fIncludeCanceled)
				{
					TextReader tr = new StreamReader(sGamesReport);
					string sLine;
					string []rgsFields;
					ReadState rs = ReadState.ScanForHeader;
					bool fCanceled = false;
					bool fOpenSlot = false;
					bool fIgnore = false;
					string sGame = "";
					string sDateTime = "";
					string sSport = "";
					string sLevel = "";
					string sSite = "";
					string sHome = "";
					string sAway = "";
					string sPosLast = "";
					string sNameLast = "";

					Dictionary<string, string> mpNamePos = new Dictionary<string, string>();
					m_mpNameSportLevelCount = new Dictionary<string, Dictionary<string, int>>();
					Umpire ump = null;
					
					while ((sLine = tr.ReadLine()) != null)
						{
						// first, change "foo, bar" into "foo bar" (get rid of quotes and the comma)
						sLine = Regex.Replace(sLine, "\"([^\",]*),([^\",]*)\"", "$1$2");

						icolGameDateTime = 2;
						if (sLine.Length < icolGameDateTime)
							continue;

						Regex rex = new Regex(",");
						rgsFields = rex.Split(sLine);

						// check for rainouts and cancellations
						if (FMatchGameCancelled(sLine))
							{
							fCanceled = true;
							// drop us back to reading officials
							rs = ReadState.ReadingOfficials1;
							continue;
							}
						// look for comments
						if (FMatchGameComment(sLine))
							{
							rs = RsHandleGameComment(rs, sLine);
							continue;
							}

						if (FMatchGameEmpty(sLine))
							continue;

						if (FMatchGameTotalLine(rgsFields))
							{
							// this is the final "total" line.  the only thing that should follow this is
							// the final page break
							// just leave the rs alone for now...
							continue;
							}

						icolGameLevel = 5;
						if (rs == ReadState.ScanForHeader)
							{
							rs = RsHandleScanForHeader(sLine, rgsFields, rs);
							continue;
							}

						if (rs == ReadState.ReadingComments)
							{
							// when reading comments, we can get text in column 1; if the line ends with commas, then this is just
							// a continuation of the comment (also be careful to look for another comment starting right after ours
							// ends
							if (FMatchGameCommentContinuation(sLine))
								{
								continue;
								}

							rs = ReadState.ReadingOfficials1;
							// drop back to reading officials
							}

						if (rs == ReadState.ReadingGame2)
						    rs = RsHandleReadingGame2(rgsFields, ref sGame, ref sDateTime, ref sLevel, ref sSite, ref sHome, ref sAway, rs);

						if (rs == ReadState.ReadingOfficials2)
						    rs = RsHandleReadingGame2(rgsFields, mpNamePos, sNameLast, sPosLast, rs);

						if (rs == ReadState.ReadingOfficials1)
						    rs = RsHandleReadingOfficials1(rst, fIncludeCanceled, sLine, rgsFields, mpNamePos, fCanceled, sSite, sGame,
						                                   sHome, sAway, sLevel, sSport, rs, ref sPosLast, ref sNameLast, ref sDateTime, ref fOpenSlot, ref ump);

						if (FMatchGameArbiterFooter(sLine))
							{
							Debug.Assert(rs == ReadState.ReadingComments || rs == ReadState.ScanForHeader || rs == ReadState.ScanForGame, String.Format("Page break at illegal position: state = {0}", rs));
							rs = ReadState.ScanForHeader;
							continue;
							}

						if (rs == ReadState.ScanForGame)
						    rs = RsHandleScanForGame(ref sGame, mpNamePos, sLine, ref sDateTime, ref sSport, ref sLevel, ref sSite,
						                             ref sHome, ref sAway, ref fCanceled, ref fIgnore, rs);

						if (rs == ReadState.ReadingGame1)
						    rs = RsHandleReadingGame1(ref sGame, rgsFields, ref sDateTime, ref sSport, ref sSite, ref sHome, ref sAway, rs);
						}

					return true;
				}

			    private ReadState RsHandleScanForHeader(string sLine, string[] rgsFields, ReadState rs)
			    {
			        if (Regex.Match(sLine, "Game.*Date.*Sport.*Level").Success == false)
			            return rs;

			        Debug.Assert(Regex.Match(rgsFields[icolGameLevel], "Sport.*Level").Success, "Sport & level not where expected!!");
			        rs = ReadState.ScanForGame;
			        return rs;
			    }

			    private ReadState RsHandleGameComment(ReadState rs, string sLine)
			    {
			        rs = ReadState.ReadingComments;
			        m_srpt.AddMessage("Reading comment: " + sLine);
			        // skip comments from officials
			        return rs;
			    }

			    private ReadState RsHandleReadingGame1(ref string sGame, string[] rgsFields, ref string sDateTime, ref string sSport,
			                                           ref string sSite, ref string sHome, ref string sAway, ReadState rs)
			    {
// reading the first line of the game.  We should always get the sport and the first part of the team names here
			        sGame = AppendCheck(sGame, rgsFields[icolGameGame]);
			        sDateTime = AppendCheck(sDateTime, rgsFields[icolGameDateTime]);
			        sSport = AppendCheck(sSport, rgsFields[icolGameLevel]);
			        sSite = AppendCheck(sSite, rgsFields[icolGameSite]);
			        sHome = AppendCheck(sHome, rgsFields[icolGameHome]);
			        sAway = AppendCheck(sAway, rgsFields[icolGameAway]);

			        rs = ReadState.ReadingGame2;
			        return rs;
			    }

			    private static ReadState RsHandleScanForGame(ref string sGame, Dictionary<string, string> mpNamePos, string sLine, ref string sDateTime,
			                                                 ref string sSport, ref string sLevel, ref string sSite, ref string sHome,
			                                                 ref string sAway, ref bool fCanceled, ref bool fIgnore, ReadState rs)
			    {
			        sGame = "";
			        sDateTime = "";
			        sSport = "";
			        sLevel = "";
			        sSite = "";
			        sHome = "";
			        sAway = "";
			        fCanceled = false;
			        fIgnore = false;
			        mpNamePos.Clear();

			        if (!(Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Baseball").Success
			              || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Interlock").Success
			              || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Tourn").Success
			              || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Softball").Success
			              || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Administrative").Success
			              || Regex.Match(sLine, ", *50/50").Success
			              || Regex.Match(sLine, ",_Events*").Success
			              || Regex.Match(sLine, ",zEvents*").Success
			              || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Training").Success))
			            Debug.Assert(false, String.Format("failed to find game as expected!: {0} ({1}", sLine, rs));
			        rs = ReadState.ReadingGame1;
			        // fallthrough to ReadingGame1
			        return rs;
			    }

			    private static bool FMatchGameArbiterFooter(string sLine)
			    {
			        return Regex.Match(sLine, ".*Created by ArbiterSports").Success;
			    }

			    private ReadState RsHandleReadingOfficials1(Roster rst, bool fIncludeCanceled, string sLine, string[] rgsFields,
			                                                Dictionary<string, string> mpNamePos, bool fCanceled, string sSite, string sGame,
			                                                string sHome, string sAway, string sLevel, string sSport, ReadState rs,
			                                                ref string sPosLast, ref string sNameLast, ref string sDateTime, ref bool fOpenSlot, ref Umpire ump)
			    {
// Games may have multiple officials, so we have to collect up the officials.
			        // we do this in mpNamePos

			        // &&&& TODO: omit dates before here...

			        if (Regex.Match(sLine, "Attached").Success)
			            return rs;

			        
			        if (rgsFields[0].Length < 1)
			            {
			            // look for possible contiuation line; if not there, then we will fall back
			            // to ReadingOfficials1
			            rs = ReadState.ReadingOfficials2;
			            sPosLast = rgsFields[1];
			            sNameLast = rgsFields[icolOfficial];

			            if (Regex.Match(rgsFields[icolOfficial], "_____").Success)
			                {
			                fOpenSlot = true;
			                mpNamePos.Add(String.Format("!!OPEN{0}", mpNamePos.Count), rgsFields[1]);
			                return rs;
			                }
			            else
			                {
			                string sName = ReverseName(rgsFields[icolOfficial]);
			                mpNamePos.Add(sName, rgsFields[1]);
			                return rs;
			                }
			            }
			        // otherwise we're done!!
			        //						m_srpt.AddMessage("recording results...");
			        if (!fCanceled || fIncludeCanceled)
			            {
			            // record our game


			            // we've got all the info for one particular game and its officials.

			            // walk through the officials that we have
			            foreach (string sName in mpNamePos.Keys)
			                {
			                string sPos = mpNamePos[sName];
			                string sEmail;
			                string sTeam;
			                string sNameUse = sName;
			                List<string> plsMisc = null;

			                if (Regex.Match(sName, "!!OPEN.*").Success)
			                    {
			                    sNameUse = null;
			                    sEmail = "";
			                    sTeam = "";
			                    }
			                else
			                    {
			                    ump = rst.UmpireLookup(sName);

			                    if (ump == null)
			                        {
			                        m_srpt.AddMessage(String.Format("Cannot find info for Umpire: {0}", sName),
			                                          StatusBox.StatusRpt.MSGT.Error);
			                        sEmail = "";
			                        sTeam = "";
			                        }
			                    else
			                        {
			                        sEmail = ump.Contact;
			                        sTeam = ump.Misc;
			                        plsMisc = ump.PlsMisc;
			                        }
			                    }
			                if (sPos != "Training")
			                    {
			                    if (sDateTime.EndsWith("TBA"))
			                        sDateTime = sDateTime.Substring(0, sDateTime.Length - icolOfficial) + "00:00";
			                    AddGame(DateTime.Parse(sDateTime), sSite, sNameUse, sTeam, sEmail, sGame, sHome, sAway, sLevel, sSport,
			                            sPos, fCanceled, plsMisc);
			                    }
			                }
			            }
			        rs = ReadState.ScanForGame;
			        return rs;
			    }

			    private static ReadState RsHandleReadingGame2(string[] rgsFields, Dictionary<string, string> mpNamePos, string sNameLast,
			                                                  string sPosLast, ReadState rs)
			    {
// we are reading the subsequent game lines.  these are not guaranteed to be there (it depends on field
			        // overflows
			        if (rgsFields[1].Length <= 1 && rgsFields[0].Length <= 1 && rgsFields[3].Length > 1)
			            {
			            // nothing in that column means we have a continuation.  now lets concatenate all our stuff
			            mpNamePos.Remove(ReverseName(sNameLast));
			            string sName = String.Format("{0} {1}", sNameLast, rgsFields[3]);
			            sName = ReverseName(sName);
			            mpNamePos.Add(sName, sPosLast);
			            return rs;
			            }
			        rs = ReadState.ReadingOfficials1;
			        // fallthrough to reading officials
			        return rs;
			    }

			    private ReadState RsHandleReadingGame2(string[] rgsFields, ref string sGame, ref string sDateTime, ref string sLevel,
			                                           ref string sSite, ref string sHome, ref string sAway, ReadState rs)
			    {
// we are reading the subsequent game lines.  these are not guaranteed to be there (it depends on field
			        // overflows
			        if (rgsFields[1].Length <= 1 && rgsFields[icolGameGame].Length <= 1)
			            {
			            // nothing in that column means we have a continuation.  now lets concatenate all our stuff
			            sGame = AppendCheck(sGame, rgsFields[icolGameGame]);
			            sDateTime = AppendCheck(sDateTime, rgsFields[icolGameDateTime]);
			            sLevel = AppendCheck(sLevel, rgsFields[icolGameLevel]);
			            sSite = AppendCheck(sSite, rgsFields[icolGameSite]);
			            sHome = AppendCheck(sHome, rgsFields[icolGameHome]);
			            sAway = AppendCheck(sAway, rgsFields[icolGameAway]);
			            return rs;
			            }
			        rs = ReadState.ReadingOfficials1;
			        // fallthrough to reading officials
			        return rs;
			    }

			    private static bool FMatchGameCommentContinuation(string sLine)
			    {
			        return Regex.Match(sLine, ",,,,,,,,,,,,,,,,,$").Success
			               && !Regex.Match(sLine, "^\\*\\*\\*").Success;
			    }

			    private static bool FMatchGameTotalLine(string[] rgsFields)
			    {
			        return Regex.Match(rgsFields[13], "Total:").Success;
			    }

			    private static bool FMatchGameEmpty(string sLine)
			    {
			        return Regex.Match(sLine, "^,,,,,,,,,,,,,,,,,,").Success
			               || Regex.Match(sLine, "^,,,,,,,,,,,,,,,,,").Success;
			    }

			    private static bool FMatchGameComment(string sLine)
			    {
			        return Regex.Match(sLine, "^\"*\\[.*by.*\\]").Success
			               || Regex.Match(sLine, "^[ \t]*\\[.*/.*/.*by.*\\]").Success;
			    }

			    private static bool FMatchGameCancelled(string sLine)
			    {
			        return Regex.Match(sLine, ".*\\*\\*\\*.*CANCEL*ED").Success
			               || Regex.Match(sLine, ".*\\*\\*\\*.*FORFEITED").Success
			               || Regex.Match(sLine, ".*\\*\\*\\*.*RAINED OUT").Success
			               || Regex.Match(sLine, ".*\\*\\*\\*.*SUSPEND*ED").Success;
			    }
		    }

			Roster m_rst;
			Games m_gms;
            StatusBox.StatusRpt m_srpt;

			public string SMiscHeader(int i)
			{
				return m_rst.SMiscHeader(i);
			}

			/* L O A D  R O S T E R */
			/*----------------------------------------------------------------------------
				%%Function: LoadRoster
				%%Qualified: GenCount.GenCounts:GenGameStats.LoadRoster
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
			public bool FLoadRoster(string sRoster, int iMiscAffiliation)
			{
				m_rst = new Roster(m_srpt);
				bool f = m_rst.LoadRoster(sRoster, iMiscAffiliation);

				if (f)
					{
					m_gms.SetMiscHeadings(m_rst.PlsMiscHeadings);
					}
                return f;
			}

			/* F  L O A D  G A M E S */
			/*----------------------------------------------------------------------------
				%%Function: FLoadGames
				%%Qualified: ArbWeb.GenCounts:GenGameStats.FLoadGames
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
			public bool FLoadGames(string sGames, bool fIncludeCanceled)
			{
				return m_gms.FLoadGames(sGames, m_rst, fIncludeCanceled);
			}

			/* G E N  R E P O R T */
			/*----------------------------------------------------------------------------
				%%Function: GenReport
				%%Qualified: ArbWeb.GenCounts:GenGameStats.GenReport
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
			public void GenReport(string sReport)
			{
				m_gms.ReduceTeams();
				m_gms.GenReport(sReport);
			}

			public void GenGamesReport(string sReport)
			{
				m_gms.GenGamesReport(sReport);
			}

			/* G E N  O P E N  S L O T S */
			/*----------------------------------------------------------------------------
				%%Function: GenOpenSlots
				%%Qualified: ArbWeb.GenCounts:GenGameStats.GenOpenSlots
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
			public void GenOpenSlots(DateTime dttmStart, DateTime dttmEnd)
			{
				m_gms.GenOpenSlots(dttmStart, dttmEnd);
			}

			/* G E N  O P E N  S L O T S */
			/*----------------------------------------------------------------------------
				%%Function: GenOpenSlots
				%%Qualified: ArbWeb.GenCounts:GenGameStats.GenOpenSlots
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
			public void GenOpenSlotsReport(string sReport, bool fDetail, bool fFuzzyTimes, bool fDatePivot, string[] rgsSportFilter, string[] rgsSportLevelFilter)
			{
				m_gms.GenOpenSlotsReport(sReport, fDetail, fFuzzyTimes, fDatePivot, rgsSportFilter, rgsSportLevelFilter);
			}

			/* G E T  O P E N  S L O T  S P O R T S */
			/*----------------------------------------------------------------------------
				%%Function: GetOpenSlotSports
				%%Qualified: ArbWeb.GenCounts:GenGameStats.GetOpenSlotSports
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
			public string[] GetOpenSlotSports()
			{
				return m_gms.GetOpenSlotSports();
			}

			/* G E T  O P E N  S L O T  S P O R T  L E V E L S */
			/*----------------------------------------------------------------------------
				%%Function: GetOpenSlotSportLevels
				%%Qualified: ArbWeb.GenCounts:GenGameStats.GetOpenSlotSportLevels
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
			public string[] GetOpenSlotSportLevels()
			{
				return m_gms.GetOpenSlotSportLevels();
			}

			/* G E N  G A M E  S T A T S */
			/*----------------------------------------------------------------------------
				%%Function: GenGameStats
				%%Qualified: GenCount.GenCounts:GenGameStats.GenGameStats
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
            public GenGameStats(StatusBox.StatusRpt srpt)
			{
				//  m_sRoster = null;
				m_srpt = srpt;
				m_rst = new Roster(srpt);
				m_gms = new Games(srpt);
			}
			} // END  GenGameStats

        StatusBox.StatusRpt m_srpt;

		/* G E N  C O U N T S */
		/*----------------------------------------------------------------------------
			%%Function: GenCounts
			%%Qualified: GenCount.GenCounts.GenCounts
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
        public GenCounts(StatusBox.StatusRpt srpt)
		{
			m_srpt = srpt;
		}

		GenGameStats m_ggs;

		/* D O  G E N  C O U N T S */
		/*----------------------------------------------------------------------------
			%%Function: DoGenCounts
			%%Qualified: GenCount.GenCounts.DoGenCounts
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		public void DoGenCounts(string sRoster, string sSource, bool fIncludeCanceled, int iMiscAffiliation)
		{
//            srpt.UnitTest();
			m_ggs = new GenGameStats(m_srpt);
			m_ggs.FLoadRoster(sRoster, iMiscAffiliation);
            m_srpt.AddMessage(String.Format("Using plsMisc[{0}] ({1}) for team affiliation", iMiscAffiliation, m_ggs.SMiscHeader(iMiscAffiliation)), StatusBox.StatusRpt.MSGT.Body);
			m_ggs.FLoadGames(sSource, fIncludeCanceled);
			// read in the roster of umpires...
		}

		public void DoGenSlotsReport(string sReport, bool fDetail, bool fFuzzyTimes, bool fDatePivot, string[] rgsSportFilter, string[] rgsSportLevelFilter)
		{
			m_ggs.GenOpenSlotsReport(sReport, fDetail, fFuzzyTimes, fDatePivot, rgsSportFilter, rgsSportLevelFilter);
		}

		public void DoGenOpenSlots(DateTime dttmStart, DateTime dttmEnd)
		{
			m_ggs.GenOpenSlots(dttmStart, dttmEnd);
		}

		public string[] GetOpenSlotSports()
		{
			return m_ggs.GetOpenSlotSports();
		}

		public string[] GetOpenSlotSportLevels()
		{
			return m_ggs.GetOpenSlotSportLevels();
		}

		public void GenReport(string sReport)
		{
			m_ggs.GenReport(sReport);
		}

		public void GenGamesReport(string sReport)
		{
			m_ggs.GenGamesReport(sReport);
		}
		}
	}
