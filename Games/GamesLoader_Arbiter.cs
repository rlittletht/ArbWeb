using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using NUnit.Framework;
using TCore.StatusBox;

namespace ArbWeb.Games
{
	public class GamesLoader_Arbiter
	{
		private const int icolGameLevelDefault = 6;
		private const int icolGameDateTimeDefault = 2;
		private const int icolGameHomeColumnDeltaFromSportLevel = 14 - icolGameLevelDefault;
		private const int icolGameSiteDeltaFromSportLevel = 11 - icolGameLevelDefault;
		private const int icolGameAwayDeltaFromSportLevel = 19 - icolGameLevelDefault;
		private const int icolOfficialDeltaFromSportLevel = 5 - icolGameLevelDefault; // yes this is a negative delta
		private const int icolSlotStatusDeltaFromSportLevel = 17 - icolGameLevelDefault; // THIS IS UNVERIFIED!
		private const int icolGameGameBase = 0;

		private int GameGameColumn => icolGameGameBase;
		private int GameSiteColumn => icolGameSiteDeltaFromSportLevel + GameLevelColumn;
		private int GameHomeColumn => icolGameHomeColumnDeltaFromSportLevel + GameLevelColumn;
		private int GameAwayColumn => icolGameAwayDeltaFromSportLevel + GameLevelColumn;
		private int OfficialColumn => icolOfficialDeltaFromSportLevel + GameLevelColumn;
		private int SlotStatusColumn => icolSlotStatusDeltaFromSportLevel + GameLevelColumn;

		private int GameLevelColumn { get; set; }
		private int GameDateTimeColumn { get; set; }

		private IStatusReporter m_srpt;
		private ScheduleGames m_gamesBuilding;
	
		public GamesLoader_Arbiter() { } // for unit tests
		public GamesLoader_Arbiter(IStatusReporter srpt)
		{
			m_srpt = srpt;
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


        /* F  L O A D  G A M E S */
        /*----------------------------------------------------------------------------
            %%Function: FLoadGames
            %%Qualified: ArbWeb.CountsData:GameData:Games.FLoadGames
            %%Contact: rlittle

            Loading the games needs a state machine -- this is a multi line report
        ----------------------------------------------------------------------------*/
        public bool FLoadGames(string sGamesReport, Roster rst, bool fIncludeCanceled, ScheduleGames games)
        {
	        m_gamesBuilding = games;
	        
            using (TextReader tr = new StreamReader(sGamesReport))
            {
                string sLine;
                string[] rgsFields;
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
                string sStatusLast = "";

                Dictionary<string, string> mpNameStatus = new Dictionary<string, string>();
                Dictionary<string, string> mpNamePos = new Dictionary<string, string>();
                // m_mpNameSportLevelCount = new Dictionary<string, Dictionary<string, int>>();
                Umpire ump = null;

                while ((sLine = tr.ReadLine()) != null)
                {
                    // first, change "foo, bar" into "foo bar" (get rid of quotes and the comma)
                    sLine = Regex.Replace(sLine, "\"([^\",]*),([^\",]*)\"", "$1$2");

                    if (sLine.Length < GameDateTimeColumn)
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
                        rs = RsHandleReadingGame2(rgsFields, ref sGame, ref sDateTime, ref sLevel, ref sSite,
                            ref sHome, ref sAway, rs);

                    if (rs == ReadState.ReadingOfficials2)
                        rs = RsHandleReadingOfficials2(rst, rgsFields, mpNamePos, mpNameStatus, sNameLast, sPosLast,
                            sStatusLast, rs);

                    if (rs == ReadState.ReadingOfficials1)
                        rs = RsHandleReadingOfficials1(rst, fIncludeCanceled, sLine, rgsFields, mpNamePos,
                            mpNameStatus, fCanceled, sSite, sGame,
                            sHome, sAway, sLevel, sSport, rs, ref sPosLast, ref sNameLast, ref sStatusLast,
                            ref sDateTime, ref fOpenSlot, ref ump);

                    if (FMatchGameArbiterFooter(sLine))
                    {
                        Debug.Assert(
                            rs == ReadState.ReadingComments || rs == ReadState.ScanForHeader ||
                            rs == ReadState.ScanForGame,
                            $"Page break at illegal position: state = {rs}");
                        rs = ReadState.ScanForHeader;
                        continue;
                    }

                    if (rs == ReadState.ScanForGame)
                        rs = RsHandleScanForGame(ref sGame, mpNamePos, mpNameStatus, sLine, ref sDateTime,
                            ref sSport, ref sLevel, ref sSite,
                            ref sHome, ref sAway, ref fCanceled, ref fIgnore, rs);

                    if (rs == ReadState.ReadingGame1)
                        rs = RsHandleReadingGame1(ref sGame, rgsFields, ref sDateTime, ref sSport, ref sSite,
                            ref sHome, ref sAway, rs);
                }
            }

            return true;
        }

        /* R S  H A N D L E  S C A N  F O R  H E A D E R */
        /*----------------------------------------------------------------------------
                %%Function: RsHandleScanForHeader
                %%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleScanForHeader
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private ReadState RsHandleScanForHeader(string sLine, string[] rgsFields, ReadState rs)
        {
            if (Regex.Match(sLine, "Game.*Date.*Sport.*Level").Success == false)
                return rs;

            // always start looking here, and adjust
            GameDateTimeColumn = icolGameDateTimeDefault;
            GameLevelColumn = icolGameLevelDefault;

            if (!Regex.Match(rgsFields[GameLevelColumn], "Sport.*Level").Success)
            {
                // check to see if the previous column is Sport/Level, and if so, automagically adjust
                if (Regex.Match(rgsFields[GameLevelColumn - 1], "Sport.*Level").Success)
                    GameLevelColumn--;
                else if (Regex.Match(rgsFields[GameLevelColumn + 1], "Sport.*Level").Success)
                    GameLevelColumn++;
            }

            if (!Regex.Match(rgsFields[GameDateTimeColumn], "Date.*Time").Success)
            {
                // check to see if the previous column is Sport/Level, and if so, automagically adjust
                if (Regex.Match(rgsFields[GameDateTimeColumn - 1], "Date.*Time").Success)
                    GameDateTimeColumn--;
                else if (Regex.Match(rgsFields[GameDateTimeColumn + 1], "Date.*Time").Success)
                    GameDateTimeColumn++;
            }

            Debug.Assert(Regex.Match(rgsFields[GameDateTimeColumn], "Date.*Time").Success, "Date & time not where expected!!");
            Debug.Assert(Regex.Match(rgsFields[GameLevelColumn], "Sport.*Level").Success, "Sport & level not where expected!!");
            rs = ReadState.ScanForGame;
            return rs;
        }

        /* R S  H A N D L E  G A M E  C O M M E N T */
        /*----------------------------------------------------------------------------
                %%Function: RsHandleGameComment
                %%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleGameComment
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private ReadState RsHandleGameComment(ReadState rs, string sLine)
        {
            rs = ReadState.ReadingComments;
            //			        m_srpt.AddMessage("Reading comment: " + sLine);
            // skip comments from officials
            return rs;
        }

        /* R S  H A N D L E  R E A D I N G  G A M E  1 */
        /*----------------------------------------------------------------------------
                %%Function: RsHandleReadingGame1
                %%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleReadingGame1
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private ReadState RsHandleReadingGame1(ref string sGame, string[] rgsFields, ref string sDateTime, ref string sSport,
            ref string sSite, ref string sHome, ref string sAway, ReadState rs)
        {
            // reading the first line of the game.  We should always get the sport and the first part of the team names here
            sGame = AppendCheck(sGame, rgsFields[GameGameColumn]);
            sDateTime = AppendCheck(sDateTime, rgsFields[GameDateTimeColumn]);
            sSport = AppendCheck(sSport, rgsFields[GameLevelColumn]);
            sSite = AppendCheck(sSite, rgsFields[GameSiteColumn]);
            sHome = AppendCheck(sHome, rgsFields[GameHomeColumn]);
            sAway = AppendCheck(sAway, rgsFields[GameAwayColumn]);

            rs = ReadState.ReadingGame2;
            return rs;
        }

        /* A P P E N D  C H E C K */
        /*----------------------------------------------------------------------------
            %%Function: AppendCheck
            %%Qualified: ArbWeb.CountsData:GameData:Games.AppendCheck
            %%Contact: rlittle

            Append s to sAppend -- deals with leading and trailing spaces as well
            as making sure there are spaces separating the arguments
        ----------------------------------------------------------------------------*/
        public static string AppendCheck(string s, string sAppend)
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

        /* R E V E R S E  N A M E  S I M P L E */
        /*----------------------------------------------------------------------------
            %%Function: ReverseNameSimple
            %%Qualified: ArbWeb.GameData.GameSlots.ReverseNameSimple
            %%Contact: rlittle

            Reverse the given "First Last" into "Last,First"

            Handles things like "van Doren, Martin"
        ----------------------------------------------------------------------------*/
        public static string ReverseNameSimple(string s)
        {
            string[] rgs;

            rgs = CountsData.RexHelper.RgsMatch(s, "^[ \t]*([^ \t]*) ([^\t]*) *$");
            if (rgs.Length < 2)
                return s;

            if (rgs[0] == null || rgs[1] == null)
                return s;
            return $"{rgs[1]},{rgs[0]}";
        }

        [Test]
        [TestCase("Rob Little", "Little,Rob")]
        [TestCase("Martin van Doren", "van Doren,Martin")]
        [TestCase("Byron (Barney) Kinzer", "(Barney) Kinzer,Byron")]
        public static void TestReverseNameSimple(string sIn, string sExpected)
        {
            string sActual = ReverseNameSimple(sIn);

            Assert.AreEqual(sExpected, sActual);
        }

        /* R E V E R S E  N A M E  S I M P L E  P A R E N T H E T I C A L */
        /*----------------------------------------------------------------------------
            %%Function: ReverseNameParenthetical
            %%Qualified: ArbWeb.GameData.GameSlots.ReverseNameParenthetical
            %%Contact: rlittle

            Same as ReverseNameSimple, but allows parenthetical first names

            Byron (Barney) Kinzer => "Kinzer,Byron (Barney)"
        ----------------------------------------------------------------------------*/
        public static string ReverseNameParenthetical(string s)
        {
            // first, see if we match a parenthetical
            if (s.IndexOf('(') <= 0)
                return ReverseNameSimple(s);

            // that was just the quick check. now we need to see if there's a
            // parenthetical match between the first and last name
            string[] rgs;

            rgs = CountsData.RexHelper.RgsMatch(s, "^[ \t]*([^ \t]*)([ \t]*\\([^ \t]*\\))[ \t]*([^\t]*) *$");
            if (rgs.Length < 3)
                return ReverseNameSimple(s);

            if (rgs[0] == null || rgs[1] == null || rgs[2] == null)
                return ReverseNameSimple(s);

            return $"{rgs[2]},{rgs[0]}{rgs[1]}";
        }



        /* R E V E R S E  N A M E */
        /*----------------------------------------------------------------------------
            %%Function: ReverseName
            %%Qualified: ArbWeb.CountsData:GameData:Games.ReverseName
            %%Contact: rlittle

            try to figure out the right reversal for the given name. Try each type of
            reversal we know about and see if it matches an entry in the roster. if not, 
            try the next one.  If all else fails, just return the simple reverse
        ----------------------------------------------------------------------------*/
        public static string ReverseName(Roster rst, string s)
        {
	        string sReverse, sFallback;

	        sFallback = sReverse = ReverseNameSimple(s);
	        if (rst == null)
		        return sFallback;

	        if (rst.UmpireLookup(sReverse) != null)
		        return sReverse;

	        sReverse = ReverseNameParenthetical(s);
	        if (rst.UmpireLookup(sReverse) != null)
		        return sReverse;

	        return sFallback;
        }
        
        #region Tests
        [Test]
        [TestCase("Rob Little", "Little,Rob")]
        [TestCase("Martin van Doren", "van Doren,Martin")]
        [TestCase("Byron (Barney) Kinzer", "Kinzer,Byron (Barney)")]
        [TestCase("Byron(Barney) Kinzer", "Kinzer,Byron(Barney)")]
        [TestCase("(Barney)Byron Kinzer", "Kinzer,(Barney)Byron")]
        [TestCase("(Barney) Byron Kinzer", "Byron Kinzer,(Barney)")]
        public static void TestReverseNameParenthetical(string sIn, string sExpected)
        {
            string sActual = ReverseNameParenthetical(sIn);

            Assert.AreEqual(sExpected, sActual);
        }

        #endregion

        /* R S  H A N D L E  S C A N  F O R  G A M E */
        /*----------------------------------------------------------------------------
                %%Function: RsHandleScanForGame
                %%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleScanForGame
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private static ReadState RsHandleScanForGame(ref string sGame, Dictionary<string, string> mpNamePos, Dictionary<string, string> mpNameStatus, string sLine, ref string sDateTime,
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
            mpNameStatus.Clear();

            if (!(Regex.Match(sLine, ", *[ a-zA-Z0-9-/]* *Baseball").Success
                  || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Interlock").Success
                  || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Interlock").Success
                  || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Tourn").Success
                  || Regex.Match(sLine, ", *[ a-zA-Z0-9-/]* *Softball").Success
                  || Regex.Match(sLine, ", *[ a-zA-Z0-9-']* *Postseason").Success
                  || Regex.Match(sLine, ", *[ a-zA-Z0-9-']* *All Stars").Success
                  || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Fall Ball").Success
                  || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Administrative").Success
                  || Regex.Match(sLine, ", *50/50").Success
                  || Regex.Match(sLine, ",_Events*").Success
                  || Regex.Match(sLine, ",zEvents*").Success
                  || Regex.Match(sLine, ", *[ a-zA-Z0-9-]* *Training").Success))
                Debug.Assert(false, $"failed to find game as expected!: {sLine} ({rs}");
            rs = ReadState.ReadingGame1;
            // fallthrough to ReadingGame1
            return rs;
        }

        /* F  M A T C H  G A M E  A R B I T E R  F O O T E R */
        /*----------------------------------------------------------------------------
                %%Function: FMatchGameArbiterFooter
                %%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameArbiterFooter
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private static bool FMatchGameArbiterFooter(string sLine)
        {
            return Regex.Match(sLine, ".*Created by ArbiterSports").Success;
        }

        /* R S  H A N D L E  R E A D I N G  O F F I C I A L S  1 */
        /*----------------------------------------------------------------------------
                %%Function: RsHandleReadingOfficials1
                %%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleReadingOfficials1
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private ReadState RsHandleReadingOfficials1(Roster rst, bool fIncludeCanceled, string sLine, string[] rgsFields,
            Dictionary<string, string> mpNamePos, Dictionary<string, string> mpNameStatus, bool fCanceled, string sSite, string sGame,
            string sHome, string sAway, string sLevel, string sSport, ReadState rs,
            ref string sPosLast, ref string sStatusLast, ref string sNameLast, ref string sDateTime, ref bool fOpenSlot, ref Umpire ump)
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
                sNameLast = rgsFields[OfficialColumn];
                sStatusLast = rgsFields[SlotStatusColumn];

                if (Regex.Match(rgsFields[OfficialColumn], "_____").Success)
                {
                    fOpenSlot = true;
                    mpNamePos.Add($"!!OPEN{mpNamePos.Count}", rgsFields[1]);
                    mpNameStatus.Add($"!!OPEN{mpNameStatus.Count}", rgsFields[SlotStatusColumn]);
                    return rs;
                }
                else
                {
                    string sName = ReverseName(rst, rgsFields[OfficialColumn]);
                    mpNamePos.Add(sName, rgsFields[1]);
                    mpNameStatus.Add(sName, rgsFields[SlotStatusColumn]);
                    return rs;
                }
            }
            // otherwise we're done!!
            //						m_srpt.AddMessage("recording results...");
            if (!fCanceled || fIncludeCanceled)
            {
                // record our game


                // we've got all the info for one particular game and its officials.

                if (mpNamePos.Count == 0)
                {
                    // invent one so we at least record this game
                    mpNamePos.Add("!!OPEN0", "!!FAKE");
                    mpNameStatus.Add("!!OPEN0", "!!FAKE");
                }
                
                // walk through the officials that we have
                foreach (string sName in mpNamePos.Keys)
                {
                    string sPos = mpNamePos[sName];
                    string sStatus = mpNameStatus[sName];
                    string sEmail;
                    string sTeam;
                    string sNameUse = sName;
                    List<string> plsMisc = null;

                    if (Regex.Match(sName, "!!OPEN.*").Success)
                    {
                        sNameUse = null;
                        sEmail = "";
                        sTeam = "";
                        sStatus = "";
                    }
                    else
                    {
                        ump = rst.UmpireLookup(sName);

                        if (ump == null)
                        {
                            if (sName != "")
                                m_srpt.AddMessage(
                                    $"Cannot find info for Umpire: {sName}",
                                                  MSGT.Error);
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
                            sDateTime = sDateTime.Substring(0, sDateTime.Length - 4) + "00:00";
                        m_gamesBuilding.AddGame(DateTime.Parse(sDateTime), sSite, sNameUse, sTeam, sEmail, sGame, sHome, sAway, sLevel, sSport,
                                sPos, sStatus, fCanceled, plsMisc);
                    }
                }
            }
            rs = ReadState.ScanForGame;
            return rs;
        }

        private static bool FEmptyField(string s)
        {
            return String.IsNullOrWhiteSpace(s);
        }
        /* R S  H A N D L E  R E A D I N G  G A M E  2 */
        /*----------------------------------------------------------------------------
                %%Function: RsHandleReadingOfficials2
                %%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleReadingOfficials2
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private static ReadState RsHandleReadingOfficials2(Roster rst, string[] rgsFields, Dictionary<string, string> mpNamePos, Dictionary<string, string> mpNameStatus, string sNameLast,
            string sPosLast, string sStatusLast, ReadState rs)
        {
            // we are reading the subsequent game lines.  these are not guaranteed to be there (it depends on field
            // overflows
            if (FEmptyField(rgsFields[1]) && FEmptyField(rgsFields[0]) && !FEmptyField(rgsFields[3]))
            {
                // nothing in that column means we have a continuation.  now lets concatenate all our stuff
                mpNamePos.Remove(ReverseName(rst, sNameLast));
                mpNameStatus.Remove(ReverseName(rst, sNameLast));
                string sName = $"{sNameLast} {rgsFields[3]}";
                sName = ReverseName(rst, sName);
                mpNamePos.Add(sName, sPosLast);
                mpNameStatus.Add(sName, sStatusLast);
                return rs;
            }
            rs = ReadState.ReadingOfficials1;
            // fallthrough to reading officials
            return rs;
        }

        /* R S  H A N D L E  R E A D I N G  G A M E  2 */
        /*----------------------------------------------------------------------------
                %%Function: RsHandleReadingOfficials2
                %%Qualified: ArbWeb.CountsData:GameData:Games.RsHandleReadingOfficials2
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private ReadState RsHandleReadingGame2(string[] rgsFields, ref string sGame, ref string sDateTime, ref string sLevel,
            ref string sSite, ref string sHome, ref string sAway, ReadState rs)
        {
            // we are reading the subsequent game lines.  these are not guaranteed to be there (it depends on field
            // overflows
            if (FEmptyField(rgsFields[1]) && FEmptyField(rgsFields[GameGameColumn]))
            {
                // nothing in that column means we have a continuation.  now lets concatenate all our stuff
                sGame = AppendCheck(sGame, rgsFields[GameGameColumn]);
                sDateTime = AppendCheck(sDateTime, rgsFields[GameDateTimeColumn]);
                sLevel = AppendCheck(sLevel, rgsFields[GameLevelColumn]);
                sSite = AppendCheck(sSite, rgsFields[GameSiteColumn]);
                sHome = AppendCheck(sHome, rgsFields[GameHomeColumn]);
                sAway = AppendCheck(sAway, rgsFields[GameAwayColumn]);
                return rs;
            }
            rs = ReadState.ReadingOfficials1;
            // fallthrough to reading officials
            return rs;
        }

        /* F  M A T C H  G A M E  C O M M E N T  C O N T I N U A T I O N */
        /*----------------------------------------------------------------------------
                %%Function: FMatchGameCommentContinuation
                %%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameCommentContinuation
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private static bool FMatchGameCommentContinuation(string sLine)
        {
            return Regex.Match(sLine, ",,,,,,,,,,,,,,,,,$").Success
                   && !Regex.Match(sLine, "^\\*\\*\\*").Success;
        }

        /* F  M A T C H  G A M E  T O T A L  L I N E */
        /*----------------------------------------------------------------------------
                %%Function: FMatchGameTotalLine
                %%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameTotalLine
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private static bool FMatchGameTotalLine(string[] rgsFields)
        {
	        return Regex.Match(rgsFields[13], "Total:").Success
	               || Regex.Match(rgsFields[14], "Total:").Success
	               || (rgsFields.Length >= 16 && Regex.Match(rgsFields[15], "Total:").Success);
        }

        /* F  M A T C H  G A M E  E M P T Y */
        /*----------------------------------------------------------------------------
                %%Function: FMatchGameEmpty
                %%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameEmpty
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private static bool FMatchGameEmpty(string sLine)
        {
            return Regex.Match(sLine, "^,,,,,,,,,,,,,,,,,,").Success
                   || Regex.Match(sLine, "^,,,,,,,,,,,,,,,,,").Success;
        }

        /* F  M A T C H  G A M E  C O M M E N T */
        /*----------------------------------------------------------------------------
                %%Function: FMatchGameComment
                %%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameComment
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private static bool FMatchGameComment(string sLine)
        {
            return Regex.Match(sLine, "^\"*\\[.*by.*\\]").Success
                   || Regex.Match(sLine, "^[ \t]*\\[.*/.*/.*by.*\\]").Success
                   || Regex.Match(sLine, "^Comments: ").Success;
        }

        /* F  M A T C H  G A M E  C A N C E L L E D */
        /*----------------------------------------------------------------------------
                %%Function: FMatchGameCancelled
                %%Qualified: ArbWeb.CountsData:GameData:Games.FMatchGameCancelled
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private static bool FMatchGameCancelled(string sLine)
        {
            return Regex.Match(sLine, ".*\\*\\*\\*.*CANCEL*ED").Success
                   || Regex.Match(sLine, ".*\\*\\*\\*.*FORFEITED").Success
                   || Regex.Match(sLine, ".*\\*\\*\\*.*POSTPONED").Success
                   || Regex.Match(sLine, ".*\\*\\*\\*.*RAINED OUT").Success
                   || Regex.Match(sLine, ".*\\*\\*\\*.*SUSPEND*ED").Success;
        }
    }
}