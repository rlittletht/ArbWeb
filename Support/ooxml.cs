using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO.Packaging;
using System.IO;
using ArbWeb.Games;
using NUnit.Framework;

namespace ArbWeb
{
    internal class OOXML
    {
        private const string s_sUriWordDoc = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string s_sUriContentTypeDoc = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        private const string s_sUriContentTypeSettings = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
        private const string s_sUriMailMergeRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource";

        private static void StartElement(XmlTextWriter xw, string sElt)
        {
            xw.WriteStartElement(sElt, s_sUriWordDoc);
        }

        private struct AttrPair
        {
            public string sAttr;
            public string sValue;

            public AttrPair(string sAttrIn, string sValueIn)
            {
                sAttr = sAttrIn;
                sValue = sValueIn;
            }
        };

        static void WriteAttributeString(XmlTextWriter xw, string sAttr, string sValue)
        {
            xw.WriteAttributeString(sAttr, s_sUriWordDoc, sValue);
        }

        private static void EndElement(XmlTextWriter xw)
        {
            xw.WriteEndElement();
        }

        private static void WriteElementFull(XmlTextWriter xw, string sElt, string[] rgsVals)
        {
            StartElement(xw, sElt);
            if (rgsVals != null)
            {
                foreach (string s in rgsVals)
                    WriteAttributeString(xw, "val", s);
            }

            EndElement(xw);
        }

        private static void WriteElementFull(XmlTextWriter xw, string sElt, AttrPair[] rgVals)
        {
            StartElement(xw, sElt);
            if (rgVals != null)
            {
                foreach (AttrPair ap in rgVals)
                    WriteAttributeString(xw, ap.sAttr, ap.sValue);
            }

            EndElement(xw);
        }

        public static bool CreateDocSettings(Package pkg, string sDataSource)
        {
            PackagePart prt;
            Stream stm = StmCreatePart(pkg, "/word/settings.xml", s_sUriContentTypeSettings, out prt);

            XmlTextWriter xw = new XmlTextWriter(stm, System.Text.Encoding.UTF8);

            StartElement(xw, "settings");
            WriteElementFull(xw, "view", new[] { "web" });
            StartElement(xw, "mailMerge");
            WriteElementFull(xw, "mainDocumentType", new[] { "email" });
            WriteElementFull(xw, "linkToQuery", (string[])null);
            WriteElementFull(xw, "dataType", new[] { "textFile" });
            WriteElementFull(xw, "connectString", new[] { "" });
            WriteElementFull(xw, "query", new[] { $"SELECT * FROM {sDataSource}" });

            PackageRelationship rel = prt.CreateRelationship(
                new System.Uri($"file:///{sDataSource}", UriKind.Absolute),
                TargetMode.External,
                s_sUriMailMergeRelType);
            StartElement(xw, "dataSource");
            xw.WriteAttributeString("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", rel.Id);
            EndElement(xw);
            StartElement(xw, "odso");
            StartElement(xw, "fieldMapData");
            WriteElementFull(xw, "type", new[] { "dbColumn" });
            WriteElementFull(xw, "name", new[] { "email" });
            WriteElementFull(xw, "mappedName", new[] { "E-mail Address" });
            WriteElementFull(xw, "column", new[] { "0" });
            WriteElementFull(xw, "lid", new[] { "en-US" });
            EndElement(xw); // fieldMapData
            EndElement(xw); // odso
            EndElement(xw); // mailMerge
            EndElement(xw); // settings
            xw.Flush();
            stm.Flush();
            return true;
        }

        static void WritePara(XmlTextWriter xw, string s)
        {
            StartElement(xw, "p");
            StartElement(xw, "r");
            do
            {
                string sub;
                bool fBreak = false;

                if (s.Contains("<br/>"))
                {
                    sub = s.Substring(0, s.IndexOf("<br/>"));
                    s = s.Substring(s.IndexOf("<br/>") + 5);
                    fBreak = true;
                }
                else
                {
                    sub = s;
                    s = "";
                }

                StartElement(xw, "t");
                xw.WriteString(sub);
                EndElement(xw); // t
                if (fBreak)
                {
                    StartElement(xw, "br");
                    EndElement(xw);
                }
            } while (s.Length > 0);

            EndElement(xw); // r
            EndElement(xw); // p
        }

        static void WriteSingleParaCell(XmlTextWriter xw, string sWidth, string sWidthType, string s)
        {
            StartElement(xw, "tc");
            StartElement(xw, "tcPr");
            WriteElementFull(xw, "tcW", new[] { new AttrPair("w", sWidth), new AttrPair("type", sWidthType) });
            EndElement(xw); // tcPr
            WritePara(xw, s);
            EndElement(xw); // tc
        }

        private static Dictionary<string, string> s_mpSportLevelFriendly = new Dictionary<string, string>()
                                                                           {
                                                                               { "All Stars SB 8/9/10's", "SB 8/9/10s" },
                                                                               { "All Stars SB 9/10/11's", "SB 9/10/11s" },
                                                                               { "All Stars SB Majors", "SB Majors" },
                                                                               { "All Stars SB Juniors", "SB Juniors" },
                                                                               { "All Stars 60' BB 9/10/11's", "BB 9/10/11s" },
                                                                               { "All Stars 60' BB 8/9/10's", "BB 8/9/10s" },
                                                                               { "All Stars 60' BB Majors", "BB Majors" },
                                                                               { "All Stars 90' BB Intermediate 70", "BB Intermediates" },
                                                                               { "All Stars 90' BB Juniors 90", "BB Juniors" },
                                                                               { "All Stars 90' BB Seniors", "BB Seniors" },
                                                                           };

        static string DescribeGames(IReadOnlyCollection<Game> games, int cTotalOpenSlots, int cGamesNoUmpires)
        {
            string sDesc;
            string sCount = "";

            if (cGamesNoUmpires > 0)
                //sDesc = $"{games.Count} games: {cTotalOpenSlots} umpires needed!<br/>{cGamesNoUmpires} GAME{(cGamesNoUmpires == 1 ? "" : "S")} WITH NO UMPIRES";
                sDesc = $"{cTotalOpenSlots} umpires needed!<br/>{cGamesNoUmpires} GAME{(cGamesNoUmpires == 1 ? "" : "S")} WITH NO UMPIRES";
            else
                sDesc = $"{cTotalOpenSlots} umpires needed";

            //sDesc = $"{gm.TotalSlots - gm.OpenSlots} UMPIRE";

            return $"{sCount}{sDesc}";
        }

        static string FriendlySport(Game gm)
        {
            if (s_mpSportLevelFriendly.ContainsKey(gm.Slots[0].SportLevel))
                return s_mpSportLevelFriendly[gm.Slots[0].SportLevel];

            return gm.Slots[0].SportLevel;
        }

        enum SlotsKind
        {
            One,
            MoreThanOne
        }

        /*----------------------------------------------------------------------------
            %%Function: GetGameSummaryFromGames
            %%Qualified: ArbWeb.OOXML.GetGameSummaryFromGames

            Every game must have at least 2 umpires on it (or the total number of
            slots if it only has 1 slot).

            We keep a running total of surplus umpires (which can go negative), and
            the count of games that only need 1 umpire.
        ----------------------------------------------------------------------------*/
        static Tuple<int, int, Game> GetGameSummaryFromGames(IReadOnlyCollection<Game> games)
        {
            int gamesOverOneSlot = 0;
            int totalUmpiresNeeded = 0;
            int totalUmpiresOverOneSlot = 0;

            Game gameFirst = null;

            foreach (Game game in games)
            {
                if (gameFirst == null)
                    gameFirst = game;

                SlotsKind kind = game.TotalSlots < 2 ? SlotsKind.One : SlotsKind.MoreThanOne;

                // we always consider open slots
                totalUmpiresNeeded += game.OpenSlots;

                if (kind != SlotsKind.One)
                {
                    gamesOverOneSlot++;
                    totalUmpiresOverOneSlot += (game.TotalSlots - game.OpenSlots);
                }
            }

            // now apportion the umpires we have
            int totalNeededForMinCoverage = gamesOverOneSlot * 2;
            int gamesWithNoUmpires = (totalNeededForMinCoverage - totalUmpiresOverOneSlot) / 2;
            if (gamesWithNoUmpires < 0)
                gamesWithNoUmpires = 0;

            return new Tuple<int, int, Game>(gamesWithNoUmpires, totalUmpiresNeeded, gameFirst);
        }

        //        [Test]
        //        [TestCase(new string[] { "2|4" })]

        [Test]
        [TestCase(new[] { "2|4" }, 0, 2)]
        [TestCase(new[] { "0|4" }, 1, 4)]
        [TestCase(new[] { "1|4" }, 0, 3)]
        [TestCase(new[] { "3|4" }, 0, 1)]
        [TestCase(new[] { "2|4", "2|4" }, 0, 4)]
        [TestCase(new[] { "3|4", "1|4" }, 0, 4)]
        [TestCase(new[] { "4|4", "0|4" }, 0, 4)]
        [TestCase(new[] { "3|4", "0|4" }, 0, 5)]
        [TestCase(new[] { "2|4", "0|4" }, 1, 6)]
        [TestCase(new[] { "1|4", "1|4", "1|4", "1|4" }, 2, 12)]
        [TestCase(new[] { "1|1", "1|4", "1|4", "1|4", "1|4" }, 2, 12)]
        [TestCase(new[] { "0|1", "1|4", "1|4", "1|4", "1|4" }, 2, 13)]
        public static void TestGetGameSummaryFromGames(string[] openAndTotals, int expectedGamesWithNoUmpires, int expectedTotalNeeded)
        {
            List<Game> games = new List<Game>();

            foreach (string openAndTotal in openAndTotals)
            {
                Game game = new Game();

                string[] nums = openAndTotal.Split('|');

                int open = int.Parse(nums[0]);
                int total = int.Parse(nums[1]);

                while (total-- > 0)
                {
                    GameSlot slot = new GameSlot(DateTime.Now, "", (open-- > 0 ? "Somebody" : null), "", "", "", "", "", "", "", "", "", false, null);

                    game.AddGameSlot(slot);
                }

                games.Add(game);
            }

            (int cGamesWithNoUmpires, int cTotalNeeded, Game gameFirst) = GetGameSummaryFromGames(games);

            Assert.AreEqual(expectedGamesWithNoUmpires, cGamesWithNoUmpires);
            Assert.AreEqual(expectedTotalNeeded, cTotalNeeded);
        }

        static void WriteGames(XmlTextWriter xw, IReadOnlyCollection<Game> games)
        {
            int cGames = games.Count;

            // see how to combine the games...

            (int cGamesWithNoUmpires, int cTotalNeeded, Game gameFirst) = GetGameSummaryFromGames(games);

            StartElement(xw, "tr");
            WriteSingleParaCell(xw, "0", "auto", gameFirst.Slots[0].Dttm.ToString("M/dd"));
            WriteSingleParaCell(xw, "0", "auto", gameFirst.Slots[0].Dttm.ToString("ddd h tt"));

            WriteSingleParaCell(xw, "0", "auto", FriendlySport(gameFirst));
            WriteSingleParaCell(xw, "0", "auto", gameFirst.Slots[0].SiteShort);

            WriteSingleParaCell(xw, "0", "auto", DescribeGames(games, cTotalNeeded, cGamesWithNoUmpires));
            EndElement(xw); // tr
        }

        static void AppendGamesToSb(IReadOnlyCollection<Game> games, StringBuilder sb)
        {
            (int cGamesWithNoUmpires, int cTotalNeeded, Game gameFirst) = GetGameSummaryFromGames(games);

            sb.Append($"<tr><td>{gameFirst.Slots[0].Dttm.ToString("M/dd")}<td>{gameFirst.Slots[0].Dttm.ToString("ddd h tt")}");
            sb.Append(
                $"<td>{FriendlySport(gameFirst)}<td>{gameFirst.Slots[0].SiteShort}<td class='bold'>{DescribeGames(games, cTotalNeeded, cGamesWithNoUmpires)}");
        }

        static void StartTable(XmlTextWriter xw, int cCols)
        {
            StartElement(xw, "tbl");
            StartElement(xw, "tblPr");
            WriteElementFull(xw, "tblStyle", new[] { new AttrPair("val", "TableGrid") });
            WriteElementFull(xw, "tblW", new[] { new AttrPair("w", "0"), new AttrPair("type", "auto") });
            WriteElementFull(
                xw,
                "tblLook",
                new[]
                {
                    new AttrPair("val", "04A0"), new AttrPair("firstRow", "1"), new AttrPair("lastRow", "0"), new AttrPair("firstColumn", "1"),
                    new AttrPair("lastColumn", "0"), new AttrPair("noHBand", "0"), new AttrPair("noVBand", "1")
                });
            EndElement(xw); // tblPr
            StartElement(xw, "tblGrid");
            while (cCols-- > 0)
                WriteElementFull(xw, "gridCol", new[] { new AttrPair("val", "1440") });
            EndElement(xw); // tblGrid
        }

        static void EndTable(XmlTextWriter xw)
        {
            EndElement(xw); // tbl
        }


        /* C R E A T E  M A I N  D O C */
        /*----------------------------------------------------------------------------
            %%Function: CreateMainDoc
            %%Qualified: ArbWeb.OOXML.CreateMainDoc
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public static bool CreateMainDoc(Package pkg, string sDataSource, ScheduleGames gms, out string sArbiterHelpNeeded)
        {
            PackagePart prt;
            StringBuilder sb = new StringBuilder();

            sb.Append(
                "<div id='ArbWebAnnounce_HelpNeeded'>"
                + "<h1>HELP NEEDED</h1>"
                + "<h4> The following upcoming games URGENTLY need help! <br>"
                + "Please <a href=\"https://www1.arbitersports.com/Official/SelfAssign.aspx\">SELF ASSIGN</a> now!</h4>"
                + "<style> "
                + "table.help td {padding-left: 2mm;padding-right: 2mm;}"
                + "td.bold {font-weight: bold;}"
                + "</style> "
                + "<table class='help' border=1 style='border-collapse: collapse'>");

            Stream stm = StmCreatePart(pkg, "/word/document.xml", s_sUriContentTypeDoc, out prt);

            XmlTextWriter xw = new XmlTextWriter(stm, System.Text.Encoding.UTF8);

            StartElement(xw, "document");
            StartElement(xw, "body");
            WritePara(xw, "The following upcoming games URGENTLY need help!");
            WritePara(
                xw,
                "If you can work ANY of these games, either sign up on Arbiter, or just reply to this mail and let me know which games you can do. Thanks!");

            StartTable(xw, 5);
            SortedDictionary<string, List<Game>> mpSlotGames = new SortedDictionary<string, List<Game>>();

            foreach (Game gm in gms.Games.Values)
            {
                if (gm.OpenSlots == 0)
                    continue;

                //if (gm.TotalSlots - gm.OpenSlots > 1)
                //    continue;

//                string s = $"{gm.Slots[0].Dttm.ToString("yyyyMMdd:HHmm")}-{gm.Slots[0].SiteShort}-{gm.TotalSlots - gm.OpenSlots}";
                string s = $"{gm.Slots[0].Dttm.ToString("yyyyMMdd:HHmm")}-{gm.Slots[0].SiteShort}";

                if (!mpSlotGames.ContainsKey(s))
                    mpSlotGames.Add(s, new List<Game>());

                mpSlotGames[s].Add(gm);
            }


            foreach (List<Game> plgm in mpSlotGames.Values)
            {
                (int cGamesWithNoUmpires, int cTotalOpenSlots, Game gameFirst) = GetGameSummaryFromGames(plgm);

                WriteGames(xw, plgm);
                AppendGamesToSb(plgm, sb);
            }

            EndTable(xw);
            sb.Append("</table></div>");
            EndElement(xw); // body
            EndElement(xw); // document

            xw.Flush();
            stm.Flush();

            sArbiterHelpNeeded = sb.ToString();
            return true;
        }

        /* S T M  C R E A T E  P A R T */
        /*----------------------------------------------------------------------------
            %%Function: StmCreatePart
            %%Qualified: ArbWeb.OOXML.StmCreatePart
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public static Stream StmCreatePart(Package pkg, string sUri, string sContentType, out PackagePart prt)
        {
            Uri uriTeams = new System.Uri(sUri, UriKind.Relative);

            prt = pkg.GetPart(uriTeams);

            List<PackageRelationship> plrel = new List<PackageRelationship>();

            foreach (PackageRelationship rel in prt.GetRelationships())
            {
                plrel.Add(rel);
            }

            prt = null;

            pkg.DeletePart(uriTeams);
            prt = pkg.CreatePart(uriTeams, sContentType);

            foreach (PackageRelationship rel in plrel)
            {
                prt.CreateRelationship(rel.TargetUri, rel.TargetMode, rel.RelationshipType, rel.Id);
            }

            return prt.GetStream(FileMode.Create, FileAccess.Write);
        }


        /* C R E A T E  M A I L  M E R G E  D O C */
        /*----------------------------------------------------------------------------
            %%Function: CreateMailMergeDoc
            %%Qualified: ArbWeb.OOXML.CreateMailMergeDoc
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public static void CreateMailMergeDoc(string sTemplate, string sDest, string sDataSource, ScheduleGames gms, out string sArbiterHelpNeeded)
        {
            System.IO.File.Copy(sTemplate, sDest);
            Package pkg = Package.Open(sDest, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

            CreateDocSettings(pkg, sDataSource);
            CreateMainDoc(pkg, sDataSource, gms, out sArbiterHelpNeeded);
            pkg.Flush();
            pkg.Close();
        }
    }
}
