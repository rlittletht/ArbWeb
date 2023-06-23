using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO.Packaging;
using System.IO;
using ArbWeb.Games;

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
            WriteElementFull(xw, "view", new[] {"web"});
            StartElement(xw, "mailMerge");
            WriteElementFull(xw, "mainDocumentType", new[] {"email"});
            WriteElementFull(xw, "linkToQuery", (string[])null);
            WriteElementFull(xw, "dataType", new[] {"textFile"});
            WriteElementFull(xw, "connectString", new[] {""});
            WriteElementFull(xw, "query", new[] {$"SELECT * FROM {sDataSource}"});
            
            PackageRelationship rel = prt.CreateRelationship( new System.Uri($"file:///{sDataSource}", UriKind.Absolute), TargetMode.External, s_sUriMailMergeRelType);
            StartElement(xw, "dataSource");
            xw.WriteAttributeString("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", rel.Id);
            EndElement(xw);
            StartElement(xw, "odso");
            StartElement(xw, "fieldMapData");
            WriteElementFull(xw, "type", new[] {"dbColumn"});
            WriteElementFull(xw, "name", new[] {"email"});
            WriteElementFull(xw, "mappedName", new[] {"E-mail Address"});
            WriteElementFull(xw, "column", new[] {"0"});
            WriteElementFull(xw, "lid", new[] {"en-US"});
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
            WriteElementFull(xw, "tcW", new[] {new AttrPair("w", sWidth), new AttrPair("type", sWidthType)});
            EndElement(xw); // tcPr
            WritePara(xw, s);
            EndElement(xw); // tc

        }

        private static Dictionary<string, string> s_mpSportLevelFriendly = new Dictionary<string, string>()
            {
            {"All Stars SB 9/10's", "SB 8/9/10s"},
            {"All Stars SB 11's", "SB 9/10/11s"},
            {"All Stars SB Majors", "SB Majors"},
            {"All Stars SB Juniors", "SB Juniors"},
            {"All Stars 60' BB 11's", "BB 9/10/11s"},
            {"All Stars 60' BB 9/10's", "BB 8/9/10s"},
            {"All Stars 60' BB Majors", "BB Majors"},
            {"All Stars 90' BB Intermediate 70", "BB Intermediates"},
            {"All Stars 90' BB Juniors 90", "BB Juniors"},
            {"All Stars 90' BB Seniors", "BB Seniors"},
            };

        static string DescribeGame(Game gm, int cGames)
        {
            string sDesc;
            string sCount = "";

#if no
            if (cGames > 1)
                sCount = $"{cGames} games, ";
            else
                sCount = $"{cGames} game, ";

            if (cGames > 1)
            {
                sCount = $"{cGames} games, ";
                if (gm.TotalSlots - gm.OpenSlots == 0)
                    sDesc = "NO UMPIRES";
                else
                    sDesc = $"{gm.TotalSlots - gm.OpenSlots} umpires!";
            }
            else
#endif
            {
                if (gm.TotalSlots - gm.OpenSlots == 0)
                    sDesc = $"{gm.OpenSlots} umpires needed!<br/>NO UMPIRES SIGNED UP";
                else
                    sDesc = $"{gm.OpenSlots} umpires needed!";
            }
            //sDesc = $"{gm.TotalSlots - gm.OpenSlots} UMPIRE";

        return $"{sCount}{sDesc}";
        }

        static string FriendlySport(Game gm)
        {
            if (s_mpSportLevelFriendly.ContainsKey(gm.Slots[0].SportLevel))
                return s_mpSportLevelFriendly[gm.Slots[0].SportLevel];

            return gm.Slots[0].SportLevel;
        }

        static void WriteGame(XmlTextWriter xw, Game gm, int cGames)
        {
            StartElement(xw, "tr");
            WriteSingleParaCell(xw, "0", "auto", gm.Slots[0].Dttm.ToString("M/dd"));
            WriteSingleParaCell(xw, "0", "auto", gm.Slots[0].Dttm.ToString("ddd h tt"));

            WriteSingleParaCell(xw, "0", "auto", FriendlySport(gm));
            WriteSingleParaCell(xw, "0", "auto", gm.Slots[0].SiteShort);

            WriteSingleParaCell(xw, "0", "auto", DescribeGame(gm, cGames));
            EndElement(xw); // tr
        }

        static void AppendGameToSb(Game gm, int cGames, StringBuilder sb)
        {
            sb.Append($"<tr><td>{gm.Slots[0].Dttm.ToString("M/dd")}<td>{gm.Slots[0].Dttm.ToString("ddd h tt")}");
            sb.Append($"<td>{FriendlySport(gm)}<td>{gm.Slots[0].SiteShort}<td class='bold'>{DescribeGame(gm, cGames)}");
        }

        static void StartTable(XmlTextWriter xw, int cCols)
        {
            StartElement(xw, "tbl");
            StartElement(xw, "tblPr");
            WriteElementFull(xw, "tblStyle", new [] { new AttrPair("val", "TableGrid") });
            WriteElementFull(xw, "tblW", new [] { new AttrPair("w", "0"),new AttrPair("type", "auto") });
            WriteElementFull(xw, "tblLook", new [] { new AttrPair("val", "04A0"), new AttrPair("firstRow", "1"), new AttrPair("lastRow", "0"), new AttrPair("firstColumn", "1"), new AttrPair("lastColumn", "0"), new AttrPair("noHBand", "0"), new AttrPair("noVBand", "1") });
            EndElement(xw); // tblPr
            StartElement(xw, "tblGrid");
            while (cCols-- > 0)
                WriteElementFull(xw, "gridCol", new [] { new AttrPair("val", "1440") });
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
                "<div id='D9UrgentHelpNeeded'>"
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
            WritePara(xw, "If you can work ANY of these games, either sign up on Arbiter, or just reply to this mail and let me know which games you can do. Thanks!");

            StartTable(xw, 5);
            SortedDictionary<string, List<Game>> mpSlotGames = new SortedDictionary<string, List<Game>>();

            foreach (Game gm in gms.Games.Values)
                {
                if (gm.TotalSlots - gm.OpenSlots > 1)
                    continue;

                string s = $"{gm.Slots[0].Dttm.ToString("yyyyMMdd:HHmm")}-{gm.Slots[0].SiteShort}-{gm.TotalSlots - gm.OpenSlots}";
                if (!mpSlotGames.ContainsKey(s))
                    mpSlotGames.Add(s, new List<Game>());

                mpSlotGames[s].Add(gm);
                }

            
            foreach (List<Game> plgm in mpSlotGames.Values)
                {
                WriteGame(xw, plgm[0], plgm.Count);
                AppendGameToSb(plgm[0], plgm.Count, sb);
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