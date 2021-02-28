using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Web;
using System.Data.SqlClient;
using System.Xml;
using System.IO.Packaging;
using System.IO;

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
            WriteElementFull(xw, "query", new[] {String.Format("SELECT * FROM {0}", sDataSource)});
            
            PackageRelationship rel = prt.CreateRelationship( new System.Uri(String.Format("file:///{0}", sDataSource), UriKind.Absolute), TargetMode.External, s_sUriMailMergeRelType);
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
            StartElement(xw, "t");
            xw.WriteString(s);
            EndElement(xw); // t
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
            {"All Stars SB 9/10's", "SB 10s ALL STARS"},
            {"All Stars SB 11's", "SB 11s ALL STARS"},
            {"All Stars SB Majors", "SB Majors ALL STARS"},
            {"All Stars 60' BB 11's", "BB 11s ALL STARS"},
            {"All Stars 60' BB 9/10's", "BB 10s ALL STARS"},
            };
 
        static string DescribeGame(GameData.Game gm, int cGames)
        {
            string sDesc;
            string sCount = "";

            if (cGames > 1)
                sCount = String.Format("{0} games, ", cGames);
            else
                sCount = String.Format("{0} game, ", cGames);

            if (gm.TotalSlots - gm.OpenSlots == 0)
                sDesc = "NO UMPIRES";
            else
                sDesc = String.Format("{0} UMPIRE", gm.TotalSlots - gm.OpenSlots);

            return String.Format("{0}{1}", sCount, sDesc);
        }

        static string FriendlySport(GameData.Game gm)
        {
            if (s_mpSportLevelFriendly.ContainsKey(gm.Slots[0].SportLevel))
                return s_mpSportLevelFriendly[gm.Slots[0].SportLevel];

            return gm.Slots[0].SportLevel;
        }

        static void WriteGame(XmlTextWriter xw, GameData.Game gm, int cGames)
        {
            StartElement(xw, "tr");
            WriteSingleParaCell(xw, "0", "auto", gm.Slots[0].Dttm.ToString("M/dd"));
            WriteSingleParaCell(xw, "0", "auto", gm.Slots[0].Dttm.ToString("ddd h tt"));

            WriteSingleParaCell(xw, "0", "auto", FriendlySport(gm));
            WriteSingleParaCell(xw, "0", "auto", gm.Slots[0].SiteShort);

            WriteSingleParaCell(xw, "0", "auto", DescribeGame(gm, cGames));
            EndElement(xw); // tr
        }

        static void AppendGameToSb(GameData.Game gm, int cGames, StringBuilder sb)
        {
            sb.Append(String.Format("<tr><td>{0}<td>{1}", gm.Slots[0].Dttm.ToString("M/dd"), gm.Slots[0].Dttm.ToString("ddd h tt")));
            sb.Append(String.Format("<td>{0}<td>{1}<td class='bold'>{2}", FriendlySport(gm), gm.Slots[0].SiteShort, DescribeGame(gm, cGames)));
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
        public static bool CreateMainDoc(Package pkg, string sDataSource, GameData.GameSlots gms, out string sArbiterHelpNeeded)
        {
            PackagePart prt;
            StringBuilder sb = new StringBuilder();

            sb.Append("<div id='D9UrgentHelpNeeded'><h1>HELP NEEDED</h1><h4> The following upcoming games URGENTLY need help! <br>Please <a href=\"https://www.arbitersports.com/Official/SelfAssign.aspx\">SELF ASSIGN</a> now!</h4><style>    table.help td {padding-left: 2mm;padding-right: 2mm;}td.bold {font-weight: bold;}</style> <table class='help' border=1 style='border-collapse: collapse'></div>");

            Stream stm = StmCreatePart(pkg, "/word/document.xml", s_sUriContentTypeDoc, out prt);

            XmlTextWriter xw = new XmlTextWriter(stm, System.Text.Encoding.UTF8);

            StartElement(xw, "document");
            StartElement(xw, "body");
            WritePara(xw, "The following upcoming games URGENTLY need help!");
            WritePara(xw, "If you can work ANY of these games, either sign up on Arbiter, or just reply to this mail and let me know which games you can do. Thanks!");

            StartTable(xw, 5);
            Dictionary<string, List<GameData.Game>> mpSlotGames = new Dictionary<string, List<GameData.Game>>();

            foreach (GameData.Game gm in gms.Games.Values)
                {
                if (gm.TotalSlots - gm.OpenSlots > 1)
                    continue;

                string s = String.Format("{0}-{1}-{2}", gm.Slots[0].Dttm.ToString("yyyyMMdd:HHmm"), gm.Slots[0].SiteShort, gm.TotalSlots - gm.OpenSlots);
                if (!mpSlotGames.ContainsKey(s))
                    mpSlotGames.Add(s, new List<GameData.Game>());

                mpSlotGames[s].Add(gm);
                }

            foreach (List<GameData.Game> plgm in mpSlotGames.Values)
                {
                WriteGame(xw, plgm[0], plgm.Count);
                AppendGameToSb(plgm[0], plgm.Count, sb);
                }
            EndTable(xw);
            sb.Append("</table>");
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
        public static void CreateMailMergeDoc(string sTemplate, string sDest, string sDataSource, GameData.GameSlots gms, out string sArbiterHelpNeeded)
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