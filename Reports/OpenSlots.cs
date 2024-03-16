using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ArbWeb.Games;
using TCore.StatusBox;
using TCore.WebControl;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ArbWeb.Reports
{
    public class OpenSlots
    {
        private IAppContext m_appContext;

        /*----------------------------------------------------------------------------
            %%Function:OpenSlots
            %%Qualified:ArbWeb.Reports.OpenSlots.OpenSlots
        ----------------------------------------------------------------------------*/
        public OpenSlots(IAppContext appContext)
        {
            m_appContext = appContext;
        }

        public SlotAggr Aggregation => m_saOpenSlots;

        /* S  H T M L  R E A D  F I L E */
        /*----------------------------------------------------------------------------
            %%Function: SHtmlReadFile
            %%Qualified: ArbWeb.AwMainForm.SHtmlReadFile
            %%Contact: rlittle
            
        ----------------------------------------------------------------------------*/
        private string SHtmlReadFile(string sFile)
        {
            string sHtml = "";
            TextReader tr = new StreamReader(sFile);
            string sLine;

            while ((sLine = tr.ReadLine()) != null)
                sHtml += " " + sLine;

            tr.Close();
            return sHtml;
        }

        /* D O  G E N  O P E N  S L O T S  M A I L */
        /*----------------------------------------------------------------------------
        	%%Function: DoGenOpenSlotsMail
        	%%Qualified: ArbWeb.AwMainForm.DoGenOpenSlotsMail
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public void DoGenOpenSlotsMail(
            CountsData gc,
            Roster rst,
            bool fTestEmail,
            string sEmailFilter,
            bool fSplitSports,
            bool fFilterSport,
            bool fIncludeOpenSlots,
            bool fFuzzyTimes,
            bool fPivotOnDates,
            bool fFilterByLevel,
            CheckedListBox sportsListbox,
            CheckedListBox sportsLevelsListbox)
        {
            string sTempFile = $"{Environment.GetEnvironmentVariable("Temp")}\\temp{System.Guid.NewGuid().ToString()}.htm";

            string sBcc = fTestEmail ? "" : rst.SBuildAddressLine(sEmailFilter);
            ;

            Outlook.Application appOlk = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");

            if (appOlk == null)
            {
                MessageBox.Show("No running instance of outlook!");
                return;
            }

            Outlook.MailItem oNote = appOlk.CreateItem(Outlook.OlItemType.olMailItem);
            // Outlook.MailItem oNote = (Outlook.MailItem)appOlk.CreateItem(Outlook.OlItemType.olMailItem);

            oNote.To = "rlittle@thetasoft.com";
            oNote.BCC = sBcc;
            oNote.Subject = "This is a test";
            oNote.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            oNote.HTMLBody = "<html><style>\r\n*#myId {\ncolor:Blue;\n}\n</style><body><p>Put your preamble here...</p>";

            if (fSplitSports)
            {
                string[] rgs;

                oNote.HTMLBody += "<h1>Baseball open slots</h1>";
                rgs = WebCore.RgsFromChlbxSport(fFilterSport, sportsListbox, "Softball", false);
                gc.GenOpenSlotsReport(
                    sTempFile,
                    fIncludeOpenSlots,
                    fFuzzyTimes,
                    fPivotOnDates,
                    rgs,
                    WebCore.RgsFromChlbx(fFilterByLevel, sportsLevelsListbox),
                    m_saOpenSlots);
                oNote.HTMLBody += SHtmlReadFile(sTempFile) + "<h1>Softball Open Slots</h1>";
                rgs = WebCore.RgsFromChlbxSport(fFilterSport, sportsListbox, "Softball", true);
                gc.GenOpenSlotsReport(
                    sTempFile,
                    fIncludeOpenSlots,
                    fFuzzyTimes,
                    fPivotOnDates,
                    rgs,
                    WebCore.RgsFromChlbx(fFilterByLevel, sportsLevelsListbox),
                    m_saOpenSlots);
                oNote.HTMLBody += SHtmlReadFile(sTempFile);
            }
            else
            {
                gc.GenOpenSlotsReport(
                    sTempFile,
                    fIncludeOpenSlots,
                    fFuzzyTimes,
                    fPivotOnDates,
                    WebCore.RgsFromChlbx(fFilterSport, sportsListbox),
                    WebCore.RgsFromChlbx(fFilterByLevel, sportsLevelsListbox),
                    m_saOpenSlots);
                oNote.HTMLBody += SHtmlReadFile(sTempFile);
            }

            oNote.Display(true);

            appOlk = null;
            System.IO.File.Delete(sTempFile);
        }

        private SlotAggr m_saOpenSlots;

        /* D O  C A L C  O P E N  S L O T S */
        /*----------------------------------------------------------------------------
        	%%Function: DoCalcOpenSlots
        	%%Qualified: ArbWeb.AwMainForm.DoCalcOpenSlots
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public async void DoCalcOpenSlots(
            string sRosterFile,
            string sGameFile,
            bool fIncludeCanceled,
            int affiliationRosterIndex,
            CheckedListBox sportsListBox,
            CheckedListBox sportsLevelsListBox,
            CheckedListBox rosterListBox,
            DateTime dttmStart,
            DateTime dttmEnd)
        {
            Task<CountsData> taskCalc = new Task<CountsData>(
                () => CalcOpenSlotsWork(sRosterFile, sGameFile, fIncludeCanceled, affiliationRosterIndex, dttmStart, dttmEnd));
            taskCalc.Start();

            CountsData cd = await taskCalc;

            m_appContext.StatusReport.PopLevel();
            m_appContext.StatusReport.AddMessage("Updating listboxes...", MSGT.Header, false);
            // update regenerate the listboxes...
            string[] rgsSports = WebCore.RgsFromChlbx(true, sportsListBox);
            string[] rgsSportLevels = WebCore.RgsFromChlbx(true, sportsLevelsListBox);

            bool fCheckAllSports = false;
            bool fCheckAllSportLevels = false;

            if (rgsSports.Length == 0 && sportsListBox.Items.Count == 0)
                fCheckAllSports = true;

            if (rgsSports.Length == 0 && sportsLevelsListBox.Items.Count == 0)
                fCheckAllSportLevels = true;

            WebCore.UpdateChlbxFromRgs(sportsListBox, cd.GetOpenSlotSports(m_saOpenSlots), rgsSports, null, fCheckAllSports);
            WebCore.UpdateChlbxFromRgs(
                sportsLevelsListBox,
                cd.GetOpenSlotSportLevels(m_saOpenSlots),
                rgsSportLevels,
                fCheckAllSports ? null : rgsSports,
                fCheckAllSportLevels);
            string[] rgsRosterSites = WebCore.RgsFromChlbx(true, rosterListBox);

            WebCore.UpdateChlbxFromRgs(rosterListBox, cd.GetSiteRosterSites(m_saOpenSlots), rgsRosterSites, null, false);
            m_appContext.StatusReport.PopLevel();
            m_appContext.DoPendingQueueUIOp();
        }

        /* C A L C  O P E N  S L O T S  W O R K */
        /*----------------------------------------------------------------------------
	    	%%Function: CalcOpenSlotsWork
	    	%%Qualified: ArbWeb.AwMainForm.CalcOpenSlotsWork
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
        private CountsData CalcOpenSlotsWork(
            string sRosterFile,
            string sGameFile,
            bool fIncludeCanceled,
            int affiliationRosterIndex,
            DateTime dttmStart,
            DateTime dttmEnd)
        {
            CountsData gc = m_appContext.GcEnsure(sRosterFile, sGameFile, fIncludeCanceled, affiliationRosterIndex);
            m_appContext.StatusReport.AddMessage("Calculating slot data...", MSGT.Header, false);

            m_appContext.StatusReport.PopLevel();
            m_appContext.StatusReport.AddMessage("Calculating open slots...", MSGT.Header, false);
            m_saOpenSlots = gc.CalcOpenSlots(dttmStart, dttmEnd);
            return gc;
        }

        /* G E N  O P E N  S L O T S  R E P O R T  D E T A I L */
        /*----------------------------------------------------------------------------
                %%Function: GenOpenSlotsReportDetail
                %%Qualified: ArbWeb.CountsData:GameData:Games.GenOpenSlotsReportDetail
                %%Contact: rlittle

                Generate a report of open slots
            ----------------------------------------------------------------------------*/
        private static void GenOpenSlotsReportDetail(
            ScheduleGames games, string sReport, string[] rgsSportFilter, string[] rgsSportLevelFilter, SlotAggr aggregation)
        {
            DateTime dttmStart = aggregation.DttmStart;
            DateTime dttmEnd = aggregation.DttmEnd;

            StreamWriter sw = new StreamWriter(sReport, false, Encoding.Default);

            string sFormat = "<tr><td>{0}<td>{1}<td>{2}<td>{3}<td>{4}<td>{5}<td>{6}<td>{7}</tr>";

            sw.WriteLine("<html><body><table>");
            sw.WriteLine(String.Format(sFormat, "Game", "Date", "Time", "Field", "Level", "Home", "Away", "Slots"));

            SortedList<string, int> plsSports = Utils.PlsUniqueFromRgs(rgsSportFilter);
            SortedList<string, int> plsLevels = Utils.PlsUniqueFromRgs(rgsSportLevelFilter);

            foreach (GameSlot gm in games.SortedGameSlots)
            {
                if (!gm.Open)
                    continue;

                if (DateTime.Compare(gm.Dttm, dttmStart) < 0 || DateTime.Compare(gm.Dttm, dttmEnd) > 0)
                    continue;

                if (plsSports != null && !(plsSports.ContainsKey(gm.Sport)))
                    continue;

                if (plsLevels != null && !(plsLevels.ContainsKey(gm.SportLevel)))
                    continue;

                sw.WriteLine(
                    String.Format(
                        sFormat,
                        gm.GameNum,
                        gm.Dttm.ToString("MM/dd/yy ddd"),
                        gm.Dttm.ToString("hh:mm tt"),
                        gm.Site,
                        $"{gm.Sport}, {gm.Level}",
                        gm.Home,
                        gm.Away,
                        gm.Pos));
            }

            sw.Close();
        }

        /* G E N  O P E N  S L O T S  R E P O R T */
        /*----------------------------------------------------------------------------
                %%Function: GenOpenSlotsReport
                %%Qualified: ArbWeb.CountsData:GameData:Games.GenOpenSlotsReport
                %%Contact: rlittle

            ----------------------------------------------------------------------------*/
        private static void GenOpenSlotsReport(
            string sReport, string[] rgsSportsFilter, string[] rgsSportLevelsFilter, bool fFuzzyTimes, bool fDatePivot, SlotAggr aggregation)
        {
            aggregation.GenReportBySite(sReport, fFuzzyTimes, fDatePivot, rgsSportsFilter, rgsSportLevelsFilter);
        }

        /* G E N  O P E N  S L O T S  R E P O R T */
        /*----------------------------------------------------------------------------
            %%Function: GenOpenSlotsReport
            %%Qualified: ArbWeb.CountsData:GameData:Games.GenOpenSlotsReport
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public static void GenOpenSlotsReport(
            ScheduleGames games, string sReport, bool fDetail, bool fFuzzyTimes, bool fDatePivot, string[] rgsSportFilter, string[] rgsSportLevelFilter,
            SlotAggr aggregation)
        {
            if (fDetail)
                GenOpenSlotsReportDetail(games, sReport, rgsSportFilter, rgsSportLevelFilter, aggregation);
            else
                GenOpenSlotsReport(sReport, rgsSportFilter, rgsSportLevelFilter, fFuzzyTimes, fDatePivot, aggregation);
        }
    }
}
