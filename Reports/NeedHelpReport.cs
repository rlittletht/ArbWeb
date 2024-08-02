using System;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using ArbWeb.Games;
using TCore.StatusBox;
using TCore.Util;
using TCore.WebControl;
using static ArbWeb.Roster;

namespace ArbWeb.Reports
{
    public class NeedHelpReport
    {
        private IAppContext m_appContext;

        /*----------------------------------------------------------------------------
            %%Function:NeedHelpReport
            %%Qualified:ArbWeb.Reports.NeedHelpReport.NeedHelpReport
        ----------------------------------------------------------------------------*/
        public NeedHelpReport(IAppContext appContext)
        {
            m_appContext = appContext;
        }


        public Roster GenerateRosterForMerge(
            ScheduleGames gms,
            Roster rst,
            bool fFilterByRanking,
            bool fFilterASOnly,
            string sMiscFilter)
        {
            if (fFilterASOnly)
                return rst.FilterToAllStarOnly(sMiscFilter);

            if (fFilterByRanking)
                return rst.FilterByRanks(gms.RequiredRanks(), sMiscFilter);

            if (string.IsNullOrEmpty(sMiscFilter))
                return rst;

            Roster rstFiltered = new Roster();

            foreach (RosterEntry entry in rst.Plrste)
            {
                if (entry.FMatchAnyMisc(sMiscFilter))
                    rstFiltered.Add(entry);
            }

            return rstFiltered;
        }

        public ScheduleGames GenerateGamesForMerge(
            CountsData gc,
            string[] rgsSports,
            string[] rgsSportLevels,
            SlotAggr aggregation)
        {
            return gc.GamesFromFilter(
                rgsSports,
                rgsSportLevels,
                false,
                aggregation);
        }

        public void DoGenMailMergeCsvOnly(
            CountsData gc,
            Roster rst,
            string[] rgsSports,
            string[] rgsSportLevels,
            SlotAggr aggregation,
            bool fFilterByRanking,
            bool fFilterASOnly,
            string sMiscFilter,
            string targetCsv)
        {
            ScheduleGames gms = null;

            m_appContext.StatusReport.ReportAction(
                "GamesFromFilter",
                () =>
                {
                    // first, generate the mailmerge source csv file.  this is either the entire roster, or just the folks 
                    // rated for the sports we are filtered to
                    gms = GenerateGamesForMerge(
                        gc,
                        rgsSports,
                        rgsSportLevels,
                        aggregation);

                });

            m_appContext.StatusReport.ReportAction(
                "Saving MailMerge roster",
                () =>
                {
                    Roster rstFiltered = GenerateRosterForMerge(gms, rst, fFilterByRanking, fFilterASOnly, sMiscFilter);
                    WriteRosterToMailMergeCsv(rstFiltered, targetCsv);
                });

            m_appContext.DoPendingQueueUIOp();
        }

        void WriteRosterToMailMergeCsv(Roster roster, string csvFile)
        {
            using StreamWriter sw = new StreamWriter(csvFile, false, System.Text.Encoding.Default);

            sw.WriteLine("email,firstname,lastname");
            foreach (RosterEntry rste in roster.Plrste)
            {
                sw.WriteLine("{0},{1},{2}", rste.Email, rste.First, rste.Last);
            }

            sw.Flush();
            sw.Close();
        }

        /*----------------------------------------------------------------------------
            %%Function:DoGenMailMergeAndAnnouce
            %%Qualified:ArbWeb.Reports.NeedHelpReport.DoGenMailMergeAndAnnouce
        ----------------------------------------------------------------------------*/
        public void DoGenMailMergeAndAnnouce(
            CountsData gc,
            Roster rst,
            string[] rgsSports,
            string[] rgsSportLevels,
            SlotAggr aggregation,
            bool fFilterByRanking,
            bool fFilterASOnly,
            string sMiscFilter,
            bool fLaunch,
            bool fSetWebAnnounce)
        {
            m_appContext.StatusReport.AddMessage("Generating mail merge documents...", MSGT.Header, false);

            m_appContext.StatusReport.LogData("GamesFromFilter...", 3, MSGT.Header);

            // first, generate the mailmerge source csv file.  this is either the entire roster, or just the folks 
            // rated for the sports we are filtered to
            ScheduleGames gms = GenerateGamesForMerge(gc, 
                rgsSports,
                rgsSportLevels,
                aggregation);

            m_appContext.StatusReport.LogData("Beginning to build rosterfiltered...", 3, MSGT.Body);
            Roster rstFiltered = GenerateRosterForMerge(gms, rst, fFilterByRanking, fFilterASOnly, sMiscFilter);

            string sCsvTemp = Filename.SBuildTempFilename("MailMergeRoster", "csv");
            m_appContext.StatusReport.LogData($"Writing filtered roster to {sCsvTemp}", 3, MSGT.Body);
            WriteRosterToMailMergeCsv(rstFiltered, sCsvTemp);

            // ok, now create the mailmerge .docx
            string sTempName;
            string sArbiterHelpNeeded;
            m_appContext.StatusReport.LogData("Filtered roster written", 3, MSGT.Body);

            string sApp = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string sFile = Path.Combine(Path.GetPathRoot(sApp), Path.GetDirectoryName(sApp), "mailmergedoc.docx");

            sTempName = Filename.SBuildTempFilename("mailmergedoc", "docx");
            m_appContext.StatusReport.LogData($"Writing mailmergedoc to {sTempName} using template at {sFile}", 3, MSGT.Body);
            OOXML.CreateMailMergeDoc(sFile, sTempName, sCsvTemp, gms, out sArbiterHelpNeeded);

            m_appContext.StatusReport.LogData($"ArbiterHelp HTML created: {sArbiterHelpNeeded}", 5, MSGT.Body);
            if (fLaunch)
            {
                System.Windows.Forms.Clipboard.SetText(sArbiterHelpNeeded);
                m_appContext.StatusReport.AddMessage("Done, launching document...", MSGT.Header, false);
                System.Diagnostics.Process.Start(sTempName);
            }

            if (fSetWebAnnounce)
            {
                WebAnnounce announce = new WebAnnounce(m_appContext);

                announce.SetArbiterAnnounce(sArbiterHelpNeeded);
            }

            m_appContext.DoPendingQueueUIOp();
        }
    }
}
