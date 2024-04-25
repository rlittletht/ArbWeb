using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;
using SeleniumExtras.WaitHelpers;
using TCore.StatusBox;
using TCore.WebControl;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace ArbWeb
{
    public partial class WebRoster
    {
        /*----------------------------------------------------------------------------
			%%Function:GetRosterInfoFromServer
			%%Qualified:ArbWeb.WebRoster.GetRosterInfoFromServer

            Get the roster information from the server.
        ----------------------------------------------------------------------------*/
        void GetRosterInfoFromServer(string sEmail, string sOfficialID, RosterEntry rste)
        {
            SyncRsteWithServer(sOfficialID, rste, null);

            if (rste.Address1 == null
                || rste.Address2 == null
                || rste.City == null
                || rste.m_sDateJoined == null
                || rste.m_sDateOfBirth == null
                || rste.Email == null
                || rste.First == null
                || rste.m_sGamesPerDay == null
                || rste.m_sGamesPerWeek == null
                || rste.Last == null
                || rste.m_sOfficialNumber == null
                || rste.State == null
                || rste.m_sTotalGames == null
                || rste.m_sWaitMinutes == null
                || rste.Zip == null)
            {
                throw new Exception("couldn't extract one more more fields from official info");
            }
        }

        /*----------------------------------------------------------------------------
        	%%Function: VOPC_UpdateLastAccess
			%%Qualified:ArbWeb.WebRoster.VOPC_UpdateLastAccess
        	
            Update the "last login" value.  since we are scraping the screen for 
			this, we have to deal with pagination
        ----------------------------------------------------------------------------*/
        private void VOPC_UpdateLastAccess(Object o)
        {
            MicroTimer timer = new MicroTimer();
            Roster rstBuilding = (Roster)o;

            UpdateLastAccessFromCoreOfficialsPage(rstBuilding);

            timer.Stop();
            m_appContext.StatusReport.LogData($"UpdateLastAccessFromPage elapsed: {timer.MsecFloat}", 1, MSGT.Body);
        }

        /*----------------------------------------------------------------------------
        	%%Function: FixupEmailAddressForNotJoinedOfficialsFromCoreOfficialsPage
			%%Qualified:ArbWeb.WebRoster.FixupEmailAddressForNotJoinedOfficialsFromCoreOfficialsPage
        	
            non-joined officials won't get an email address in the downloaded
            quickroster. we can fix this up as we visit the officials page by matching
            the full name of the official.

            this is n^2 for now (every non-joined official searches every roster
            entry to try to fixup)
        ----------------------------------------------------------------------------*/
        private void FixupEmailAddressForNotJoinedOfficialsFromCoreOfficialsPage(Roster rstBuilding)
        {
            // grab the info from the current navigated page
            IWebElement table = m_appContext.WebControl.Driver.FindElement(By.Id(WebCore._sid_OfficialsView_ContentTable));

            string sHtml = table.GetAttribute("outerHTML");
            HtmlDocument html = new HtmlDocument();

            html.LoadHtml(sHtml);
            ;

            string sSelect = "//tr";
            HtmlNodeCollection rows = html.DocumentNode.SelectSingleNode(".").SelectNodes(sSelect);

            foreach (HtmlNode row in rows)
            {
                if (!row.HasClass("notJoinedItems"))
                    continue;

                HtmlNodeCollection cells = row.SelectNodes("td");

                if (cells.Count < 4)
                    continue;

                HtmlNode cellFullName = cells[2];
                HtmlNode cellEmail = cells[3];

                string sEmail = cellEmail.InnerText.Trim();
                string sRosterName = cellFullName.InnerText.Trim();

                RosterEntry rste = rstBuilding.RsteLookupRosterNameNoEmail(sRosterName);
                if (rste == null)
                {
                    m_appContext.StatusReport.AddMessage(
                        $"Could not find no email entry for nonJoinedOfficial {sRosterName} with email {sEmail}",
                        MSGT.Error);

                    continue;
                }

                m_appContext.StatusReport.AddMessage(
                    $"Updating email address for {sRosterName} to {sEmail}",
                    MSGT.Body);

                rste.Email = sEmail;
            }
        }

        /*----------------------------------------------------------------------------
			%%Function:VOPC_FixupNonJoinedEmailAddress
			%%Qualified:ArbWeb.WebRoster.VOPC_FixupNonJoinedEmailAddress
        ----------------------------------------------------------------------------*/
        private void VOPC_FixupNonJoinedEmailAddress(Object o)
        {
            MicroTimer timer = new MicroTimer();

            Roster rstBuilding = (Roster)o;

            FixupEmailAddressForNotJoinedOfficialsFromCoreOfficialsPage(rstBuilding);

            timer.Stop();
            m_appContext.StatusReport.LogData($"FixupEmailAddressForNotJoinedFromPage elapsed: {timer.MsecFloat}", 1, MSGT.Body);
        }

        /*----------------------------------------------------------------------------
        	%%Function: UpdateLastAccessFromCoreOfficialsPage
			%%Qualified:ArbWeb.WebRoster.UpdateLastAccessFromCoreOfficialsPage
        	
            Assuming we are on the core officials page...
        ----------------------------------------------------------------------------*/
        private void UpdateLastAccessFromCoreOfficialsPage(Roster rstBuilding)
        {
            IWebElement table = m_appContext.WebControl.Driver.FindElement(By.Id(WebCore._sid_OfficialsView_ContentTable));

            string sHtml = table.GetAttribute("outerHTML");
            HtmlDocument html = new HtmlDocument();
            html.LoadHtml(sHtml);

            string sSelect = "//tr";

            HtmlNodeCollection rows = html.DocumentNode.SelectNodes(sSelect);

            foreach (HtmlNode row in rows)
            {
                sSelect = "td";
                HtmlNodeCollection cells = row.SelectNodes(sSelect);

                if (cells.Count < 5)
                    continue;

                HtmlNode cellEmail = cells[3];
                HtmlNode cellSignedIn = cells[4];

                string sEmail = cellEmail.InnerText.Trim();
                string sSignedIn = cellSignedIn.InnerText.Trim();

                if (sEmail == "Email")
                    continue;

                RosterEntry rste = rstBuilding.RsteLookupEmail(sEmail);
                if (rste == null)
                {
                    m_appContext.StatusReport.AddMessage(
                        $"Lookup failed during ProcessAllOfficialPages for official '{cells[2].InnerText.Trim()}'({sEmail})",
                        MSGT.Error);
                    continue;
                }

                m_appContext.StatusReport.LogData(
                    $"Updating last access for official '{rste.Name}', {sSignedIn}",
                    5,
                    MSGT.Body);
                rste.m_sLastSignin = sSignedIn;
            }
        }

        /*----------------------------------------------------------------------------
			%%Function:DoProcessQuickRosterOfficials
			%%Qualified:ArbWeb.WebRoster.DoProcessQuickRosterOfficials
        ----------------------------------------------------------------------------*/
        Roster DoProcessQuickRosterOfficials(string sDownloadedRoster, bool fIncludeRankings, bool fIncludeLastAccess, bool fRankOnly, bool fAddOfficialsOnly)
        {
            Roster rstBuilding = new Roster();

            rstBuilding.ReadRoster(sDownloadedRoster);

            // always fixup the email addresses for non-joined officials
            {
                MicroTimer timer = new MicroTimer();

                OfficialsRosterWebInterop gr = new OfficialsRosterWebInterop(
                    m_appContext,
                    !fRankOnly, // !m_cbRankOnly.Checked,
                    fAddOfficialsOnly, // m_cbAddOfficialsOnly.Checked,
                    null,
                    null,
                    null);
                gr.ProcessAllOfficialPages(VOPC_FixupNonJoinedEmailAddress, rstBuilding);

                timer.Stop();
                m_appContext.StatusReport.LogData($"ProcessAllOfficialPages For EmailFixup elapsed: {timer.MsecFloat}", 1, MSGT.Body);
            }

            if (fIncludeLastAccess)
            {
                MicroTimer timer = new MicroTimer();
                OfficialsRosterWebInterop gr = new OfficialsRosterWebInterop(
                    m_appContext,
                    !fRankOnly, // !m_cbRankOnly.Checked,
                    fAddOfficialsOnly, // m_cbAddOfficialsOnly.Checked,
                    null,
                    null,
                    null);

                gr.ProcessAllOfficialPages(VOPC_UpdateLastAccess, rstBuilding);

                timer.Stop();
                m_appContext.StatusReport.LogData($"ProcessAllOfficialPages For LastAccess elapsed: {timer.MsecFloat}", 1, MSGT.Body);
            }

            if (fIncludeRankings)
            {
                MicroTimer timer = new MicroTimer();

                HandleRankings(null, rstBuilding);

                timer.Stop();
                m_appContext.StatusReport.LogData($"Handle rankings elapsed: {timer.MsecFloat}", 1, MSGT.Body);
            }

            return rstBuilding;
        }

        /*----------------------------------------------------------------------------
			%%Function:RosterQuickBuildFromDownloadedRoster
			%%Qualified:ArbWeb.WebRoster.RosterQuickBuildFromDownloadedRoster
        ----------------------------------------------------------------------------*/
        private Roster RosterQuickBuildFromDownloadedRoster(
            string sDownloadedRoster, bool fIncludeRankings, bool fIncludeLastAccess, bool fRankOnly, bool fAddOfficialsOnly)
        {
            return DoProcessQuickRosterOfficials(sDownloadedRoster, fIncludeRankings, fIncludeLastAccess, fRankOnly, fAddOfficialsOnly);
        }


        /*----------------------------------------------------------------------------
        	%%Function: DownloadQuickRosterToFile
			%%Qualified:ArbWeb.WebRoster.DownloadQuickRosterToFile
			
        ----------------------------------------------------------------------------*/
        string DownloadQuickRosterToFile()
        {
            m_appContext.StatusReport.AddMessage("Starting Quick Roster download to temp file...");
            m_appContext.StatusReport.PushLevel();

            m_appContext.PushCursor(Cursors.WaitCursor);
            string sTempFile = SRosterFileDownload();


            m_appContext.PopCursor();
            m_appContext.StatusReport.PopLevel();
            return sTempFile;
        }

        /* D O  D O W N L O A D  Q U I C K  R O S T E R  W O R K */
        /*----------------------------------------------------------------------------
        	%%Function: DoDownloadQuickRosterWork
			%%Qualified:ArbWeb.WebRoster.DoDownloadQuickRosterWork
        ----------------------------------------------------------------------------*/
        Roster DoDownloadQuickRosterWork(bool fRankOnly, bool fAddOfficialsOnly)
        {
            m_appContext.StatusReport.AddMessage("Starting Quick Roster download...");
            m_appContext.StatusReport.PushLevel();

            string sTempFile = DownloadQuickRosterToFile();

            // now, update the last access date and fetch the rankings and update the last access date
            Roster rst = RosterQuickBuildFromDownloadedRoster(sTempFile, true, true, fRankOnly, fAddOfficialsOnly);

            m_appContext.StatusReport.PopLevel();

            m_appContext.StatusReport.AddMessage("Completed Quick Roster download.");
            return rst;
        }


        /*----------------------------------------------------------------------------
			%%Function:DoDownloadQuickRosterOfficialsOnlyWork
			%%Qualified:ArbWeb.WebRoster.DoDownloadQuickRosterOfficialsOnlyWork
        ----------------------------------------------------------------------------*/
        Roster DoDownloadQuickRosterOfficialsOnlyWork(bool fRankOnly, bool fAddOfficialsOnly)
        {
            m_appContext.StatusReport.AddMessage("Starting Quick Roster download (officials only, no rankings)...");
            m_appContext.StatusReport.PushLevel();

            string sTempFile = DownloadQuickRosterToFile();

            // now, update the last access date and fetch the rankings and update the last access date
            Roster rst = RosterQuickBuildFromDownloadedRoster(sTempFile, false, false, fRankOnly, fAddOfficialsOnly);

            m_appContext.StatusReport.PopLevel();

            m_appContext.StatusReport.AddMessage("Completed Quick Roster download.");
            return rst;
        }

        /*----------------------------------------------------------------------------
			%%Function:DoRosterDownload
			%%Qualified:ArbWeb.WebRoster.DoRosterDownload
        ----------------------------------------------------------------------------*/
        void DoRosterDownload(string sTempFile)
        {
            Utils.ThrowIfNot(m_appContext.WebControl.FNavToPage(WebCore._s_Page_OfficialsView), "Couldn't nav to officials view!");
            m_appContext.WebControl.WaitForPageLoad();

            // now we are on the PrintRoster screen
            Utils.ThrowIfNot(
                m_appContext.WebControl.FClickControlId(WebCore._sid_OfficialsView_PrintCustomRoster, WebCore._sid_CustomRosterPrint_UserFilter),
                "Can't click on roster control");
            // check a whole bunch of config checkboxes

            // select All Officials
            WebControl.FSetSelectedOptionTextForControlId(
                m_appContext.WebControl.Driver,
                m_appContext.StatusReport,
                WebCore._sid_CustomRosterPrint_UserFilter,
                "All Officials");
            m_appContext.WebControl.WaitForPageLoad();

            WebControl.FSetCheckboxControlIdVal(m_appContext.WebControl.Driver, true, WebCore._sid_CustomRosterPrint_DateJoined);
            WebControl.FSetCheckboxControlIdVal(m_appContext.WebControl.Driver, true, WebCore._sid_CustomRosterPrint_OfficialNumber);
            WebControl.FSetCheckboxControlIdVal(m_appContext.WebControl.Driver, true, WebCore._sid_CustomRosterPrint_DateOfBirth);
            WebControl.FSetCheckboxControlIdVal(m_appContext.WebControl.Driver, true, WebCore._sid_CustomRosterPrint_UserID);
            WebControl.FSetCheckboxControlIdVal(m_appContext.WebControl.Driver, true, WebCore._sid_CustomRosterPrint_MiddleName);

            m_appContext.WebControl.FClickControlId(WebCore._sid_CustomRosterPrint_CustomFieldListDropdown); // dropdown the menu
            m_appContext.WebControl.WaitForCondition(
                ExpectedConditions.ElementToBeClickable(
                    m_appContext.WebControl.Driver.FindElement(By.Id(WebCore._sid_CustomRosterPrint_SelectAllCustomFields))));

            WebControl.FSetCheckboxControlIdVal(m_appContext.WebControl.Driver, true, WebCore._sid_CustomRosterPrint_SelectAllCustomFields);
            m_appContext.WebControl.WaitForPageLoad();
            m_appContext.WebControl.FClickControlId(WebCore._sid_CustomRosterPrint_CustomFieldListDropdown); // dismiss the menu

            WebControl.FileDownloader downloader = new WebControl.FileDownloader(
                m_appContext.WebControl,
                "RosterReport.xlsx",
                null,
                () => m_appContext.WebControl.FClickControlId(WebCore._sid_CustomRosterPrint_GenerateRosterReport));

            string sDownloadedFile = downloader.GetDownloadedFile();

            DownloadGenericExcelReport.ConvertExcelFileToCsv(sDownloadedFile, sTempFile);
            File.Delete(sDownloadedFile);
        }

        /*----------------------------------------------------------------------------
			%%Function:SRosterFileDownload
			%%Qualified:ArbWeb.WebRoster.SRosterFileDownload
        ----------------------------------------------------------------------------*/
        private string SRosterFileDownload()
        {
            // navigate to the officials page...
            m_appContext.EnsureLoggedIn();

            string sTempFile = $"{Environment.GetEnvironmentVariable("Temp")}\\temp{System.Guid.NewGuid().ToString()}.csv";


            DoRosterDownload(sTempFile);

            return sTempFile;
        }


        /*----------------------------------------------------------------------------
			%%Function:HandleRosterPostUpdateForDownload
			%%Qualified:ArbWeb.WebRoster.HandleRosterPostUpdateForDownload
        ----------------------------------------------------------------------------*/
        private void HandleRosterPostUpdateForDownload(OfficialsRosterWebInterop gr, IRoster irstBuilding)
        {
            // get the last login date from the officials main page
            gr.NavigateOfficialsPageAllOfficials();
            throw new Exception("NYI");

            // gr.ProcessAllOfficialPages(VOPC_UpdateLastAccess, irstBuilding);
        }

        /*----------------------------------------------------------------------------
			%%Function:HandleRosterPass1VisitForUploadDownload
			%%Qualified:ArbWeb.WebRoster.HandleRosterPass1VisitForUploadDownload
        ----------------------------------------------------------------------------*/
        void HandleRosterPass1VisitForUploadDownload(
            string sEmail,
            string sOfficialID,
            IRoster irstUploading,
            IRoster irstServer,
            IRosterEntry irste,
            IRoster irstBuilding,
            bool fNotJustAdded,
            bool fMarkOnly)
        {
            // if we're just marking officials, don't update the misc fields
            if (!fMarkOnly)
                UpdateMisc(sEmail, sOfficialID, irstUploading, irstServer, (RosterEntry)irste, irstBuilding);

            // if we just added the official, then we already entered all of their information, don't do it again
            if (fNotJustAdded)
                UpdateInfo(sEmail, sOfficialID, irstUploading, irstServer, (RosterEntry)irste, fMarkOnly);
        }

        /*----------------------------------------------------------------------------
        	%%Function: DoDownloadRoster
			%%Qualified:ArbWeb.WebRoster.DoDownloadRoster
        ----------------------------------------------------------------------------*/
        public void DoDownloadRoster(bool fRankOnly, bool fAddOfficialsOnly)
        {
            m_appContext.StatusReport.AddMessage("Starting FULL Roster download...");
            m_appContext.StatusReport.PushLevel();

            m_appContext.PushCursor(Cursors.WaitCursor);
            string sOutFile = OfficialsRosterWebInterop.SBuildRosterFilename(m_appContext.Profile.Roster);

            m_appContext.Profile.Roster = sOutFile;
            OfficialsRosterWebInterop gr = new OfficialsRosterWebInterop(
                m_appContext,
                !fRankOnly, // fNeedPass1OnUpload
                fAddOfficialsOnly, // only add officials
                HandleRosterPass1VisitForUploadDownload,
                AddOfficials,
                HandleRankings
            );

            Roster rstBuilding = new Roster();
            gr.GenericVisitRoster(null, rstBuilding, sOutFile, null, HandleRosterPostUpdateForDownload);

            m_appContext.PopCursor();
            m_appContext.StatusReport.PopLevel();
            System.IO.File.Delete(m_appContext.Profile.RosterWorking);
            System.IO.File.Copy(sOutFile, m_appContext.Profile.RosterWorking);
            m_appContext.StatusReport.AddMessage("Completed FULL Roster download.");
        }

        /*----------------------------------------------------------------------------
        	%%Function: DoDownloadQuickRoster
			%%Qualified:ArbWeb.WebRoster.DoDownloadQuickRoster
        ----------------------------------------------------------------------------*/
        public async void DoDownloadQuickRoster(bool fRankOnly, bool fAddOfficialsOnly)
        {
            Task<Roster> tsk = new Task<Roster>(
                () => DoDownloadQuickRosterWork(fRankOnly, fAddOfficialsOnly));

            tsk.Start();

            string sOutFile = OfficialsRosterWebInterop.SBuildRosterFilename(m_appContext.Profile.Roster);
            m_appContext.Profile.Roster = sOutFile;

            // make sure directories exist
            Directory.CreateDirectory(Path.GetDirectoryName(m_appContext.Profile.Roster));
            Directory.CreateDirectory(Path.GetDirectoryName(m_appContext.Profile.RosterWorking));

            Roster rst = await tsk;

            rst.WriteRoster(sOutFile);
            System.IO.File.Delete(m_appContext.Profile.RosterWorking);
            System.IO.File.Copy(sOutFile, m_appContext.Profile.RosterWorking);
        }
    }
}
