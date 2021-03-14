using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ArbWeb;
using mshtml;
using Win32Win;
using System.IO;
using HtmlAgilityPack;
using OpenQA.Selenium;
using StatusBox;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace ArbWeb
{
    public partial class AwMainForm : System.Windows.Forms.Form
    {
        /* G E T  R O S T E R  I N F O  F R O M  S E R V E R */
        /*----------------------------------------------------------------------------
        	%%Function: GetRosterInfoFromServer
        	%%Qualified: ArbWeb.AwMainForm.GetRosterInfoFromServer
        	%%Contact: rlittle
        	
            Get the roster information from the server.
        ----------------------------------------------------------------------------*/
        void GetRosterInfoFromServer(string sEmail, string sOfficialID, RosterEntry rste)
        {
            SyncRsteWithServer(m_awc.Document2, sOfficialID, rste, null);

            if (rste.Address1 == null || rste.Address2 == null || rste.City == null || rste.m_sDateJoined == null
                || rste.m_sDateOfBirth == null || rste.Email == null || rste.First == null
                || rste.m_sGamesPerDay == null || rste.m_sGamesPerWeek == null || rste.Last == null 
                || rste.m_sOfficialNumber == null
                || rste.State == null || rste.m_sTotalGames == null || rste.m_sWaitMinutes == null
                || rste.Zip == null)
            {
                throw new Exception("couldn't extract one more more fields from official info");
            }
        }

        /*----------------------------------------------------------------------------
        	%%Function: VOPC_UpdateLastAccess
        	%%Qualified: ArbWeb.AwMainForm.VOPC_UpdateLastAccess
        	%%Contact: rlittle
        	
            Update the "last login" value.  since we are scraping the screen for this, we have to deal with pagination
        ----------------------------------------------------------------------------*/
        private void VOPC_UpdateLastAccess(Object o)
        {
	        MicroTimer timer = new MicroTimer();
            Roster rstBuilding = (Roster)o;

            UpdateLastAccessFromCoreOfficialsPage(rstBuilding);
            
            timer.Stop();
            m_srpt.LogData($"UpdateLastAccessFromPage elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
        }

        /*----------------------------------------------------------------------------
        	%%Function: FixupEmailAddressForNotJoinedOfficialsFromCoreOfficialsPage
        	%%Qualified: ArbWeb.AwMainForm.FixupEmailAddressForNotJoinedOfficialsFromCoreOfficialsPage
        	
            non-joined officials won't get an email address in the downloaded
            quickroster. we can fix this up as we visit the officials page by matching
            the full name of the official.

            this is n^2 for now (every non-joined official searches every roster
            entry to try to fixup)
        ----------------------------------------------------------------------------*/
        private void FixupEmailAddressForNotJoinedOfficialsFromCoreOfficialsPage(Roster rstBuilding)
        {
	        // grab the info from the current navigated page
	        IWebElement table = m_webControl.Driver.FindElement(By.Id(WebCore._sid_OfficialsView_ContentTable));

	        string sHtml = table.GetAttribute("outerHTML");
	        HtmlDocument html = new HtmlDocument();
	        
	        html.LoadHtml(sHtml);;

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
                    m_srpt.AddMessage(
                        $"Could not find no email entry for nonJoinedOfficial {sRosterName} with email {sEmail}",
                        StatusBox.StatusRpt.MSGT.Error);

                    continue;
                }

                m_srpt.AddMessage($"Updating email address for {sRosterName} to {sEmail}",
                                  StatusBox.StatusRpt.MSGT.Body);

                rste.Email = sEmail;
            }
        }

        private void VOPC_FixupNonJoinedEmailAddress(Object o)
        {
	        MicroTimer timer = new MicroTimer();
	        
            Roster rstBuilding = (Roster)o;

            FixupEmailAddressForNotJoinedOfficialsFromCoreOfficialsPage(rstBuilding);
            
            timer.Stop();
            m_srpt.LogData($"FixupEmailAddressForNotJoinedFromPage elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
        }

        /*----------------------------------------------------------------------------
        	%%Function: UpdateLastAccessFromCoreOfficialsPage
        	%%Qualified: ArbWeb.AwMainForm.UpdateLastAccessFromCoreOfficialsPage
        	%%Contact: rlittle
        	
            Assuming we are on the core officials page...
        ----------------------------------------------------------------------------*/
        private void UpdateLastAccessFromCoreOfficialsPage(Roster rstBuilding)
        {
	        IWebElement table = m_webControl.Driver.FindElement(By.Id(WebCore._sid_OfficialsView_ContentTable));

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
                    m_srpt.AddMessage(
                        String.Format("Lookup failed during ProcessAllOfficialPages for official '{0}'({1})",
                                      cells[2].InnerText, sEmail), StatusBox.StatusRpt.MSGT.Error);
                    continue;
                    }

                m_srpt.LogData(String.Format("Updating last access for official '{0}', {1}", rste.Name, sSignedIn), 5,
                                  StatusBox.StatusRpt.MSGT.Body);
                rste.m_sLastSignin = sSignedIn;
                }
        }

        delegate Roster ProcessQuickRosterOfficialsDel(string sDownloadedRoster, bool fIncludeRankings, bool fIncludeLastAccess);

        Roster DoProcessQuickRosterOfficials(string sDownloadedRoster, bool fIncludeRankings, bool fIncludeLastAccess)
        {
            Roster rstBuilding = new Roster();

            rstBuilding.ReadRoster(sDownloadedRoster);

            // always fixup the email addresses for non-joined officials
            {
	            MicroTimer timer = new MicroTimer();
	            
                HandleGenericRoster gr = new HandleGenericRoster(
                    this,
                    !m_cbRankOnly.Checked,
                    m_cbAddOfficialsOnly.Checked,
                    null,
                    null,
                    null);
                gr.ProcessAllOfficialPages(VOPC_FixupNonJoinedEmailAddress, rstBuilding);

                timer.Stop();
                m_srpt.LogData($"ProcessAllOfficialPages For EmailFixup elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
            }

            if (fIncludeLastAccess)
            {
	            MicroTimer timer = new MicroTimer();
	            HandleGenericRoster gr = new HandleGenericRoster(
		            this,
		            !m_cbRankOnly.Checked,
		            m_cbAddOfficialsOnly.Checked,
		            null,
		            null,
		            null);

	            gr.ProcessAllOfficialPages(VOPC_UpdateLastAccess, rstBuilding);
	            
	            timer.Stop();
	            m_srpt.LogData($"ProcessAllOfficialPages For LastAccess elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
            }

            if (fIncludeRankings)
            {
	            MicroTimer timer = new MicroTimer();
	            
	            HandleRankings(null, rstBuilding);
	            
	            timer.Stop();
	            m_srpt.LogData($"Handle rankings elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
            }

            return rstBuilding;
        }

        private Roster RosterQuickBuildFromDownloadedRoster(string sDownloadedRoster, bool fIncludeRankings, bool fIncludeLastAccess)
        {
            return DoProcessQuickRosterOfficials(sDownloadedRoster, fIncludeRankings, fIncludeLastAccess);
        }


        /*----------------------------------------------------------------------------
        	%%Function: DownloadQuickRosterToFile
        	%%Qualified: ArbWeb.AwMainForm.DownloadQuickRosterToFile
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        string DownloadQuickRosterToFile()
        {
            m_srpt.AddMessage("Starting Quick Roster download to temp file...");
            m_srpt.PushLevel();

            PushCursor(Cursors.WaitCursor);
            string sTempFile = @"C:\Users\rlittle\AppData\Local\Temp\tempcd316a92-2120-4234-9427-7d3965076999.csv";
            // SRosterFileDownload();

            PopCursor();
            m_srpt.PopLevel();
            return sTempFile;
        }

        /* D O  D O W N L O A D  Q U I C K  R O S T E R  W O R K */
        /*----------------------------------------------------------------------------
        	%%Function: DoDownloadQuickRosterWork
        	%%Qualified: ArbWeb.AwMainForm.DoDownloadQuickRosterWork
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        Roster DoDownloadQuickRosterWork()
        {
            m_srpt.AddMessage("Starting Quick Roster download...");
            m_srpt.PushLevel();

            string sTempFile = DownloadQuickRosterToFile();

            // now, update the last access date and fetch the rankings and update the last access date
            Roster rst = RosterQuickBuildFromDownloadedRoster(sTempFile, true, true);

            m_srpt.PopLevel();

            m_srpt.AddMessage("Completed Quick Roster download.");
            return rst;
        }

        Roster DoDownloadQuickRosterOfficialsOnlyWork()
        {
            m_srpt.AddMessage("Starting Quick Roster download (officials only, no rankings)...");
            m_srpt.PushLevel();

            string sTempFile = DownloadQuickRosterToFile();

            // now, update the last access date and fetch the rankings and update the last access date
            Roster rst = RosterQuickBuildFromDownloadedRoster(sTempFile, false, false);

            m_srpt.PopLevel();

            m_srpt.AddMessage("Completed Quick Roster download.");
            return rst;
        }

        private delegate AutoResetEvent LaunchRosterFileDownloadDel(string sTempFile);

        AutoResetEvent LaunchRosterFileDownload(string sTempFile)
        {
			return DoLaunchRosterFileDownload(sTempFile);
        }

        void DoRosterDownload(string sTempFile)
        {
            ThrowIfNot(m_webControl.FNavToPage(WebCore._s_Page_OfficialsView), "Couldn't nav to officials view!");
            
            // now we are on the PrintRoster screen
            ThrowIfNot(m_webControl.FClickControlId(WebCore._sid_OfficialsView_PrintCustomRoster, WebCore._sid_CustomRosterPrint_UserFilter), "Can't click on roster control");
            // check a whole bunch of config checkboxes

            // select All Officials
            ArbWebControl_Selenium.FSetSelectControlText(m_webControl.Driver, this, WebCore._sid_CustomRosterPrint_UserFilter, "All Officials");

            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_DateJoined);
            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_OfficialNumber);
            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_DateOfBirth);
            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_UserID);
            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_MiddleName);

            m_webControl.FClickControlId(WebCore._sid_CustomRosterPrint_CustomFieldListDropdown); // dropdown the menu
            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_SelectAllCustomFields);
            m_webControl.FClickControlId(WebCore._sid_CustomRosterPrint_CustomFieldListDropdown); // dismiss the menu

            m_webControl.FClickControlId(WebCore._sid_CustomRosterPrint_GenerateRosterReport);
            
            // now wait for the file to be available and non-zero
            string sExpectedFile = Path.Combine(m_webControl.DownloadPath, "RosterReport.xlsx");
            int cRetry = 100;
            while (--cRetry > 0)
            {
	            Thread.Sleep(100);
	            if (File.Exists(sExpectedFile))
	            {
		            FileInfo info = new FileInfo(sExpectedFile);

		            if (info.Length > 0)
			            break;
	            }
            }

            if (cRetry <= 0)
	            throw new Exception("file never downloaded?");
            
            DownloadGenericExcelReport.ConvertExcelFileToCsv(sExpectedFile, sTempFile);
            File.Delete(sExpectedFile);
        }

        AutoResetEvent DoLaunchRosterFileDownload(string sTempFile)
        {
            ThrowIfNot(m_webControl.FNavToPage(WebCore._s_Page_OfficialsView), "Couldn't nav to officials view!");
            
            // now we are on the PrintRoster screen
            ThrowIfNot(m_webControl.FClickControlId(WebCore._sid_OfficialsView_PrintCustomRoster, WebCore._sid_CustomRosterPrint_UserFilter), "Can't click on roster control");

            ArbWebControl_Selenium.FSetSelectControlText(m_webControl.Driver, this, WebCore._sid_CustomRosterPrint_UserFilter, "All Officials");

            // check a whole bunch of config checkboxes
            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_DateJoined);

            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_DateJoined);
            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_OfficialNumber);
            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_DateOfBirth);
            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_UserID);
            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_MiddleName);

            m_webControl.FClickControlId(WebCore._sid_CustomRosterPrint_CustomFieldListDropdown); // dropdown the menu
            ArbWebControl_Selenium.FSetCheckboxControlIdVal(m_webControl.Driver, true, WebCore._sid_CustomRosterPrint_SelectAllCustomFields);
            m_webControl.FClickControlId(WebCore._sid_CustomRosterPrint_CustomFieldListDropdown); // dismiss the menu

            AutoResetEvent evtDownload = new AutoResetEvent(false);
            Win32Win.TrapFileDownload aww = new TrapFileDownload(m_srpt, "RosterReport.xlsx", "RosterReport", sTempFile, "of OfficialsView.aspx from", evtDownload);

            m_webControl.FClickControlId(WebCore._sid_CustomRosterPrint_GenerateRosterReport);
            return evtDownload;
        }

        private string SRosterFileDownload()
        {
            // navigate to the officials page...
            EnsureLoggedIn();

            string sTempFile = String.Format("{0}\\temp{1}.csv", Environment.GetEnvironmentVariable("Temp"),
                                             System.Guid.NewGuid().ToString());


            DoRosterDownload(sTempFile);
            
            return sTempFile;
        }


        private void HandleRosterPostUpdateForDownload(HandleGenericRoster gr, IRoster irstBuilding)
        {
            // get the last login date from the officials main page
            gr.NavigateOfficialsPageAllOfficials();
            throw new Exception("NYI");

            // gr.ProcessAllOfficialPages(VOPC_UpdateLastAccess, irstBuilding);
        }

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
        	%%Qualified: ArbWeb.AwMainForm.DoDownloadRoster
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void DoDownloadRoster()
        {
            m_srpt.AddMessage("Starting FULL Roster download...");
            m_srpt.PushLevel();

            PushCursor(Cursors.WaitCursor);
            string sOutFile = HandleGenericRoster.SBuildRosterFilename(m_pr.Roster);

            m_pr.Roster = sOutFile;
            HandleGenericRoster gr = new HandleGenericRoster(
                this, 
                !m_cbRankOnly.Checked, // fNeedPass1OnUpload
                m_cbAddOfficialsOnly.Checked, // only add officials
                new HandleGenericRoster.delDoPass1Visit(HandleRosterPass1VisitForUploadDownload), 
                new HandleGenericRoster.delAddOfficials(AddOfficials),
                new HandleGenericRoster.delDoPostHandleRoster(HandleRankings)
                );

            Roster rstBuilding = new Roster();
            gr.GenericVisitRoster(null, rstBuilding, sOutFile, null, HandleRosterPostUpdateForDownload);
            PopCursor();
            m_srpt.PopLevel();
            System.IO.File.Delete(m_pr.RosterWorking);
            System.IO.File.Copy(sOutFile, m_pr.RosterWorking);
            m_srpt.AddMessage("Completed FULL Roster download.");
        }

        /*----------------------------------------------------------------------------
        	%%Function: DoDownloadQuickRoster
        	%%Qualified: ArbWeb.AwMainForm.DoDownloadQuickRoster
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        async void DoDownloadQuickRoster()
        {
            var x = m_awc.Handle; // this makes sure that m_awc has a handle before we ask it if invoke is required.(forces it to get created on the correct thread)

            Task<Roster> tsk = new Task<Roster>(DoDownloadQuickRosterWork);

            tsk.Start();

            string sOutFile = HandleGenericRoster.SBuildRosterFilename(m_pr.Roster);
            m_pr.Roster = sOutFile;

            Roster rst = await tsk;

            rst.WriteRoster(sOutFile);
            System.IO.File.Delete(m_pr.RosterWorking);
            System.IO.File.Copy(sOutFile, m_pr.RosterWorking);
        }
    }
}
