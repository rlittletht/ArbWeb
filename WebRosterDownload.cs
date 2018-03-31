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
                || rste.m_sDateOfBirth == null || rste.Email == null || rste.First == null || rste.m_sGamesPerDay == null
                || rste.m_sGamesPerWeek == null || rste.Last == null || rste.m_sOfficialNumber == null
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
        private void VOPC_UpdateLastAccess(IHTMLDocument2 oDoc2, Object o)
        {
            Roster rstBuilding = (Roster)o;

            UpdateLastAccessFromCoreOfficialsPage(rstBuilding, oDoc2);
        }

        /*----------------------------------------------------------------------------
        	%%Function: UpdateLastAccessFromCoreOfficialsPage
        	%%Qualified: ArbWeb.AwMainForm.UpdateLastAccessFromCoreOfficialsPage
        	%%Contact: rlittle
        	
            Assuming we are on the core officials page...
        ----------------------------------------------------------------------------*/
        private void UpdateLastAccessFromCoreOfficialsPage(Roster rstBuilding, IHTMLDocument2 oDoc2)
        {
            IHTMLTable ihtbl;

            // misc field info.  every text input field is a misc field we want to save
            ihtbl = (IHTMLTable)oDoc2.all.item(WebCore._sid_OfficialsView_ContentTable, 0);

            foreach (IHTMLTableRow ihtr in ihtbl.rows)
                {
                IHTMLElement iheEmail = (IHTMLElement)ihtr.cells.item(3);
                IHTMLElement iheSignedIn = (IHTMLElement)ihtr.cells.item(4);

                if (iheEmail == null || iheSignedIn == null)
                    continue;

                string sEmail = iheEmail.innerText;
                string sSignedIn = iheSignedIn.innerText;

                RosterEntry rste = rstBuilding.RsteLookupEmail(sEmail);
                if (rste == null)
                    {
                    m_srpt.AddMessage(
                        String.Format("Lookup failed during ProcessAllOfficialPages for official '{0}'({1})",
                                      ((IHTMLElement)ihtr.cells.item(2)).innerText, sEmail), StatusBox.StatusRpt.MSGT.Error);
                    continue;
                    }

                m_srpt.AddMessage(String.Format("Updating last access for official '{0}', {1}", rste.Name, sSignedIn),
                                  StatusBox.StatusRpt.MSGT.Body);
                rste.m_sLastSignin = sSignedIn;
                }
        }

        delegate Roster ProcessQuickRosterOfficialsDel(string sDownloadedRoster, bool fIncludeRankings, bool fIncludeLastAccess);

        Roster DoProcessQuickRosterOfficials(string sDownloadedRoster, bool fIncludeRankings, bool fIncludeLastAccess)
        {
            Roster rstBuilding = new Roster();

            rstBuilding.ReadRoster(sDownloadedRoster);

            if (fIncludeLastAccess)
                {
                HandleGenericRoster gr = new HandleGenericRoster(
                    this,
                    !m_cbRankOnly.Checked,
                    m_cbAddOfficialsOnly.Checked,
                    null,
                    null,
                    null);
                gr.ProcessAllOfficialPages(VOPC_UpdateLastAccess, rstBuilding);
                }

            if (fIncludeRankings)
                HandleRankings(null, rstBuilding);

            return rstBuilding;
        }

        private Roster RosterQuickBuildFromDownloadedRoster(string sDownloadedRoster, bool fIncludeRankings, bool fIncludeLastAccess)
        {
            Roster rst;

            if (m_awc.InvokeRequired)
            {
                IAsyncResult rslt = m_awc.BeginInvoke(new ProcessQuickRosterOfficialsDel(DoProcessQuickRosterOfficials),
                                                      new object[] { sDownloadedRoster, fIncludeRankings, fIncludeLastAccess });
                rst = (Roster)m_awc.EndInvoke(rslt);
            }
            else
                rst = DoProcessQuickRosterOfficials(sDownloadedRoster, fIncludeRankings, fIncludeLastAccess);

            return rst;
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
            string sTempFile = SRosterFileDownload();

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
            if (m_awc.InvokeRequired)
            {
                IAsyncResult rslt = m_awc.BeginInvoke(new LaunchRosterFileDownloadDel(DoLaunchRosterFileDownload), new object[] { sTempFile });
                return (AutoResetEvent)m_awc.EndInvoke(rslt);
            }
            else
                return DoLaunchRosterFileDownload(sTempFile);
        }

        AutoResetEvent DoLaunchRosterFileDownload(string sTempFile)
        {
            System.Windows.Forms.Clipboard.SetText(sTempFile);

            IHTMLDocument2 oDoc2;
            m_awc.ResetNav();
            ThrowIfNot(m_awc.FNavToPage(WebCore._s_Page_OfficialsView), "Couldn't nav to officials view!");
            m_awc.FWaitForNavFinish();

            oDoc2 = m_awc.Document2;

            // from the officials view, make sure we are looking at active officials
            m_awc.ResetNav();
            m_awc.FSetSelectControlText(oDoc2, WebCore._s_OfficialsView_Select_Filter, WebCore._sid_OfficialsView_Select_Filter, "All Officials", true);
            m_awc.FWaitForNavFinish();

            oDoc2 = m_awc.Document2;
            // now we have all officials showing.  download the report

            // sometimes running the javascript takes a while, but the page isn't busy
            int cTry = 3;
            while (cTry > 0)
            {
                m_awc.ResetNav();
                m_awc.ReportNavState("Before click on PrintRoster: ");
                ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_OfficialsView_PrintRoster), "Can't click on roster control");
                m_awc.FWaitForNavFinish();

                oDoc2 = m_awc.Document2;
                if (ArbWebControl.FCheckForControl(oDoc2, WebCore._sid_RosterPrint_MergeStyle))
                    break;

                cTry--;
            }

            // now we are on the PrintRoster screen

            // clicking on the Merge Style control will cause a page refresh
            m_awc.ResetNav();
            ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_RosterPrint_MergeStyle), "Can't click on roster control");
            m_awc.FWaitForNavFinish();

            oDoc2 = m_awc.Document2;

            ThrowIfNot(ArbWebControl.FCheckForControl(oDoc2, WebCore._sid_RosterPrint_DateJoined),
                       "Couldn't find expected control on roster print config!");

            // check a whole bunch of config checkboxes
            ArbWebControl.FSetCheckboxControlVal(oDoc2, true, WebCore._s_RosterPrint_DateJoined);
            ArbWebControl.FSetCheckboxControlVal(oDoc2, true, WebCore._s_RosterPrint_OfficialNumber);
            ArbWebControl.FSetCheckboxControlVal(oDoc2, true, WebCore._s_RosterPrint_MiscFields);
            ArbWebControl.FSetCheckboxControlVal(oDoc2, true, WebCore._s_RosterPrint_NonPublicPhone);
            ArbWebControl.FSetCheckboxControlVal(oDoc2, true, WebCore._s_RosterPrint_NonPublicAddress);

            m_awc.ResetNav();


            AutoResetEvent evtDownload = new AutoResetEvent(false);
            Win32Win.TrapFileDownload aww = new TrapFileDownload(m_srpt, "roster.csv", "roster", sTempFile, "of OfficialsView.aspx from", evtDownload);

            ((IHTMLElement)(oDoc2.all.item(WebCore._sid_RosterPrint_BeginPrint, 0))).click();

            return evtDownload;
        }

        private string SRosterFileDownload()
        {
            // navigate to the officials page...
            EnsureLoggedIn();

            string sTempFile = String.Format("{0}\\temp{1}.csv", Environment.GetEnvironmentVariable("Temp"),
                                             System.Guid.NewGuid().ToString());


            var evtDownload = LaunchRosterFileDownload(sTempFile);
            evtDownload.WaitOne();
            return sTempFile;
        }


        private void HandleRosterPostUpdateForDownload(HandleGenericRoster gr, IRoster irstBuilding)
        {
            // get the last login date from the officials main page
            gr.NavigateOfficialsPageAllOfficials();
            gr.ProcessAllOfficialPages(VOPC_UpdateLastAccess, irstBuilding);
        }

        void HandleRosterPass1VisitForUploadDownload(
            string sEmail,
            string sOfficialID,
            IRoster irstUploading,
            IRoster irstServer,
            IRosterEntry irste,
            IRoster irstBuilding,
            bool fJustAdded,
            bool fMarkOnly)
        {
            if (!fMarkOnly)
                UpdateMisc(sEmail, sOfficialID, irstUploading, irstServer, (RosterEntry)irste, irstBuilding);

            if (!fJustAdded)
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
