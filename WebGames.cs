using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using System.Net;
using System.Threading.Tasks;
using mshtml;
using StatusBox;
using TCore.Util;
using Win32Win;

namespace ArbWeb
{
    /// <summary>
    /// Summary description for AwMainForm.
    /// </summary>
    public partial class AwMainForm : System.Windows.Forms.Form
    {
        /*----------------------------------------------------------------------------
        	%%Function: TestDownload
        	%%Qualified: ArbWeb.AwMainForm.TestDownload
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void TestDownload()
        {
            m_srpt.AddMessage("Starting test download...");
            m_srpt.PushLevel();
            string sTempFile = Filename.SBuildTempFilename("temp", "xls");

            sTempFile = TestDownload(sTempFile, "http://thetasoft2.azurewebsites.net/rwp/TeamsReport.aspx");

            m_srpt.PopLevel();
            m_srpt.AddMessage("Completed test download.");
            DoPendingQueueUIOp();

            return;
        }

        /*----------------------------------------------------------------------------
        	%%Function: DownloadGames
        	%%Qualified: ArbWeb.AwMainForm.DownloadGames
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DownloadGames(string sFilterReq)
        {
            m_srpt.AddMessage("Starting games download...");
            m_srpt.PushLevel();
            string sTempFile = Filename.SBuildTempFilename("temp", "xls");

            sTempFile = DownloadGamesToFile(sTempFile, sFilterReq);
            HandleDownloadGames(sTempFile);

            System.IO.File.Delete(sTempFile);

//          m_awc.PopNewWindow3Delegate();

            // set the view to "all games"
//            fNavDone = false;
//            FSetSelectControlText(oDoc2, "pgeGamesView$conGamesView$ddlSavedFilters", "All Games", false);
//            FWaitForNavFinish();

            // ok, now we have all games selected...
            // time to try to download a report

            m_srpt.PopLevel();
            m_srpt.AddMessage("Completed downloading games.");
            DoPendingQueueUIOp();

            return;
        }

        Dictionary<string, string> MpFetchGameFilters()
        {
            if (!m_awc.FNavToPage(WebCore._s_Assigning))
                throw (new Exception("could not navigate to games view"));

            return ArbWebControl.MpGetSelectValues(m_srpt, m_awc.Document2, WebCore._s_Assigning_Select_Filters);
        }

        /*----------------------------------------------------------------------------
        	%%Function: TestDownload
        	%%Qualified: ArbWeb.AwMainForm.TestDownload
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private string TestDownload(string sTempFile, string sTestAddress)
        {
            m_srpt.LogData("LaunchTestDownload async task launched", 3, StatusRpt.MSGT.Body);
            var evtDownload = LaunchTestDownload(sTempFile, sTestAddress);
            m_srpt.LogData("Before evtDownload.Wait()", 3, StatusRpt.MSGT.Body);
            evtDownload.WaitOne();
            m_srpt.LogData("evtDownload.WaitOne() complete", 3, StatusRpt.MSGT.Body);

            return sTempFile;
        }

        /*----------------------------------------------------------------------------
        	%%Function: DownloadGamesToFile
        	%%Qualified: ArbWeb.AwMainForm.DownloadGamesToFile
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private string DownloadGamesToFile(string sTempFile, string sFilterReq)
        {
            EnsureLoggedIn();

            m_srpt.LogData("LaunchDownloadGames async task launched", 3, StatusRpt.MSGT.Body);
            var evtDownload = LaunchDownloadGames(sTempFile, sFilterReq);
            m_srpt.LogData("Before evtDownload.Wait()", 3, StatusRpt.MSGT.Body);
            evtDownload.WaitOne();
            m_srpt.LogData("evtDownload.WaitOne() complete", 3, StatusRpt.MSGT.Body);
//                {
                //Application.DoEvents();
                //Thread.Sleep(500);
                //Application.DoEvents();
                //}

#if nomore
            if (MessageBox.Show(String.Format("Please download the games file to {0}. This path is on the clipboard, so you can just paste it into the file/save dialog when you click Save.\n\nWhen the download is complete, click OK.\n\n(NOTE: Click CANCEL if you need the clipboard reset with the pathname)",
                                              sTempFile), "ArbWeb", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                {
                System.Windows.Forms.Clipboard.SetText(sTempFile);
                MessageBox.Show(String.Format("Please download the games file to {0}. This path is on the clipboard, so you can just paste it into the file/save dialog when you click Save.\n\nWhen the download is complete, click OK.",
                                              sTempFile), "ArbWeb", MessageBoxButtons.OKCancel);
                }
#endif // nomore

            return sTempFile;
        }

        private delegate AutoResetEvent LaunchTestDownloadDel(string sTempFile, string sTestAddress);

        /*----------------------------------------------------------------------------
        	%%Function: LaunchTestDownload
        	%%Qualified: ArbWeb.AwMainForm.LaunchTestDownload
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        AutoResetEvent LaunchTestDownload(string sTempFile, string sTestAddress)
        {
            if (m_awc.InvokeRequired)
            {
                m_srpt.LogData("InvokeRequired true for DoLaunchTestDownload", 3, StatusRpt.MSGT.Body);

                IAsyncResult rsl = m_awc.BeginInvoke(new LaunchTestDownloadDel(DoLaunchTestDownload), sTempFile, sTestAddress);
                return (AutoResetEvent)m_awc.EndInvoke(rsl);
            }
            else
            {
                m_srpt.LogData("InvokeRequired false for DoLaunchTestDownload", 3, StatusRpt.MSGT.Body);
                return DoLaunchTestDownload(sTempFile, sTestAddress);
            }
        }

        /*----------------------------------------------------------------------------
        	%%Function: DoLaunchTestDownload
        	%%Qualified: ArbWeb.AwMainForm.DoLaunchTestDownload
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private AutoResetEvent DoLaunchTestDownload(string sTempFile, string sTestAddress)
        {
            m_srpt.LogData(String.Format("Setting clipboard data: {0}", sTempFile), 3, StatusRpt.MSGT.Body);
            System.Windows.Forms.Clipboard.SetText(sTempFile);

            m_awc.ResetNav();
            
            AutoResetEvent evtDownload = new AutoResetEvent(false);

            m_srpt.LogData("Setting up TrapFileDownload", 3, StatusRpt.MSGT.Body);

            Win32Win.TrapFileDownload aww = new TrapFileDownload(m_srpt, "Teams.csv", "Teams", sTempFile, null, evtDownload);
            if (!m_awc.FNavToPage(sTestAddress))
                throw (new Exception("could not navigate to the test page!"));

            return evtDownload;
        }

        private delegate AutoResetEvent LaunchDownloadGamesDel(string sTempFile, string sFilterReq);

        /*----------------------------------------------------------------------------
        	%%Function: LaunchDownloadGames
        	%%Qualified: ArbWeb.AwMainForm.LaunchDownloadGames
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        AutoResetEvent LaunchDownloadGames(string sTempFile, string sFilterReq)
        {
            if (m_awc.InvokeRequired)
                {
                m_srpt.LogData("InvokeRequired true for DoLaunchDownloadGames", 3, StatusRpt.MSGT.Body);

                IAsyncResult rsl = m_awc.BeginInvoke(new LaunchDownloadGamesDel(DoLaunchDownloadGames), sTempFile, sFilterReq);
                return (AutoResetEvent) m_awc.EndInvoke(rsl);
                }
            else
                {
                m_srpt.LogData("InvokeRequired false for DoLaunchDownloadGames", 3, StatusRpt.MSGT.Body);
                return DoLaunchDownloadGames(sTempFile, sFilterReq);
                }
        }

        /*----------------------------------------------------------------------------
        	%%Function: DoLaunchDownloadGames
        	%%Qualified: ArbWeb.AwMainForm.DoLaunchDownloadGames
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private AutoResetEvent DoLaunchDownloadGames(string sTempFile, string sFilterReq)
        {
            IHTMLDocument2 oDoc2 = m_awc.Document2;
            int count = 0;
            string sFilter = null;

            while (count < 2)
                {
                // ok, now we're at the main assigner page...
                if (!m_awc.FNavToPage(WebCore._s_Assigning))
                    throw (new Exception("could not navigate to games view"));

                oDoc2 = m_awc.Document2;
                sFilter = m_awc.SGetFilterID(oDoc2, WebCore._s_Assigning_Select_Filters, sFilterReq);
                if (sFilter != null)
                    break;

                count++;
                }

            if (sFilter == null)
                throw (new Exception("there is no 'all games' filter"));

            // now set that filter

            m_awc.ResetNav();
            m_awc.FSetSelectControlText(oDoc2, WebCore._s_Assigning_Select_Filters, null, sFilterReq, false);
            m_awc.FWaitForNavFinish();

            if (!m_awc.FNavToPage(WebCore._s_Assigning_PrintAddress + sFilter))
                throw (new Exception("could not navigate to the reports page!"));

            // setup the file formats and go!

            oDoc2 = m_awc.Document2;
            m_awc.FSetSelectControlText(oDoc2, WebCore._s_Assigning_Reports_Select_Format, WebCore._sid_Assigning_Reports_Select_Format, "Excel Worksheet Format (.xls)", false);

            m_srpt.LogData(String.Format("Setting clipboard data: {0}", sTempFile), 3, StatusRpt.MSGT.Body);
            System.Windows.Forms.Clipboard.SetText(sTempFile);

            m_awc.ResetNav();
            //          m_awc.PushNewWindow3Delegate(new DWebBrowserEvents2_NewWindow3EventHandler(DownloadGamesNewWindowDelegate));

            AutoResetEvent evtDownload = new AutoResetEvent(false);

            m_srpt.LogData("Setting up TrapFileDownload", 3, StatusRpt.MSGT.Body);
            Win32Win.TrapFileDownload aww = new TrapFileDownload(m_srpt, "Schedule.xls", "Schedule", sTempFile, null, evtDownload);
            m_awc.FClickControlNoWait(oDoc2, WebCore._s_Assigning_Reports_Submit_Print);
            return evtDownload;
        }

        private delegate void SetTextDel(TextBox eb, string s);

        private void DoSetText(TextBox eb, string s)
        {
            eb.Text = s;
        }

        /*----------------------------------------------------------------------------
        	%%Function: SetText
        	%%Qualified: ArbWeb.AwMainForm.SetText
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void SetText(TextBox eb, string s)
        {
            if (eb.InvokeRequired)
                eb.BeginInvoke(new SetTextDel(DoSetText), new object[] {eb, s});
            else
                DoSetText(eb, s);
        }

        /*----------------------------------------------------------------------------
        	%%Function: HandleDownloadGames
        	%%Qualified: ArbWeb.AwMainForm.HandleDownloadGames
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void HandleDownloadGames(string sFile)
        {
            object missing = System.Type.Missing;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook wkb;

            wkb = app.Workbooks.Open(sFile, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);

            string sOutFile = "";
            string sPrefix = "";

            if (m_pr.GameFile.Length < 1)
                {
                sOutFile = String.Format("{0}", Environment.GetEnvironmentVariable("temp"));
                }
            else
                {
                sOutFile = System.IO.Path.GetDirectoryName(m_pr.GameFile);
                string[] rgs;
                if (m_pr.GameFile.Length > 5 && sOutFile.Length > 0)
                    {
                    rgs = CountsData.RexHelper.RgsMatch(m_pr.GameFile.Substring(sOutFile.Length + 1), "([.*])games");
                    if (rgs != null && rgs.Length > 0 && rgs[0] != null)
                        sPrefix = rgs[0];
                    }
                }


            sOutFile = String.Format("{0}\\{2}games_{1:MM}{1:dd}{1:yy}_{1:HH}{1:mm}.csv", sOutFile, DateTime.Now, sPrefix);

            if (wkb != null)
                {
                wkb.SaveAs(sOutFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                wkb.Close(0, missing, missing);
                }
            app.Quit();
            app = null;
            m_pr.GameFile = sOutFile;
            System.IO.File.Delete(m_pr.GameCopy);
            System.IO.File.Copy(sOutFile, m_pr.GameCopy);
        }

        private void contextMenu1_Popup(object sender, System.EventArgs e)
        {

        }

		private void InvalGameCount()
		{
			m_gc = null;
		}

		private CountsData GcEnsure(string sRoster, string sGameFile, bool fIncludeCanceled)
		{
			if (m_gc != null)
				return m_gc;

			CountsData gc = new CountsData(m_srpt);

			gc.LoadData(sRoster, sGameFile, fIncludeCanceled, Int32.Parse(m_ebAffiliationIndex.Text));
			m_gc = gc;
			return gc;
		}

        /*----------------------------------------------------------------------------
        	%%Function: DoDownloadGames
        	%%Qualified: ArbWeb.AwMainForm.DoDownloadGames
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void DoDownloadGames()
        {
            var x = m_awc.Handle;
            string sFilterReq = (string)m_cbxGameFilter.SelectedItem;
            if (sFilterReq == null)
                sFilterReq = "All Games";

            // let's make sure the webbrowser handle is created

            m_srpt.LogData("Starting DoDownloadGames", 3, StatusRpt.MSGT.Header);

            Task tskDownloadGames = new Task(() => DownloadGames(sFilterReq));

            tskDownloadGames.Start();
        }

    }
}