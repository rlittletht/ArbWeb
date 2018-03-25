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

            DownloadGenericExcelReport dg = new DownloadGenericExcelReport(sFilterReq,
                                                                           "games", 
                                                                           WebCore._s_Assigning, 
                                                                           WebCore._s_Assigning_Select_Filters, 
                                                                           WebCore._s_Assigning_PrintAddress,
                                                                           WebCore._s_Assigning_Reports_Select_Format,
                                                                           WebCore._sid_Assigning_Reports_Select_Format,
                                                                           WebCore._s_Assigning_Reports_Submit_Print,
                                                                           this);

            Task tskDownloadGames = new Task(() =>
                {
                dg.DownloadGeneric();
                DoPendingQueueUIOp();
                });

            tskDownloadGames.Start();
        }

    }
}