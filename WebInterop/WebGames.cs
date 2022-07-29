using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Threading.Tasks;
using TCore.StatusBox;
using TCore.WebControl;

namespace ArbWeb
{
    /// <summary>
    /// Summary description for AwMainForm.
    /// </summary>
    public class WebGames
    {
	    public IAppContext m_appContext;

	    private WebControl m_webControl => m_appContext.WebControl;
	    private IStatusReporter m_srpt => m_appContext.StatusReport;
	    
	    public WebGames(IAppContext appContext)
	    {
		    m_appContext = appContext;
	    }
	    
        /*----------------------------------------------------------------------------
			%%Function:MpFetchGameFilters
			%%Qualified:ArbWeb.AwMainForm.MpFetchGameFilters
        ----------------------------------------------------------------------------*/
        public Dictionary<string, string> FetchOptionValueTextMapForGameFilter()
        {
            if (!m_webControl.FNavToPage(WebCore._s_Assigning))
                throw (new Exception("could not navigate to games view"));

            return m_webControl.GetOptionsValueTextMappingFromControlId(WebCore._sid_Assigning_Select_Filters);
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
        	%%Function: DoDownloadGames
        	%%Qualified: ArbWeb.AwMainForm.DoDownloadGames
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public void DoDownloadGames(string sFilterOptionTextReq)
        {
            if (sFilterOptionTextReq == null)
                sFilterOptionTextReq = "All Games";

            // let's make sure the webbrowser handle is created

            m_srpt.LogData("Starting DoDownloadGames", 3, MSGT.Header);

            DownloadGenericExcelReport dg =
                new DownloadGenericExcelReport(
                    sFilterOptionTextReq,
                    "games",
                    WebCore._s_Assigning,
                    WebCore._s_Assigning_Select_Filters,
                    WebCore._sid_Assigning_Select_Filters,
                    WebCore._s_Assigning_PrintAddress,
                    WebCore._s_Assigning_Reports_Submit_Print,
                    new [] {"Schedule.xls", "Schedule{0}.xls", "Schedule({0}).xls", "Schedule ({0}).xls"},
                    "Schedule",
                    new[]
                        {
                        new DownloadGenericExcelReport.ControlSetting<string>(WebCore._s_Assigning_Reports_Select_Format,
                                                                              WebCore._sid_Assigning_Reports_Select_Format,
                                                                              "Excel Worksheet Format (.xls)")
                        },
                    m_appContext.Profile.GameFile,
                    m_appContext.Profile.GameCopy,
                    m_appContext);

            Task tskDownloadGames = new Task(
	            () =>
	            {
		            dg.DownloadGeneric(out var sGameFileNew);
		            m_appContext.Profile.GameFile = sGameFileNew;
		            m_appContext.DoPendingQueueUIOp();
	            });

            tskDownloadGames.Start();
        }

    }
}