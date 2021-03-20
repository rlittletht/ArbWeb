using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using System.Net;
using System.Threading.Tasks;
using mshtml;
using StatusBox;
using TCore.Util;

namespace ArbWeb
{
    /// <summary>
    /// Summary description for AwMainForm.
    /// </summary>
    public partial class AwMainForm : System.Windows.Forms.Form
    {
        /*----------------------------------------------------------------------------
			%%Function:MpFetchGameFilters
			%%Qualified:ArbWeb.AwMainForm.MpFetchGameFilters
        ----------------------------------------------------------------------------*/
        Dictionary<string, string> FetchOptionValueTextMapForGameFilter()
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
            string sFilterOptionTextReq = (string) m_cbxGameFilter.SelectedItem;
            if (sFilterOptionTextReq == null)
                sFilterOptionTextReq = "All Games";

            // let's make sure the webbrowser handle is created

            m_srpt.LogData("Starting DoDownloadGames", 3, StatusRpt.MSGT.Header);

            DownloadGenericExcelReport dg =
                new DownloadGenericExcelReport(
                    sFilterOptionTextReq,
                    "games",
                    WebCore._s_Assigning,
                    WebCore._s_Assigning_Select_Filters,
                    WebCore._sid_Assigning_Select_Filters,
                    WebCore._s_Assigning_PrintAddress,
                    WebCore._s_Assigning_Reports_Submit_Print,
                    "Schedule.xls",
                    "Schedule",
                    new[]
                        {
                        new DownloadGenericExcelReport.ControlSetting<string>(WebCore._s_Assigning_Reports_Select_Format,
                                                                              WebCore._sid_Assigning_Reports_Select_Format,
                                                                              "Excel Worksheet Format (.xls)")
                        },
                    Profile.GameFile,
                    Profile.GameCopy,
                    this);

            Task tskDownloadGames = new Task(() =>
                {
                string sGameFileNew;

                dg.DownloadGeneric(out sGameFileNew);
                Profile.GameFile = sGameFileNew;
                DoPendingQueueUIOp();
                });

            tskDownloadGames.Start();
        }

    }
}