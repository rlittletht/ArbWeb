using System;
using System.Diagnostics;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Net;
using System.IO;
using Microsoft.Win32;
using AxSHDocVw;
using StatusBox;
using mshtml;
using System.Text.RegularExpressions;
using Microsoft.Office;
using System.Runtime.InteropServices;
using Outlook=Microsoft.Office.Interop.Outlook;
using Excel=Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;

namespace ArbWeb
{
    /// <summary>
    /// Summary description for AwMainForm.
    /// </summary>
    public partial class AwMainForm : System.Windows.Forms.Form
    {

        private void DownloadGames()
        {
            m_srpt.AddMessage("Starting games download...");
            m_srpt.PushLevel();
            string sTempFile = SBuildTempFilename("temp", "xls");

            sTempFile = DownloadGamesToFile(sTempFile);
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
            return;
        }

        private string DownloadGamesToFile(string sTempFile)
        {
            IHTMLDocument2 oDoc2 = null;

            EnsureLoggedIn();

            int count = 0;
            string sFilter = null;
            string sFilterReq = m_cbFutureOnly.Checked ? "Future Games" : "All Games";

            while (count < 2)
                {
                // ok, now we're at the main assigner page...
                if (!m_awc.FNavToPage(_s_Assigning))
                    throw (new Exception("could not navigate to games view"));

                oDoc2 = m_awc.Document2;
                sFilter = ArbWebControl.SGetFilterID(oDoc2, _s_Assigning_Select_Filters, sFilterReq);
                if (sFilter != null)
                    break;

                count++;
                }

            if (sFilter == null)
                throw (new Exception("there is no 'all games' filter"));

            // now set that filter

            m_awc.ResetNav();
            ArbWebControl.FSetSelectControlText(oDoc2, _s_Assigning_Select_Filters, sFilterReq, false);
            m_awc.FWaitForNavFinish();

            if (!m_awc.FNavToPage(_s_Assigning_PrintAddress + sFilter))
                throw (new Exception("could not navigate to the reports page!"));

            // setup the file formats and go!

            oDoc2 = m_awc.Document2;
            ArbWebControl.FSetSelectControlText(oDoc2, _s_Assigning_Reports_Select_Format, "Excel Worksheet Format (.xls)", false);


            System.Windows.Forms.Clipboard.SetText(sTempFile);

            m_awc.ResetNav();
            //          m_awc.PushNewWindow3Delegate(new DWebBrowserEvents2_NewWindow3EventHandler(DownloadGamesNewWindowDelegate));

            ArbWebControl.FClickControlNoWait(oDoc2, _s_Assigning_Reports_Submit_Print);

            if (MessageBox.Show(String.Format("Please download the games file to {0}. This path is on the clipboard, so you can just paste it into the file/save dialog when you click Save.\n\nWhen the download is complete, click OK.\n\n(NOTE: Click CANCEL if you need the clipboard reset with the pathname)",
                                              sTempFile), "ArbWeb", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                {
                System.Windows.Forms.Clipboard.SetText(sTempFile);
                MessageBox.Show(String.Format("Please download the games file to {0}. This path is on the clipboard, so you can just paste it into the file/save dialog when you click Save.\n\nWhen the download is complete, click OK.",
                                              sTempFile), "ArbWeb", MessageBoxButtons.OKCancel);
                }
            return sTempFile;
        }

        private void HandleDownloadGames(string sFile)
        {
            object missing = System.Type.Missing;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook wkb;

            wkb = app.Workbooks.Open(sFile, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);

            string sOutFile = "";
            string sPrefix = "";

            if (m_ebGameFile.Text.Length < 1)
                {
                sOutFile = String.Format("{0}", Environment.GetEnvironmentVariable("temp"));
                }
            else
                {
                sOutFile = System.IO.Path.GetDirectoryName(m_ebGameFile.Text);
                string[] rgs;
                if (m_ebGameFile.Text.Length > 5 && sOutFile.Length > 0)
                    {
                    rgs = CountsData.RexHelper.RgsMatch(m_ebGameFile.Text.Substring(sOutFile.Length + 1), "([.*])games");
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
            m_ebGameFile.Text = sOutFile;
            System.IO.File.Delete(m_ebGameCopy.Text);
            System.IO.File.Copy(sOutFile, m_ebGameCopy.Text);
        }

        private void DownloadGamesNewWindowDelegate(object sender, DWebBrowserEvents2_NewWindow3Event e)
        {
            // at this point, e.bstrUrlContext has the URL to the XLS schedule file!!!
            WebClient wc = new WebClient();

            string sFile = String.Format("{0}\\{1}", Environment.GetEnvironmentVariable("temp"), System.Guid.NewGuid().ToString());
            wc.DownloadFile(e.bstrUrl, sFile);
            HandleDownloadGames(sFile);
            System.IO.File.Delete(sFile);
        }

        private void DownloadQuickRosterNewWindowDelegate(object sender, DWebBrowserEvents2_NewWindow3Event e)
        {
            // at this point, e.bstrUrlContext has the URL to the CSV schedule file!!!
            WebClient wc = new WebClient();
            object missing = System.Type.Missing;

            // copy the file directly to the output filenames
            string sFile = m_ebRoster.Text;
            wc.DownloadFile(e.bstrUrl, sFile);
        }
#if no
		private void TriggerDocumentDone(object sender, AxSHDocVw.DWebBrowserEvents2_DocumentCompleteEvent e)
		{
			fNavDone = true;
		}
		bool m_fInternalBrowserChange = false;

		private void ShowBrowserStateChange(object sender, System.EventArgs e) {
			if (!m_fInternalBrowserChange)
				m_axWebBrowser1.Visible = true;
			
		}
#endif // no
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

    }
}