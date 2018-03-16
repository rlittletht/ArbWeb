using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using StatusBox;
using mshtml;
using System.Text.RegularExpressions;
using NUnit.Framework;
using Win32Win;

namespace ArbWeb
{
    public partial class AwMainForm : System.Windows.Forms.Form
    {
        /*----------------------------------------------------------------------------
        	%%Function: DoUploadRosterWork
        	%%Qualified: ArbWeb.AwMainForm.DoUploadRosterWork
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private async void DoUploadRosterWork(string sRosterToUpload)
        {
            m_srpt.AddMessage("Starting Roster upload...");
            m_srpt.PushLevel();
            Roster rstServer = null;

            if (m_pr.DownloadRosterOnUpload)
            {
                // first thing to do is download a new (temporary) roster copy
                Task<Roster> tsk = new Task<Roster>(DoDownloadQuickRosterOfficialsOnlyWork);

                tsk.Start();
                rstServer = await tsk;
            }

            // now, check the roster we just downloaded against the roster we just already have

            InvalRoster();
            string sInFile = sRosterToUpload;

            Roster rst = RstEnsure(sInFile);

            if (rst.IsQuick && (!m_cbRankOnly.Checked || !rst.HasRankings))
            {
                //				MessageBox.Show("Cannot upload a quick roster.  Please perform a full roster download before uploading.\n\nIf you want to upload rankings only, please check 'Upload Rankings Only'");
                //    			m_srpt.PopLevel();
                m_srpt.AddMessage("Detected QuickRoster...", StatusBox.StatusRpt.MSGT.Warning);
            }

            // compare the two rosters to find differences

            if (m_awc.InvokeRequired)
            {
                IAsyncResult rslt = m_awc.BeginInvoke(new AwMainForm.HandleRosterDel(HandleRoster), new object[] { rst, sInFile, rstServer, null });
                m_awc.EndInvoke(rslt);
            }
            else
            {
                HandleRoster(rst, null, rstServer, null);
            }
            m_srpt.PopLevel();
            m_srpt.AddMessage("Completed Roster upload.");
        }

        string SGetRosterToUpload()
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.InitialDirectory = Path.GetDirectoryName(m_pr.RosterWorking);
            ofd.Filter = "CSV Files|*.csv";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                return ofd.FileName;
            }
            return null;
        }

        bool FLoadRosterToUpload()
        {
            return false;
        }

        /*----------------------------------------------------------------------------
        	%%Function: DoRosterUpload
        	%%Qualified: ArbWeb.AwMainForm.DoRosterUpload
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void DoRosterUpload()
        {
            string sRosterToUpload = SGetRosterToUpload();

            if (sRosterToUpload == null)
                return;


            Task tsk = new Task(() => DoUploadRosterWork(sRosterToUpload));

            tsk.Start();
        }
    }
}
