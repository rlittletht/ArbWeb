using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using TCore.UI;

namespace ArbWeb
{
    public partial class EditProfile : Form
    {
        public EditProfile()
        {
            InitializeComponent();
        }

        static void SetProfileFromUI(Profile pr, EditProfile ep)
        {
            pr.UserID = ep.m_ebUserID.Text;
            pr.Password = ep.m_ebPassword.Text;
            pr.GameFile = ep.m_ebGameFile.Text;
            pr.Roster = ep.m_ebRoster.Text;
            pr.Contacts = ep.m_ebContacts.Text;
            pr.GameCopy = ep.m_ebGameCopy.Text;
            pr.RosterWorking = ep.m_ebRosterWorking.Text;
            pr.ContactsWorking = ep.m_ebContactsWorking.Text;
            pr.LogToFile = ep.m_cbLogToFile.Checked;
            pr.TestOnly = ep.m_cbTestOnly.Checked;
            pr.SkipZ = ep.m_cbIgnoreZSports.Checked;
            pr.DownloadRosterOnUpload = ep.m_cbDownloadRosterOnUpload.Checked;
            pr.LogLevel = Int32.Parse(ep.m_ebLogLevel.Text);
            pr.SchedSpoSite = ep.m_ebSpoSite.Text;
            pr.SchedSpoSubsite = ep.m_ebSpoSubsite.Text;
            List<string> schedules = new List<string>();
            foreach (ListViewItem item in ep.m_lvBaseballSchedules.Items)
	            schedules.Add(item.Text);

            pr.BaseballScheduleFiles = schedules.ToArray();

            schedules.Clear();
            foreach (ListViewItem item in ep.m_lvSoftballSchedules.Items)
	            schedules.Add(item.Text);

            pr.SoftballScheduleFiles = schedules.ToArray();
        }

        
        public static bool FShowEditProfile(Profile pr)
        {
            EditProfile ep = new EditProfile();
            ep.m_ebUserID.Text = pr.UserID;
            ep.m_ebPassword.Text = pr.Password;
            ep.m_ebProfileName.Text = pr.ProfileName;
    	    ep.m_ebGameFile.Text = pr.GameFile;
            ep.m_ebRoster.Text = pr.Roster;
            ep.m_ebContacts.Text = pr.Contacts;
	        ep.m_ebGameCopy.Text = pr.GameCopy;
	        ep.m_ebRosterWorking.Text = pr.RosterWorking;
            ep.m_ebContactsWorking.Text = pr.ContactsWorking;
            ep.m_ebProfileName.Enabled = false;
            ep.m_cbLogToFile.Checked = pr.LogToFile;
            ep.m_cbTestOnly.Checked = pr.TestOnly;
            ep.m_cbIgnoreZSports.Checked = pr.SkipZ;
            ep.m_cbDownloadRosterOnUpload.Checked = pr.DownloadRosterOnUpload;
            ep.m_ebLogLevel.Text = pr.LogLevel.ToString();
            ep.m_ebSpoSite.Text = pr.SchedSpoSite;
            ep.m_ebSpoSubsite.Text = pr.SchedSpoSubsite;
            ep.m_lvBaseballSchedules.Items.Clear();

            foreach (string s in pr.BaseballScheduleFiles)
	            ep.m_lvBaseballSchedules.Items.Add(s);

            ep.m_lvSoftballSchedules.Items.Clear();

            foreach (string s in pr.SoftballScheduleFiles)
	            ep.m_lvSoftballSchedules.Items.Add(s);

            if (ep.ShowDialog() == DialogResult.OK)
                {
                SetProfileFromUI(pr, ep);
                pr.Save();
                return true;
                }
            return false;
        }

        public static string AddProfile()
        {
            EditProfile ep = new EditProfile();

            ep.m_ebProfileName.Enabled = true;

            if (ep.ShowDialog() == DialogResult.OK)
                {
                Profile pr = new Profile();
                pr.ProfileName = ep.m_ebProfileName.Text;
                SetProfileFromUI(pr, ep);

                pr.Save();
                return pr.ProfileName;
                }
            return null;
        }

        public enum FNC // FileName Control
            {
            GameFile,
            GameFile2,
            RosterFile,
            RosterFile2,
            ReportFile,
            AnalysisFile
            };
        
        /* E B  F R O M  F N C */
        /*----------------------------------------------------------------------------
        	%%Function: EbFromFnc
        	%%Qualified: ArbWeb.AwMainForm.EbFromFnc
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        TextBox EbFromFnc(FNC fnc)
        {
            switch (fnc)
                {
                case FNC.GameFile:
                    return m_ebGameFile;
                case FNC.GameFile2:
                    return m_ebGameCopy;
                case FNC.RosterFile:
                    return m_ebRoster;
                case FNC.RosterFile2:
                    return m_ebRosterWorking;
#if NOTHERE                
                case FNC.AnalysisFile:
                    return m_ebOutputFile;
                case FNC.ReportFile:
                    return m_ebGameOutput;
#endif // NOTHERE
                }
            return null;
        }
        
        /* D O  B R O W S E  O P E N */
        /*----------------------------------------------------------------------------
        	%%Function: DoBrowseOpen
        	%%Qualified: ArbWeb.AwMainForm.DoBrowseOpen
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoBrowseOpen(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            TextBox eb = EbFromFnc((FNC)(((Button)sender).Tag));
            
            ofd.InitialDirectory = Path.GetDirectoryName(eb.Text);
            if (ofd.ShowDialog() == DialogResult.OK)
                {
                eb.Text = ofd.FileName;
                }
        }

		private void AddBaseballScheduleFileItem(object sender, EventArgs e)
		{
			ListViewItem item = m_lvBaseballSchedules.Items.Add("");
			item.BeginEdit();
		}

		private void DeleteBaseballScheduleFileItem(object sender, EventArgs e)
		{
			foreach (int i in m_lvBaseballSchedules.SelectedIndices)
				m_lvBaseballSchedules.Items.RemoveAt(i);
		}

		private void RenderHeadingLine(object sender, PaintEventArgs e)
		{
			RenderSupp.RenderHeadingLine(sender, e);
		}

		private void AddSoftballScheduleFileItem(object sender, EventArgs e)
		{
			ListViewItem item = m_lvSoftballSchedules.Items.Add("");
			item.BeginEdit();
		}

		private void DeleteSoftballScheduleFileItem(object sender, EventArgs e)
		{
			foreach (int i in m_lvSoftballSchedules.SelectedIndices)
				m_lvSoftballSchedules.Items.RemoveAt(i);
		}
	}
}
