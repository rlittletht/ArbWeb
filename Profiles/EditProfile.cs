﻿using System;
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
            m_lvBaseballSchedules.Columns.Add("Schedule");
            m_lvBaseballSchedules.Columns[0].Width = -1;
            m_lvSoftballSchedules.Columns.Add("Schedule");
            m_lvSoftballSchedules.Columns[0].Width = -1;
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
            pr.MergeCsvWorking = ep.m_ebMergeCsvWorking.Text;
            pr.MergeCsv = ep.m_ebMergeCsv.Text;
            pr.Announcements = ep.m_ebAnnouncements.Text;
            pr.AnnouncementsWorking = ep.m_ebAnnouncementsWorking.Text;

            pr.LogToFile = ep.m_cbLogToFile.Checked;
            pr.TestOnly = ep.m_cbTestOnly.Checked;
            pr.SkipZ = ep.m_cbIgnoreZSports.Checked;
            pr.DownloadRosterOnUpload = ep.m_cbDownloadRosterOnUpload.Checked;
            pr.LogLevel = Int32.Parse(ep.m_ebLogLevel.Text);
            pr.NoHonorificRanks = ep.m_cbNoHonorificRanks.Checked;
            pr.AllowAdvancedArbiterFunctions = ep.m_cbAdvancedFeatures.Checked;

            pr.SchedSpoSite = ep.m_ebSpoSite.Text;
            pr.SchedSpoSubsite = ep.m_ebSpoSubsite.Text;
            pr.SchedDownloadFolder = ep.m_ebSchedDownloadFolder.Text;
            pr.SchedWorkingFolder = ep.m_ebSchedWorkingFolder.Text;

            List<string> schedules = new List<string>();
            foreach (ListViewItem item in ep.m_lvBaseballSchedules.Items)
                schedules.Add(item.Text);

            pr.BaseballSchedFiles = schedules.ToArray();

            schedules.Clear();
            foreach (ListViewItem item in ep.m_lvSoftballSchedules.Items)
                schedules.Add(item.Text);

            pr.SoftballSchedFiles = schedules.ToArray();
        }


        public static bool FShowEditProfile(Profile pr)
        {
            EditProfile ep = new EditProfile();
            ep.m_ebUserID.Text = pr.UserID;
            ep.m_ebPassword.Text = pr.Password;
            ep.m_ebProfileName.Text = pr.ProfileName;
            ep.m_ebGameFile.Text = pr.GameFile;
            ep.m_ebRoster.Text = pr.Roster;
            ep.m_ebAnnouncements.Text = pr.Announcements;
            ep.m_ebAnnouncementsWorking.Text = pr.AnnouncementsWorking;

            ep.m_ebContacts.Text = pr.Contacts;
            ep.m_ebGameCopy.Text = pr.GameCopy;
            ep.m_ebRosterWorking.Text = pr.RosterWorking;
            ep.m_ebContactsWorking.Text = pr.ContactsWorking;
            ep.m_ebMergeCsv.Text = pr.MergeCsv;
            ep.m_ebMergeCsvWorking.Text = pr.MergeCsvWorking;

            ep.m_ebProfileName.Enabled = false;
            ep.m_cbLogToFile.Checked = pr.LogToFile;
            ep.m_cbTestOnly.Checked = pr.TestOnly;
            ep.m_cbIgnoreZSports.Checked = pr.SkipZ;
            ep.m_cbDownloadRosterOnUpload.Checked = pr.DownloadRosterOnUpload;
            ep.m_cbNoHonorificRanks.Checked = pr.NoHonorificRanks;
            ep.m_ebLogLevel.Text = pr.LogLevel.ToString();
            ep.m_ebSpoSite.Text = pr.SchedSpoSite;
            ep.m_ebSpoSubsite.Text = pr.SchedSpoSubsite;
            ep.m_ebSchedDownloadFolder.Text = pr.SchedDownloadFolder;
            ep.m_ebSchedWorkingFolder.Text = pr.SchedWorkingFolder;
            ep.m_cbAdvancedFeatures.Checked = pr.AllowAdvancedArbiterFunctions;

            ep.m_lvBaseballSchedules.Items.Clear();

            foreach (string s in pr.BaseballSchedFiles)
                ep.m_lvBaseballSchedules.Items.Add(s);

            ep.m_lvSoftballSchedules.Items.Clear();

            foreach (string s in pr.SoftballSchedFiles)
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
            string year = DateTime.Now.ToString("yyyy");

            ep.m_ebProfileName.Enabled = true;
            ep.m_ebGameFile.Text = $"c:\\baseball\\{year}\\arb\\archive\\games_010123_0000.csv";
            ep.m_ebGameCopy.Text = $"c:\\baseball\\{year}\\arb\\gamesLatest.csv";
            ep.m_ebRoster.Text = $"c:\\baseball\\{year}\\arb\\archive\\roster_010123_0000.csv";
            ep.m_ebRosterWorking.Text = $"c:\\baseball\\{year}\\arb\\rosterLatest.csv";
            ep.m_ebContacts.Text = $"c:\\baseball\\{year}\\arb\\archive\\contacts_010123_0000.csv";
            ep.m_ebContactsWorking.Text = $"c:\\baseball\\{year}\\arb\\contactsLatest.csv";
            ep.m_ebSchedDownloadFolder.Text = $"c:\\baseball\\{year}\\arb\\archive\\schedules";
            ep.m_ebSchedWorkingFolder.Text = $"c:\\baseball\\{year}\\arb";
            ep.m_ebMergeCsv.Text = $"c:\\baseball\\{year}\\arb\\archive\\merge_010123_0000.csv";
            ep.m_ebMergeCsvWorking.Text = $"c:\\baseball\\{year}\\arb\\mergeLatest.csv";
            ep.m_ebAnnouncements.Text = $"c:\\baseball\\{year}\\arb\\archive\\announcements_010123_0000.html";
            ep.m_ebAnnouncementsWorking.Text = $"c:\\baseball\\{year}\\arb\\announcementsLatest.html";
            ep.m_ebSpoSite.Text = "washdist9.sharepoint.com";
            ep.m_ebSpoSubsite.Text = "sched";
            ep.m_ebLogLevel.Text = "0";

            ep.m_cbDownloadRosterOnUpload.Checked = true;
            ep.m_cbNoHonorificRanks.Checked = true;

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
