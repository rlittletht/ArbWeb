using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ArbWeb.Announcements;

public partial class EditAnnouncement : Form
{
    public EditAnnouncement()
    {
        InitializeComponent();
    }

    private void BtnSave_Click(object sender, EventArgs e)
    {
        this.DialogResult = DialogResult.OK;
        this.Close();
    }

    private void BtnCancel_Click(object sender, EventArgs e)
    {
        this.DialogResult = DialogResult.Cancel;
        this.Close();
    }

    public void SelectOfficialsFilter(string officials)
    {
        for (int i = 0; i < m_officialsFilter.Items.Count; i++)
        {
            if (m_officialsFilter.Items[i].ToString() == officials)
            {
                m_officialsFilter.SelectedIndex = i;
                return;
            }
        }

        m_officialsFilter.SelectedIndex = 0;
    }

    public void SetFromAnnouncement(Announcement announcement)
    {
        txtContent.Text = announcement.ExtractEditHtml();
        m_announcementID.Text = announcement.AnnouncementID();
        m_showToAssigners.Checked = announcement.ShowAssigners;
        m_showToContacts.Checked = announcement.ShowContacts;
        if (announcement.Officials == "")
            m_officialsFilter.SelectedIndex = 0; // select "None"
        else
            SelectOfficialsFilter(announcement.Officials);
    }

    public Announcement GetAnnouncement()
    {
        Announcement announcement = new Announcement();
        announcement.AnnouncementHtml = announcement.BuildAnnounceHtml(m_announcementID.Text, txtContent.Text);
        announcement.ShowAssigners = m_showToAssigners.Checked;
        announcement.ShowContacts = m_showToContacts.Checked;
        announcement.Officials = m_officialsFilter.SelectedItem.ToString();

        return announcement;
    }

    public static Announcement DoEditAnnouncement(Announcement announcement)
    {
        using EditAnnouncement form = new EditAnnouncement();

        form.SetFromAnnouncement(announcement);

        if (form.ShowDialog() == DialogResult.OK)
        {
            return form.GetAnnouncement();
        }

        return null;
    }
}