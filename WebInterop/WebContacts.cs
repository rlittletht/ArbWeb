using System.Windows.Forms;

namespace ArbWeb
{
    /// <summary>
    /// Summary description for AwMainForm.
    /// </summary>
    public class WebContacts
    {
        private readonly IAppContext m_appContext;

        /*----------------------------------------------------------------------------
			%%Function:WebContacts
			%%Qualified:ArbWeb.WebRoster.WebContacts
		----------------------------------------------------------------------------*/
        public WebContacts(IAppContext appContext)
        {
            m_appContext = appContext;
        }

        /*----------------------------------------------------------------------------
			%%Function:DoDownloadFullContacts
			%%Qualified:ArbWeb.WebContacts.DoDownloadFullContacts
        ----------------------------------------------------------------------------*/
        public void DoDownloadFullContacts()
        {
            m_appContext.StatusReport.AddMessage("Starting FULL Contact download...");
            m_appContext.StatusReport.PushLevel();

            m_appContext.PushCursor(Cursors.WaitCursor);
            string sOutFile = HandleGenericRoster.SBuildRosterFilename(m_appContext.Profile.Contacts);

            m_appContext.Profile.Contacts = sOutFile;
#if no
            HandleGenericRoster gr = new HandleGenericRoster(this,);

            gr.HandleRoster(null, sOutFile, null, HandleRosterPostUpdateForDownload);
            PopCursor();
            m_srpt.PopLevel();
            System.IO.File.Delete(m_pr.RosterWorking);
            System.IO.File.Copy(sOutFile, m_pr.RosterWorking);
            m_srpt.AddMessage("Completed FULL Roster download.");
#endif // no
        }

        /*----------------------------------------------------------------------------
			%%Function:DoDownloadContacts
			%%Qualified:ArbWeb.WebContacts.DoDownloadContacts
        ----------------------------------------------------------------------------*/
        public void DoDownloadContacts()
        {
#if NYI
            var x = m_awc.Handle; // let's make sure the webbrowser handle is created

            m_srpt.LogData("Starting DoDownloadContacts", 3, StatusRpt.MSGT.Header);

            DownloadGenericExcelReport dg =
                new DownloadGenericExcelReport(
                    "contacts",
                    WebCore._s_ContactsView,
                    WebCore._sid_Contacts_Anchor_ContactsReport,
                    WebCore._s_Contacts_Roster_Submit_Print,
                    "Roster.xls",
                    "Roster",
                    new[]
                        {
                        new DownloadGenericExcelReport.ControlSetting<string>(WebCore._s_Contacts_Roster_Select_Format,
                                                                              WebCore._sid_Contacts_Roster_Select_Format,
                                                                              "Excel Worksheet Format (.xls)")
                        },
                    new[]
                        {
                        new DownloadGenericExcelReport.ControlSetting<bool>(WebCore._s_Contacts_Roster_Check_Address, true),
                        new DownloadGenericExcelReport.ControlSetting<bool>(WebCore._s_Contacts_Roster_Check_Email, true),
                        new DownloadGenericExcelReport.ControlSetting<bool>(WebCore._s_Contacts_Roster_Check_PageHeader, false),
                        new DownloadGenericExcelReport.ControlSetting<bool>(WebCore._s_Contacts_Roster_Check_Phone, true),
                        new DownloadGenericExcelReport.ControlSetting<bool>(WebCore._s_Contacts_Roster_Check_Team, true),
                        new DownloadGenericExcelReport.ControlSetting<bool>(WebCore._s_Contacts_Roster_Check_Site, true)
                        },
                    Profile.Contacts,
                    Profile.ContactsWorking,
                    this
                );

            string sContactsNew;
            Task tskDownloadContacts = new Task(() =>
                {
                if (!m_cbSkipContactDownload.Checked)
                    {
                    dg.DownloadGeneric(out sContactsNew);
                    Profile.Contacts = sContactsNew;
                    }

                DoPendingQueueUIOp();
                });

            tskDownloadContacts.Start();
#endif
        }
    }
}
