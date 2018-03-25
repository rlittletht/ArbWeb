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
        void DoDownloadContacts()
        {
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
        }
    }
}
