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

            DownloadGenericExcelReport dg = new DownloadGenericExcelReport("Contacts", this);
            Task tskDownloadContacts = new Task(() =>
                {
                dg.DownloadGeneric();
                DoPendingQueueUIOp();
                });

            tskDownloadContacts.Start();
        }
    }
}
