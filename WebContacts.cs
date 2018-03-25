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
            var x = m_awc.Handle;
            string sFilterReq = (string)m_cbxGameFilter.SelectedItem;
            if (sFilterReq == null)
                sFilterReq = "All Games";

            // let's make sure the webbrowser handle is created

            m_srpt.LogData("Starting DoDownloadGames", 3, StatusRpt.MSGT.Header);

            Task tskDownloadGames = new Task(() => DownloadGames(sFilterReq));

            tskDownloadGames.Start();
        }

    }
}
