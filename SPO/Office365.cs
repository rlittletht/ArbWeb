using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Identity.Client;
using TCore.MsalWeb;

namespace ArbWeb
{
    public class Office365
    {
        private Auth m_auth = null;
        private WebApiInterop m_api = null;

        class CollectionInfo
        {
            public string hostname { get; set; }
        }

        class SharepointSiteInfo
        {
            public DateTime createdDateTime { get; set; }
            public string description { get; set; }
            public string id { get; set; }
            public DateTime lastModifiedDateTime { get; set; }
            public string name { get; set; }
            public string webUrl { get; set; }
            public string displayName { get; set; }
            public CollectionInfo siteCollection { get; set; }
        }

        class DriveInfo
        {
            public DateTime createdDateTime { get; set; }
            public string id { get; set; }
            public string name { get; set; }
            public string webUrl { get; set; }
        }

        /*----------------------------------------------------------------------------
            %%Function: Office365
            %%Qualified: ArbWeb.Office365.Office365
        ----------------------------------------------------------------------------*/
        public Office365(string sClientID)
        {
            if (m_auth == null)
                m_auth = new Auth(sClientID, new[] { "https://graph.microsoft.com/.default" /* "Files.Read.All", "Sites.Read.All"*/ });
        }

        /*----------------------------------------------------------------------------
            %%Function: EnsureLoggedIn
            %%Qualified: ArbWeb.Office365.EnsureLoggedIn
        ----------------------------------------------------------------------------*/
        public async Task EnsureLoggedIn()
        {
            if (!m_auth.IsLoggedIn)
                await m_auth.Login();

            if (m_api == null)
                m_api = new WebApiInterop("https://graph.microsoft.com/v1.0", m_auth);
        }

        public async Task GetFormInfo()
        {
            // HttpResponseMessage resp = m_api.CallServiceApiDirect("https://forms.office.com/formapi/api/forms", true);
            HttpResponseMessage resp = m_api.CallServiceApiDirect("https://graph.microsoft.com/v1.0/me/forms", true);

            if (resp.StatusCode == HttpStatusCode.Unauthorized)
            {
                throw new Exception("Service returned 'user is unauthorized'");
            }

            string json = m_api.GetContentAsString(resp);
        }

        /*----------------------------------------------------------------------------
            %%Function: DownloadFile
            %%Qualified: ArbWeb.Office365.DownloadFile
        ----------------------------------------------------------------------------*/
        public async Task DownloadFile(string sSiteRoot, string sSubSite, string sFile, string sOutfile)
        {
            SharepointSiteInfo siteInfo = m_api.CallService<SharepointSiteInfo>($"sites/{sSiteRoot}:/{sSubSite}", true);

            string sQuery = $"sites/{siteInfo.id}/drive/root:/{sFile}:/content";

            HttpResponseMessage response = m_api.CallService(sQuery, true);

            // and now we have a stream to copy to a file
            using (FileStream fs = new FileStream(sOutfile, FileMode.CreateNew))
            {
                await m_api.CopyContentToFileStream(response, fs);
            }
        }
    }
}
