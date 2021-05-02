using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Identity.Client;
using TCore.MsalWeb;

namespace ArbWeb
{
	class Office365
	{
		IPublicClientApplication m_app;

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

		public Office365(string sClientID)
		{
			if (m_auth == null)
				m_auth = new Auth(sClientID, new[] {"https://graph.microsoft.com/.default" /* "Files.Read.All", "Sites.Read.All"*/});
		}

		public async Task EnsureLoggedIn()
		{
			if (!m_auth.IsLoggedIn)
				await m_auth.Login();

			if (m_api == null)
				m_api = new WebApiInterop("https://graph.microsoft.com/v1.0", m_auth);
		}
		
		public async Task DownloadFile(string sSiteRoot, string sSubSite, string sFile, string sOutfile)
		{ 
		// HttpResponseMessage resp = m_api.CallService("sites/washdist9.sharepoint.com:/sites/scheduling", true);
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
