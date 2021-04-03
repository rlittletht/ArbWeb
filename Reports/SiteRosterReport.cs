using System;
using System.Diagnostics;
using TCore.StatusBox;
using TCore.WebControl;

namespace ArbWeb.Reports
{
	public class SiteRosterReport
	{
		private IAppContext m_appContext;

		private WebControl m_webControl => m_appContext.WebControl;
		private IStatusReporter m_srpt => m_appContext.StatusReport;
		void ThrowIfNot(bool f, string s) => m_appContext.ThrowIfNot(f, s);
		void EnsureLoggedIn() => m_appContext.EnsureLoggedIn();

		public SiteRosterReport(IAppContext appContext)
		{
			m_appContext = appContext;
		}

		public  void DoGenSiteRosterReport(CountsData gc, Roster rst, string []rgsRoster, DateTime dttmStart, DateTime dttmEnd)
		{
			string sTempFile = $"{Environment.GetEnvironmentVariable("Temp")}\\temp{System.Guid.NewGuid().ToString()}.doc";

			gc.GenSiteRosterReport(sTempFile, rst, rgsRoster, dttmStart, dttmEnd);
			// launch word with the file
			Process.Start(sTempFile);
			// System.IO.File.Delete(sTempFile);
		}
    }
}