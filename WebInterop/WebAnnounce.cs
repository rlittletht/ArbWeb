using System;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using TCore.StatusBox;
using TCore.WebControl;

namespace ArbWeb
{
	public class WebAnnounce
	{
		private IAppContext m_appContext;
		
		/*----------------------------------------------------------------------------
			%%Function:WebAnnounce
			%%Qualified:ArbWeb.WebAnnounce.WebAnnounce
		----------------------------------------------------------------------------*/
		public WebAnnounce(IAppContext appContext)
		{
			m_appContext = appContext;
		}
		
		/*----------------------------------------------------------------------------
			%%Function:SetArbiterAnnounce
			%%Qualified:ArbWeb.WebAnnounce.SetArbiterAnnounce
		----------------------------------------------------------------------------*/
		public void SetArbiterAnnounce(string sArbiterHelpNeeded)
		{
			m_appContext.StatusReport.AddMessage("Starting Announcement Set...");
			m_appContext.StatusReport.PushLevel();

			m_appContext.EnsureLoggedIn();
			Utils.ThrowIfNot(m_appContext.WebControl.FNavToPage(WebCore._s_Announcements), "Couldn't nav to announcements page!");
			m_appContext.WebControl.WaitForPageLoad();

			// now we need to find the URGENT HELP NEEDED row
			string sHtml = m_appContext.WebControl.Driver.FindElement(By.XPath("//body")).GetAttribute("innerHTML");
			HtmlDocument html = new HtmlDocument();
			html.LoadHtml(sHtml);

			string sXpath = "//div[@id='D9UrgentHelpNeeded']";

			HtmlNode node = html.DocumentNode.SelectSingleNode(sXpath);

			string sCtl = null;

			m_appContext.StatusReport.LogData("Found D9UrgentHelpNeeded DIV, looking for parent TR element", 3, MSGT.Body);


			// ok, go up to the parent TR.

			HtmlNode nodeFind = node;

			while (nodeFind.Name.ToLower() != "tr")
			{
				nodeFind = nodeFind.ParentNode;
				Utils.ThrowIfNot(nodeFind != null, "Can't find HELP announcement");
			}

			m_appContext.StatusReport.LogData("Found D9UrgentHelpNeeded parent TR", 3, MSGT.Body);

			// now find one of our controls and get its control number
			string s = nodeFind.InnerHtml;
			int ich = s.IndexOf(WebCore._s_Announcements_Button_Edit_Prefix);
			if (ich > 0)
			{
				sCtl = s.Substring(ich + WebCore._s_Announcements_Button_Edit_Prefix.Length, 5);
			}

			m_appContext.StatusReport.LogData($"Extracted ID for announcment to set: {sCtl}", 3, MSGT.Body);

			Utils.ThrowIfNot(sCtl != null, "Can't find HELP announcement");

			string sidControl = BuildAnnouncementNameOrIdString(WebCore._sid_Announcements_Button_Edit_Prefix, WebCore._sid_Announcements_Button_Edit_Suffix, sCtl);

			Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(sidControl), "Couldn't find edit button");
			m_appContext.WebControl.WaitForPageLoad();

			// find the CKEDITOR control, it will be a peer of the textarea control that it is editing
            {
    			// string sidCkeDiv = WebCore._sid_cke_Prefix + BuildAnnouncementNameOrIdString(WebCore._sid_Announcements_Textarea_Text_Prefix, WebCore._sid_Announcements_Textarea_Text_Suffix, sCtl);

				// we need to find the id to use to click to get the source control

			}

			// wait for CKEDITOR to load and init...wait for the control
            WebDriverWait wait = new WebDriverWait(m_appContext.WebControl.Driver, TimeSpan.FromSeconds(5));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("cke_12")));

            Utils.ThrowIfNot(WebControl.WaitForControl(m_appContext.WebControl.Driver, m_appContext.StatusReport, "cke_12"), "CKEDITOR never loaded");

			// select source mode
            Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId("cke_12"), "Couldn't find <SOURCE> button");

			m_appContext.WebControl.FSetTextAreaTextForControlAsChildOfDivId("cke_1_contents", sArbiterHelpNeeded, true);
			m_appContext.WebControl.WaitForPageLoad();

			sidControl = BuildAnnouncementNameOrIdString(WebCore._sid_Announcements_Button_Save_Prefix, WebCore._sid_Announcements_Button_Save_Suffix, sCtl);

			Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(sidControl), "Couldn't find save button");
			m_appContext.WebControl.WaitForPageLoad();

			// and now save it.

			m_appContext.StatusReport.PopLevel();
			m_appContext.StatusReport.AddMessage("Completed Announcement Set.");
		}

		/*----------------------------------------------------------------------------
			%%Function:BuildAnnouncementNameOrIdString
			%%Qualified:ArbWeb.WebAnnounce.BuildAnnouncementNameOrIdString
		----------------------------------------------------------------------------*/
		private static string BuildAnnouncementNameOrIdString(string sPrefix, string sSuffix, string sCtl)
		{
			return $"{sPrefix}{sCtl}{sSuffix}";
		}
	}
}