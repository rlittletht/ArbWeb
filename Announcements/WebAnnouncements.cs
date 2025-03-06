using System.Collections.Generic;
using System.Web.UI.WebControls;
using HtmlAgilityPack;
using OpenQA.Selenium;

namespace ArbWeb.Announcements;

// In order for an announcement to be able to be managed by ArbWeb, the entire announcement MUST be wrapped with
// <div id="ArbWebAnnounce_[SomeUniqueID]"> ... </div>
public class WebAnnouncements
{
    private List<Announcement> m_announcements;
    private IAppContext m_appContext;

    public static string s_AnnounceDivIdPrefix = "ArbWebAnnounce_";

    public WebAnnouncements(IAppContext appContext)
    {
        m_announcements = new List<Announcement>();
        m_appContext = appContext;
    }

    public static void NavigateToAnnouncementsPage(IAppContext appContext)
    {
        appContext.EnsureLoggedIn();
        Utils.ThrowIfNot(appContext.WebControl.FNavToPage(WebCore._s_Announcements), "Couldn't nav to announcements page!");
        appContext.WebControl.WaitForPageLoad();
    }

    public static HtmlDocument GetHtmlDocumentForAnnouncementsPage(IAppContext appContext)
    {
        NavigateToAnnouncementsPage(appContext);

        // now we need to find the URGENT HELP NEEDED row
        string sHtml = appContext.WebControl.Driver.FindElement(By.XPath("//body")).GetAttribute("innerHTML");
        HtmlDocument html = new HtmlDocument();
        html.LoadHtml(sHtml);
        return html;
    }

    public void ReadCurrentAnnouncements()
    {
        // navigate to the announcements page
        HtmlDocument html = GetHtmlDocumentForAnnouncementsPage(m_appContext);

        string xpath = $"//div[contains(@id, '{s_AnnounceDivIdPrefix}')]";
        html.DocumentNode.SelectNodes(xpath);


    }
}
