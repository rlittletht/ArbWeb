using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using HtmlAgilityPack;
using OpenQA.Selenium;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace ArbWeb.Announcements;

// In order for an announcement to be able to be managed by ArbWeb, the entire announcement MUST be wrapped with
// <div id="ArbWebAnnounce_[SomeUniqueID]"> ... </div>
public class WebAnnouncements
{
    private List<Announcement> m_announcements;
    private IAppContext m_appContext;
    private WebPagination m_pagination;

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

    public static HtmlDocument GetHtmlDocumentForCurrentPage(IAppContext appContext)
    {
        string sHtml = appContext.WebControl.Driver.FindElement(By.XPath("//body")).GetAttribute("innerHTML");
        HtmlDocument html = new HtmlDocument();
        html.LoadHtml(sHtml);
        return html;
    }

    public static HtmlDocument GetHtmlDocumentForAnnouncementsPage(IAppContext appContext)
    {
        NavigateToAnnouncementsPage(appContext);

        return GetHtmlDocumentForCurrentPage(appContext);
    }

    public void ReadCurrentAnnouncements()
    {
        m_pagination = new(m_appContext);

        WebAnnouncements.NavigateToAnnouncementsPage(m_appContext);
        m_pagination.ResetLinksFromStartingPage();

        m_pagination.ProcessAllPagesFromStartingPage(VisitAnnouncementsPage);
    }

    public bool VisitAnnouncementsPage(int pageIndex)
    {
        // navigate to the announcements page
        HtmlDocument html = GetHtmlDocumentForCurrentPage(m_appContext);

        string xpath = $"//tr[./td/div/div[contains(@id, '{s_AnnounceDivIdPrefix}')]]";
        HtmlNodeCollection nodes = html.DocumentNode.SelectNodes(xpath);

        if (nodes == null)
            return true;

        foreach (HtmlNode node in nodes)
        {
            Announcement announce = CreateAnnouncementForMatchedAnnouncementRow(node, pageIndex);

            if (announce != null)
                m_announcements.Add(announce);
        }

        return true;
    }

    HtmlNode AnnouncementHtmlNodeFromAnnouncementRow(HtmlNode node)
    {
        return node.SelectSingleNode("./td[2]");
    }

    bool GetCheckboxValueForIdSubstringInRowNode(HtmlNode row, string idSubstring)
    {
        string xpath = $".//input[contains(@id, '{idSubstring}') and @checked='checked']";

       HtmlNodeCollection nodes = row.SelectNodes(xpath);

        return nodes != null && nodes.Count == 1;
    }

    string GetOfficialsVisibleToFromRowNode(HtmlNode row)
    {
        string xpath = "./td[4]";

        HtmlNode node = row.SelectSingleNode(xpath);

        return node.InnerText;
    }

    string CollapseWhitespace(string s)
    {
        s = s.Trim();
        s = Regex.Replace(s, @"\s+", " ");

        return s;
    }

    public Announcement CreateAnnouncementForMatchedAnnouncementRow(HtmlNode node, int pageIndex)
    {
        Announcement announce = new Announcement();

        announce.PageIndex = pageIndex;

        string xpath = "//input[contains(@id, 'btnEdit')]";
        announce.Editable = node.SelectNodes(xpath).Count != 0;

        xpath = "//input[contains(@id, 'btnDelete')]";
        announce.Deletable = node.SelectNodes(xpath).Count != 0;

        HtmlNode htmlNode = AnnouncementHtmlNodeFromAnnouncementRow(node);

        announce.AnnouncementHtml = htmlNode.InnerHtml;
        string text = htmlNode.InnerText;
        announce.Hint = CollapseWhitespace(text.Substring(0, Math.Min(80, text.Length)));

        Regex hintMatch = new Regex(@$"{s_AnnounceDivIdPrefix}[^ '""]+");

        Match match = hintMatch.Match(announce.AnnouncementHtml);

        if (match.Success == false)
            return null;

        announce.MatchString = match.Value;
        announce.ShowAssigners = GetCheckboxValueForIdSubstringInRowNode(node, "chkToAssigners");
        announce.ShowContacts = GetCheckboxValueForIdSubstringInRowNode(node, "chkToContacts");

        announce.Officials = CollapseWhitespace(GetOfficialsVisibleToFromRowNode(node));


        return announce;
    }

    private ListView m_listView;

    public void InitializeListView(ListView list)
    {
        m_listView = list;

        list.Columns.Add("ID", 75);
        list.Columns.Add("Hint", 300);
        list.Columns.Add("ShowAssigners", 25);
        list.Columns.Add("ShowContacts", 25);
        list.Columns.Add("Officials", 50);
    }

    public void SyncListViewItems()
    {
        m_listView.Items.Clear();

        foreach (Announcement announce in m_announcements)
        {
            ListViewItem item = new ListViewItem(announce.AnnouncementID());
            item.SubItems.Add(announce.Hint);
            item.SubItems.Add(announce.ShowAssigners ? "Yes" : "No");
            item.SubItems.Add(announce.ShowContacts ? "Yes" : "No");
            item.SubItems.Add(announce.Officials);
            item.Tag = announce;

            m_listView.Items.Add(item);
        }
    }

}
