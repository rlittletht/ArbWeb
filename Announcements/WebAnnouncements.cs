using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using TCore.StatusBox;
using TCore.WebControl;
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

    /*----------------------------------------------------------------------------
        %%Function: WebAnnouncements
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.WebAnnouncements
    ----------------------------------------------------------------------------*/
    public WebAnnouncements(IAppContext appContext)
    {
        m_announcements = new List<Announcement>();
        m_appContext = appContext;
    }

    /*----------------------------------------------------------------------------
        %%Function: NavigateToAnnouncementsPage
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.NavigateToAnnouncementsPage
    ----------------------------------------------------------------------------*/
    public static void NavigateToAnnouncementsPage(IAppContext appContext)
    {
        appContext.EnsureLoggedIn();
        Utils.ThrowIfNot(appContext.WebControl.FNavToPage(WebCore._s_Announcements), "Couldn't nav to announcements page!");
        appContext.WebControl.WaitForPageLoad();
    }

    public static void EnsureNavigatedToAnnouncementsPage(IAppContext appContext)
    {
        if (string.Compare(appContext.WebControl.Driver.Url, WebCore._s_Announcements, StringComparison.CurrentCultureIgnoreCase) == 0)
            return;

        appContext.EnsureLoggedIn();

        Utils.ThrowIfNot(appContext.WebControl.FNavToPage(WebCore._s_Announcements), "Couldn't nav to announcements page!");
        appContext.WebControl.WaitForPageLoad();
    }

    /*----------------------------------------------------------------------------
        %%Function: GetHtmlDocumentForCurrentPage
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.GetHtmlDocumentForCurrentPage
    ----------------------------------------------------------------------------*/
    public static HtmlDocument GetHtmlDocumentForCurrentPage(IAppContext appContext)
    {
        string sHtml = appContext.WebControl.Driver.FindElement(By.XPath("//body")).GetAttribute("innerHTML");
        HtmlDocument html = new HtmlDocument();
        html.LoadHtml(sHtml);
        return html;
    }

    /*----------------------------------------------------------------------------
        %%Function: GetHtmlDocumentForAnnouncementsPage
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.GetHtmlDocumentForAnnouncementsPage
    ----------------------------------------------------------------------------*/
    public static HtmlDocument GetHtmlDocumentForAnnouncementsPage(IAppContext appContext)
    {
        NavigateToAnnouncementsPage(appContext);

        return GetHtmlDocumentForCurrentPage(appContext);
    }

    /*----------------------------------------------------------------------------
        %%Function: ReadCurrentAnnouncements
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.ReadCurrentAnnouncements
    ----------------------------------------------------------------------------*/
    public void ReadCurrentAnnouncements()
    {
        m_pagination = new(m_appContext);
        m_announcements.Clear();

        WebAnnouncements.NavigateToAnnouncementsPage(m_appContext);
        m_pagination.ResetLinksFromStartingPage();

        m_pagination.ProcessAllPagesFromStartingPage(VisitAnnouncementsPage);
    }

    /*----------------------------------------------------------------------------
        %%Function: VisitAnnouncementsPage
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.VisitAnnouncementsPage
    ----------------------------------------------------------------------------*/
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

    /*----------------------------------------------------------------------------
        %%Function: AnnouncementHtmlNodeFromAnnouncementRow
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.AnnouncementHtmlNodeFromAnnouncementRow
    ----------------------------------------------------------------------------*/
    HtmlNode AnnouncementHtmlNodeFromAnnouncementRow(HtmlNode node)
    {
        return node.SelectSingleNode("./td[2]");
    }

    /*----------------------------------------------------------------------------
        %%Function: GetCheckboxValueForIdSubstringInRowNode
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.GetCheckboxValueForIdSubstringInRowNode
    ----------------------------------------------------------------------------*/
    bool GetCheckboxValueForIdSubstringInRowNode(HtmlNode row, string idSubstring)
    {
        string xpath = $".//input[contains(@id, '{idSubstring}') and @checked='checked']";

        HtmlNodeCollection nodes = row.SelectNodes(xpath);

        return nodes != null && nodes.Count == 1;
    }

    /*----------------------------------------------------------------------------
        %%Function: GetOfficialsVisibleToFromRowNode
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.GetOfficialsVisibleToFromRowNode
    ----------------------------------------------------------------------------*/
    string GetOfficialsVisibleToFromRowNode(HtmlNode row)
    {
        string xpath = "./td[4]";

        HtmlNode node = row.SelectSingleNode(xpath);

        return node.InnerText;
    }

    /*----------------------------------------------------------------------------
        %%Function: CollapseWhitespace
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.CollapseWhitespace
    ----------------------------------------------------------------------------*/
    string CollapseWhitespace(string s)
    {
        s = s.Trim();
        s = Regex.Replace(s, @"\s+", " ");

        return s;
    }

    /*----------------------------------------------------------------------------
        %%Function: CreateAnnouncementForMatchedAnnouncementRow
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.CreateAnnouncementForMatchedAnnouncementRow
    ----------------------------------------------------------------------------*/
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

    int GetPageForAnnouncement(Announcement announcement)
    {
        foreach (Announcement check in m_announcements)
        {
            if (check.AnnouncementID() == announcement.AnnouncementID())
                return check.PageIndex;
        }

        return -1;
    }


    /*----------------------------------------------------------------------------
        %%Function:BuildAnnouncementNameOrIdString
        %%Qualified:ArbWeb.WebAnnounce.BuildAnnouncementNameOrIdString
    ----------------------------------------------------------------------------*/
    public static string BuildAnnouncementNameOrIdString(string sPrefix, string sSuffix, string sCtl)
    {
        return $"{sPrefix}{sCtl}{sSuffix}";
    }

    static void CheckAnnounceControl(IAppContext appContext, string prefix, string suffix, string ctl, bool enabled)
    {
        string sidEnable = BuildAnnouncementNameOrIdString(prefix, suffix, ctl);

        appContext.WebControl.FSetCheckboxControlIdVal(enabled, sidEnable);
    }

    public static void UpdateAnnouncementOnCurrentPage(IAppContext appContext, Announcement announcement)
    {
        string containingDiv = announcement.MatchString;

        appContext.StatusReport.AddMessage($"Starting Announcement Set for '{containingDiv}'...");
        appContext.StatusReport.PushLevel();

        HtmlDocument html = GetHtmlDocumentForCurrentPage(appContext);

        string sXpath = $"//tr[./td/div/div[contains(@id, '{containingDiv}')]]";
        HtmlNode nodeFind = html.DocumentNode.SelectSingleNode(sXpath);

        string sCtl = null;

        appContext.StatusReport.LogData($"Found {containingDiv} DIV, looking for parent TR element", 3, MSGT.Body);

        // now find one of our controls and get its control number
        string s = nodeFind.InnerHtml;
        int ich = s.IndexOf(WebCore._s_Announcements_Button_Edit_Prefix, StringComparison.Ordinal);
        if (ich > 0)
        {
            sCtl = s.Substring(ich + WebCore._s_Announcements_Button_Edit_Prefix.Length, 5);
        }

        appContext.StatusReport.LogData($"Extracted ID for announcment to set: {sCtl}", 3, MSGT.Body);

        Utils.ThrowIfNot(sCtl != null, $"Can't find {containingDiv} edit control");

        string sidControl = BuildAnnouncementNameOrIdString(
            WebCore._sid_Announcements_Button_Edit_Prefix,
            WebCore._sid_Announcements_Button_Edit_Suffix,
            sCtl);

        Utils.ThrowIfNot(appContext.WebControl.FClickControlId(sidControl), "Couldn't find edit button");
        appContext.WebControl.WaitForPageLoad();

        // wait for CKEDITOR to load and init...wait for the control
        WebDriverWait wait = new WebDriverWait(appContext.WebControl.Driver, TimeSpan.FromSeconds(5));
        wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("cke_12")));

        Utils.ThrowIfNot(WebControl.WaitForControl(appContext.WebControl.Driver, appContext.StatusReport, "cke_12"), "CKEDITOR never loaded");

        // we should already be in source mode -- when we initialize we inject source into the page to make sure we always
        // startup in source mode

        string buttonClass = appContext.WebControl.GetAttributeValueFromId("cke_12", "class");
        Utils.ThrowIfNot(buttonClass != null, "can't find source button");
        bool sourceOn = true;

        if (buttonClass.Contains("button_off"))
        {
            if (MessageBox.Show("CKE Editor didn't start in source mode. This edit may be lossy. Continue", "ArbWeb", MessageBoxButtons.YesNo)
                != DialogResult.Yes)
            {
                string cancel = BuildAnnouncementNameOrIdString(
                    WebCore._sid_Announcements_Button_Cancel_Prefix,
                    WebCore._sid_Announcements_Button_Cancel_Suffix,
                    sCtl);

                Utils.ThrowIfNot(appContext.WebControl.FClickControlId(cancel), "Couldn't find cancel button");
                appContext.WebControl.WaitForPageLoad();

                appContext.StatusReport.PopLevel();
                appContext.StatusReport.AddMessage("Aborted Announcement Set.");
                return;
            }

            sourceOn = false;
        }
        else
        {
            Utils.ThrowIfNot(buttonClass.Contains("button_on"), $"could not determine CKE button state: {buttonClass}");
        }

        if (!string.IsNullOrEmpty(announcement.AnnouncementHtml))
        {
            if (!sourceOn)
            {
                // select source mode
                Utils.ThrowIfNot(appContext.WebControl.FClickControlId("cke_12"), "Couldn't find <SOURCE> button");
                appContext.WebControl.WaitForPageLoad();
            }

            appContext.WebControl.FSetTextAreaTextForControlAsChildOfDivId("cke_1_contents", announcement.AnnouncementHtml, true);
            appContext.WebControl.WaitForPageLoad();
        }
        // and now save it.

        // now set the visibility as requested
        CheckAnnounceControl(
            appContext,
            WebCore._sid_Announcements_Button_ToAssigners_Prefix,
            WebCore._sid_Announcements_Button_ToAssigners_Suffix,
            sCtl,
            announcement.ShowAssigners);

        CheckAnnounceControl(
            appContext,
            WebCore._sid_Announcements_Button_ToContacts_Prefix,
            WebCore._sid_Announcements_Button_ToContacts_Suffix,
            sCtl,
            announcement.ShowContacts);

        // and lastly, choose the officials this is visible to
        string sidEnable = BuildAnnouncementNameOrIdString(
            WebCore._s_Announcements_Button_ToOfficials_Prefix,
            WebCore._s_Announcements_Button_ToOfficials_Suffix,
            sCtl);

        string value =
            appContext.WebControl.GetOptionValueFromFilterOptionTextForControlName(
                sidEnable,
                announcement.Officials);

        Utils.ThrowIfNot(value != null, $"Can't set announcement visible to {announcement.Officials}");

        string current = appContext.WebControl.GetSelectedOptionValueFromSelectControlName(sidEnable);
        if (current != value)
        {
            Utils.ThrowIfNot(
                appContext.WebControl.FSetSelectedOptionValueForControlName(sidEnable, value),
                $"can't set officials visible to to {value}");
        }

        sidControl = BuildAnnouncementNameOrIdString(
            WebCore._sid_Announcements_Button_Save_Prefix,
            WebCore._sid_Announcements_Button_Save_Suffix,
            sCtl);

        Utils.ThrowIfNot(appContext.WebControl.FClickControlId(sidControl), "Couldn't find save button");
        appContext.WebControl.WaitForPageLoad();

        appContext.StatusReport.PopLevel();
        appContext.StatusReport.AddMessage("Completed Announcement Set.");
    }

    public void UpdateArbiterAnnouncement(Announcement baseAnnouncement, Announcement updatedAnnouncement)
    {
        int page;

        if (baseAnnouncement.AnnouncementID() != updatedAnnouncement.AnnouncementID())
            throw new Exception("Announcement IDs don't match!");

        EnsureNavigatedToAnnouncementsPage(m_appContext);

        page = baseAnnouncement.PageIndex;
        if (page == -1)
        {
            // need to find what page its on
            page = GetPageForAnnouncement(baseAnnouncement);

            if (page == -1)
            {
                // reload the announcements and try again
                ReadCurrentAnnouncements();

                page = GetPageForAnnouncement(baseAnnouncement);

                if (page == -1)
                    throw new Exception("Couldn't find announcement on any page");
            }
        }

        // now navigate to the page
        m_pagination.NavigateToPageByIndex(page);

        UpdateAnnouncementOnCurrentPage(m_appContext, updatedAnnouncement);

        for (int i = 0; i < m_announcements.Count; i++)
        {
            if (m_announcements[i].AnnouncementID() == baseAnnouncement.AnnouncementID())
            {
                m_announcements[i] = updatedAnnouncement;
                break;
            }
        }
    }

    private ListView m_listView;

/*----------------------------------------------------------------------------
    %%Function: InitializeListView
    %%Qualified: ArbWeb.Announcements.WebAnnouncements.InitializeListView
----------------------------------------------------------------------------*/
    public void InitializeListView(ListView list)
    {
        m_listView = list;

        list.Columns.Add("ID", 75);
        list.Columns.Add("Hint", 500);
        list.Columns.Add("ShowAssigners", 50);
        list.Columns.Add("ShowContacts", 50);
        list.Columns.Add("Officials", 50);
    }

/*----------------------------------------------------------------------------
    %%Function: SyncListViewItems
    %%Qualified: ArbWeb.Announcements.WebAnnouncements.SyncListViewItems
----------------------------------------------------------------------------*/
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
