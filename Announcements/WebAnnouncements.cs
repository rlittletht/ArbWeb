using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using TCore.StatusBox;
using TCore.WebControl;
using static System.Net.Mime.MediaTypeNames;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace ArbWeb.Announcements;

// In order for an announcement to be able to be managed by ArbWeb, the entire announcement MUST be wrapped with
// <div id="ArbWebAnnounce_[SomeUniqueID]"> ... </div>
public class WebAnnouncements
{
    private List<Announcement> m_announcements;
    private IAppContext m_appContext;
    private WebPagination m_pagination;
    public string CommonStylesheet { get; set; } = string.Empty;
    public int CommonStylesheetVectorClock { get; set; } = 0;

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

    void EnsureSupportFileExists(string workingTargetDir, string supportName)
    {
        // make sure the supporting CSS exists
        if (!File.Exists(Path.Combine(workingTargetDir, supportName)))
            File.Copy(Path.Combine(AppContext.BaseDirectory, supportName), Path.Combine(workingTargetDir, supportName));
    }

    private string htmlPrologue1 =
        @"<html>
            <link id=""target-preview"" href=""./announcementsOfficials.css"" rel=""stylesheet"">
            <link href=""./announcements.css"" rel=""stylesheet"">
            <head>
                <script type=""text/javascript"" src=""./announcements.js""></script>";

    private string htmlPrologue2 =
        @"</head>
            <body>
                <label for=""previewSelect"">Preview for:</label>
                <select id=""previewSelect"">
                    <option value=""announcementsAssigners.css"">Assigners</option>
                    <option value=""announcementsOfficials.css"">Officials</option>
                </select>
                <div class=""outerWrapper"">
                    <div class=""announceHeader"">Announcements</div>";

    private string htmlEpilogue =
        @"</div></body></html>";

    void WriteAnnouncementToHtml(StreamWriter sw, Announcement announce)
    {
        sw.WriteLine($"<div class='wrapper'>");
        sw.WriteLine($"<div class='postHeader'>");
        sw.WriteLine($"<div class='postBox'>");
        sw.WriteLine($"<div class='postItem'>Posted by Some Assigner</div>");
        sw.WriteLine($"<div class='postItem'>Assigners: {(announce.ShowAssigners ? "Yes" : "No")}</div>");
        sw.WriteLine($"<div class='postItem'>Contacts: {(announce.ShowContacts ? "Yes" : "No")}</div>");
        sw.WriteLine($"<div class='postItem'>Officials: {announce.Officials}</div>");
        sw.WriteLine("</div>");
        sw.WriteLine("</div>");
        sw.WriteLine("<!-- ======================================================================================== -->");
        sw.WriteLine($"<!-- {announce.AnnouncementID().ToUpper()} -->");
        sw.WriteLine("<!-- ======================================================================================== -->");
        sw.WriteLine(announce.AnnouncementHtml);
        sw.WriteLine("</div>");
    }

    /*----------------------------------------------------------------------------
        %%Function: SaveAnnouncementsToHtmlFile
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.SaveAnnouncementsToHtmlFile
    ----------------------------------------------------------------------------*/
    public void SaveAnnouncementsToHtmlFile(string target, string working)
    {
        string workingTargetDir = Path.GetDirectoryName(working);

        // make sure directories exist
        Directory.CreateDirectory(Path.GetDirectoryName(target));
        Directory.CreateDirectory(workingTargetDir);

        EnsureSupportFileExists(workingTargetDir, "announcements.css");
        EnsureSupportFileExists(workingTargetDir, "announcements.js");
        EnsureSupportFileExists(workingTargetDir, "announcementsAssigners.css");
        EnsureSupportFileExists(workingTargetDir, "announcementsOfficials.css");

        using (StreamWriter sw = new StreamWriter(target))
        {
            sw.WriteLine(htmlPrologue1);
            sw.WriteLine(CommonStylesheet);
            sw.WriteLine(htmlPrologue2);

            foreach (Announcement announce in m_announcements)
            {
                if (announce.NonArbweb)
                    continue;
                WriteAnnouncementToHtml(sw, announce);
            }

            sw.WriteLine(htmlEpilogue);
        }

        System.IO.File.Delete(working);
        File.Copy(target, working);
    }

    bool GetTrElementTypeAndClock(HtmlNode styleNode, out string elementType, out int clock)
    {
        elementType = null;
        clock = 0;

        if (styleNode == null)
            return false;

        string styleValue = styleNode.GetAttributeValue("style", "");

        if (string.IsNullOrEmpty(styleValue))
            return false;

        Regex regexType = new Regex(@"tr-element:\s*['""]*([\w-_]+)['""]*");

        Match match = regexType.Match(styleValue);
        if (match.Success)
        {
            elementType = match.Groups[1].Value;
        }
        else
        {
            return false;
        }

        Regex regexClock = new Regex(@"tr-element-clock:\s*['""]*(\d+)['""]*");

        match = regexClock.Match(styleValue);
        if (match.Success)
        {
            clock = Int32.Parse(match.Groups[1].Value);
        }
        else
        {
            return false;
        }

        return true;
    }

    /*----------------------------------------------------------------------------
        %%Function: GetInlineStylesheetAndClock
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.GetInlineStylesheetAndClock

        Find the inline stylesheet under the given node and return the entire
        stylesheet as a string.  Also return the vector clock for the stylesheet.

        this can be used for each announcement (when reading from arbiter) OR
        for finding the common stylesheet in a saved announcements file
    ----------------------------------------------------------------------------*/
    HtmlNode GetInlineStylesheetAndClock(HtmlNode node, out string stylesheet, out int vectorClock)
    {
        stylesheet = null;
        string xpath = ".//style[contains(@style, 'tr-element')]";
        HtmlNode styleNode = node.SelectSingleNode(xpath);

        if (!GetTrElementTypeAndClock(styleNode, out string elementType, out vectorClock))
            return null;

        if (elementType != "inline-stylesheet")
            return null;

        stylesheet = styleNode.OuterHtml;

        return styleNode;
    }

    public bool UpdatePaginationBasedOnPage(int pageIndex)
    {
        // navigate to the announcements page
        HtmlDocument html = GetHtmlDocumentForCurrentPage(m_appContext);

        // we want all of the announcements on the page, even those that we can't edit
        // (we need placeholders for non-arbweb friendly announcements so our page counts
        // get updated appropriately)

        string xpathToTableRows = $"//table[@id='{WebCore._sid_AnnouncementsTable}']/tbody/tr";
        HtmlNodeCollection announcementRows = html.DocumentNode.SelectNodes(xpathToTableRows);

        if (announcementRows == null)
            return true;

        bool fSkipHeader = true;

        string lastMatchedOnPage = null;
        int itemsPulledBackwards = 0;

        foreach (HtmlNode row in announcementRows)
        {
            if (fSkipHeader)
            {
                fSkipHeader = false;
                continue;
            }

            string xpath = $"./td[div/div[contains(@id, '{s_AnnounceDivIdPrefix}')]]";
            HtmlNode announcementRow = row.SelectSingleNode(xpath);
            HtmlNode htmlNode = AnnouncementHtmlNodeFromAnnouncementRow(row);

            string announcementId = Announcement.AnnouncementIdFromMatchString(GetAnnouncementMatchStringFromNode(htmlNode));

            if (announcementId == null)
                // we can't synchronize if its not one of our announcements
                continue;

            // match this announcement in our collection and see if the pagination has changed
            Announcement match = LookupAnnouncementById(announcementId);
            lastMatchedOnPage = announcementId;

            Utils.ThrowIfNot(match != null, $"Couldn't find announcement {announcementId} in collection");
            Utils.ThrowIfNot(match.PageIndex <= pageIndex, $"announcement moved backwards?! {match.PageIndex} < {pageIndex}");

            if (match.PageIndex > pageIndex)
            {
                // we pulled a later item onto this page
                match.PageIndex = pageIndex;
                itemsPulledBackwards++;
            }

            if (match.PageIndex < pageIndex)
                // this moved to the next page. update it
                match.PageIndex = pageIndex;
        }

        // now, all items AFTER the last item we matched should move to the next page (and try to count how many that is
        // so we can update subsequent pages as well)
        int announcementsMoved = 0;
        bool movingToNewPage = false;

        for (int i = 0; i < m_announcements.Count; i++)
        {
            Announcement check = m_announcements[i];

            if (check.AnnouncementID() == lastMatchedOnPage)
            {
                movingToNewPage = true;
                // subsequent items that think they are still on this page should get bumped
                continue;
            }

            if (movingToNewPage)
            {
                if (check.PageIndex == pageIndex)
                {
                    // this can't be on this page because we didn't find it on the current page
                    announcementsMoved++;
                    check.PageIndex++;
                    continue;
                }

                if (announcementsMoved > 0)
                {
                    if (i >= announcementsMoved)
                    {
                        // check to see if we are one of the last XX items on the page that have to get moved
                        if (m_announcements[i - announcementsMoved].PageIndex < check.PageIndex)
                            m_announcements[i - announcementsMoved].PageIndex++;
                    }
                }

                if (itemsPulledBackwards > 0)
                {
                    if (i >= itemsPulledBackwards)
                    {
                        // check to see if we are one of the last XX items on the page that have to get moved
                        if (m_announcements[i - itemsPulledBackwards].PageIndex > check.PageIndex)
                            m_announcements[i - itemsPulledBackwards].PageIndex--;
                    }
                }
            }
        }

        return true;
    }

    /*----------------------------------------------------------------------------
        %%Function: VisitAnnouncementsPage
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.VisitAnnouncementsPage
    ----------------------------------------------------------------------------*/
    public bool VisitAnnouncementsPage(int pageIndex)
    {
        // navigate to the announcements page
        HtmlDocument html = GetHtmlDocumentForCurrentPage(m_appContext);

        // we want all of the announcements on the page, even those that we can't edit
        // (we need placeholders for non-arbweb friendly announcements so our page counts
        // get updated appropriately)

        string xpathToTableRows = $"//table[@id='{WebCore._sid_AnnouncementsTable}']/tbody/tr";
        HtmlNodeCollection announcementRows = html.DocumentNode.SelectNodes(xpathToTableRows);

        if (announcementRows == null)
            return true;

        foreach (HtmlNode row in announcementRows)
        {
            string className = row.GetAttributeValue("class", "");

            if (className == "headers" || className == "numericPaging")
                continue;

            string xpath = $"./td[div/div[contains(@id, '{s_AnnounceDivIdPrefix}')]]";
            HtmlNode announcementRow = row.SelectSingleNode(xpath);
            Announcement announce;

            if (announcementRow == null)
            {

                announce = Announcement.CreateNonArbweb();
                announce.PageIndex = pageIndex;
            }
            else
            {
                announce = CreateAnnouncementForMatchedAnnouncementRow(announcementRow.ParentNode, pageIndex);
            }

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
        string xpath = $"./td[2]//div[contains(@id, '{s_AnnounceDivIdPrefix}')]";

        HtmlNode announcement = node.SelectSingleNode(xpath);

        return announcement;
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
        %%Function: GetAnnouncementMatchStringFromNode
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.GetAnnouncementMatchStringFromNode
    ----------------------------------------------------------------------------*/
    public string GetAnnouncementMatchStringFromNode(HtmlNode htmlNode)
    {
        if (htmlNode == null)
            return null;

        Regex hintMatch = new Regex(@$"{s_AnnounceDivIdPrefix}[^ '""]+");

        Match match = hintMatch.Match(htmlNode.OuterHtml);

        if (match.Success == false)
            return null;

        return match.Value;
    }

    /*----------------------------------------------------------------------------
        %%Function: CreateAnnouncementForMatchedAnnouncementRow
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.CreateAnnouncementForMatchedAnnouncementRow
    ----------------------------------------------------------------------------*/
    public Announcement CreateAnnouncementForMatchedAnnouncementRow(HtmlNode node, int pageIndex)
    {
        Announcement announce = new Announcement();
        HtmlNode htmlNode = AnnouncementHtmlNodeFromAnnouncementRow(node);
        HtmlNode styleNode = GetInlineStylesheetAndClock(htmlNode, out string stylesheet, out int clock);

        if (styleNode != null)
        {
            announce.InlineStylesheet = stylesheet;
            announce.StyleClock = clock;

            if (clock > CommonStylesheetVectorClock)
            {
                CommonStylesheet = stylesheet;
                CommonStylesheetVectorClock = clock;
            }

            HtmlNode commonStylesheetComment = node.OwnerDocument.CreateElement("CommonStylesheet");
            styleNode.ParentNode.ReplaceChild(commonStylesheetComment, styleNode);
        }

        announce.PageIndex = pageIndex;

        string xpath = "//input[contains(@id, 'btnEdit')]";
        announce.Editable = node.SelectNodes(xpath).Count != 0;

        xpath = "//input[contains(@id, 'btnDelete')]";
        announce.Deletable = node.SelectNodes(xpath).Count != 0;

        announce.AnnouncementHtml = htmlNode.OuterHtml;
        string text = htmlNode.InnerText;
        announce.Hint = CollapseWhitespace(text.Substring(0, Math.Min(80, text.Length)));

        announce.MatchString = GetAnnouncementMatchStringFromNode(htmlNode);
        if (announce.MatchString == null)
            return null;

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

    /*----------------------------------------------------------------------------
        %%Function: CheckAnnounceControl
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.CheckAnnounceControl
    ----------------------------------------------------------------------------*/
    static void CheckAnnounceControl(IAppContext appContext, string prefix, string suffix, string ctl, bool enabled)
    {
        string sidEnable = BuildAnnouncementNameOrIdString(prefix, suffix, ctl);

        appContext.WebControl.FSetCheckboxControlIdVal(enabled, sidEnable);
    }

    /*----------------------------------------------------------------------------
        %%Function: ReplaceClockInCommonStylesheet
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.ReplaceClockInCommonStylesheet
    ----------------------------------------------------------------------------*/
    string ReplaceClockInCommonStylesheet()
    {
        Regex regex = new Regex(@"tr-element-clock:\s*\d+\s*");

        // first try without quotes
        if (regex.IsMatch(CommonStylesheet))
            return regex.Replace(CommonStylesheet, $"tr-element-clock: {CommonStylesheetVectorClock}");

        // then try with quotes.
        regex = new Regex(@"tr-element-clock:\s*['""]\s*\d+\s*['""]");
        return regex.Replace(CommonStylesheet, $"tr-element-clock: {CommonStylesheetVectorClock}");
    }

    /*----------------------------------------------------------------------------
        %%Function: BuildHtmlForUpdate
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.BuildHtmlForUpdate
    ----------------------------------------------------------------------------*/
    public string BuildHtmlForUpdate(Announcement announcement)
    {
        string inlineStylesheet = ReplaceClockInCommonStylesheet();
        string updateHtml = announcement.AnnouncementHtml.Replace("<commonstylesheet></commonstylesheet>", inlineStylesheet);

        return updateHtml;
    }

    /*----------------------------------------------------------------------------
        %%Function: UpdateAnnouncementOnCurrentPage
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.UpdateAnnouncementOnCurrentPage
    ----------------------------------------------------------------------------*/
    public void UpdateAnnouncementOnCurrentPage(Announcement announcement)
    {
        string containingDiv = announcement.MatchString;

        m_appContext.StatusReport.AddMessage($"Starting Announcement Set for '{containingDiv}'...");
        m_appContext.StatusReport.PushLevel();

        HtmlDocument html = GetHtmlDocumentForCurrentPage(m_appContext);

        string sXpath = $"//tr[./td/div/div[contains(@id, '{containingDiv}')]]";
        HtmlNode nodeFind = html.DocumentNode.SelectSingleNode(sXpath);

        string sCtl = null;

        m_appContext.StatusReport.LogData($"Found {containingDiv} DIV, looking for parent TR element", 3, MSGT.Body);

        // now find one of our controls and get its control number
        string s = nodeFind.InnerHtml;
        int ich = s.IndexOf(WebCore._s_Announcements_Button_Edit_Prefix, StringComparison.Ordinal);
        if (ich > 0)
        {
            sCtl = s.Substring(ich + WebCore._s_Announcements_Button_Edit_Prefix.Length, 5);
        }

        m_appContext.StatusReport.LogData($"Extracted ID for announcment to set: {sCtl}", 3, MSGT.Body);

        Utils.ThrowIfNot(sCtl != null, $"Can't find {containingDiv} edit control");

        string sidControl = BuildAnnouncementNameOrIdString(
            WebCore._sid_Announcements_Button_Edit_Prefix,
            WebCore._sid_Announcements_Button_Edit_Suffix,
            sCtl);

        Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(sidControl), "Couldn't find edit button");
        m_appContext.WebControl.WaitForPageLoad();

        // wait for CKEDITOR to load and init...wait for the control
        WebDriverWait wait = new WebDriverWait(m_appContext.WebControl.Driver, TimeSpan.FromSeconds(5));
        wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("cke_12")));

        Utils.ThrowIfNot(WebControl.WaitForControl(m_appContext.WebControl.Driver, m_appContext.StatusReport, "cke_12"), "CKEDITOR never loaded");

        // we should already be in source mode -- when we initialize we inject source into the page to make sure we always
        // startup in source mode

        string buttonClass = m_appContext.WebControl.GetAttributeValueFromId("cke_12", "class");
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

                Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(cancel), "Couldn't find cancel button");
                m_appContext.WebControl.WaitForPageLoad();

                m_appContext.StatusReport.PopLevel();
                m_appContext.StatusReport.AddMessage("Aborted Announcement Set.");
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
                Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId("cke_12"), "Couldn't find <SOURCE> button");
                m_appContext.WebControl.WaitForPageLoad();
            }

            string htmlForUpdate = BuildHtmlForUpdate(announcement);
            m_appContext.WebControl.FSetTextAreaTextForControlAsChildOfDivId("cke_1_contents", htmlForUpdate, true);
            m_appContext.WebControl.WaitForPageLoad();
        }
        // and now save it.

        // now set the visibility as requested
        CheckAnnounceControl(
            m_appContext,
            WebCore._sid_Announcements_Button_ToAssigners_Prefix,
            WebCore._sid_Announcements_Button_ToAssigners_Suffix,
            sCtl,
            announcement.ShowAssigners);

        CheckAnnounceControl(
            m_appContext,
            WebCore._sid_Announcements_Button_ToContacts_Prefix,
            WebCore._sid_Announcements_Button_ToContacts_Suffix,
            sCtl,
            announcement.ShowContacts);

        // and lastly, choose the officials this is visible to
        string sidEnable = BuildAnnouncementNameOrIdString(
            WebCore._s_Announcements_Button_ToOfficials_Prefix,
            WebCore._s_Announcements_Button_ToOfficials_Suffix,
            sCtl);

        string officials = announcement.Officials;

        if (string.IsNullOrEmpty(officials))
            officials = "None";

        string value =
            m_appContext.WebControl.GetOptionValueFromFilterOptionTextForControlName(
                sidEnable,
                officials);

        Utils.ThrowIfNot(value != null, $"Can't set announcement visible to {officials}");

        string current = m_appContext.WebControl.GetSelectedOptionValueFromSelectControlName(sidEnable);
        if (current != value)
        {
            Utils.ThrowIfNot(
                m_appContext.WebControl.FSetSelectedOptionValueForControlName(sidEnable, value),
                $"can't set officials visible to to {value}");
        }

        sidControl = BuildAnnouncementNameOrIdString(
            WebCore._sid_Announcements_Button_Save_Prefix,
            WebCore._sid_Announcements_Button_Save_Suffix,
            sCtl);

        Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(sidControl), "Couldn't find save button");
        m_appContext.WebControl.WaitForPageLoad();

        m_appContext.StatusReport.PopLevel();
        m_appContext.StatusReport.AddMessage("Completed Announcement Set.");
    }

    /*----------------------------------------------------------------------------
        %%Function: UpdateArbiterAnnouncement
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.UpdateArbiterAnnouncement
    ----------------------------------------------------------------------------*/
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

        UpdateAnnouncementOnCurrentPage(updatedAnnouncement);

        ChangeAndTouchAnnouncement(updatedAnnouncement);
    }

    Announcement LookupAnnouncementById(string id)
    {
        foreach (Announcement announce in m_announcements)
        {
            if (announce.AnnouncementID() == id)
                return announce;
        }

        return null;
    }

    /*----------------------------------------------------------------------------
        %%Function: ChangeAndTouchAnnouncement
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.ChangeAndTouchAnnouncement

        find the given announcement in our collection

        remove it from the collection at its position and insert the new announcement
        in its position collection at the beginning.

        BUT FIRST, this is going to go to page 0 -- see if it will change the
        pagination of the other items and if so, update them.
    ----------------------------------------------------------------------------*/
    void ChangeAndTouchAnnouncement(Announcement announcement)
    {
        int currentIndex = 0;

        while (currentIndex < m_announcements.Count)
        {
            if (m_announcements[currentIndex].AnnouncementID() == announcement.AnnouncementID())
                break;

            currentIndex++;
        }

        // ok, this is going to go to page 0 -- see if it will change the pagination of the other items and if so, update them
        Utils.ThrowIfNot(currentIndex < m_announcements.Count, "could not find announcement to update");

        int oldPageIndex = m_announcements[currentIndex].PageIndex;

        m_announcements.RemoveAt(currentIndex);

        announcement.PageIndex = 0;
        m_announcements.Insert(0, announcement);

        UpdatePaginationBasedOnPage(oldPageIndex);
    }

    /*----------------------------------------------------------------------------
        %%Function: UpdateAllAnnouncements
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.UpdateAllAnnouncements

        update all the announcements we have. The rank, if set, determines the
        order they announcements will show in.  we rank if the item will show to
        anyone
    ----------------------------------------------------------------------------*/
    public void UpdateAllAnnouncements()
    {
        // first, get the current announcements
        WebAnnouncements currentAnnouncements = new WebAnnouncements(m_appContext);
        currentAnnouncements.ReadCurrentAnnouncements();

        currentAnnouncements.CommonStylesheet = CommonStylesheet;
        currentAnnouncements.CommonStylesheetVectorClock = CommonStylesheetVectorClock;

        // we want to update in reverse order so that we get the ordering we want
        for (int i = m_announcements.Count - 1; i >= 0; i--)
        {
            Announcement announce = m_announcements[i];
            if (announce.NonArbweb)
                continue;

            Announcement current = currentAnnouncements.LookupAnnouncementById(announce.AnnouncementID());
            Utils.ThrowIfNot(current != null, "did not find current announcement to update");

            if (announce.Equals(current))
                continue;

            // this is an existing announcement -- update it
            currentAnnouncements.UpdateArbiterAnnouncement(current, announce);
        }
    }

    public void SortAnnouncementsByRank()
    {
        m_announcements.Sort((a, b) => a.Rank.CompareTo(b.Rank));
    }

    string GetArgValueFromKvp(HtmlNode node, string name)
    {
        string xpath = $".//div[child::text()[contains(., '{name}')]]";

        HtmlNode kvpNode = node.SelectSingleNode(xpath);
        if (kvpNode != null)
        {
            string kvp = kvpNode.InnerText;
            string[] parts = kvp.Split(':');

            if (parts.Length == 2)
                return parts[1].Trim();
        }

        return null;
    }

    bool GetArgBoolValueFromKvp(HtmlNode node, string name)
    {
        string value = GetArgValueFromKvp(node, name);
        if (value == null)
            return false;
        return value.ToUpper() == "YES";
    }

    public void LoadAnnouncementsFromHtmlFile(string source)
    {
        HtmlDocument htmlDoc = new HtmlDocument();

        htmlDoc.Load(source, Encoding.UTF8);

        // read the common stylesheet, if present
        HtmlNode commonStylesheet = GetInlineStylesheetAndClock(htmlDoc.DocumentNode, out string stylesheet, out int vectorClock);

        if (commonStylesheet != null)
        {
            CommonStylesheet = stylesheet;
            CommonStylesheetVectorClock = vectorClock;
        }

        // now find all the announcements
        string xpathWrappers = $"//div[@class='wrapper']";
        HtmlNodeCollection wrappers = htmlDoc.DocumentNode.SelectNodes(xpathWrappers);

        foreach (HtmlNode node in wrappers)
        {
            string xpath = $".//div[contains(@id, '{s_AnnounceDivIdPrefix}')]";
            HtmlNode announcementNode = node.SelectSingleNode(xpath);

            Announcement announce =
                new Announcement
                {
                    AnnouncementHtml = announcementNode.OuterHtml,
                    ShowAssigners = GetArgBoolValueFromKvp(node, "Assigners"),
                    ShowContacts = GetArgBoolValueFromKvp(node, "Contacts"),
                    Officials = GetArgValueFromKvp(node, "Officials"),
                    MatchString = announcementNode.GetAttributeValue("id", "")
                };
            string text = announcementNode.InnerText;
            announce.Hint = CollapseWhitespace(text.Substring(0, Math.Min(80, text.Length)));

            m_announcements.Add(announce);
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

        list.Columns.Add("ID", 110);
        list.Columns.Add("Hint", 400);
        list.Columns.Add("ShowAssigners", 50);
        list.Columns.Add("ShowContacts", 50);
        list.Columns.Add("Officials", 50);
        list.Columns.Add("Style Clock", 50);
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
            item.SubItems.Add(announce.StyleClock.ToString());
            item.Tag = announce;

            m_listView.Items.Add(item);
        }
    }

    /*----------------------------------------------------------------------------
        %%Function: UpdateRanksFromListView
        %%Qualified: ArbWeb.Announcements.WebAnnouncements.UpdateRanksFromListView
    ----------------------------------------------------------------------------*/
    public void UpdateRanksFromListView()
    {
        int rank = 0;

        foreach (ListViewItem lvi in m_listView.Items)
        {
            if (lvi.Tag is Announcement announce)
                announce.Rank = rank++;
        }
    }
}
