using System;
using OpenQA.Selenium;
using System.Collections.Generic;
using System.Windows.Forms;
using SeleniumExtras.WaitHelpers;
using TCore.StatusBox;

namespace ArbWeb;

public class WebPagination
{
    public delegate bool VisitPageDelegate(int pageIndex);

    private readonly IAppContext m_appContext;
    public List<string> PaginationLinks = new();
    private string m_firstPageUrl;

    public WebPagination(IAppContext appContext)
    {
        m_appContext = appContext;
    }

    void EnsurePaginationEntries(int targetIndex)
    {
        while (targetIndex >= PaginationLinks.Count)
        {
            PaginationLinks.Add(null);
        }
    }

    void SetPaginationEntry(int targetIndex, string href)
    {
        EnsurePaginationEntries(targetIndex);
        PaginationLinks[targetIndex] = href;
    }

    bool TryGetPageHref(int targetIndex, out string href)
    {
        if (targetIndex >= PaginationLinks.Count)
        {
            href = null;
            return false;
        }

        href = PaginationLinks[targetIndex];
        return href != null;
    }

    /*----------------------------------------------------------------------------
        %%Function: ReadLinksFromPageIndex
        %%Qualified: ArbWeb.WebPagination.ReadLinksFromPageIndex

        we know that we are on page index currentPage.  read all of the pagination
        links from this page and remember them (remember, there won't be a link
        for currentPage -- this will be a gap in the list).

        most of these links should already exist (and we will just validate they
        are correct), but we likely will be able to fill in the link for page 0
    ----------------------------------------------------------------------------*/
    public void ReadLinksFromPageIndex(int currentPage, bool updateExistingValues)
    {
        string xpath = $"//tr[@class='numericPaging']//a[contains(@href, '{WebCore._s_GenericPagination_PaginationHrefPostbackHrefSubstr}') and contains(@href, '{WebCore._s_GenericPagination_PaginationHrefPostbackSubstr}')]";

        // figure out how many pages we have
        // find all of the <a> tags with an href that targets a pagination postback
        IList<IWebElement> anchors = m_appContext.WebControl.Driver.FindElements(
            By.XPath(xpath));

        int pageNum = 0;

        foreach (IWebElement anchor in anchors)
        {
            if (pageNum == currentPage)
                pageNum++;

            EnsurePaginationEntries(pageNum);

            string href = anchor.GetAttribute("href");

            if (href != null && href.Contains(WebCore._s_GenericPagination_PaginationHrefPostbackHrefSubstr))
            {
                string currentValue = PaginationLinks[pageNum];

                if (currentValue != null)
                    Utils.ThrowIfNot(currentValue == href, $"pagination href changed for page {pageNum}: {href} != {currentValue}");

                // we can't just remember this element because we will be navigating around.  instead we will
                // just remember the entire target so we can find it again
                SetPaginationEntry(pageNum, href);
                pageNum++;
            }
        }
    }

    public void ResetLinksFromStartingPage()
    {
        PaginationLinks.Clear();
        m_firstPageUrl = m_appContext.WebControl.Driver.Url;

        // when we read the links, its OK to update the existing values -- we are starting from the first
        // page and we are possible reloading after an edit.
        ReadLinksFromPageIndex(0, false);
        // we won't have page 0 because its the current page. we remembered the URL for it above so we 
        // can get back to page one.

        // when we navigate to a different page, we will take the opportunity to fill in the gaps.
    }

    public static void NavigateToGivenPaginatedPageIfNecessary(IAppContext appContext, string pageLink, out double msecFindElement, out double msecNavigate)
    {
        msecFindElement = 0.0;
        msecNavigate = 0.0;

        if (pageLink == "")
            return;

        MicroTimer timer = new MicroTimer();

        string sHref = pageLink;

        string sXpath = $"//a[@href={WebNav.ToXPath(sHref)}]";

        IWebElement anchor;

        try
        {
            anchor = appContext.WebControl.Driver.FindElement(By.XPath(sXpath));
            appContext.WebControl.WaitForCondition(
                ExpectedConditions.ElementToBeClickable(By.XPath(sXpath)),
                10000);

        }
        catch
        {
            anchor = null;
        }

        timer.Stop();
        msecFindElement = timer.MsecFloat;

        appContext.StatusReport.LogData($"Process All GivenPage find item by xpath) elapsedTime: {timer.MsecFloat}", 1, MSGT.Body);

        timer.Reset();
        timer.Start();

        anchor?.Click();
        timer.Stop();
        msecNavigate = timer.MsecFloat;
    }

    public void NavigateToPageByIndex(int pageIndex)
    {
        if (!TryGetPageHref(pageIndex, out string pageLink))
        {
            if (pageIndex == 0)
            {
                // navigate to the first page we have saved
                Utils.ThrowIfNot(m_appContext.WebControl.FNavToPage(m_firstPageUrl), "couldn't nav to first page");
                m_appContext.WebControl.WaitForPageLoad();
                return;
            }
            throw new Exception($"No link found for page index {pageIndex}");
        }

        NavigateToGivenPaginatedPageIfNecessary(m_appContext, pageLink, out double msecFindElement, out double msecNavigate);

        if (pageIndex != 0)
        {
            if (PaginationLinks[0] == null)
            {
                // get the pagination link for page 0 (and validate other links)
                ReadLinksFromPageIndex(pageIndex, false);
            }
        }
    }

    /*----------------------------------------------------------------------------
        %%Function: ProcessAllPagesFromStartingPage
        %%Qualified: ArbWeb.WebPagination.ProcessAllPagesFromStartingPage

        Visit all of the pages starting from the currently loaded page.
    ----------------------------------------------------------------------------*/
    public void ProcessAllPagesFromStartingPage(VisitPageDelegate visit)
    {
        // make sure the pagination row is visible (even if it has no clickable links)
        string XPath = "//tr[@class='numericPaging']";

        try
        {
            m_appContext.WebControl.WaitForCondition(ExpectedConditions.ElementIsVisible(By.XPath(XPath)), 10000);
        }
        catch (Exception e)
        {
            MessageBox.Show($"no pagination row on current page: {e}");
            throw;
        }

        // read the links from the current page
        ResetLinksFromStartingPage();

        if (!visit(0))
            return;

        int currentPage = 1;

        while (currentPage < PaginationLinks.Count)
        {
            NavigateToPageByIndex(currentPage);
            if (!visit(currentPage))
                return;

            currentPage++;
        }
    }
}
