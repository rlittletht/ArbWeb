using System.Collections.Generic;
using System.Web.UI.WebControls;
using OpenQA.Selenium;
using TCore.StatusBox;

namespace ArbWeb
{
    public class WebTeams
    {
        public IAppContext m_appContext;

        // private WebControl m_webControl => m_appContext.WebControl;
        private IStatusReporter m_srpt => m_appContext.StatusReport;

        public WebTeams(IAppContext appContext)
        {
            m_appContext = appContext;
        }

        public void DeleteUnusedTeams()
        {
            m_appContext.StatusReport.AddMessage("Deleting teams with no games associated...");
            m_appContext.StatusReport.PushLevel();

            m_appContext.EnsureLoggedIn();
            Utils.ThrowIfNot(
                m_appContext.WebControl.FNavToPage(WebCore._s_TeamsView),
                "Couldn't nav to teams!");

            m_appContext.WebControl.WaitForPageLoad();

            // figure out how many pages we have
            // find all of the <a> tags with an href that targets a pagination postback
            IList<IWebElement> anchors = m_appContext.WebControl.Driver.FindElements(
                By.XPath(
                    $"//tr[@class='numericPaging']//a[contains(@href, '{WebCore._s_OfficialsView_PaginationHrefPostbackSubstr}')]"));
            List<string> plsHrefs = new List<string>();

            foreach (IWebElement anchor in anchors)
            {
                string href = anchor.GetAttribute("href");

                if (href != null && href.Contains(WebCore._s_OfficialsView_PaginationHrefPostbackSubstr))
                {
                    // we can't just remember this element because we will be navigating around.  instead we will
                    // just remember the entire target so we can find it again
                    plsHrefs.Add(href);
                }
            }


        }
    }
}
