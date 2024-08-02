using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.DevTools.V127.Network;
using SeleniumExtras.WaitHelpers;
using TCore.StatusBox;

namespace ArbWeb
{
    public class OfficialsRosterWebInterop
    {
        private IAppContext m_appContext;
        private readonly bool m_fAddOfficialsOnly;
        private readonly OfficialsVisitorDelegate m_pass1Vistor;
        private readonly OfficialsVisitorDelegate m_pass2Vistor;
        private readonly AddOfficialsVisitorDelegate m_addOfficialsVisitorDelegate;

        private readonly PostHandleRosterDelegate m_postHandleRosterDelegate;

        // if fNeedPaginationTraversal is true, then we need to make sure we are on the correct paginated
        // page before calling the pass1Visitor
        private readonly bool m_fNeedPaginationTraversal;

        public delegate void OfficialsVisitorDelegate(
            OfficialsRosterWebInterop interop, OfficialLinkInfo linkInfo, IRoster irst, IRoster irstServer, IRosterEntry irste, IRoster irstBuilding,
            bool fNotJustAdded, bool fMarkOnly);

        public delegate void AddOfficialsVisitorDelegate(List<IRosterEntry> plirste);
        public delegate void PostHandleRosterDelegate(IRoster irstUpload, IRoster irstBuilding);

        public OfficialsRosterWebInterop(
            IAppContext appContext, bool fNeedPass1OnUpload, bool fAddOfficialsOnly, OfficialsVisitorDelegate pass1Visitor,
            AddOfficialsVisitorDelegate doAddOfficialsVisitorDelegate,
            PostHandleRosterDelegate doPostHandleRosterDelegate, bool fNeedPaginationTraversal,
            OfficialsVisitorDelegate pass2Visitor = null)
        {
            m_appContext = appContext;
            m_fNeedPass1OnUpload = fNeedPass1OnUpload;
            m_fAddOfficialsOnly = fAddOfficialsOnly;
            m_pass1Vistor = pass1Visitor;
            m_pass2Vistor = pass2Visitor;
            m_addOfficialsVisitorDelegate = doAddOfficialsVisitorDelegate;
            m_postHandleRosterDelegate = doPostHandleRosterDelegate;
            m_fNeedPaginationTraversal = fNeedPaginationTraversal;
        }

        public class PageLinks
        {
            public PageLinks()
            {
                officialLinkInfos = new List<OfficialLinkInfo>();
            }

            public List<OfficialLinkInfo> officialLinkInfos;

            //	        public List<string> rgsLinks;
            //	        public List<string> rgsData;
            public int iCur;
        };

        /* P O P U L A T E  P G L  F R O M  P A G E  C O R E */
        /*----------------------------------------------------------------------------
            %%Function: PopulatePglFromPageCore
            %%Qualified: ArbWeb.AwMainForm.PopulatePglFromPageCore
            %%Contact: rlittle

            Return a PageLinks (page of links) from the give sUrl.

                rx3 is a match for either the link name or the link text
                rx4 is a match for the link name always (will supercede rx3)
                rxData, if set, is the match that will be used to populat the rgsData

            on exit, rgsLinkNames, rgsLinks, and (optionally) rgsData will be
            populated in the pglLinks

            we need to collect information from two separate places in the DOM --
            the Official Name (Last, First) will be in an anchor linking to the
            offical page (which we can get the official ID from).  Then we have the
            email address from which we can get the actual email address.

            because of this, we collect the email first, then note that we
            are looking for the official ID.  essentially, a state machine.. (albeit 2
            states)
        ----------------------------------------------------------------------------*/
        private void PopulatePglOfficialsFromPageCore(PageLinks pageLinks, string pageNavLink)
        {
            // grab the info from the current navigated page
            IWebElement table = m_appContext.WebControl.Driver.FindElement(By.XPath("//body"));

            string sHtml = table.GetAttribute("outerHTML");
            HtmlAgilityPack.HtmlDocument html = new HtmlAgilityPack.HtmlDocument();

            html.LoadHtml(sHtml);

            Regex rxMatchEditUser = new Regex(@"showEditUserModal\( *\d+ *, *\d+,[^\)]*\);");
            Regex rxMatchDeleteOfficial = new Regex(@"deleteOfficial\(\d* *, *(\d*) *\)");
            Regex rxData = new Regex("mailto:.*");

            bool fLookingForEmail = false;
            bool fLookingForOfficialID = false;

            string sXpath = $"//a"; // [contains(@href, '{sHrefSubstringMatch}')]

            // build up a list of probable index links
            HtmlNodeCollection links = html.DocumentNode.SelectSingleNode(".")
               .SelectNodes(sXpath);

            foreach (HtmlNode link in links)
            {
                string sLinkTarget = link.GetAttributeValue("href", "");
                string onClick = link.GetAttributeValue("onclick", "");

                if (sLinkTarget != null && rxData.IsMatch(sLinkTarget))
                {
                    if (fLookingForEmail)
                    {
                        // adjust the top item in officialLinkInfos...
                        pageLinks.officialLinkInfos[pageLinks.officialLinkInfos.Count - 1].sEmail = sLinkTarget;
                        fLookingForEmail = false;
                    }
                    else
                    {
                        m_appContext.StatusReport.AddMessage("Found (" + sLinkTarget + ") when not looking for email!", MSGT.Error);
                    }
                }

                if (sLinkTarget != null && rxMatchDeleteOfficial.IsMatch(sLinkTarget))
                {
                    if (fLookingForOfficialID)
                    {
                        GroupCollection groups = rxMatchDeleteOfficial.Match(sLinkTarget).Groups;

                        if (groups.Count == 2)
                        {
                            // adjust the top item in officialLinkInfos...
                            pageLinks.officialLinkInfos[pageLinks.officialLinkInfos.Count - 1].sOfficialID = groups[1].Value;
                            fLookingForOfficialID = false;
                        }
                        else
                        {
                            m_appContext.StatusReport.AddMessage($"Bad group match for official id: {sLinkTarget}, groups.count={groups.Count}", MSGT.Error);
                        }
                    }
                    else
                    {
                        m_appContext.StatusReport.AddMessage("Found (" + sLinkTarget + ") when not looking for official ID!", MSGT.Error);
                    }
                }

                if (rxMatchEditUser.IsMatch(onClick))
                {
                    OfficialLinkInfo officialLinkInfo = new OfficialLinkInfo();

                    officialLinkInfo.sEmail = "";
                    officialLinkInfo.PageNavLink = pageNavLink;
                    officialLinkInfo.sOfficialEditLink = onClick;
                    pageLinks.officialLinkInfos.Add(officialLinkInfo);

                    fLookingForEmail = true;
                    fLookingForOfficialID = true;
                }
            }

            pageLinks.iCur = 0;
        }

        private bool m_fNeedPass1OnUpload;

        /*----------------------------------------------------------------------------
            %%Function: DoCoreRosterSync
            %%Qualified: ArbWeb.AwMainForm.DoCoreRosterSync
            %%Contact: rlittle

            Do the core roster syncing.

            We are either syncing server->local (download)
            or local->server (upload).

            We are being given the list of links on
            the official's edit page, the roster that we are uploading (if any),
            and a list of officials to limit our handling to (this is used when
            we just added new officials and we just want to update their info/misc
            fields...)

            rstServer is the latest roster from the server -- useful for quickly
            determining what we need to update (without having to check the
            server again)
        ----------------------------------------------------------------------------*/
        private void DoCoreRosterSync(
            PageLinks pageLinks, IRoster rosterUploading, IRoster rosterBuilding, IRoster rosterServer, List<IRosterEntry> rosterEntriesToLimitTo)
        {
            Dictionary<string, bool> mpOfficials = new Dictionary<string, bool>();

            if (rosterEntriesToLimitTo != null)
            {
                foreach (IRosterEntry irsteCheck in rosterEntriesToLimitTo)
                {
                    mpOfficials.Add("MAILTO:" + irsteCheck.Email.ToUpper(), true);
                }
            }

            // we are going to make two passes over the links. the first pass must be done
            // page by page in order to allow us to click on links on the page. the second pass
            // can go to arbitrary other pages (breaking the back-chain).

            WebRoster.NavigateOfficialsPageAllOfficials(m_appContext);
            string currentPageNavLink = ""; // we start out on the paginated first page

            pageLinks.iCur = 0;

            while (pageLinks.iCur < pageLinks.officialLinkInfos.Count // we have links left to visit
                   && (rosterUploading == null
                       || m_fNeedPass1OnUpload)) // why is this condition part of the while?! rst and cbRankOnly never changes in the loop)
            {
                OfficialLinkInfo linkInfo = pageLinks.officialLinkInfos[pageLinks.iCur];

                // if we aren't uploading, or if we are uploading and we have values for this email address AND the current link has an email address
                if (rosterUploading == null
                    || (!String.IsNullOrEmpty(linkInfo.sEmail)
                        && rosterUploading.IrsteLookupEmail(linkInfo.sEmail) != null))
                {
                    if (m_fNeedPaginationTraversal && string.Compare(linkInfo.PageNavLink, currentPageNavLink, StringComparison.OrdinalIgnoreCase) != 0)
                    {
                        // got to the next paginated page
                        NavigateToPaginatedPageIfNecessary(linkInfo.PageNavLink, out _, out _);
                        currentPageNavLink = linkInfo.PageNavLink;
                    }

                    IRosterEntry irste;

                    // we need to build a RosterEntry for this. if we are downloading, then we will add it to the building roster
                    // otherwise, we will use it to compare with the one we are uploading to see if we need to update anything

                    if (rosterBuilding != null)
                        irste = rosterBuilding.CreateRosterEntry();
                    else
                        irste = rosterUploading.CreateRosterEntry();

                    bool fMarkOnly = false;

                    irste.SetEmail(linkInfo.sEmail);
                    m_appContext.StatusReport.AddMessage($"Processing roster info for {pageLinks.officialLinkInfos[pageLinks.iCur].sEmail}...");

                    if (m_fAddOfficialsOnly && rosterEntriesToLimitTo == null)
                        fMarkOnly = true;

                    if (rosterEntriesToLimitTo != null)
                    {
                        if (!mpOfficials.ContainsKey(linkInfo.sEmail.ToUpper()))
                        {
                            pageLinks.iCur++;
                            continue; // it doesn't match an official in the "limit-to" list.
                        }

                        fMarkOnly = false; // we want to process this one.
                    }

                    // if we don't have a limit list, then we aren't in pass right after adding officials
                    // (we build the limit list when we actually add officials)
                    bool fNotJustAdded = rosterEntriesToLimitTo == null && (rosterUploading == null || rosterUploading.IsUploadableRoster);
                    m_pass1Vistor(
                        this,
                        linkInfo,
                        rosterUploading,
                        rosterServer,
                        irste,
                        rosterBuilding,
                        fNotJustAdded,
                        fMarkOnly);

                    if (rosterUploading == null && !String.IsNullOrEmpty(irste.Email))
                    {
                        rosterBuilding.Add(irste);
                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(linkInfo.sEmail))
                        {
                            IRosterEntry irsteT = rosterUploading.IrsteLookupEmail(linkInfo.sEmail);

                            if (irsteT != null)
                                irsteT.Marked = true;
                        }
                    }

                    if (m_appContext.Profile.TestOnly)
                    {
                        break;
                    }
                }

                pageLinks.iCur++;
            }

            // don't do pass 2 if:
            //  We don't have a pass2 Visitor
            // OR We are adding officials and we DON'T have a rosterEntriesToLimitTo

            // (if we have rosterEntriesToLimitTo, that means we ARE adding officials, and we know exactly which officials we are adding because
            // we have already done the add and are trying to set the misc info and custom fields for them...)
            if (m_pass2Vistor != null && (!m_fAddOfficialsOnly || rosterEntriesToLimitTo != null))
            {
                // we have a pass 2. we won't do pagination navigation here
                pageLinks.iCur = 0;

                // we will have already created the roster entries, so just get the entry this time

                while (pageLinks.iCur < pageLinks.officialLinkInfos.Count // we have links left to visit
                       && (rosterUploading == null
                           || m_fNeedPass1OnUpload)) // why is this condition part of the while?! rst and cbRankOnly never changes in the loop)
                {
                    OfficialLinkInfo linkInfo = pageLinks.officialLinkInfos[pageLinks.iCur];

                    // if we aren't uploading, or if we are uploading and we have values for this email address AND the current link has an email address
                    if (rosterUploading == null
                        || (!String.IsNullOrEmpty(linkInfo.sEmail)
                            && rosterUploading.IrsteLookupEmail(linkInfo.sEmail) != null))
                    {
                        IRosterEntry irste;

                        // we need to build a RosterEntry for this. if we are downloading, then we will add it to the building roster
                        // otherwise, we will use it to compare with the one we are uploading to see if we need to update anything

                        if (rosterBuilding != null)
                            irste = rosterBuilding.IrsteLookupEmail(linkInfo.sEmail);
                        else
                            irste = rosterUploading.IrsteLookupEmail(linkInfo.sEmail);

                        bool fMarkOnly = false;

                        m_appContext.StatusReport.AddMessage($"Pass 2 Processing roster info for {pageLinks.officialLinkInfos[pageLinks.iCur].sEmail}...");

                        if (rosterEntriesToLimitTo != null)
                        {
                            if (!mpOfficials.ContainsKey(((string)pageLinks.officialLinkInfos[pageLinks.iCur].sEmail.ToUpper())))
                            {
                                pageLinks.iCur++;
                                continue; // it doesn't match an official in the "limit-to" list.
                            }
                        }

                        // if we don't have a limit list, then we aren't in pass right after adding officials
                        // (we build the limit list when we actually add officials)
                        bool fNotJustAdded = rosterEntriesToLimitTo == null && (rosterUploading == null || rosterUploading.IsUploadableRoster);
                        m_pass2Vistor(
                            this,
                            pageLinks.officialLinkInfos[pageLinks.iCur],
                            rosterUploading,
                            rosterServer,
                            irste,
                            rosterBuilding,
                            fNotJustAdded,
                            false /*fMarkOnly*/);
                    }

                    pageLinks.iCur++;
                }
            }
        }

        private void VOPC_PopulatePgl(Object o, string pageNavLink)
        {
            PopulatePglOfficialsFromPageCore((PageLinks)o, pageNavLink);
        }


        // object could be RST or PageLinks
        public delegate void VisitOfficialsPageCallback(Object o, string pageNavLink);

        public static string ToXPath(string value)
        {
            const string apostrophe = "'";
            const string quote = "\"";

            if (value.Contains(quote))
            {
                if (value.Contains(apostrophe))
                {
                    throw new Exception("Illegal XPath string literal.");
                }
                else
                {
                    return apostrophe + value + apostrophe;
                }
            }
            else
            {
                return quote + value + quote;
            }
        }

        List<string> GetAllPaginationLinksOnPage()
        {
            // figure out how many pages we have
            // find all of the <a> tags with an href that targets a pagination postback
            IList<IWebElement> anchors = m_appContext.WebControl.Driver.FindElements(
                By.XPath($"//tr[@class='numericPaging']//a[contains(@href, '{WebCore._s_OfficialsView_PaginationHrefPostbackSubstr}')]"));
            List<string> paginationLinks = new List<string>();

            foreach (IWebElement anchor in anchors)
            {
                string href = anchor.GetAttribute("href");

                if (href != null && href.Contains(WebCore._s_OfficialsView_PaginationHrefPostbackSubstr))
                {
                    // we can't just remember this element because we will be navigating around.  instead we will
                    // just remember the entire target so we can find it again
                    paginationLinks.Add(href);
                }
            }

            return paginationLinks;
        }

        public void NavigateToPaginatedPageIfNecessary(string pageLink, out double msecFindElement, out double msecNavigate)
        {
            msecFindElement = 0.0;
            msecNavigate = 0.0;

            if (pageLink == "")
                return;

            MicroTimer timer = new MicroTimer();

            string sHref = pageLink;

            string sXpath = $"//a[@href={ToXPath(sHref)}]";

            IWebElement anchor;

            try
            {
                anchor = m_appContext.WebControl.Driver.FindElement(By.XPath(sXpath));
            }
            catch
            {
                anchor = null;
            }

            timer.Stop();
            msecFindElement = timer.MsecFloat;

            m_appContext.StatusReport.LogData($"Process All Officials(find item by xpath) elapsedTime: {timer.MsecFloat}", 1, MSGT.Body);

            timer.Reset();
            timer.Start();

            anchor?.Click();
            timer.Stop();
            msecNavigate = timer.MsecFloat;
        }


        /*----------------------------------------------------------------------------
            %%Function: ProcessAllOfficialPages
            %%Qualified: ArbWeb.OfficialsRosterWebInterop.ProcessAllOfficialPages
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public void ProcessAllOfficialPages(VisitOfficialsPageCallback visit, Object o)
        {
            MicroTimer timer = new MicroTimer();

            WebRoster.NavigateOfficialsPageAllOfficials(m_appContext);

            // first, get the first pages and callback
            timer.Stop();
            m_appContext.StatusReport.LogData($"Process All Officials(NavigateAtStart) elapsedTime: {timer.MsecFloat}", 1, MSGT.Body);

            visit(o, "");

            timer.Reset();
            timer.Start();

            List<string> paginationLinks = GetAllPaginationLinksOnPage();

            timer.Stop();
            m_appContext.StatusReport.LogData($"Process All Officials(buildAnchorList) elapsedTime: {timer.MsecFloat}", 1, MSGT.Body);

            // now, we are going to navigate to each page by finding and clicking each pagination link in turn
            foreach (string sHref in paginationLinks)
            {
                NavigateToPaginatedPageIfNecessary(sHref, out double msecFind, out double msecNavigate);
                m_appContext.StatusReport.LogData($"Process All Officials(find item by xpath) elapsedTime: {msecFind}", 1, MSGT.Body);
                m_appContext.StatusReport.LogData($"Process All Officials(anchor click) elapsedTime: {msecNavigate}", 1, MSGT.Body);

                visit(o, sHref);
            }
        }

        private PageLinks PglGetOfficialsFromWeb()
        {
            m_appContext.EnsureLoggedIn();

            PageLinks pageLinks = new PageLinks();

            ProcessAllOfficialPages(VOPC_PopulatePgl, pageLinks);

            // if there are no links, then we aren't logged in yet
            if (pageLinks.officialLinkInfos.Count == 0)
            {
                throw (new Exception("Not logged in after EnsureLoggedIn()!!"));
            }

            return pageLinks;
        }


        public delegate void HandleRosterPostUpdateDelegate(OfficialsRosterWebInterop gr, IRoster irst);

        /*----------------------------------------------------------------------------
            %%Function: GenericVisitRoster
            %%Qualified: ArbWeb.OfficialsRosterWebInterop.GenericVisitRoster
            %%Contact: rlittle

            If rst == null, then we're downloading the roster.  Otherwise, we are
            uploading

            FUTURE: Make this a generic "VisitRoster" with callbacks or methods
            specific to upload or download (i.e. core out the code shared by
            upload and download, then make separate upload and download functions
            with no duplication)
        ----------------------------------------------------------------------------*/
        public void GenericVisitRoster(
            IRoster irstUpload, IRoster irstBuilding, string sOutFile, IRoster irstServer, HandleRosterPostUpdateDelegate handleRosterPostUpdate)
        {
            //Roster rstBuilding = null;
            PageLinks pageLinks;

            if (irstUpload != null && irstBuilding != null)
                throw new Exception("cannot upload AND download at the same time");

            // we're not going to write the roster out until the end now...

            //if (rstUpload == null)
            //rstBuilding = new Roster();

            pageLinks = PglGetOfficialsFromWeb();
            DoCoreRosterSync(pageLinks, irstUpload, irstBuilding, irstServer, null /*plrsteLimit*/);

            handleRosterPostUpdate?.Invoke(this, irstBuilding);

            if (irstUpload != null)
            {
                List<IRosterEntry> plirsteUnmarked = irstUpload.PlirsteUnmarked();

                // we might have some officials left "unmarked".  These need to be added

                // at this point, all the officials have either been marked or need to 
                // be added

                if (plirsteUnmarked.Count > 0)
                {
                    if (MessageBox.Show($"There are {plirsteUnmarked.Count} new officials.  Add these officials?", "ArbWeb", MessageBoxButtons.YesNo)
                        == DialogResult.Yes)
                    {
                        m_addOfficialsVisitorDelegate?.Invoke(plirsteUnmarked);
                        // now we have to reload the page of links and do the whole thing again (updating info, etc)
                        // so we get the misc fields updated.  Then fall through to the rankings and do everyone at
                        // once
                        pageLinks = PglGetOfficialsFromWeb(); // refresh to get new officials
                        DoCoreRosterSync(pageLinks, irstUpload, null /*rstBuilding*/, irstServer, plirsteUnmarked);
                        // now we can fall through to our core ranking handling...
                    }
                }
            }

            // now, do the rankings.  this is easiest done in the bulk rankings tool...
            m_postHandleRosterDelegate?.Invoke(irstUpload, irstBuilding);
            // lastly, if we're downloading, then output the roster

            if (irstUpload == null)
                irstBuilding.WriteRoster(sOutFile);

            if (m_appContext.Profile.TestOnly)
            {
                MessageBox.Show("Stopping after 1 roster item");
            }
        }

        /*----------------------------------------------------------------------------
            %%Function: SBuildRosterFilename
            %%Qualified: ArbWeb.AwMainForm.SBuildRosterFilename
            %%Contact: rlittle

        ----------------------------------------------------------------------------*/
        public static string SBuildRosterFilename(string sRosterName)
        {
            return WebCore.BuildDownloadFilenameFromTemplate(sRosterName, "games");
        }

        class WebPhoneInfo
        {
            public string Number;
            public string Extension;
            public string Type;
        }

        /*----------------------------------------------------------------------------
            %%Function: GetPhoneNumbersFromAntDialog
            %%Qualified: ArbWeb.OfficialsRosterWebInterop.GetPhoneNumbersFromAntDialog
        ----------------------------------------------------------------------------*/
        List<WebPhoneInfo> GetPhoneNumbersFromAntDialog()
        {
            List<WebPhoneInfo> phones = new();

            string _sid_OfficialsView_EditAccount_Phones_Prefix = "phones_";
            string _sid_OfficialsView_EditAccount_Phones_NumberSuffix = "_phoneNumber";
            string _sid_OfficialsView_EditAccount_Phones_ExtensionSuffix = "_extension";
            string _sid_OfficialsView_EditAccount_Phones_TypeIdSuffix = "_phoneTypeId";

            int phoneNum = 0;

            // keep reading phones until we can't read any more
            try
            {
                while (true)
                {
                    string numberId = $"{_sid_OfficialsView_EditAccount_Phones_Prefix}{phoneNum}{_sid_OfficialsView_EditAccount_Phones_NumberSuffix}";
                    string extensionId = $"{_sid_OfficialsView_EditAccount_Phones_Prefix}{phoneNum}{_sid_OfficialsView_EditAccount_Phones_ExtensionSuffix}";
                    string typeId = $"{_sid_OfficialsView_EditAccount_Phones_Prefix}{phoneNum}{_sid_OfficialsView_EditAccount_Phones_TypeIdSuffix}";

                    WebPhoneInfo phone = new WebPhoneInfo();

                    // try to get the phone number
                    phone.Number = m_appContext.WebControl.GetValueForControlId($"{numberId}");
                    phone.Extension = m_appContext.WebControl.GetValueForControlId($"{extensionId}");

                    // to get the type, we have to get the control for the type, then look at its sibling for the selection item
                    phone.Type = m_appContext.WebControl.GetSelectedValueFromAntControlId(
                        m_appContext.WebControl.Driver,
                        typeId,
                        "span[@class='ant-select-selection-item']");

                    phones.Add(phone);
                    phoneNum++;
                }
            }
            catch
            {
                return phones;
            }
        }


        /*----------------------------------------------------------------------------
            %%Function: FQueryAssignTextByIdSibling
            %%Qualified: ArbWeb.OfficialsRosterWebInterop.FQueryAssignTextByIdSibling

            Given an id for a control and optionally new value for it, check if the
            control needs updated. return the original value in sOriginalValue.

            if we made a change, set fNeedSave, otherwise leave untouched.
            if we failed to access the control, set fFailAssign
        ----------------------------------------------------------------------------*/
        private bool FQueryAssignTextByIdSibling(string id, string sNewValue, out string sOriginalValue, ref bool fNeedSave, ref bool fFailAssign)
        {
            string antDropdown = $"{id}_list";
            string xpathDropdown = $"//div[@id='{antDropdown}']/div[@aria-label='{sNewValue}']";

            sOriginalValue = null;

            IWebElement input = m_appContext.WebControl.Driver.FindElement(By.Id(id));
            IWebElement sibling = m_appContext.WebControl.GetSiblingFromAntControlId(
                m_appContext.WebControl.Driver,
                id,
                "span[@class='ant-select-selection-item']");

            if (sibling == null)
            {
                fFailAssign = true;
                return false;
            }

            string value = sibling.Text;

            sOriginalValue = value;

            if (sOriginalValue == null)
                sOriginalValue = "";

            if (sNewValue != null)
            {
                // check to see if it matches what we have
                // find a match on email address first
                if (String.Compare(sNewValue, sOriginalValue, StringComparison.OrdinalIgnoreCase /*ignoreCase*/) != 0)
                {
                    TCore.WebControl.WebControl.FSetTextForAntInputControl(input, sNewValue, false);
                    fNeedSave = true;
                }
            }

            return true;
        }

        /*----------------------------------------------------------------------------
            %%Function: FMatchAssignTextById
            %%Qualified: ArbWeb.OfficialsRosterWebInterop.FMatchAssignTextById

           Get the value for the control sMatch. Determine the current value of
           that control (which will be returned in sOriginalValue).

           sNewValue is not null and does not match the controls value, then
           set the control to sNewValue (which will strangely leave sOriginalValue
           set to the OLD value)

           return true if we found the control, false if we didn't
        ----------------------------------------------------------------------------*/
        private bool FMatchAssignTextById(IWebElement input, string sMatch, string sNewValue, out string sAssign, ref bool fNeedSave, ref bool fFailAssign)
        {
            string id = input.GetAttribute("id");
            string value = input.GetAttribute("value");

            sAssign = null;

            if (!id.Contains(sMatch))
                return false;

            sAssign = value;

            if (sAssign == null)
                sAssign = "";

            if (sNewValue != null)
            {
                // check to see if it matches what we have
                // find a match on email address first
                if (String.Compare(sNewValue, sAssign, StringComparison.OrdinalIgnoreCase /*ignoreCase*/) != 0)
                {
                    if (!input.Enabled)
                    {
                        fFailAssign = true;
                    }
                    else
                    {
                        m_appContext.WebControl.FSetTextForAntInputControlId(id, sNewValue, false);
                        fNeedSave = true;
                    }
                }
            }

            return true;
        }

        /*----------------------------------------------------------------------------
            %%Function: DismissOfficialsEditPopup
            %%Qualified: ArbWeb.OfficialsRosterWebInterop.DismissOfficialsEditPopup
        ----------------------------------------------------------------------------*/
        private void DismissOfficialsEditPopup()
        {
            IWebElement cancelButton = m_appContext.WebControl.Driver.FindElement(By.Id(WebCore._sid_OfficialsView_EditAccount_ButtonCancel));
            cancelButton.Click();

            m_appContext.WebControl.WaitForCondition(
                ExpectedConditions.InvisibilityOfElementLocated(By.XPath(WebCore._xpath_modalDialogRoot)),
                2000);

            {
                string xpath = WebCore._xpath_modalDialogRoot;

                IWebElement element = m_appContext.WebControl.GetElementBy(By.XPath(xpath));
            }
        }

        private void SaveOfficialsInfoFromPopup(bool fNeedSave, bool fHadSaveInfo)
        {
            IWebElement saveButton = m_appContext.WebControl.Driver.FindElement(By.Id(WebCore._sid_OfficialsView_EditAccount_ButtonSave));

            if (!fNeedSave && !saveButton.Enabled)
                return;

            if (!fHadSaveInfo)
            {
                MessageBox.Show("Huh? no new roster entry -- shouldn't be saving");
                return;
            }

            // the save process is:
            // 1) Make sure the previous sibling to the save button is nbsp
            // 2) Click save button
            // 3) Wait for sibling to be Saving
            // 4) Wait for sibling to be nbsp again


            if (!saveButton.Enabled)
            {
                m_appContext.StatusReport.AddMessage($"Save button not enabled", MSGT.Error);
                return;
            }

            string text = "";
            string tagName = "";

            TCore.WebControl.StaleSafeAccess safeAccessPreviousSibling = new(
                () => m_appContext.WebControl.GetPreviousSiblingFromAntControlId(WebCore._sid_OfficialsView_EditAccount_ButtonSave, "*[3]"),
                (IWebElement element) =>
                {
                    text = element?.Text;
                    tagName = element?.TagName;
                    return true;
                });

            safeAccessPreviousSibling.Access();

            m_appContext.StatusReport.AddMessage($"Text before save click: {text}", MSGT.Body);

            saveButton.Click();
            int cRetry = 0;

            do
            {
                safeAccessPreviousSibling.Access();
                m_appContext.StatusReport.AddMessage($"Text waiting for span: {text}", MSGT.Body);
                if (tagName != "span")
                    Thread.Sleep(100);
            } while (tagName != "span" && cRetry++ < 200);

            m_appContext.StatusReport.AddMessage($"Done looking for span", MSGT.Body);

            // now wait for the div to come back
            cRetry = 0;

            do
            {
                safeAccessPreviousSibling.Access();
                m_appContext.StatusReport.AddMessage($"Text waiting for span to disappear: {text}", MSGT.Body);
                if (tagName == "span")
                    Thread.Sleep(100);
            } while (tagName == "span" && cRetry++ < 200);

            m_appContext.StatusReport.AddMessage($"Done looking for span to disappear", MSGT.Body);
        }

        /*----------------------------------------------------------------------------
            %%Function: SyncRosterEntryWithServerNew
            %%Qualified: ArbWeb.OfficialsRosterWebInterop.SyncRosterEntryWithServerNew
        ----------------------------------------------------------------------------*/
        private void SyncRosterEntryWithServerNew(OfficialLinkInfo linkInfo, RosterEntry rsteOut, RosterEntry rsteNew)
        {
            bool fFailUpdate = false;
            bool fNeedSave = false;

            // now make sure we are on the correct page

            // the linkInfo will tell us what pagination link we have to click on to get here. If that
            // link isn't present, then we are already on the correct page.  (NOTE: We should already be
            // on the correct page!)
            NavigateToPaginatedPageIfNecessary(linkInfo.PageNavLink, out _, out _);

            // before we click on edit for the official, get some info from the main page
            string readyText = m_appContext.WebControl.GetTextFromControlId(
                m_appContext.WebControl.Driver,
                $"{WebCore._sid_OfficialsView_IsReadyStatusPrefix}{linkInfo.sOfficialID}");

            rsteOut.m_fReady = string.Compare(readyText, "Mark Not Ready", StringComparison.OrdinalIgnoreCase) == 0;

            // now start harvesting from the edit page...

            // simulate a click on the edit official link
            if (!m_appContext.WebControl.FClickControlByXpath($"//a[@onclick='{linkInfo.sOfficialEditLink}']"))
                throw (new Exception($"couldn't click on control for {linkInfo.sOfficialEditLink}"));

            {
                string xpath = WebCore._xpath_modalDialogRoot;

                IWebElement element = m_appContext.WebControl.GetElementBy(By.XPath(xpath));

                bool visible = element.Displayed;
                visible = element.Enabled;
            }

            try
            {
                m_appContext.WebControl.WaitForCondition(
                    ExpectedConditions.ElementIsVisible(By.XPath(WebCore._xpath_modalDialogRoot)),
                    10000);

                // can't wait for FirstName here because if its read only it will never be true
                m_appContext.WebControl.WaitForCondition(
                    ExpectedConditions.ElementIsVisible(By.Id(WebCore._sid_OfficialsView_EditAccount_NickName)),
                    10000);

                m_appContext.WebControl.WaitForCondition(
                    ExpectedConditions.ElementToBeClickable(By.Id(WebCore._sid_OfficialsView_EditAccount_NickName)),
                    10000);
            }
            catch (Exception e)
            {
                MessageBox.Show($"{e}");
            }

            List<WebPhoneInfo> phones = GetPhoneNumbersFromAntDialog();

            string email = m_appContext.WebControl.GetValueForControlIdOrNull(WebCore._sid_OfficialsView_EditAccount_Email0);

            if (email != null && rsteOut.Email != null)
            {
                if (String.Compare(email, rsteOut.Email, StringComparison.OrdinalIgnoreCase) != 0)
                    throw new Exception("email addresses don't match!");
            }
            else
            {
                m_appContext.StatusReport.AddMessage($"NULL Email address for {rsteOut.First},{rsteOut.Last}", MSGT.Error);
                DismissOfficialsEditPopup();
                return;
            }

            // now iterate over all the input controls
            ReadOnlyCollection<IWebElement> inputs = m_appContext.WebControl.Driver.FindElements(By.XPath("//div[@role='dialog']//input"));
            string assigned;
            foreach (IWebElement input in inputs)
            {
                if (FMatchAssignTextById(input, WebCore._sid_OfficialsView_EditAccount_FirstName, rsteNew?.First, out assigned, ref fNeedSave, ref fFailUpdate))
                    rsteOut.First = assigned;
                if (FMatchAssignTextById(input, WebCore._sid_OfficialsView_EditAccount_LastName, rsteNew?.Last, out assigned, ref fNeedSave, ref fFailUpdate))
                    rsteOut.Last = assigned;
                if (FMatchAssignTextById(input, WebCore._sid_OfficialsView_EditAccount_MiddleName, rsteNew?.Middle, out assigned, ref fNeedSave, ref fFailUpdate))
                    rsteOut.Middle = assigned;
                if (FMatchAssignTextById(input, WebCore._sid_OfficialsView_EditAccount_Address1, rsteNew?.Address1, out assigned, ref fNeedSave, ref fFailUpdate))
                    rsteOut.Address1 = assigned;
                if (FMatchAssignTextById(input, WebCore._sid_OfficialsView_EditAccount_Address2, rsteNew?.Address2, out assigned, ref fNeedSave, ref fFailUpdate))
                    rsteOut.Address2 = assigned;
                if (FMatchAssignTextById(input, WebCore._sid_OfficialsView_EditAccount_CityName, rsteNew?.City, out assigned, ref fNeedSave, ref fFailUpdate))
                    rsteOut.City = assigned;
                if (FMatchAssignTextById(input, WebCore._sid_OfficialsView_EditAccount_ZipCode, rsteNew?.Zip, out assigned, ref fNeedSave, ref fFailUpdate))
                    rsteOut.Zip = assigned;
                // read only
                if (FMatchAssignTextById(input, WebCore._sid_OfficialsView_EditAccount_BirthDate, null, out assigned, ref fNeedSave, ref fFailUpdate))
                    rsteOut.m_sDateOfBirth = assigned;
            }

            // and query the ones that have to come from the sibling
            if (FQueryAssignTextByIdSibling(WebCore._sid_OfficialsView_EditAccount_State, rsteNew?.State, out assigned, ref fNeedSave, ref fFailUpdate))
                rsteOut.State = assigned;

            SaveOfficialsInfoFromPopup(fNeedSave, rsteNew != null);

            // and lastly, cancel to get back
            DismissOfficialsEditPopup();
        }


        /*----------------------------------------------------------------------------
            %%Function:SetServerRosterInfo
            %%Qualified:ArbWeb.WebRoster.SetServerRosterInfo
        ----------------------------------------------------------------------------*/
        public void SetServerRosterInfo(OfficialLinkInfo linkInfo, IRoster irst, IRoster irstServer, RosterEntry rste, bool fMarkOnly)
        {
            RosterEntry rsteNew = null;
            RosterEntry rsteServer = null;

            if (irst != null)
                rsteNew = (RosterEntry)irst.IrsteLookupEmail(linkInfo.sEmail);

            if (rsteNew == null)
                rsteNew = new RosterEntry(); // just to get nulls filled in to the member variables
            else
                rsteNew.Marked = true;

            if (fMarkOnly)
                return;

            if (irstServer != null)
            {
                rsteServer = (RosterEntry)irstServer.IrsteLookupEmail(linkInfo.sEmail);
                if (rsteServer == null)
                {
                    m_appContext.StatusReport.AddMessage($"NULL Server entry for {linkInfo.sEmail}, SKIPPING", MSGT.Error);
                    return;
                }

                if (rsteNew.FEquals(rsteServer))
                    return;
            }

            SyncRosterEntryWithServerNew(linkInfo, rste, rsteNew);
        }

        /*----------------------------------------------------------------------------
            %%Function:GetRosterInfoFromServer
            %%Qualified:ArbWeb.WebRoster.GetRosterInfoFromServer

            Get the roster information from the server.
        ----------------------------------------------------------------------------*/
        public void GetRosterInfoFromServer(OfficialLinkInfo linkInfo, RosterEntry rste)
        {
            SyncRosterEntryWithServerNew(linkInfo, rste, null);

            if (rste.Address1 == null
                || rste.Address2 == null
                || rste.City == null
                || rste.m_sDateOfBirth == null
                || rste.Email == null
                || rste.First == null
                || rste.Last == null
                || rste.State == null
                || rste.Zip == null)
            {
                throw new Exception("couldn't extract one more more fields from official info");
            }
        }
    }
}
