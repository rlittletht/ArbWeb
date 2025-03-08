using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ArbWeb.Announcements;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using TCore.StatusBox;
using TCore.WebControl;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

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

        void CheckAnnounceControl(string prefix, string suffix, string ctl, bool enabled)
        {
            string sidEnable = WebAnnouncements.BuildAnnouncementNameOrIdString(prefix, suffix, ctl);

            m_appContext.WebControl.FSetCheckboxControlIdVal(enabled, sidEnable);
        }

        public void UpdateArbiterAnnouncement(
            string containingDiv,
            string announcement,
            bool? assignersEnabled,
            bool? contactsEnabled,
            string officialsVisibleTo)
        {
            if (announcement != null && !announcement.Contains(containingDiv))
            {
                Utils.ThrowIfNot(false, "announcement doesn't contain containing div");
            }

            m_appContext.StatusReport.AddMessage($"Starting Announcement Set for '{containingDiv}'...");
            m_appContext.StatusReport.PushLevel();

            // now we need to find the URGENT HELP NEEDED row
            HtmlDocument html = WebAnnouncements.GetHtmlDocumentForAnnouncementsPage(m_appContext);

            string sXpath = $"//div[@id='{containingDiv}']";

            HtmlNode node = html.DocumentNode.SelectSingleNode(sXpath);

            string sCtl = null;

            m_appContext.StatusReport.LogData($"Found {containingDiv} DIV, looking for parent TR element", 3, MSGT.Body);

            // ok, go up to the parent TR.

            HtmlNode nodeFind = node;

            while (nodeFind.Name.ToLower() != "tr")
            {
                nodeFind = nodeFind.ParentNode;
                Utils.ThrowIfNot(nodeFind != null, "Can't find HELP announcement");
            }

            m_appContext.StatusReport.LogData($"Found {containingDiv} parent TR", 3, MSGT.Body);

            // now find one of our controls and get its control number
            string s = nodeFind.InnerHtml;
            int ich = s.IndexOf(WebCore._s_Announcements_Button_Edit_Prefix);
            if (ich > 0)
            {
                sCtl = s.Substring(ich + WebCore._s_Announcements_Button_Edit_Prefix.Length, 5);
            }

            m_appContext.StatusReport.LogData($"Extracted ID for announcment to set: {sCtl}", 3, MSGT.Body);

            Utils.ThrowIfNot(sCtl != null, $"Can't find {containingDiv} announcement");

            string sidControl = WebAnnouncements.BuildAnnouncementNameOrIdString(
                WebCore._sid_Announcements_Button_Edit_Prefix,
                WebCore._sid_Announcements_Button_Edit_Suffix,
                sCtl);

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
                    string cancel = WebAnnouncements.BuildAnnouncementNameOrIdString(
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

            if (announcement != null)
            {
                if (!sourceOn)
                {
                    // select source mode
                    Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId("cke_12"), "Couldn't find <SOURCE> button");
                    m_appContext.WebControl.WaitForPageLoad();
                }

                m_appContext.WebControl.FSetTextAreaTextForControlAsChildOfDivId("cke_1_contents", announcement, true);
                m_appContext.WebControl.WaitForPageLoad();
            }
            // and now save it.

            // now set the visibility as requested
            if (assignersEnabled != null)
            {
                CheckAnnounceControl(
                    WebCore._sid_Announcements_Button_ToAssigners_Prefix,
                    WebCore._sid_Announcements_Button_ToAssigners_Suffix,
                    sCtl,
                    assignersEnabled.Value);
            }

            if (contactsEnabled != null)
            {
                CheckAnnounceControl(
                    WebCore._sid_Announcements_Button_ToContacts_Prefix,
                    WebCore._sid_Announcements_Button_ToContacts_Suffix,
                    sCtl,
                    contactsEnabled.Value);
            }

            // and lastly, choose the officials this is visible to
            if (officialsVisibleTo != null)
            {
                string sidEnable = WebAnnouncements.BuildAnnouncementNameOrIdString(
                    WebCore._s_Announcements_Button_ToOfficials_Prefix,
                    WebCore._s_Announcements_Button_ToOfficials_Suffix,
                    sCtl);

                string value =
                    m_appContext.WebControl.GetOptionValueFromFilterOptionTextForControlName(
                        sidEnable,
                        officialsVisibleTo);

                Utils.ThrowIfNot(value != null, $"Can't set announcement visible to {officialsVisibleTo}");

                string current = m_appContext.WebControl.GetSelectedOptionValueFromSelectControlName(sidEnable);
                if (current != value)
                {
                    Utils.ThrowIfNot(
                        m_appContext.WebControl.FSetSelectedOptionValueForControlName(sidEnable, value),
                        $"can't set officials visible to to {value}");
                }
            }

            sidControl = WebAnnouncements.BuildAnnouncementNameOrIdString(
                WebCore._sid_Announcements_Button_Save_Prefix,
                WebCore._sid_Announcements_Button_Save_Suffix,
                sCtl);

            Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(sidControl), "Couldn't find save button");
            m_appContext.WebControl.WaitForPageLoad();

            m_appContext.StatusReport.PopLevel();
            m_appContext.StatusReport.AddMessage("Completed Announcement Set.");
        }

        /*----------------------------------------------------------------------------
			%%Function:SetArbiterAnnounce
			%%Qualified:ArbWeb.WebAnnounce.SetArbiterAnnounce
		----------------------------------------------------------------------------*/
        public void SetArbiterAnnounce(string sArbiterHelpNeeded)
        {
            UpdateArbiterAnnouncement("D9UrgentHelpNeeded", sArbiterHelpNeeded, null, null, "All Officials");
        }
    }
}
