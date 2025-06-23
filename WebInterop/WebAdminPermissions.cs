using System;
using System.Collections.Generic;
using HtmlAgilityPack;
using Microsoft.Identity.Client.Platforms.Features.WinFormsLegacyWebUi;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using TCore.StatusBox;
using TCore.UI;
using TCore.WebControl;
using static ArbWeb.Roster;

namespace ArbWeb
{
    public class WebAdminPermissions
    {
        private IAppContext m_appContext;

        private const string s_announceDiv = "ArbWebAnnounce_D9BulkMaintentance";

        public WebAdminPermissions(IAppContext appContext)
        {
            m_appContext = appContext;
        }

        Dictionary<string, string> BuildPermissionsValueMap()
        {
            return
                new Dictionary<string, string>()
                {
                    { "Rank officials", null },
                    { "Add/Delete officials", null },
                    { "Mark Official Active/Inactive", null }
                };
        }

        public void AddAdminLockout()
        {
            if (!InputBox.ShowInputBox("When will you be done?", DateTime.Now.AddHours(3).ToString("g"), out string response))
                return;

            SetLockoutState(true);

            string announceText = $@"
                <div id='{s_announceDiv}'>
                <h1>Bulk Maintenance</h1>

                <p>I&#39;m performing some bulk edits, so I have temporarily turned off roster permissions for ALL Assigners (you can still edit games).</p>

                <p>ETA: {response}</p>
                </div>";

            WebAnnounce announce = new WebAnnounce(m_appContext);

            announce.UpdateArbiterAnnouncement(s_announceDiv, announceText, true, false, "None");
        }

        public void RemoveAdminLockout()
        {
            SetLockoutState(false);

            WebAnnounce announce = new WebAnnounce(m_appContext);

            announce.UpdateArbiterAnnouncement(s_announceDiv, null, false, false, "None");
        }

        void SetLockoutState(bool setLockout)
        {
            m_appContext.EnsureLoggedIn();

            m_appContext.StatusReport.AddMessage("Removing group admin lockouts...");

            Utils.ThrowIfNot(m_appContext.WebControl.FNavToPage(WebCore.s_RightsEdit), "couldn't navigate to rights page");
            m_appContext.WebControl.WaitForPageLoad();

            string groupAdminValue = m_appContext.WebControl.GetOptionValueFromFilterOptionTextForControlName(
                WebCore.s_RightsEdit_PermissionsTarget,
                "Group Admin");

            m_appContext.WebControl.WaitForXpath($"//option[contains(text(), 'Group Admin')]", 1000);

            Utils.ThrowIfNot(
                m_appContext.WebControl.FSetSelectedOptionValueForControlName(WebCore.s_RightsEdit_PermissionsTarget, groupAdminValue),
                "Can't select Group Admin for permissions target");
            m_appContext.WebControl.WaitForPageLoad();

            Dictionary<string, string> permissionsValueMap = BuildPermissionsValueMap();

            m_appContext.WebControl.WaitForXpath($"//option[contains(text(), 'Rank officials')]", 1000);

            List<string> permissionsList = new List<string>(permissionsValueMap.Keys);

            foreach (string name in permissionsList)
            {
                string value =
                    m_appContext.WebControl.GetOptionValueFromFilterOptionTextForControlName(
                        WebCore.s_RightsEdit_AllowUsersTo,
                        name);

                Utils.ThrowIfNot(value != null, $"could not find options value for {name}");
                permissionsValueMap[name] = value;
            }

            foreach (string name in permissionsValueMap.Keys)
            {
                Utils.ThrowIfNot(
                    m_appContext.WebControl.FSetSelectedOptionValueForControlName(WebCore.s_RightsEdit_AllowUsersTo, permissionsValueMap[name]),
                    $"can't switch to permission '{name}'({permissionsValueMap[name]})");
                m_appContext.WebControl.WaitForPageLoad();

                if (setLockout)
                {
                    // we are disabline all, so find selectall(true)
                    Utils.ThrowIfNot(
                        m_appContext.WebControl.FClickControlByXpath("//a[contains(@onclick, 'SelectAll(true)')]"),
                        $"could not click select all for {name}");
                    Utils.ThrowIfNot(
                        m_appContext.WebControl.FClickControlId(WebCore.sid_RightsEdit_RemovePermission));
                }
                else
                {
                    // we are enabling all, so find selectall(false)
                    Utils.ThrowIfNot(
                        m_appContext.WebControl.FClickControlByXpath("//a[contains(@onclick, 'SelectAll(false)')]"),
                        $"could not click select all for {name}");
                    Utils.ThrowIfNot(
                        m_appContext.WebControl.FClickControlId(WebCore.sid_RightsEdit_AddPermission));
                }
            }
        }
    }
}
