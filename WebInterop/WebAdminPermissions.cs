using System;
using System.Collections.Generic;
using HtmlAgilityPack;
using Microsoft.Identity.Client.Platforms.Features.WinFormsLegacyWebUi;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using TCore.StatusBox;
using TCore.WebControl;
using static ArbWeb.Roster;

namespace ArbWeb
{
    public class WebAdminPermissions
    {
        private IAppContext m_appContext;

        public WebAdminPermissions(IAppContext appContext)
        {
            m_appContext = appContext;
        }

        Dictionary<string, string> BuildPermissionsValueMap()
        {
            return
                new Dictionary<string, string>()
                {
                    { "Edit permissions", null },
                    { "Rank officials", null },
                    { "Add/Delete officials", null },
                    { "Mark Official Active/Inactive", null }
                };
        }

        public void RemoveAdminLockout()
        {
            m_appContext.EnsureLoggedIn();

            m_appContext.StatusReport.AddMessage("Removing group admin lockouts...");

            string groupAdminValue = null;

            //            m_appContext.WebControl.RepeatIfNotCondition(
            //                d =>
            //                {
            //                    Utils.ThrowIfNot(m_appContext.WebControl.FNavToPage(WebCore.s_RightsEdit), "couldn't navigate to rights page");
            //                    return true;
            //                },
            //                d =>
            //                {
            //                    sGroupAdminValue = m_appContext.WebControl.GetOptionValueFromFilterOptionTextForControlName(
            //                        WebCore.s_RightsEdit_PermissionsTarget,
            //                        "Group Admin");
            //                    return sGroupAdminValue != null;
            //                },
            //                2,
            //                500);

            Utils.ThrowIfNot(m_appContext.WebControl.FNavToPage(WebCore.s_RightsEdit), "couldn't navigate to rights page");
            m_appContext.WebControl.WaitForPageLoad();

            groupAdminValue =
                m_appContext.WebControl.GetOptionValueFromFilterOptionTextForControlName(
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

                // we are enabling all, so find select all
                Utils.ThrowIfNot(
                    m_appContext.WebControl.FClickControlByXpath("//a[contains(@onclick, 'SelectAll(false)')]"),
                    $"could not click select all for {name}");
            }
        }
    }
}
