﻿using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OpenQA.Selenium;
using TCore.StatusBox;
using TCore.WebControl;

namespace ArbWeb
{
    public class WebNav
    {
	    private IAppContext m_appContext;

	    /*----------------------------------------------------------------------------
			%%Function:WebNav
			%%Qualified:ArbWeb.WebNav.WebNav
	    ----------------------------------------------------------------------------*/
	    public WebNav(IAppContext appContext)
	    {
		    m_appContext = appContext;
	    }
	    
        #region Navigation Support
        /*----------------------------------------------------------------------------
			%%Function:EnsureAdminLoggedIn
			%%Qualified:ArbWeb.WebNav.EnsureAdminLoggedIn
        ----------------------------------------------------------------------------*/
        private void EnsureAdminLoggedIn()
        {
	        List<IWebElement> elements = new List<IWebElement>(m_appContext.WebControl.Driver.FindElements(By.ClassName("accountsGridRow")));
	        bool fClickedAdmin = false;
	        
	        foreach (IWebElement element in elements)
	        {
		        if (element.Text.Contains("Admin"))
		        {
			        element.Click();
			        fClickedAdmin = true;
			        break;
		        }
	        }
	        
            Utils.ThrowIfNot(fClickedAdmin, "Can't find Admin account link");
            m_appContext.WebControl.WaitForPageLoad(200);
        }

        /*----------------------------------------------------------------------------
			%%Function:EnsureLoggedIn
			%%Qualified:ArbWeb.WebNav.EnsureLoggedIn
        ----------------------------------------------------------------------------*/
        public void EnsureLoggedIn()
        {
	        DoEnsureLoggedIn();
        }

        bool m_fLoggedIn;

        /*----------------------------------------------------------------------------
			%%Function:DoEnsureLoggedIn
			%%Qualified:ArbWeb.WebNav.DoEnsureLoggedIn
        
			BE CAREFUL - appContext delegates DoLogin to us, so don't get caught in
			an infinite loop
        ----------------------------------------------------------------------------*/
        private void DoEnsureLoggedIn()
        {
	        MicroTimer timer = new MicroTimer();
	        
	        m_appContext.EnsureWebControl();
	        
	        timer.Stop();
	        m_appContext.StatusReport.LogData($"EnsureWebControl elapsed: {timer.MsecFloat}", 1, MSGT.Body);
	        timer.Reset();
	        
            if (m_fLoggedIn == false)
            {
	            timer.Start();
	            
                m_appContext.StatusReport.AddMessage("Logging in...");
                m_appContext.StatusReport.PushLevel();
                
                // login to arbiter
                // nav to the main arbiter login page
                if (!m_appContext.WebControl.FNavToPage(WebCore._s_Home))
                    throw (new Exception("could not navigate to arbiter homepage!"));

                if (!WebControl.FCheckForControlId(m_appContext.WebControl.Driver, WebCore._sid_Home_Anchor_NeedHelpLink))
                {
                    WebControl.FSetTextForInputControlName(m_appContext.WebControl.Driver, WebCore._s_Home_Input_Email, m_appContext.Profile.UserID, false);
                    WebControl.FSetTextForInputControlName(m_appContext.WebControl.Driver, WebCore._s_Home_Input_Password, m_appContext.Profile.Password, false);

                    m_appContext.WebControl.FClickControlName(WebCore._s_Home_Button_SignIn);
                }

                if (WebControl.FCheckForControlId(m_appContext.WebControl.Driver, WebCore._sid_Home_Div_PnlAccounts))
                {
                    EnsureAdminLoggedIn();
                }

#if cantdo
                // at this point, we are either going to get to the main page, or we are going to get a
                // page asking us which account to login to
                while (count < 100 && (ArbWebControl.FCheckForControl(oDoc2, WebCore._sid_Home_Div_PnlAccounts) || !ArbWebControl.FCheckForControl(oDoc2, WebCore._sid_Home_Anchor_NeedHelpLink)))
                {
                    if (m_cbShowBrowser.Checked == false)
                    {
                        m_cbShowBrowser.Checked = true;
                        ChangeShowBrowser(null, null);
                        fToggledBrowser = true;
                    }

                    Application.DoEvents();
                    System.Threading.Thread.Sleep(100);
                    oDoc2 = m_awc.Document2;
                    count++;
                }

                if (fToggledBrowser)
                {
                    m_cbShowBrowser.Checked = false;
                    ChangeShowBrowser(null, null);
                }
#endif

                if (!WebControl.FCheckForControlId(m_appContext.WebControl.Driver, WebCore._sid_Home_MessagingText))
                    MessageBox.Show("Login failed for ArbiterOne!");
                else
                    m_fLoggedIn = true;

                // and wait for nav to complete
                m_appContext.WebControl.WaitForPageLoad(300);
                m_appContext.StatusReport.PopLevel();
                m_appContext.StatusReport.AddMessage("Completed login.");
                
                timer.Stop();
                m_appContext.StatusReport.LogData($"EnsureLoggedIn elapsed: {timer.MsecFloat}", 1, MSGT.Body);
            }

        }

        #endregion

    }
}
