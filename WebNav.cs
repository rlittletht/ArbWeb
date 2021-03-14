using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using mshtml;
using OpenQA.Selenium;
using StatusBox;

namespace ArbWeb
{
    public partial class AwMainForm
    {

        #region Navigation Support
        /* E N S U R E  A D M I N  L O G G E D  I N */
        /*----------------------------------------------------------------------------
	    	%%Function: EnsureAdminLoggedIn
	    	%%Qualified: ArbWeb.AwMainForm.EnsureAdminLoggedIn
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
        private void EnsureAdminLoggedIn()
        {
	        List<IWebElement> elements = new List<IWebElement>(m_webControl.Driver.FindElements(By.ClassName("accountsGridRow")));
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
	        
            ThrowIfNot(fClickedAdmin, "Can't find Admin account link");
            m_webControl.WaitForPageLoad(m_webControl.Driver, 200);
        }

        public delegate void EnsureLoggedInDel();

        /* E N S U R E  L O G G E D  I N */
        /*----------------------------------------------------------------------------
        	%%Function: EnsureLoggedIn
        	%%Qualified: ArbWeb.AwMainForm.EnsureLoggedIn
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public void EnsureLoggedIn()
        {
	        DoEnsureLoggedIn();
        }

        public void EnsureWebControl()
        {
	        if (m_webControl == null)
				m_webControl = new ArbWebControl_Selenium(this);
        }
        
        /* E N S U R E  L O G G E D  I N */
        /*----------------------------------------------------------------------------
        	%%Function: EnsureLoggedIn
        	%%Qualified: ArbWeb.AwMainForm.EnsureLoggedIn
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoEnsureLoggedIn()
        {
	        MicroTimer timer = new MicroTimer();
	        
	        EnsureWebControl();
	        
	        timer.Stop();
	        m_srpt.LogData($"EnsureWebControl elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
	        timer.Reset();
	        
            if (m_fLoggedIn == false)
            {
	            timer.Start();
	            
                m_srpt.AddMessage("Logging in...");
                m_srpt.PushLevel();
                
                // login to arbiter
                // nav to the main arbiter login page
                if (!m_webControl.FNavToPage(WebCore._s_Home))
                    throw (new Exception("could not navigate to arbiter homepage!"));

                if (!ArbWebControl_Selenium.FCheckForControl(m_webControl.Driver, WebCore._sid_Home_Anchor_NeedHelpLink))
                {
                    ArbWebControl_Selenium.FSetInputControlText(m_webControl.Driver, WebCore._s_Home_Input_Email, m_pr.UserID, false);
                    ArbWebControl_Selenium.FSetInputControlText(m_webControl.Driver, WebCore._s_Home_Input_Password, m_pr.Password, false);

                    m_webControl.FClickControl(WebCore._s_Home_Button_SignIn);
                }

                int count = 0;

                bool fToggledBrowser = false;

                if (ArbWebControl_Selenium.FCheckForControl(m_webControl.Driver, WebCore._sid_Home_Div_PnlAccounts))
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

                if (!ArbWebControl_Selenium.FCheckForControl(m_webControl.Driver, WebCore._sid_Home_Anchor_NeedHelpLink))
                    MessageBox.Show("Login failed for ArbiterOne!");
                else
                    m_fLoggedIn = true;

                // and wait for nav to complete
                m_webControl.WaitForPageLoad(m_webControl.Driver, 300);
                m_srpt.PopLevel();
                m_srpt.AddMessage("Completed login.");
                
                timer.Stop();
                m_srpt.LogData($"EnsureLoggedIn elapsed: {timer.MsecFloat}", 1, StatusRpt.MSGT.Body);
            }

        }

        /*----------------------------------------------------------------------------
        	%%Function: DebugModelessWait
        	%%Qualified: ArbWeb.AwMainForm.DebugModelessWait
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static void DebugModelessWait()
        {
            string s;

            TCore.UI.InputBox.ShowInputBoxModelessWait("DEBUG PAUSE", "Press OK or Cancel to continue", "String?", out s);
        }
        #endregion

    }
}
