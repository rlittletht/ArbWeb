using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using mshtml;

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
            // now we need to find the URGENT HELP NEEDED row
            IHTMLDocument2 oDoc2 = m_awc.Document2;
            IHTMLElementCollection hec = (IHTMLElementCollection)oDoc2.all.tags("span");

            string sCtl = null;

            foreach (IHTMLElement he in hec)
            {
                if (he.id == null)
                    continue;

                if (he.id.StartsWith(WebCore._sid_Login_Span_Type_Prefix) && he.innerText == "Admin")
                {
                    // found a span with the right prefix and innerText, now figure out its control
                    int ich = he.id.IndexOf(WebCore._sid_Login_Span_Type_Prefix);
                    if (ich >= 0)
                    {
                        sCtl = he.id.Substring(ich + WebCore._sid_Login_Span_Type_Prefix.Length, 5);
                    }
                    break;
                }
            }

            ThrowIfNot(sCtl != null, "Can't find Admin account link");

            m_awc.ResetNav();
            string sControl = BuildAnnName(WebCore._sid_Login_Anchor_TypeLink_Prefix, WebCore._sid_Login_Anchor_TypeLink_Suffix, sCtl);

            ThrowIfNot(m_awc.FClickControl(oDoc2, sControl), "Couldn't admin link");
            m_awc.FWaitForNavFinish();
        }

        public delegate void EnsureLoggedInDel();

        /* E N S U R E  L O G G E D  I N */
        /*----------------------------------------------------------------------------
        	%%Function: EnsureLoggedIn
        	%%Qualified: ArbWeb.AwMainForm.EnsureLoggedIn
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void EnsureLoggedIn()
        {
            if (m_awc.InvokeRequired)
            {
                m_awc.Invoke(new EnsureLoggedInDel(DoEnsureLoggedIn));
            }
            else
            {
                DoEnsureLoggedIn();
            }
        }

        /* E N S U R E  L O G G E D  I N */
        /*----------------------------------------------------------------------------
        	%%Function: EnsureLoggedIn
        	%%Qualified: ArbWeb.AwMainForm.EnsureLoggedIn
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoEnsureLoggedIn()
        {
            if (m_fLoggedIn == false)
            {
                m_srpt.AddMessage("Logging in...");
                m_srpt.PushLevel();
                // login to arbiter
                // nav to the main arbiter login page
                if (!m_awc.FNavToPage(WebCore._s_Home))
                    throw (new Exception("could not navigate to arbiter homepage!"));

                // if this control is already there, then we were auto-logged in...
                IHTMLDocument2 oDoc2 = m_awc.Document2;
                if (!ArbWebControl.FCheckForControl(oDoc2, WebCore._sid_Home_Anchor_NeedHelpLink))
                {
                    IHTMLDocument oDoc = m_awc.Document;
                    IHTMLDocument3 oDoc3 = m_awc.Document3;

                    ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_Home_Input_Email, m_pr.UserID, false);
                    ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_Home_Input_Password, m_pr.Password, false);

                    m_awc.ResetNav();
                    m_awc.FClickControl(oDoc2, WebCore._s_Home_Button_SignIn);
                    m_awc.FWaitForNavFinish();
                }

                int count = 0;

                oDoc2 = m_awc.Document2;
                bool fToggledBrowser = false;

                if (ArbWebControl.FCheckForControl(oDoc2, WebCore._sid_Home_Div_PnlAccounts))
                {
                    EnsureAdminLoggedIn();
                }

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

                oDoc2 = m_awc.Document2;
                if (!ArbWebControl.FCheckForControl(oDoc2, WebCore._sid_Home_Anchor_NeedHelpLink))
                    MessageBox.Show("Login failed for ArbiterOne!");
                else
                    m_fLoggedIn = true;

                // and wait for nav to complete
                m_awc.FWaitForNavFinish();
                m_awc.ReportNavState("after login complete");
                m_srpt.PopLevel();
                m_srpt.AddMessage("Completed login.");
            }
        }
        #endregion

    }
}
