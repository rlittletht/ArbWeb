using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using TCore.StatusBox;
using OpenQA.Selenium;
using TCore.WebControl;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace ArbWeb
{
	public partial class WebRoster
    {
        /*----------------------------------------------------------------------------
			%%Function:SetServerMiscFields
			%%Qualified:ArbWeb.WebRoster.SetServerMiscFields
        ----------------------------------------------------------------------------*/
	    void SetServerMiscFields(string sEmail, string sOfficialID, IRoster irstUploading, IRoster irstServer, RosterEntry rste)
        {
            Roster rstServer = (Roster) irstServer;
            Roster rstUploading = (Roster) irstUploading;

            RosterEntry rsteNew = (RosterEntry)irstUploading.IrsteLookupEmail(rste.Email);
            RosterEntry rsteServer = (RosterEntry)irstServer?.IrsteLookupEmail(rste.Email);

            if (rsteNew.FEqualsMisc(rsteServer))
                return;

            List<string> plsMiscServer = rstServer.PlsMisc;

            List<string> plsValue = SyncPlsMiscWithServer(sEmail, sOfficialID, rsteNew.Misc, rstUploading.PlsMisc, ref plsMiscServer);

            rstServer.PlsMisc = plsMiscServer;

            if (plsValue.Count == 0)
                throw new Exception("couldn't extract misc field for official");

            rste.m_plsMisc = plsValue;
        }

        /*----------------------------------------------------------------------------
			%%Function:SetServerRosterInfo
			%%Qualified:ArbWeb.WebRoster.SetServerRosterInfo
        ----------------------------------------------------------------------------*/
        void SetServerRosterInfo(string sEmail, string sOfficialID, IRoster irst, IRoster irstServer, RosterEntry rste, bool fMarkOnly)
        {
            RosterEntry rsteNew = null;
            RosterEntry rsteServer = null;

            if (irst != null)
                rsteNew = (RosterEntry)irst.IrsteLookupEmail(sEmail);

            if (rsteNew == null)
                rsteNew = new RosterEntry(); // just to get nulls filled in to the member variables
            else
                rsteNew.Marked = true;

            if (fMarkOnly)
                return;

            if (irstServer != null)
                {
                rsteServer = (RosterEntry) irstServer.IrsteLookupEmail(sEmail);
                if (rsteServer == null)
                    {
                    m_appContext.StatusReport.AddMessage($"NULL Server entry for {sEmail}, SKIPPING", MSGT.Error);
                    return;
                    }
                if (rsteNew.FEquals(rsteServer))
                    return;
                }

            SyncRsteWithServer(sOfficialID, rste, rsteNew);

        }

        /*----------------------------------------------------------------------------
			%%Function:FConfirmExistingArbiterUserAdd
			%%Qualified:ArbWeb.WebRoster.FConfirmExistingArbiterUserAdd
        
			The user is already in the system (another association)
        ----------------------------------------------------------------------------*/
        bool FConfirmExistingArbiterUserAdd(RosterEntry rsteNewUser)
        {
            m_appContext.StatusReport.AddMessage($"Email {rsteNewUser.Email} already in use", MSGT.Warning);

            // this email is member of another group.  we can't change their personal info
            // do a quick sanity match to make sure this is the same user
            string sHtml = m_appContext.WebControl.Driver.FindElement(By.XPath("//body")).GetAttribute("innerHTML");
            
            HtmlDocument html = new HtmlDocument();
            html.LoadHtml(sHtml);

            string sText = html.DocumentNode.InnerText;
            string sPrefix = "is already being used in the system by ";
            int iFirst = sText.IndexOf(sPrefix);

            Utils.ThrowIfNot(iFirst > 0, "Failed hierarchy on assumed 'in use' email name");
            iFirst += sPrefix.Length;

            int iLast = sText.IndexOf(".  Click", iFirst);
            Utils.ThrowIfNot(iLast > iFirst, "couldn't find the end of the users name on 'in use' email page");

            string sName = sText.Substring(iFirst, iLast - iFirst);
            if (String.Compare(sName, rsteNewUser.Name, true /*ignoreCase*/) != 0)
                {
                if (MessageBox.Show(
	                $"Trying to add office {rsteNewUser.Name} and found a mismatch with existing official {sName}, with email {rsteNewUser.Email}",
                        "ArbWeb", MessageBoxButtons.YesNo) != DialogResult.Yes)
                    {
                    // ok, then just cancel...
                    Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(WebCore._sid_AddUser_Button_Cancel), "Can't click cancel button on adduser");
                    return false;
                    }
                }

            // cool, just go on...
            Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(WebCore._sid_AddUser_Button_Next), "Can't click next button on adduser");

            // sigh, now we're being asked whether we want to add local personal info.  of course
            // we don't since it will be thrown away when they choose to join our group!

            // but make sure that we're really on that page...
            sHtml = m_appContext.WebControl.Driver.FindElement(By.XPath("//body")).GetAttribute("innerHTML");
            html.LoadHtml(sHtml);
            sText = html.DocumentNode.InnerText;
            
            Utils.ThrowIfNot(sText.IndexOf("as a fully integrated user") > 0, "Didn't find the confirmation text on 'personal info' portion of existing user sequence");

            // cool, let's just move on again...
            Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(WebCore._sid_AddUser_Button_Next), "Can't click next button on adduser");

            // now fallthrough to the "Official's info" page handling, which is common
            return true;
        }

        /*----------------------------------------------------------------------------
			%%Function:AddOfficials
			%%Qualified:ArbWeb.WebRoster.AddOfficials
        ----------------------------------------------------------------------------*/
        private void AddOfficials(List<IRosterEntry> plirsteNew)
        {
            foreach (IRosterEntry irste in plirsteNew)
                {
                RosterEntry rste = (RosterEntry)irste;
                // add the official rste
                m_appContext.StatusReport.AddMessage($"Adding official '{rste.Name}', {rste.Email}", MSGT.Body);
                m_appContext.StatusReport.PushLevel();

                // go to the add user page
                if (!m_appContext.WebControl.FNavToPage(WebCore._s_AddUser))
                    {
                    throw (new Exception("could not navigate to the add user page"));
                    }

                // Set the basic user info + email address
                Utils.ThrowIfNot(m_appContext.WebControl.FSetTextForInputControlName(WebCore._s_AddUser_Input_FirstName, rste.First, false /*fCheck*/), "Failed to find first name control");
                Utils.ThrowIfNot(m_appContext.WebControl.FSetTextForInputControlName(WebCore._s_AddUser_Input_LastName, rste.Last, false /*fCheck*/), "Failed to find last name control");
                Utils.ThrowIfNot(m_appContext.WebControl.FSetTextForInputControlName(WebCore._s_AddUser_Input_Email, rste.Email, false /*fCheck*/), "Failed to find email control");

                Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(WebCore._sid_AddUser_Button_Next), "Can't click next button on adduser");

                // we are either adding a new user, or a user that arbiter already knows
                // about...
                // 
                if (!m_appContext.WebControl.FCheckForControlId(WebCore._sid_AddUser_Input_Address1))
                    {
                    if (!FConfirmExistingArbiterUserAdd(rste))
                        continue; // don't add this user, they cancelled
                    }
                else
                    {
                    // once we set the country, we will be able to set the zip code. note that we cleverly 
                    // set the other info after the country, so we will commit the change to the country.
                    Utils.ThrowIfNot(m_appContext.WebControl.FSetSelectedOptionTextForControlId(WebCore._sid_AddUser_Input_Country, "United States"), "Failed to set country control");

                    m_appContext.WebControl.WaitForXpath($"//option[contains(text(), '{rste.State}')]", 1000);
                    
                    // if there's an address control, then this is a brand new official
                    Utils.ThrowIfNot(m_appContext.WebControl.FSetTextForInputControlName(WebCore._s_AddUser_Input_City, rste.City, false /*fCheck*/), "Failed to find city control");
                    Utils.ThrowIfNot(m_appContext.WebControl.FSetSelectedOptionTextForControlId(WebCore._sid_AddUser_Input_State, rste.State), "Failed to find state control");

                    Utils.ThrowIfNot(m_appContext.WebControl.FSetTextForInputControlName(WebCore._s_AddUser_Input_Zip, rste.Zip, false /*fCheck*/), "Failed to find zip control");

                    Utils.ThrowIfNot(m_appContext.WebControl.FSetTextForInputControlName(WebCore._s_AddUser_Input_Address1, rste.Address1, false /*fCheck*/), "Failed to find address1 control");
                    Utils.ThrowIfNot(m_appContext.WebControl.FSetTextForInputControlName(WebCore._s_AddUser_Input_Address1, rste.Address2, false /*fCheck*/), "Failed to find address2 control");

                    string[] rgsPhoneNums = new string[] {WebCore._s_AddUser_Input_PhoneNum1, WebCore._s_AddUser_Input_PhoneNum2, WebCore._s_AddUser_Input_PhoneNum3};
                    string[] rgsPhoneTypes = new string[] {WebCore._s_AddUser_Input_PhoneType1, WebCore._s_AddUser_Input_PhoneType2, WebCore._s_AddUser_Input_PhoneType3};

                    int iPhone = 0;
                    while (iPhone < 3)
                        {
                        string sPhoneNum, sPhoneType;

                        rste.GetPhoneNumber(iPhone + 1 /*convert to 1 based*/, out sPhoneNum, out sPhoneType);
                        if (sPhoneNum != null)
                            {
                            Utils.ThrowIfNot(m_appContext.WebControl.FSetTextForInputControlName(rgsPhoneNums[iPhone], sPhoneNum, false /*fCheck*/), "Failed to find phonenum* control");

                            string sNewTypeOptionValue = m_appContext.WebControl.GetOptionValueForSelectControlNameOptionText(rgsPhoneTypes[iPhone], sPhoneType);
                            m_appContext.WebControl.FSetSelectedOptionValueForControlName(rgsPhoneTypes[iPhone], sNewTypeOptionValue);
                            }

                        iPhone++;
                        }

                    Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(WebCore._sid_AddUser_Button_Next), "Can't click next button on adduser");

                    // fallthrough to the common handling below
                    }

                // now we are on the last add official page
                // the only thing that *might* be interesting on this page is the active button which is
                // not checked by default...
                Utils.ThrowIfNot(m_appContext.WebControl.FCheckForControlId(WebCore._sid_AddUser_Input_IsActive),
                           "bad hierarchy in add user.  expected screen with 'active' checkbox, didn't find it.");

                // don't worry about Active for now...Just click next again
                Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(WebCore._sid_AddUser_Button_Next), "Can't click next button on adduser");

                // and now we're on the finish page.  oddly enough, the finish button has the "Cancel" ID
                Utils.ThrowIfNot(String.Compare("Finish", m_appContext.WebControl.GetValueForControlId(WebCore._sid_AddUser_Button_Cancel)) == 0, "Finish screen didn't have a finish button");

                Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(WebCore._sid_AddUser_Button_Cancel), "Can't click finish/cancel button on adduser");
                m_appContext.StatusReport.PopLevel();
                // and now we're back somewhere (probably officials edit page)
                // continue to the next one...
                }

            // and that's it...simple n'est pas?
        }

        /*----------------------------------------------------------------------------
			%%Function:VisitRankCallbackUpload
			%%Qualified:ArbWeb.WebRoster.VisitRankCallbackUpload
        ----------------------------------------------------------------------------*/
        private static void VisitRankCallbackUpload(IRoster irst, string sRankPosition, Dictionary<string, int> mpNameRank, Dictionary<string, string> mpNameOptionValue, WebControl webControl, IStatusReporter srpt)
        {
	        BuildRankingJobs(
		        irst,
		        sRankPosition,
		        mpNameRank,
		        out List<string> plsUnrankNames,
		        out Dictionary<int, List<string>> mpRankNames,
		        out Dictionary<int, List<string>> mpRankNamesRerank);

            // at this point, we have a list of jobs to do.

            // first, unrank everyone that needs unranked
            if (plsUnrankNames.Count > 0)
            {
                webControl.FResetMultiSelectOptionsForControlName(WebCore._s_RanksEdit_Select_Ranked);
                foreach (string s in plsUnrankNames)
                {
                    if (!webControl.FSelectMultiSelectOptionValueForControlName(WebCore._s_RanksEdit_Select_Ranked, mpNameOptionValue[s]))
                        throw new Exception("couldn't select an official for unranking!");
                }

                // now, do the unrank
                webControl.FClickControlName(WebCore._s_RanksEdit_Button_Unrank);
            }

            // now, let's rerank the folks that need to be re-ranked
            // we will do this once for every new rank we are setting
            foreach (int nRank in mpRankNamesRerank.Keys)
            {
                webControl.FResetMultiSelectOptionsForControlName(WebCore._s_RanksEdit_Select_Ranked);
                foreach (string s in mpRankNamesRerank[nRank])
                {
                    if (!webControl.FSelectMultiSelectOptionValueForControlName(WebCore._s_RanksEdit_Select_Ranked, mpNameOptionValue[s]))
                        throw new Exception("couldn't select an official for reranking!");
                }
                webControl.FSetTextForInputControlName(WebCore._s_RanksEdit_Input_Rank, nRank.ToString(), false);

                // now, rank'em
                webControl.FClickControlName(WebCore._s_RanksEdit_Button_ReRank);
            }

            // finally, let's rank the folks that weren't ranked before

            Dictionary<string, string> mpNameOptionValueUnranked =
	            webControl.GetOptionsTextValueMappingFromControlId(WebCore._sid_RanksEdit_Select_NotRanked);
            
            foreach (int nRank in mpRankNames.Keys)
            {
                webControl.FResetMultiSelectOptionsForControlName(WebCore._s_RanksEdit_Select_NotRanked);
                foreach (string s in mpRankNames[nRank])
                {
                    if (!webControl.FSelectMultiSelectOptionValueForControlName(WebCore._s_RanksEdit_Select_NotRanked, mpNameOptionValueUnranked[s]))
                        srpt.AddMessage(
	                        $"Could not select an official for ranking: {s}",
                                        MSGT.Error);
                    // throw new Exception("couldn't select an official for ranking!");
                }

                webControl.FSetTextForInputControlName(WebCore._s_RanksEdit_Input_Rank, nRank.ToString(), false);

                // now, rank'em
                webControl.FClickControlName(WebCore._s_RanksEdit_Button_Rank);
            }
        }

        /*----------------------------------------------------------------------------
			%%Function:InvokeHandleRoster
			%%Qualified:ArbWeb.WebRoster.InvokeHandleRoster
        ----------------------------------------------------------------------------*/
        void InvokeHandleRoster(Roster rstUpload, string sInFile, Roster rstServer, bool fRankOnly, bool fAddOfficialsOnly, HandleGenericRoster.HandleRosterPostUpdateDelegate hrpu)
        {
            HandleGenericRoster gr = new HandleGenericRoster(
                m_appContext,
                !fRankOnly, // !m_cbRankOnly.Checked, // fNeedPass1OnUpload
                fAddOfficialsOnly, // m_cbAddOfficialsOnly.Checked, // only add officials
                HandleRosterPass1VisitForUploadDownload,
                AddOfficials,
                HandleRankings
            );

            gr.GenericVisitRoster(rstUpload, null, sInFile, rstServer, hrpu);
        }

        /*----------------------------------------------------------------------------
			%%Function:DoUploadRosterWork
			%%Qualified:ArbWeb.WebRoster.DoUploadRosterWork
        ----------------------------------------------------------------------------*/
        private async void DoUploadRosterWork(string sRosterToUpload, bool fRankOnly, bool fAddOfficialsOnly)
        {
            m_appContext.StatusReport.AddMessage("Starting Roster upload...");
            m_appContext.StatusReport.PushLevel();
            Roster rstServer = null;

            if (m_appContext.Profile.DownloadRosterOnUpload)
                {
                // first thing to do is download a new (temporary) roster copy
                Task<Roster> tsk = new Task<Roster>(
	                () => DoDownloadQuickRosterOfficialsOnlyWork(fRankOnly, fAddOfficialsOnly));

                tsk.Start();
                rstServer = await tsk;
            }

            // now, check the roster we just downloaded against the roster we just already have

            m_appContext.InvalRoster();
            string sInFile = sRosterToUpload;

            Roster rst = m_appContext.RstEnsure(sInFile);

            if (rst.IsQuick && (!fRankOnly || !rst.HasRankings))
            {
                //				MessageBox.Show("Cannot upload a quick roster.  Please perform a full roster download before uploading.\n\nIf you want to upload rankings only, please check 'Upload Rankings Only'");
                //    			m_srpt.PopLevel();
                m_appContext.StatusReport.AddMessage("Detected QuickRoster...", MSGT.Warning);
            }

            // compare the two rosters to find differences

			InvokeHandleRoster(rst, null, rstServer, fRankOnly, fAddOfficialsOnly, null);

            m_appContext.StatusReport.PopLevel();
            m_appContext.StatusReport.AddMessage("Completed Roster upload.");
        }

        /*----------------------------------------------------------------------------
			%%Function:SGetRosterToUpload
			%%Qualified:ArbWeb.WebRoster.SGetRosterToUpload
        ----------------------------------------------------------------------------*/
        string SGetRosterToUpload()
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.InitialDirectory = Path.GetDirectoryName(m_appContext.Profile.RosterWorking);
            ofd.Filter = "CSV Files|*.csv";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                return ofd.FileName;
            }
            return null;
        }

        /*----------------------------------------------------------------------------
			%%Function:DoRosterUpload
			%%Qualified:ArbWeb.WebRoster.DoRosterUpload
        ----------------------------------------------------------------------------*/
        public void DoRosterUpload(bool fRankOnly, bool fAddOfficialsOnly)
        {
            string sRosterToUpload = SGetRosterToUpload();

            if (sRosterToUpload == null)
                return;


            Task tsk = new Task(() => DoUploadRosterWork(sRosterToUpload, fRankOnly, fAddOfficialsOnly));

            tsk.Start();
        }
    }
}
