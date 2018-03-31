using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using StatusBox;
using mshtml;
using System.Text.RegularExpressions;
using NUnit.Framework;
using Win32Win;

namespace ArbWeb
{
    public partial class AwMainForm : System.Windows.Forms.Form
    {
        /*----------------------------------------------------------------------------
        	%%Function: SetServerMiscFields
        	%%Qualified: ArbWeb.AwMainForm.SetServerMiscFields
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void SetServerMiscFields(string sEmail, string sOfficialID, IRoster irst, IRoster irstServer, RosterEntry rste)
        {
            RosterEntry rsteNew = (RosterEntry)irst.IrsteLookupEmail(rste.Email);
            RosterEntry rsteServer = (RosterEntry)irstServer?.IrsteLookupEmail(rste.Email);

            if (rsteNew.FEqualsMisc(rsteServer))
                return;

            List<string> plsMiscServer = irstServer.PlsMisc;

            List<string> plsValue = SyncPlsMiscWithServer(m_awc.Document2, sEmail, sOfficialID, rsteNew.Misc, irst.PlsMisc, ref plsMiscServer);

            irstServer.PlsMisc = plsMiscServer;

            if (plsValue.Count == 0)
                throw new Exception("couldn't extract misc field for official");

            rste.m_plsMisc = plsValue;
        }

        /* S E T  S E R V E R  R O S T E R  I N F O */
        /*----------------------------------------------------------------------------
        	%%Function: SetServerRosterInfo
        	%%Qualified: ArbWeb.AwMainForm.SetServerRosterInfo
        	%%Contact: rlittle
        	
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
                    m_srpt.AddMessage(String.Format("NULL Server entry for {0}, SKIPPING", sEmail), StatusBox.StatusRpt.MSGT.Error);
                    return;
                    }
                if (rsteNew.FEquals(rsteServer))
                    return;
                }

            SyncRsteWithServer(m_awc.Document2, sOfficialID, rste, rsteNew);

        }

        bool FConfirmExistingArbiterUserAdd(IHTMLDocument2 oDoc2, RosterEntry rsteNewUser)
        {
            m_srpt.AddMessage(String.Format("Email {0} already in use", rsteNewUser.Email), StatusBox.StatusRpt.MSGT.Warning);

            // this email is member of another group.  we can't change their personal info
            // do a quick sanity match to make sure this is the same user
            string sText = oDoc2.body.innerText;
            string sPrefix = "is already being used in the system by ";
            int iFirst = sText.IndexOf(sPrefix);

            ThrowIfNot(iFirst > 0, "Failed hierarchy on assumed 'in use' email name");
            iFirst += sPrefix.Length;

            int iLast = sText.IndexOf(". Click", iFirst);
            ThrowIfNot(iLast > iFirst, "couldn't find the end of the users name on 'in use' email page");

            string sName = sText.Substring(iFirst, iLast - iFirst);
            if (String.Compare(sName, rsteNewUser.Name, true /*ignoreCase*/) != 0)
                {
                if (MessageBox.Show(
                        String.Format("Trying to add office {0} and found a mismatch with existing official {1}, with email {2}", rsteNewUser.Name, sName, rsteNewUser.Email),
                        "ArbWeb", MessageBoxButtons.YesNo) != DialogResult.Yes)
                    {
                    // ok, then just cancel...
                    m_awc.ResetNav();
                    ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_AddUser_Button_Cancel), "Can't click cancel button on adduser");
                    m_awc.FWaitForNavFinish();
                    return false;
                    }
                }

            // cool, just go on...
            m_awc.ResetNav();
            ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_AddUser_Button_Next), "Can't click next button on adduser");
            m_awc.FWaitForNavFinish();

            // sigh, now we're being asked whether we want to add local personal info.  of course
            // we don't since it will be thrown away when they choose to join our group!

            // but make sure that we're really on that page...
            sText = oDoc2.body.innerText;
            ThrowIfNot(sText.IndexOf("as a fully integrated user") > 0, "Didn't find the confirmation text on 'personal info' portion of existing user sequence");

            // cool, let's just move on again...
            m_awc.ResetNav();
            ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_AddUser_Button_Next), "Can't click next button on adduser");
            m_awc.FWaitForNavFinish();

            // now fallthrough to the "Official's info" page handling, which is common
            return true;
        }
        /*----------------------------------------------------------------------------
        	%%Function: AddOfficials
        	%%Qualified: ArbWeb.AwMainForm.AddOfficials
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void AddOfficials(List<IRosterEntry> plirsteNew)
        {
            foreach (IRosterEntry irste in plirsteNew)
                {
                RosterEntry rste = (RosterEntry)irste;
                // add the official rste
                m_srpt.AddMessage(String.Format("Adding official '{0}', {1}", rste.Name, rste.Email), StatusBox.StatusRpt.MSGT.Body);
                m_srpt.PushLevel();

                // go to the add user page
                m_awc.ResetNav();
                if (!m_awc.FNavToPage(WebCore._s_AddUser))
                    {
                    throw (new Exception("could not navigate to the add user page"));
                    }

                m_awc.FWaitForNavFinish();

                IHTMLDocument2 oDoc2 = m_awc.Document2;

                // Set the basic user info + email address
                ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_AddUser_Input_FirstName, rste.First, false /*fCheck*/), "Failed to find first name control");
                ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_AddUser_Input_LastName, rste.Last, false /*fCheck*/), "Failed to find last name control");
                ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_AddUser_Input_Email, rste.Email, false /*fCheck*/), "Failed to find email control");

                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_AddUser_Button_Next), "Can't click next button on adduser");
                m_awc.FWaitForNavFinish();

                // we are either adding a new user, or a user that arbiter already knows
                // about...
                // 
                if (!ArbWebControl.FCheckForControl(oDoc2, WebCore._sid_AddUser_Input_Address1))
                    {
                    if (!FConfirmExistingArbiterUserAdd(oDoc2, rste))
                        continue; // don't add this user, they cancelled
                    }
                else
                    {
                    // if there's an address control, then this is a brand new official
                    ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_AddUser_Input_Address1, rste.Address1, false /*fCheck*/), "Failed to find address1 control");
                    ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_AddUser_Input_Address1, rste.Address2, false /*fCheck*/), "Failed to find address2 control");
                    ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_AddUser_Input_City, rste.City, false /*fCheck*/), "Failed to find city control");
                    ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_AddUser_Input_State, rste.State, false /*fCheck*/), "Failed to find state control");

                    // DebugModelessWait();
                    // once we set the country, we will be able to set the zip code
                    ThrowIfNot(ArbWebControl.FSetSelectControlTextFromDoc(m_awc, oDoc2, WebCore._s_AddUser_Input_Country, WebCore._sid_AddUser_Input_Country, "United States", true), "Failed to set country control");
                    // DebugModelessWait();

                    ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_AddUser_Input_Zip, rste.Zip, false /*fCheck*/), "Failed to find zip control");
                    ArbWebControl.DispatchChangeEventCore(m_awc, WebCore._sid_AddUser_Input_Zip, "keyup");


                    string[] rgsPhoneNums = new string[] {WebCore._s_AddUser_Input_PhoneNum1, WebCore._s_AddUser_Input_PhoneNum2, WebCore._s_AddUser_Input_PhoneNum3};
                    string[] rgsPhoneTypes = new string[] {WebCore._s_AddUser_Input_PhoneType1, WebCore._s_AddUser_Input_PhoneType2, WebCore._s_AddUser_Input_PhoneType3};

                    int iPhone = 0;
                    while (iPhone < 3)
                        {
                        string sPhoneNum, sPhoneType;

                        rste.GetPhoneNumber(iPhone + 1 /*convert to 1 based*/, out sPhoneNum, out sPhoneType);
                        if (sPhoneNum != null)
                            {
                            ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, rgsPhoneNums[iPhone], sPhoneNum, false /*fCheck*/), "Failed to find phonenum* control");
                            string sNewTypeID = ArbWebControl.SGetSelectIDFromDoc(oDoc2, rgsPhoneTypes[iPhone], sPhoneType);
                            ArbWebControl.FSetSelectControlValue(oDoc2, rgsPhoneTypes[iPhone], sNewTypeID, false);
                            }

                        iPhone++;
                        }

                    m_awc.ResetNav();
                    // DebugModelessWait();

                    ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_AddUser_Button_Next), "Can't click next button on adduser");
                    // DebugModelessWait();
                    m_awc.FWaitForNavFinish();

                    // fallthrough to the common handling below
                    }

                // now we are on the last add official page
                // the only thing that *might* be interesting on this page is the active button which is
                // not checked by default...
                ThrowIfNot(ArbWebControl.FCheckForControl(oDoc2, WebCore._sid_AddUser_Input_IsActive),
                           "bad hierarchy in add user.  expected screen with 'active' checkbox, didn't find it.");

                // don't worry about Active for now...Just click next again
                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_AddUser_Button_Next), "Can't click next button on adduser");
                m_awc.FWaitForNavFinish();

                // and now we're on the finish page.  oddly enough, the finish button has the "Cancel" ID
                ThrowIfNot(String.Compare("Finish", ArbWebControl.SGetControlValue(oDoc2, WebCore._sid_AddUser_Button_Cancel)) == 0, "Finish screen didn't have a finish button");

                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_AddUser_Button_Cancel), "Can't click finish/cancel button on adduser");
                m_awc.FWaitForNavFinish();
                m_srpt.PopLevel();
                // and now we're back somewhere (probably officials edit page)
                // continue to the next one...
                }

            // and that's it...simple n'est pas?
        }

        /*----------------------------------------------------------------------------
        	%%Function: VisitRankCallbackUpload
        	%%Qualified: ArbWeb.AwMainForm.VisitRankCallbackUpload
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private static void VisitRankCallbackUpload(IRoster irst, string sRankPosition, Dictionary<string, int> mpRanked, Dictionary<string, string> mpRankedId, ArbWebControl awc, StatusBox.StatusRpt srpt)
        {
            IHTMLDocument2 oDoc2;
            oDoc2 = awc.Document2;

            List<string> plsUnrank;
            Dictionary<int, List<string>> mpRank;
            Dictionary<int, List<string>> mpRerank;

            BuildRankingJobs(irst, sRankPosition, mpRanked, out plsUnrank, out mpRank, out mpRerank);

            // at this point, we have a list of jobs to do.

            // first, unrank everyone that needs unranked
            if (plsUnrank.Count > 0)
            {
                ArbWebControl.FResetMultiSelectOptions(oDoc2, WebCore._s_RanksEdit_Select_Ranked);
                foreach (string s in plsUnrank)
                {
                    if (!ArbWebControl.FSelectMultiSelectOption(oDoc2, WebCore._s_RanksEdit_Select_Ranked, mpRankedId[s], true))
                        throw new Exception("couldn't select an official for unranking!");
                }

                // now, do the unrank
                awc.ResetNav();
                awc.FClickControl(oDoc2, WebCore._s_RanksEdit_Button_Unrank);
                awc.FWaitForNavFinish();
                oDoc2 = awc.Document2;
            }

            // now, let's rerank the folks that need to be re-ranked
            // we will do this once for every new rank we are setting
            foreach (int nRank in mpRerank.Keys)
            {
                ArbWebControl.FResetMultiSelectOptions(oDoc2, WebCore._s_RanksEdit_Select_Ranked);
                foreach (string s in mpRerank[nRank])
                {
                    if (!ArbWebControl.FSelectMultiSelectOption(oDoc2, WebCore._s_RanksEdit_Select_Ranked, mpRankedId[s], true))
                        throw new Exception("couldn't select an official for reranking!");
                }
                ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_RanksEdit_Input_Rank, nRank.ToString(), false);

                // now, rank'em
                awc.ResetNav();
                awc.FClickControl(oDoc2, WebCore._s_RanksEdit_Button_ReRank);
                awc.FWaitForNavFinish();
                oDoc2 = awc.Document2;
            }

            // finally, let's rank the folks that weren't ranked before

            foreach (int nRank in mpRank.Keys)
            {
                ArbWebControl.FResetMultiSelectOptions(oDoc2, WebCore._s_RanksEdit_Select_NotRanked);
                foreach (string s in mpRank[nRank])
                {
                    if (!ArbWebControl.FSelectMultiSelectOption(oDoc2, WebCore._s_RanksEdit_Select_NotRanked, s, false))
                        srpt.AddMessage(String.Format("Could not select an official for ranking: {0}", s),
                                        StatusRpt.MSGT.Error);
                    // throw new Exception("couldn't select an official for ranking!");
                }

                ArbWebControl.FSetInputControlText(oDoc2, WebCore._s_RanksEdit_Input_Rank, nRank.ToString(), false);

                // now, rank'em
                awc.ResetNav();
                awc.FClickControl(oDoc2, WebCore._s_RanksEdit_Button_Rank);
                awc.FWaitForNavFinish();
                oDoc2 = awc.Document2;
            }
        }

        void InvokeHandleRoster(Roster rstUpload, string sInFile, Roster rstServer, HandleGenericRoster.HandleRosterPostUpdateDelegate hrpu)
        {
            HandleGenericRoster gr = new HandleGenericRoster(
                this,
                !m_cbRankOnly.Checked, // fNeedPass1OnUpload
                m_cbAddOfficialsOnly.Checked, // only add officials
                new HandleGenericRoster.delDoPass1Visit(HandleRosterPass1VisitForDownload),
                new HandleGenericRoster.delAddOfficials(AddOfficials),
                new HandleGenericRoster.delDoPostHandleRoster(HandleRankings)
            );

            gr.GenericVisitRoster(rstUpload, null, sInFile, rstServer, hrpu);
        }

        /*----------------------------------------------------------------------------
        	%%Function: DoUploadRosterWork
        	%%Qualified: ArbWeb.AwMainForm.DoUploadRosterWork
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private async void DoUploadRosterWork(string sRosterToUpload)
        {
            m_srpt.AddMessage("Starting Roster upload...");
            m_srpt.PushLevel();
            Roster rstServer = null;

            if (m_pr.DownloadRosterOnUpload)
                {
                // first thing to do is download a new (temporary) roster copy
                Task<Roster> tsk = new Task<Roster>(DoDownloadQuickRosterOfficialsOnlyWork);

                tsk.Start();
                rstServer = await tsk;
            }

            // now, check the roster we just downloaded against the roster we just already have

            InvalRoster();
            string sInFile = sRosterToUpload;

            Roster rst = RstEnsure(sInFile);

            if (rst.IsQuick && (!m_cbRankOnly.Checked || !rst.HasRankings))
            {
                //				MessageBox.Show("Cannot upload a quick roster.  Please perform a full roster download before uploading.\n\nIf you want to upload rankings only, please check 'Upload Rankings Only'");
                //    			m_srpt.PopLevel();
                m_srpt.AddMessage("Detected QuickRoster...", StatusBox.StatusRpt.MSGT.Warning);
            }

            // compare the two rosters to find differences

            if (m_awc.InvokeRequired)
                {
                IAsyncResult rslt = m_awc.BeginInvoke(new AwMainForm.HandleRosterDel(InvokeHandleRoster), new object[] {rst, sInFile, rstServer, null});
                m_awc.EndInvoke(rslt);
                }
            else
                {
                InvokeHandleRoster(rst, null, rstServer, null);
                }

            m_srpt.PopLevel();
            m_srpt.AddMessage("Completed Roster upload.");
        }

        string SGetRosterToUpload()
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.InitialDirectory = Path.GetDirectoryName(m_pr.RosterWorking);
            ofd.Filter = "CSV Files|*.csv";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                return ofd.FileName;
            }
            return null;
        }

        bool FLoadRosterToUpload()
        {
            return false;
        }

        /*----------------------------------------------------------------------------
        	%%Function: DoRosterUpload
        	%%Qualified: ArbWeb.AwMainForm.DoRosterUpload
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void DoRosterUpload()
        {
            string sRosterToUpload = SGetRosterToUpload();

            if (sRosterToUpload == null)
                return;


            Task tsk = new Task(() => DoUploadRosterWork(sRosterToUpload));

            tsk.Start();
        }
    }
}
