using System;
using System.Collections.Generic;
using TCore.StatusBox;
using HtmlAgilityPack;
using NUnit.Framework;
using OpenQA.Selenium;
using TCore.WebControl;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace ArbWeb
{
    /// <summary>
    /// Summary description for AwMainForm.
    /// </summary>
    public partial class AwMainForm : System.Windows.Forms.Form
    {
        /*----------------------------------------------------------------------------
        	%%Function: FetchMiscFieldsFromServer
        	%%Qualified: ArbWeb.AwMainForm.FetchMiscFieldsFromServer
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        void FetchMiscFieldsFromServer(string sEmail, string sOfficialID, RosterEntry rste, IRoster irstBuilding)
        {
            Roster rstBuilding = (Roster) irstBuilding;
            List<string> plsMiscBuilding = rstBuilding.PlsMisc;

            rste.m_plsMisc = SyncPlsMiscWithServer(sEmail, sOfficialID, null, null, ref plsMiscBuilding);
            rstBuilding.PlsMisc = plsMiscBuilding;

            if (rste.m_plsMisc.Count == 0)
                throw new Exception("couldn't extract misc field for official");
        }

        /* U P D A T E  M I S C */
        /*----------------------------------------------------------------------------
			%%Function: UpdateMisc
			%%Qualified: ArbWeb.AwMainForm.UpdateMisc
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
        private void UpdateMisc(string sEmail, string sOfficialID, IRoster irst, IRoster irstServer, RosterEntry rste, IRoster irstBuilding)
        {
            if (irst == null)
                FetchMiscFieldsFromServer(sEmail, sOfficialID, rste, irstBuilding);
            else
                SetServerMiscFields(sEmail, sOfficialID, irst, irstServer, rste);
        }

        /*----------------------------------------------------------------------------
        	%%Function: MiscLabelFromControl
        	%%Qualified: ArbWeb.AwMainForm.MiscLabelFromControl
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private static string MiscLabelFromControl(HtmlNode input)
        {
	        HtmlNode parent = input;
            do
            {
	            parent = parent.ParentNode;
                } while (parent != null && String.Compare(parent.Name, "TR", StringComparison.InvariantCultureIgnoreCase) != 0);

            if (parent == null)
                return null;

            // now, find the first TD child
            foreach (HtmlNode child in parent.ChildNodes)
                {
                if (String.Compare(child.Name, "TD", StringComparison.InvariantCultureIgnoreCase) == 0)
                    return child.InnerText.Trim();
                }

            return null;

        }

        static int IMiscFromMiscName(List<string> plsMiscMap, string sMiscName)
        {
            for (int i = 0; i < plsMiscMap.Count; i++)
                if (String.Compare(plsMiscMap[i], sMiscName, StringComparison.InvariantCultureIgnoreCase) == 0)
                    return i;

            return -1;
        }
        /* S Y N C  P L S  M I S C  W I T H  S E R V E R */
        /*----------------------------------------------------------------------------
        	%%Function: SyncPlsMiscWithServer
        	%%Qualified: ArbWeb.AwMainForm.SyncPlsMiscWithServer
        	%%Contact: rlittle
        	
            navigate to the custom fields page and return the values. if plsMiscNew
            is supplied, then make sure the server matches that, and return fNeedSave
            to let caller know that the page needs to be saved

            This will make sure to use plsMiscMapNew to get the right misc value into
            the right server control.  It will also build up plsMiscMapServer so
            we have a correct order map for the pls that we return to the caller.
        ----------------------------------------------------------------------------*/
        private List<string> SyncPlsMiscWithServer(
	        string sEmail, 
	        string sOfficialID, 
	        List<string> plsMiscNew, 
	        List<string> plsMiscMapNew,
            ref List<string> plsMiscMapServer)
        {
            bool fNeedSave = false;
            string sValue;

            if (!m_webControl.FNavToPage(WebCore._s_EditUser_MiscFields + sOfficialID))
                {
                throw (new Exception("could not navigate to the officials page"));
                }

            // misc field info.  every text input field is a misc field we want to save
            List<string> plsValue = new List<string>();
            string sHtml = m_webControl.Driver.FindElement(By.Id(WebCore._sid_MiscFields_MainBodyContentTable)).GetAttribute("outerHTML");
            HtmlDocument html = new HtmlDocument();
            
            html.LoadHtml(sHtml);
            sValue = null;

            HtmlNodeCollection inputs = html.DocumentNode.SelectNodes(
	            $"//input[@type='text' and contains(@name, '{WebCore.s_MiscField_EditControlSubstring}')]");
//            IList <IWebElement> inputs =
//	            m_webControl.Driver.FindElements(By.XPath($"//input[@type='text' and contains(@name, '{WebCore.s_MiscField_EditControlSubstring}')]";

            
            foreach (HtmlNode input in inputs)
                {
                    // figure out which misc field this is
                    string sMiscLabel = MiscLabelFromControl(input);

                    // cool, extract the value
                    sValue = input.GetAttributeValue("value", null);
                    string sName = input.GetAttributeValue("name", null);
                    
                    if (sValue == null)
                        sValue = "";

                    if (plsMiscNew != null)
                    {
	                    int iMisc = IMiscFromMiscName(plsMiscMapNew, sMiscLabel);

	                    if (iMisc == -1)
		                    throw new Exception(
			                    "couldn't find misc field name! (OR maybe this is a new misc field that the roster doesn't know about, in which case we should just set it to empty, but this isn't debugged yet so we don't trust that decision yet, hence the exception");

	                    if (iMisc == -1 // couldn't find this server misc field in the roster's list of misc fields...set to empty
	                        && !String.IsNullOrEmpty(sValue))
	                    {
		                    // null means empty which replaces non-empty
		                    m_webControl.FSetTextForInputControlName(sName, "", false);
		                    fNeedSave = true;
	                    }
	                    else if (iMisc != -1
	                             && String.Compare(plsMiscNew[iMisc], sValue, true /*ignoreCase*/) != 0)
	                    {
		                    m_webControl.FSetTextForInputControlName(sName, plsMiscNew[iMisc], false);
		                    fNeedSave = true;
	                    }
                    }

                    // we always keep the server plsMiscMap up to date
                    int iMiscServer = IMiscFromMiscName(plsMiscMapServer, sMiscLabel);

                    if (iMiscServer == -1)
                        {
                        plsValue.Add(sValue);
                        plsMiscMapServer.Add(sMiscLabel);
                        }
                    else
                        {
                        if (plsValue.Count <= iMiscServer)
                            {
                            while (plsValue.Count <= iMiscServer)
                                plsValue.Add("");
                            }

                        plsValue[iMiscServer] = sValue;
                        }

                    // don't break here -- just get the next misc value...
                }

            // before we return, commit the change or cancel (so we are no longer on the page)
            if (fNeedSave)
                {
                m_srpt.AddMessage(String.Format("Updating misc info...", sEmail));
                ThrowIfNot(m_webControl.FClickControlId(WebCore._sid_MiscFields_Button_Save), "Couldn't find save button");
                }
            else
                {
                ThrowIfNot(m_webControl.FClickControlId(WebCore._sid_MiscFields_Button_Cancel), "Couldn't find cancel button");
                }

            return plsValue;
        }

        bool FIsInputNodeDisabled(HtmlNode input)
        {
	        string sDisabled = input.GetAttributeValue("disabled", null);

	        if (sDisabled == null)
		        return false;

	        if (String.Compare(sDisabled, "true", true) == 0
	            || String.Compare(sDisabled, "on", true) == 0
	            || sDisabled == "1")
	        {
		        return true;
	        }

	        return false;
        }
        
        /* M A T C H  A S S I G N  T E X T */
        /*----------------------------------------------------------------------------
        	%%Function: MatchAssignText
        	%%Qualified: ArbWeb.AwMainForm.MatchAssignText
        	%%Contact: rlittle
        	
            Get the value for the control sMatch. Determine the current value of
            that control (which will be returned in sAssign). 

            sNewValue is not null and does not match the controls value, then
            set the control to sNewValue (which will strangely leave sAssign
            set to the OLD value)

            return true if we found the control, false if we didn't
        ----------------------------------------------------------------------------*/
        private bool MatchAssignText(HtmlNode input, string sMatch, string sNewValue, out string sAssign, ref bool fNeedSave, ref bool fFailAssign)
        {
	        string name = input.GetAttributeValue("name", null);
	        string value = input.GetAttributeValue("value", null);
	        
            sAssign = null;
            if (name.Contains(sMatch))
                {
                sAssign = value;

                if (sAssign == null)
                    sAssign = "";

                if (sNewValue != null)
                    {
                    // check to see if it matches what we have
                    // find a match on email address first
                    if (sNewValue != null
                        && String.Compare(sNewValue, sAssign, true /*ignoreCase*/) != 0)
                        {
                        if (FIsInputNodeDisabled(input))
                            {
                            fFailAssign = true;
                            }
                        else
                        {
	                        m_webControl.FSetTextForInputControlName(name, sNewValue, false);
                            fNeedSave = true;
                            }
                        }
                    }

                return true;
                }

            return false;
        }


        /* U P D A T E  I N F O */
        /*----------------------------------------------------------------------------
			%%Function: UpdateInfo
			%%Qualified: ArbWeb.AwMainForm.UpdateInfo
			%%Contact: rlittle

			when we leave, if rst was null, then rste will have the values as we
			fetched from arbiter

            rstServer UNUSED right now
		----------------------------------------------------------------------------*/
        private void UpdateInfo(string sEmail, string sOfficialID, IRoster irst, IRoster irstServer, RosterEntry rste, bool fMarkOnly)
        {
            if (irst == null)
                GetRosterInfoFromServer(sEmail, sOfficialID, rste);
            else
                SetServerRosterInfo(sEmail, sOfficialID, irst, irstServer, rste, fMarkOnly);
        }

        /*----------------------------------------------------------------------------
        	%%Function: SetPhoneNames
        	%%Qualified: ArbWeb.AwMainForm.SetPhoneNames
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        static void SetPhoneNames(int iPhoneRow, out string sPhoneNum, out string sidPhoneNum, out string sPhoneType, out string sPhoneCarrier, out string sPhonePublicNext)
        {
            sPhoneNum = $"{WebCore._s_EditUser_PhoneNumber_Prefix}ctl{iPhoneRow:00}{WebCore._s_EditUser_PhoneNumber_Suffix}";
            sidPhoneNum = $"{WebCore._sid_EditUser_PhoneNumber_Prefix}ctl{iPhoneRow:00}{WebCore._sid_EditUser_PhoneNumber_Suffix}";
            sPhoneType = $"{WebCore._s_EditUser_PhoneType_Prefix}ctl{iPhoneRow:00}{WebCore._s_EditUser_PhoneType_Suffix}";
            sPhoneCarrier = $"{WebCore._s_EditUser_PhoneCarrier_Prefix}ctl{iPhoneRow:00}{WebCore._s_EditUser_PhoneCarrier_Suffix}";
            sPhonePublicNext = $"{WebCore._s_EditUser_PhonePublic_Prefix}ctl{iPhoneRow:00}{WebCore._s_EditUser_PhonePublic_Suffix}";
        }

        /* S Y N C  R S T E  W I T H  S E R V E R */
        /*----------------------------------------------------------------------------
        	%%Function: SyncRsteWithServer
        	%%Qualified: ArbWeb.AwMainForm.SyncRsteWithServer
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void SyncRsteWithServer(string sOfficialID, RosterEntry rsteOut, RosterEntry rsteNew)
        {
            bool fFailUpdate = false;
            bool fNeedSave = false;

            // ok, nav to the page and scrape
            if (!m_webControl.FNavToPage(WebCore._s_EditUser + sOfficialID))
                {
                throw (new Exception("could not navigate to the officials page"));
                }

            string sHtml = m_webControl.Driver.FindElement(By.XPath("//body")).GetAttribute("outerHTML");
            
            HtmlDocument body = new HtmlDocument();
            body.LoadHtml(sHtml);

            HtmlNodeCollection inputs = body.DocumentNode.SelectNodes("//input");

            string sNamePhoneNumberNext;
            string sidPhoneNumberNext;
            string sNamePhoneTypeNext;
            string sNamePhoneCarrierNext;
            string sNamePhonePublicNext;
            int iNextPhone = 1;
            SetPhoneNames(iNextPhone, out sNamePhoneNumberNext, out sidPhoneNumberNext, out sNamePhoneTypeNext, out sNamePhoneCarrierNext, out sNamePhonePublicNext);

            foreach (HtmlNode input in inputs)
            {
	            string type = input.GetAttributeValue("type", null);
	            string name = input.GetAttributeValue("name", null);
	            string value = input.GetAttributeValue("value", null);// checked for checkbox
	            
                if (String.Compare(type, "checkbox", true) == 0)
                {
	                string checkedAttr = input.GetAttributeValue("checked", null);
	                
                    // checkboxes are either ready or active
                    if (name.Contains("Active"))
                        rsteOut.m_fActive = String.Compare(checkedAttr, "checked", true) == 0;
                    else if (name.Contains("Ready"))
                        rsteOut.m_fReady = String.Compare(checkedAttr, "checked", true) == 0;
                    }

                if (String.Compare(type, "text", true) == 0 && name != null)
                    {
                    if (name.Contains(WebCore._s_EditUser_Email))
                        {
//						if (ihie.value == null && rsteOut.m_sEmail != null && rsteOut.m_sEmail != "")
                        // continue;

                        if (value != null && rsteOut.Email != null)
                            {
                            if (String.Compare(value, rsteOut.Email, true) != 0)
                                throw new Exception("email addresses don't match!");
                            }
                        else
                            {
                            m_srpt.AddMessage($"NULL Email address for {rsteOut.First},{rsteOut.Last}", MSGT.Error);
                            }
                        }

                    string s; // add middle name
                    if (MatchAssignText(input, WebCore._s_EditUser_FirstName, rsteNew?.First, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.First = s;
                    if (MatchAssignText(input, WebCore._s_EditUser_MiddleName, rsteNew?.Middle, out s, ref fNeedSave, ref fFailUpdate))
	                    rsteOut.Middle = s;
                    if (MatchAssignText(input, WebCore._s_EditUser_LastName, rsteNew?.Last, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.Last = s;
                    if (MatchAssignText(input, WebCore._s_EditUser_Address1, rsteNew?.Address1, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.Address1 = s;
                    if (MatchAssignText(input, WebCore._s_EditUser_Address2, rsteNew?.Address2, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.Address2 = s;
                    if (MatchAssignText(input, WebCore._s_EditUser_City, rsteNew?.City, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.City = s;
                    if (MatchAssignText(input, WebCore._s_EditUser_State, rsteNew?.State, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.State = s;
                    if (MatchAssignText(input, WebCore._s_EditUser_PostalCode, rsteNew?.Zip, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.Zip = s;

                    if (MatchAssignText(input, WebCore._s_EditUser_OfficialNumber, rsteNew?.m_sOfficialNumber, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.m_sOfficialNumber = s;
                    if (MatchAssignText(input, WebCore._s_EditUser_DateJoined, rsteNew?.m_sDateJoined, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.m_sDateJoined = s;


                    if (rsteNew == null || rsteNew.IsUploadableQuickroster)
                        {
                        if (MatchAssignText(input, WebCore._s_EditUser_DateOfBirth, rsteNew?.m_sDateOfBirth, out s, ref fNeedSave, ref fFailUpdate))
                            rsteOut.m_sDateOfBirth = s;
                        if (MatchAssignText(input, WebCore._s_EditUser_GamesPerDay, rsteNew?.m_sGamesPerDay, out s, ref fNeedSave, ref fFailUpdate))
                            rsteOut.m_sGamesPerDay = s;
                        if (MatchAssignText(input, WebCore._s_EditUser_GamesPerWeek, rsteNew?.m_sGamesPerWeek, out s, ref fNeedSave, ref fFailUpdate))
                            rsteOut.m_sGamesPerWeek = s;
                        if (MatchAssignText(input, WebCore._s_EditUser_GamesTotal, rsteNew?.m_sTotalGames, out s, ref fNeedSave, ref fFailUpdate))
                            rsteOut.m_sTotalGames = s;
                        if (MatchAssignText(input, WebCore._s_EditUser_WaitMinutes, rsteNew?.m_sWaitMinutes, out s, ref fNeedSave, ref fFailUpdate))
                            rsteOut.m_sWaitMinutes = s;
                        }

                    if (name.Contains(sNamePhoneNumberNext))
                        {
                        // we have a phone control.  Make sure it matches.
                        // NOTE: We don't delete phone numbers, so if it turns out we don't have this number, just skip...
                        if (MatchAssignPhoneNumber(m_webControl, rsteOut, rsteNew, iNextPhone, sNamePhoneNumberNext, sidPhoneNumberNext, sNamePhoneTypeNext, sNamePhonePublicNext))
                            fNeedSave = true;

                        iNextPhone++;
                        SetPhoneNames(iNextPhone, out sNamePhoneNumberNext, out sidPhoneNumberNext, out sNamePhoneTypeNext, out sNamePhoneCarrierNext, out sNamePhonePublicNext);
                        }
                    }
                }

            if (iNextPhone < 4 && rsteNew != null)
                {
                while (iNextPhone < 4)
                    {
                    SetPhoneNames(iNextPhone, out sNamePhoneNumberNext, out sidPhoneNumberNext, out sNamePhoneTypeNext, out sNamePhoneCarrierNext, out sNamePhonePublicNext);
                    if (rsteNew.FHasPhoneNumber(iNextPhone))
                        {
                        // add this phone...
                        ThrowIfNot(m_webControl.FClickControlId(WebCore._sid_EditUser_PhoneNumber_AddNew, sidPhoneNumberNext), "could not add new phone number");
                        if (MatchAssignPhoneNumber(m_webControl, rsteOut, rsteNew, iNextPhone, sNamePhoneNumberNext, sidPhoneNumberNext, sNamePhoneTypeNext, sNamePhonePublicNext))
                            fNeedSave = true;
                        }

                    iNextPhone++;
                    }
                }

            if (fFailUpdate)
                {
                m_srpt.AddMessage($"FAILED to update some general info!  '{rsteOut.Email}' was read only", MSGT.Error);
                }

            if (fNeedSave)
                {
                m_srpt.AddMessage($"Updating general info for {rsteOut.Email}...");
                ThrowIfNot(m_webControl.FClickControlId(WebCore._sid_OfficialsEdit_Button_Save), "couldn't find save button");
                }
            else
                {
                ThrowIfNot(m_webControl.FClickControlId(WebCore._sid_OfficialsEdit_Button_Cancel), "Couldn't find cancel button!");
                }
        }

        /*----------------------------------------------------------------------------
        	%%Function: MatchAssignPhoneNumber
        	%%Qualified: ArbWeb.AwMainForm.MatchAssignPhoneNumber
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private static bool MatchAssignPhoneNumber(
	        WebControl webControl, 
	        RosterEntry rsteOut, 
	        RosterEntry rsteNew, 
	        int iNextPhone, 
	        string sNamePhoneNumberNext,
            string sidPhoneNumberNext, 
	        string sNamePhoneTypeNext, 
	        string sNamePhonePublicNext)
        {
            string sNumberNew = null;
            string sTypeNew = null;
            string sNumber = null;
            string sTypeOptionText = null;
            bool fNeedSave = false;

            sNumber = webControl.GetValueForControlId(sidPhoneNumberNext);

            // handle the phone number
            if (rsteNew != null)
                {
                rsteNew.GetPhoneNumber(iNextPhone, out sNumberNew, out sTypeNew);
                if (webControl.FSetTextForInputControlName(sNamePhoneNumberNext, sNumberNew, true))
                    {
                    // new numbers are public by default
                    webControl.FSetCheckboxControlNameVal(true, sNamePhonePublicNext);
                    fNeedSave = true;
                    }
                }

            // handle the type
            // get the selected item first
            string sTypeOptionValue = webControl.GetSelectedOptionValueFromSelectControlName(sNamePhoneTypeNext);
            // convert the type into the name
            sTypeOptionText = webControl.GetOptionTextFromOptionValueForControlName(sNamePhoneTypeNext, sTypeOptionValue);

            rsteOut.SetPhoneNumber(iNextPhone, sNumber, sTypeOptionText);

            if (rsteNew != null && String.Compare(sTypeOptionText, sTypeNew) != 0)
                {
                // now set the type if we have a new number
                string sNewTypeOptionValue = webControl.GetOptionValueForSelectControlNameOptionText(sNamePhoneTypeNext, sTypeNew);
                webControl.FSetSelectedOptionValueForControlName(sNamePhoneTypeNext, sNewTypeOptionValue);
                fNeedSave = true;
                }

            return fNeedSave;
        }

        // make the rankings on the page match the rankings in our roster
        /*----------------------------------------------------------------------------
        	%%Function: BuildRankingJobs
        	%%Qualified: ArbWeb.AwMainForm.BuildRankingJobs
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private static void BuildRankingJobs(
            IRoster irst,
            string sRankPosition,
            Dictionary<string, int> mpRanked,
            out List<string> plsUnrank, // officials that need to be unranked
            out Dictionary<int, List<string>> mpRank, // officials that need to be ranked
            out Dictionary<int, List<string>> mpRerank) // officials that need to be re-ranked
        {
            List<IRosterEntry> plirste = irst.Plirste;

            // there are 3 things we can potentially do-
            //  1) unrank
            //  2) rank
            //  3) re-rank

            // all of these are most optimally done by multi-selecting and 
            // doing like-item things together
            // 
            // so, we will collect all the stuff together

            // just keep a list of officials to unrank.
            // for rank and re-rank, we want a mapping of (rank -> list of officials)

            plsUnrank = new List<string>();
            mpRank = new Dictionary<int, List<string>>();
            mpRerank = new Dictionary<int, List<string>>();

            // first, unrank any officials that should now become unranked
            foreach (RosterEntry rste in plirste)
                {
                string sReversed = $"{rste.Last}, {rste.First}";

                if (!rste.FRanked(sRankPosition))
                    {
                    if (mpRanked.ContainsKey(sReversed))
                        {
                        // need to unrank
                        plsUnrank.Add(sReversed);
                        }

                    // else, we're cool..we're both unranked
                    }
                else
                    {
                    int nRank = rste.Rank(sRankPosition);

                    // see if we need to rank or rerank
                    if (mpRanked.ContainsKey(sReversed))
                        {
                        // may need to rerank
                        if (mpRanked[sReversed] != nRank)
                            {
                            // need to rerank
                            if (!mpRerank.ContainsKey(nRank))
                                mpRerank.Add(nRank, new List<string>());

                            mpRerank[nRank].Add(sReversed);
                            }
                        }
                    else
                        {
                        // need to rank
                        if (!mpRank.ContainsKey(nRank))
                            mpRank.Add(nRank, new List<string>());

                        mpRank[nRank].Add(sReversed);
                        }
                    }
                }
        }

        /*----------------------------------------------------------------------------
        	%%Function: VisitRankCallbackDownload
        	%%Qualified: ArbWeb.AwMainForm.VisitRankCallbackDownload
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private static void VisitRankCallbackDownload(
	        IRoster irst, 
	        string sRank, 
	        Dictionary<string, int> mpNameRank, 
	        Dictionary<string, string> mpNameOptionValue, 
	        WebControl webControl,
            StatusBox srpt)
        {
	        MicroTimer timer = new MicroTimer();
	        
            // don't do anything with unranked
            // just add the rankings
            foreach (string s in mpNameRank.Keys)
            {
	            ((Roster)irst).FAddRanking(s, sRank, mpNameRank[s]);
            }
            
            timer.Stop();
            srpt.LogData($"VisitRankingDownload({sRank}: {timer.MsecFloat}", 1, MSGT.Body);
        }

        private delegate void VisitRankCallback(
	        IRoster irst, 
	        string sRank, 
	        Dictionary<string, int> mpRanked, 
	        Dictionary<string, string> mpRankedId, 
	        WebControl webControl,
            StatusBox srpt);

        /*----------------------------------------------------------------------------
        	%%Function: HandleRankings
        	%%Qualified: ArbWeb.AwMainForm.HandleRankings
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void HandleRankings(IRoster irst, IRoster irstBuilding)
        {
            if (irst != null && ((Roster)irst).PlsRankings == null)
                return;

            NavigateArbiterRankings();

            Dictionary<string, string> mpPositionOptionsValueText = m_webControl.GetOptionsValueTextMappingFromControlId(WebCore._sid_RanksEdit_Select_PosNames);
            List<string> plsRankingPositions = RankingPositionsBuildFromRst(irst, irstBuilding, mpPositionOptionsValueText);

            if (m_pr.SkipZ)
                {
                List<string> plsKeysToRemove = new List<string>();
                foreach (string sKey in mpPositionOptionsValueText.Keys)
                    {
                    if (mpPositionOptionsValueText[sKey].StartsWith("z"))
                        plsKeysToRemove.Add(sKey);
                    }

                foreach (string sKey in plsKeysToRemove)
                    mpPositionOptionsValueText.Remove(sKey);

                int i = plsRankingPositions.Count;

                while (--i >= 0)
                    {
                    if (plsRankingPositions[i].StartsWith("z"))
                        plsRankingPositions.RemoveAt(i);
                    }
                }

            if (irst == null)
	            VisitRankings(plsRankingPositions, mpPositionOptionsValueText, VisitRankCallbackDownload, irstBuilding, false /*fVerbose*/);
            else
	            VisitRankings(plsRankingPositions, mpPositionOptionsValueText, VisitRankCallbackUpload, irst, false /*fVerbose*/); // true
        }

        /* V I S I T  R A N K I N G S */
        /*----------------------------------------------------------------------------
	    	%%Function: VisitRankings
	    	%%Qualified: ArbWeb.AwMainForm.VisitRankings
	    	%%Contact: rlittle
         
            Visit a rankings page. Used for both upload and download, with the
            callback interface used to differentiate up/down.
	    ----------------------------------------------------------------------------*/
        private void VisitRankings(List<string> plsRankedPositions, IDictionary<string, string> mpPositionOptionsValueText, VisitRankCallback visit, IRoster irstParam, bool fVerboseLog)
        {
            // now, navigate to every ranked positions' page and either fetch or sync every
            // official
            m_srpt.LogData("Visit Rankings", 3, MSGT.Header);
            m_srpt.LogData("plsRankedPositions:", 3, MSGT.Body, plsRankedPositions);

            foreach (string sRankPosition in plsRankedPositions)
                {
                m_srpt.AddMessage($"Processing ranks for {sRankPosition}...");

                if (!FNavigateToRankPosition(mpPositionOptionsValueText, sRankPosition))
                    {
                    m_srpt.AddMessage("Ranks for position '{0}' do not exist on Arbiter!  Skipping...",
                                      MSGT.Error);
                    continue;
                    }

                BuildRankingMapFromPage(sRankPosition, out Dictionary<string, int> mpNameRank, out Dictionary<string, string> mpNameOptionValue);

                m_srpt.LogData("Rankings built: mpRanked:", 5, MSGT.Body, mpNameRank);
                m_srpt.LogData("Rankings built: mpRankedId:", 5, MSGT.Body, mpNameOptionValue);

                visit(irstParam, sRankPosition, mpNameRank, mpNameOptionValue, m_webControl, m_srpt);

                if (fVerboseLog)
                    {
                    BuildRankingMapFromPage(sRankPosition, out Dictionary<string, int> mpNameRankCheck, out Dictionary<string, string> _);

                    List<string> plsUnrank;
                    Dictionary<int, List<string>> mpRank;
                    Dictionary<int, List<string>> mpRerank;
                    BuildRankingJobs(irstParam, sRankPosition, mpNameRankCheck, out plsUnrank, out mpRank, out mpRerank);

                    if (plsUnrank.Count != 0)
                        m_srpt.LogData("plsUnrank not empty: ", 3, MSGT.Error, plsUnrank);
                    else
                        m_srpt.LogData("plsUnrank empty after upload", 5, MSGT.Header);

                    if (mpRank.Count != 0)
                        m_srpt.LogData("mpRank not empty: ", 3, MSGT.Error, mpRank);
                    else
                        m_srpt.LogData("mpRank empty after upload", 5, MSGT.Header);
                    if (mpRerank.Count != 0)
                        m_srpt.LogData("mpRerank not empty: ", 3, MSGT.Error, mpRerank);
                    else
                        m_srpt.LogData("mpRerank empty after upload", 5, MSGT.Header);


                    }
                }
        }

        /*----------------------------------------------------------------------------
        	%%Function: FNavigateToRankPosition
        	%%Qualified: ArbWeb.AwMainForm.FNavigateToRankPosition
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private bool FNavigateToRankPosition(IDictionary<string, string> mpPositionOptionsValueText, string sRankPosition)
        {
			// make sure we have the position
			bool fFound = false;
			
			foreach (string s in mpPositionOptionsValueText.Values)
			{
				if (String.Compare(s, sRankPosition, true) == 0)
					fFound = true;
			}
			
            if (!fFound)
                return false;

            // make sure we have the right checkbox states 
            // (Show unranked only = false, Show Active only = false)
            
            m_webControl.FSetCheckboxControlIdVal(false, WebCore._sid_RanksEdit_Checkbox_Active);
            m_webControl.FSetCheckboxControlIdVal(false, WebCore._sid_RanksEdit_Checkbox_Rank);

            m_webControl.FSetSelectedOptionTextForControlId(WebCore._sid_RanksEdit_Select_PosNames, sRankPosition);
            return true;
        }

        /*----------------------------------------------------------------------------
        	%%Function: BuildRankingMapFromPage
        	%%Qualified: ArbWeb.AwMainForm.BuildRankingMapFromPage
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void BuildRankingMapFromPage(
	        string sRankPosition,
	        out Dictionary<string, int> mpNameRank,
	        out Dictionary<string, string> mpNameOptionValue)
        {
	        List<string> plsUnranked = new List<string>();
	        mpNameRank = new Dictionary<string, int>();
	        mpNameOptionValue = new Dictionary<string, string>();

	        Dictionary<string, string> mpT;

	        // unranked officials
	        mpT = m_webControl.GetOptionsValueTextMappingFromControlId(WebCore._sid_RanksEdit_Select_NotRanked);

	        // for each of the option values, add the text for it (this is the name of the official)
	        foreach (string s in mpT.Keys)
		        plsUnranked.Add(mpT[s]);

	        // ranked officials
	        mpT = m_webControl.GetOptionsValueTextMappingFromControlId(WebCore._sid_RanksEdit_Select_Ranked);

	        foreach (string sKey in mpT.Keys)
	        {
		        string sRankAndName = mpT[sKey];
		        
		        int iColon = sRankAndName.IndexOf(":");
		        if (iColon == -1)
			        throw new Exception("bad format for ranked official on arbiter!");

		        int nRank = Int32.Parse(sRankAndName.Substring(0, iColon));

		        int iStart = iColon + 1;
		        while (Char.IsWhiteSpace(sRankAndName.Substring(iStart, 1)[0]))
			        iStart++;

		        string sName = sRankAndName.Substring(iStart);
		        if (!mpNameRank.ContainsKey(sName))
			        mpNameRank.Add(sName, nRank);
		        else
		        {
			        m_srpt.AddMessage(
				        $"Duplicate key {sName} adding rank {nRank} to rank {sRankPosition}",
				        MSGT.Error);
		        }

		        if (!mpNameOptionValue.ContainsKey(sName))
			        mpNameOptionValue.Add(sName, sKey);
		        else
		        {
			        m_srpt.AddMessage(
				        $"Duplicate key {sName} adding rankid {sRankAndName} to rank {sRankPosition}",
				        MSGT.Error);
		        }
	        }
        }

        /*----------------------------------------------------------------------------
        	%%Function: PlsRankingsBuildFromRst
        	%%Qualified: ArbWeb.AwMainForm.PlsRankingsBuildFromRst
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private static List<string> RankingPositionsBuildFromRst(IRoster irst, IRoster irstBuilding, Dictionary<string, string> mpPositionsOptionsValueText)
        {
            List<string> plsRankingPositions;
            if (irst == null)
                {
                // now, build up our plsRankedPositions
                plsRankingPositions = new List<string>();

                foreach (string s in mpPositionsOptionsValueText.Keys)
                    plsRankingPositions.Add(mpPositionsOptionsValueText[s]);

                ((Roster)irstBuilding).PlsRankings = plsRankingPositions;
                }
            else
                plsRankingPositions = ((Roster)irst).PlsRankings;

            return plsRankingPositions;
        }

        /*----------------------------------------------------------------------------
        	%%Function: NavigateArbiterRankings
        	%%Qualified: ArbWeb.AwMainForm.NavigateArbiterRankings
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void NavigateArbiterRankings()
        {
            if (!m_webControl.FNavToPage(WebCore._s_RanksEdit))
                throw (new Exception("could not navigate to the bulk rankings page"));
        }

        /*----------------------------------------------------------------------------
			%%Function: InvalRoster
			%%Qualified: ArbWeb.AwMainForm.InvalRoster
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
        private void InvalRoster()
        {
            m_rst = null;
        }

        /*----------------------------------------------------------------------------
            %%Function: RstEnsure
            %%Qualified: ArbWeb.AwMainForm.RstEnsure
            %%Contact: rlittle
            
        ----------------------------------------------------------------------------*/
        private Roster RstEnsure(string sInFile)
        {
            if (m_rst != null)
                return m_rst;

            m_rst = new Roster();

            m_rst.ReadRoster(sInFile);
            return m_rst;
        }


        #region Tests

        [TestCase(1, "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclPhones$rptPhone$ctl01$txtPhone",
            "ctl00_ContentHolder_pgeOfficialEdit_conOfficialEdit_uclPhones_rptPhone_ctl01_txtPhone",
            "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclPhones$rptPhone$ctl01$ddlPhoneType",
            "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclPhones$rptPhone$ctl01$ddlCarrier",
            "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclPhones$rptPhone$ctl01$chkPublic")]
        [Test]
        public static void TestSetPhoneNames(int iPhoneRow, string sExpectedNum, string sidExpectedNum, string sExpectedType, string sExpectedCarrier, string sExpectedPublic)
        {
            string sidNumActual, sNumActual, sTypeActual, sCarrierActual, sPublicActual;

            SetPhoneNames(iPhoneRow, out sNumActual, out sidNumActual, out sTypeActual, out sCarrierActual, out sPublicActual);
            Assert.AreEqual(sExpectedNum, sNumActual);
            Assert.AreEqual(sExpectedType, sTypeActual);
            Assert.AreEqual(sExpectedCarrier, sCarrierActual);
            Assert.AreEqual(sExpectedPublic, sPublicActual);
        }

        #endregion
    }

}