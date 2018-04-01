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

            rste.m_plsMisc = SyncPlsMiscWithServer(m_awc.Document2, sEmail, sOfficialID, null, null, ref plsMiscBuilding);
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
        private static string MiscLabelFromControl(IHTMLElement ihe)
        {
            IHTMLElement iheParent;

            iheParent = ihe;
            do
                {
                iheParent = iheParent.parentElement;
                } while (iheParent != null && String.Compare(iheParent.tagName, "TR", StringComparison.InvariantCultureIgnoreCase) != 0);

            if (iheParent == null)
                return null;

            // now, find the first TD child
            IHTMLElementCollection ihecChildren;

            ihecChildren = (IHTMLElementCollection) iheParent.children;

            foreach (IHTMLElement iheChild in ihecChildren)
                {
                if (String.Compare(iheChild.tagName, "TD", StringComparison.InvariantCultureIgnoreCase) == 0)
                    return iheChild.innerText.Trim();
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
        private List<string> SyncPlsMiscWithServer(IHTMLDocument2 oDoc2, string sEmail, string sOfficialID, List<string> plsMiscNew, List<string> plsMiscMapNew,
            ref List<string> plsMiscMapServer)
        {
            bool fNeedSave = false;
            string sValue;

            if (!m_awc.FNavToPage(WebCore._s_EditUser_MiscFields + sOfficialID))
                {
                throw (new Exception("could not navigate to the officials page"));
                }

            IHTMLDocument oDoc = m_awc.Document;
            IHTMLDocument3 oDoc3 = m_awc.Document3;

            IHTMLElementCollection hec;

            // misc field info.  every text input field is a misc field we want to save
            hec = (IHTMLElementCollection) oDoc2.all.tags("input");
            List<string> plsValue = new List<string>();

            sValue = null;

            foreach (IHTMLInputElement ihie in hec)
                {
                if (String.Compare(ihie.type, "text", true) == 0 && ihie.name != null && ihie.name.Contains(WebCore.s_MiscField_EditControlSubstring))
                    {
                    // figure out which misc field this is
                    string sMiscLabel = MiscLabelFromControl((IHTMLElement) ihie);

                    // cool, extract the value
                    sValue = ihie.value;
                    if (sValue == null)
                        sValue = "";

                    if (plsMiscNew != null)
                        {
                        int iMisc = IMiscFromMiscName(plsMiscMapNew, sMiscLabel);

                        if (iMisc == -1)
                            throw new Exception(
                                "couldn't find misc field name! (OR maybe this is a new misc field that the roster doesn't know about, in which case we should just set it to empty, but this isn't debugged yet so we don't trust that decision yet, hence the exception");

                        if (iMisc == -1 // couldn't find this server misc field in the roster's list of misc fields...set to empty
                            && ihie.value != null
                            && ihie.value.Length > 0)
                            {
                            // null means empty which replaces non-empty
                            ihie.value = "";
                            fNeedSave = true;
                            }
                        else if (iMisc != -1
                                 && String.Compare(plsMiscNew[iMisc], sValue, true /*ignoreCase*/) != 0)
                            {
                            ihie.value = plsMiscNew[iMisc];
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
                }

            // before we return, commit the change or cancel (so we are no longer on the page)
            if (fNeedSave)
                {
                m_srpt.AddMessage(String.Format("Updating misc info...", sEmail));
                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_MiscFields_Button_Save), "Couldn't find save button");

                m_awc.FWaitForNavFinish();
                }
            else
                {
                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_MiscFields_Button_Cancel), "Couldn't find cancel button");

                m_awc.FWaitForNavFinish();
                }

            return plsValue;
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
        private bool MatchAssignText(IHTMLInputElement ihie, string sMatch, string sNewValue, out string sAssign, ref bool fNeedSave, ref bool fFailAssign)
        {
            sAssign = null;
            if (ihie.name.Contains(sMatch))
                {
                sAssign = ihie.value;

                if (sAssign == null)
                    sAssign = "";

                if (sNewValue != null)
                    {
                    // check to see if it matches what we have
                    // find a match on email address first
                    if (sNewValue != null
                        && String.Compare(sNewValue, sAssign, true /*ignoreCase*/) != 0)
                        {
                        if (ihie.disabled)
                            {
                            fFailAssign = true;
                            }
                        else
                            {
                            ihie.value = sNewValue;
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
            sPhoneNum = String.Format("{0}ctl{1:00}{2}", WebCore._s_EditUser_PhoneNumber_Prefix, iPhoneRow, WebCore._s_EditUser_PhoneNumber_Suffix);
            sidPhoneNum = String.Format("{0}ctl{1:00}{2}", WebCore._sid_EditUser_PhoneNumber_Prefix, iPhoneRow, WebCore._sid_EditUser_PhoneNumber_Suffix);
            sPhoneType = String.Format("{0}ctl{1:00}{2}", WebCore._s_EditUser_PhoneType_Prefix, iPhoneRow, WebCore._s_EditUser_PhoneType_Suffix);
            sPhoneCarrier = String.Format("{0}ctl{1:00}{2}", WebCore._s_EditUser_PhoneCarrier_Prefix, iPhoneRow, WebCore._s_EditUser_PhoneCarrier_Suffix);
            sPhonePublicNext = String.Format("{0}ctl{1:00}{2}", WebCore._s_EditUser_PhonePublic_Prefix, iPhoneRow, WebCore._s_EditUser_PhonePublic_Suffix);
        }

        /* S Y N C  R S T E  W I T H  S E R V E R */
        /*----------------------------------------------------------------------------
        	%%Function: SyncRsteWithServer
        	%%Qualified: ArbWeb.AwMainForm.SyncRsteWithServer
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void SyncRsteWithServer(IHTMLDocument2 oDoc2, string sOfficialID, RosterEntry rsteOut, RosterEntry rsteNew)
        {
            bool fFailUpdate = false;
            bool fNeedSave = false;

            // ok, nav to the page and scrape
            if (!m_awc.FNavToPage(WebCore._s_EditUser + sOfficialID))
                {
                throw (new Exception("could not navigate to the officials page"));
                }

            IHTMLDocument oDoc = m_awc.Document;
            IHTMLDocument3 oDoc3 = m_awc.Document3;

            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection) oDoc2.all.tags("input");

            string sPhoneNumberNext;
            string sidPhoneNumberNext;
            string sPhoneTypeNext;
            string sPhoneCarrierNext;
            string sPhonePublicNext;
            int iNextPhone = 1;
            SetPhoneNames(iNextPhone, out sPhoneNumberNext, out sidPhoneNumberNext, out sPhoneTypeNext, out sPhoneCarrierNext, out sPhonePublicNext);

            foreach (IHTMLInputElement ihie in hec)
                {
                if (String.Compare(ihie.type, "checkbox", true) == 0)
                    {
                    // checkboxes are either ready or active
                    if (ihie.name.Contains("Active"))
                        rsteOut.m_fActive = String.Compare(ihie.value, "on", true) == 0;
                    else if (ihie.name.Contains("Ready"))
                        rsteOut.m_fReady = String.Compare(ihie.value, "on", true) == 0;
                    }

                if (String.Compare(ihie.type, "text", true) == 0 && ihie.name != null)
                    {
                    if (ihie.name.Contains(WebCore._s_EditUser_Email))
                        {
//						if (ihie.value == null && rsteOut.m_sEmail != null && rsteOut.m_sEmail != "")
                        // continue;

                        if (ihie.value != null && rsteOut.Email != null)
                            {
                            if (String.Compare(ihie.value, rsteOut.Email, true) != 0)
                                throw new Exception("email addresses don't match!");
                            }
                        else
                            {
                            m_srpt.AddMessage(String.Format("NULL Email address for {0},{1}", rsteOut.First, rsteOut.Last), StatusBox.StatusRpt.MSGT.Error);
                            }
                        }

                    string s;
                    if (MatchAssignText(ihie, WebCore._s_EditUser_FirstName, rsteNew?.First, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.First = s;
                    if (MatchAssignText(ihie, WebCore._s_EditUser_LastName, rsteNew?.Last, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.Last = s;
                    if (MatchAssignText(ihie, WebCore._s_EditUser_Address1, rsteNew?.Address1, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.Address1 = s;
                    if (MatchAssignText(ihie, WebCore._s_EditUser_Address2, rsteNew?.Address2, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.Address2 = s;
                    if (MatchAssignText(ihie, WebCore._s_EditUser_City, rsteNew?.City, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.City = s;
                    if (MatchAssignText(ihie, WebCore._s_EditUser_State, rsteNew?.State, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.State = s;
                    if (MatchAssignText(ihie, WebCore._s_EditUser_PostalCode, rsteNew?.Zip, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.Zip = s;

                    if (MatchAssignText(ihie, WebCore._s_EditUser_OfficialNumber, rsteNew?.m_sOfficialNumber, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.m_sOfficialNumber = s;
                    if (MatchAssignText(ihie, WebCore._s_EditUser_DateJoined, rsteNew?.m_sDateJoined, out s, ref fNeedSave, ref fFailUpdate))
                        rsteOut.m_sDateJoined = s;


                    if (rsteNew == null || rsteNew.IsUploadableQuickroster)
                        {
                        if (MatchAssignText(ihie, WebCore._s_EditUser_DateOfBirth, rsteNew?.m_sDateOfBirth, out s, ref fNeedSave, ref fFailUpdate))
                            rsteOut.m_sDateOfBirth = s;
                        if (MatchAssignText(ihie, WebCore._s_EditUser_GamesPerDay, rsteNew?.m_sGamesPerDay, out s, ref fNeedSave, ref fFailUpdate))
                            rsteOut.m_sGamesPerDay = s;
                        if (MatchAssignText(ihie, WebCore._s_EditUser_GamesPerWeek, rsteNew?.m_sGamesPerWeek, out s, ref fNeedSave, ref fFailUpdate))
                            rsteOut.m_sGamesPerWeek = s;
                        if (MatchAssignText(ihie, WebCore._s_EditUser_GamesTotal, rsteNew?.m_sTotalGames, out s, ref fNeedSave, ref fFailUpdate))
                            rsteOut.m_sTotalGames = s;
                        if (MatchAssignText(ihie, WebCore._s_EditUser_WaitMinutes, rsteNew?.m_sWaitMinutes, out s, ref fNeedSave, ref fFailUpdate))
                            rsteOut.m_sWaitMinutes = s;
                        }

                    if (ihie.name.Contains(sPhoneNumberNext))
                        {
                        // we have a phone control.  Make sure it matches.
                        // NOTE: We don't delete phone numbers, so if it turns out we don't have this number, just skip...
                        if (MatchAssignPhoneNumber(oDoc2, rsteOut, rsteNew, iNextPhone, sPhoneNumberNext, sidPhoneNumberNext, sPhoneTypeNext, sPhonePublicNext))
                            fNeedSave = true;

                        iNextPhone++;
                        SetPhoneNames(iNextPhone, out sPhoneNumberNext, out sidPhoneNumberNext, out sPhoneTypeNext, out sPhoneCarrierNext, out sPhonePublicNext);
                        }
                    }
                }

            if (iNextPhone < 4 && rsteNew != null)
                {
                while (iNextPhone < 4)
                    {
                    SetPhoneNames(iNextPhone, out sPhoneNumberNext, out sidPhoneNumberNext, out sPhoneTypeNext, out sPhoneCarrierNext, out sPhonePublicNext);
                    if (rsteNew.FHasPhoneNumber(iNextPhone))
                        {
                        // add this phone...
                        m_awc.ResetNav();
                        m_awc.ReportNavState("Before click control");
                        ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_EditUser_PhoneNumber_AddNew, sidPhoneNumberNext), "could not add new phone number");
                        m_awc.ReportNavState("After click control");
                        m_awc.FWaitForNavFinish();
                        oDoc2 = m_awc.Document2;
                        if (MatchAssignPhoneNumber(oDoc2, rsteOut, rsteNew, iNextPhone, sPhoneNumberNext, sidPhoneNumberNext, sPhoneTypeNext, sPhonePublicNext))
                            fNeedSave = true;
                        }

                    iNextPhone++;
                    }
                }

            if (fFailUpdate)
                {
                m_srpt.AddMessage(String.Format("FAILED to update some general info!  '{0}' was read only", rsteOut.Email), StatusBox.StatusRpt.MSGT.Error);
                }

            if (fNeedSave)
                {
                m_srpt.AddMessage(String.Format("Updating general info...", rsteOut.Email));
                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_OfficialsEdit_Button_Save), "couldn't find save button");
                m_awc.FWaitForNavFinish();
                }
            else
                {
                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, WebCore._sid_OfficialsEdit_Button_Cancel), "Couldn't find cancel button!");
                m_awc.FWaitForNavFinish();
                }
        }

        /*----------------------------------------------------------------------------
        	%%Function: MatchAssignPhoneNumber
        	%%Qualified: ArbWeb.AwMainForm.MatchAssignPhoneNumber
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private static bool MatchAssignPhoneNumber(IHTMLDocument2 oDoc2, RosterEntry rsteOut, RosterEntry rsteNew, int iNextPhone, string sPhoneNumberNext,
            string sidPhoneNumberNext, string sPhoneTypeNext, string sPhonePublicNext)
        {
            string sNumberNew = null;
            string sTypeNew = null;
            string sNumber = null;
            string sType = null;
            bool fNeedSave = false;

            // handle the phone number
            if (rsteNew != null)
                {
                rsteNew.GetPhoneNumber(iNextPhone, out sNumberNew, out sTypeNew);
                if (ArbWebControl.FSetInputControlText(oDoc2, sPhoneNumberNext, sNumberNew, true))
                    {
                    // new numbers are public by default
                    ArbWebControl.FSetCheckboxControlVal(oDoc2, true, sPhonePublicNext);
                    fNeedSave = true;
                    }
                }

            sNumber = ArbWebControl.SGetControlValue(oDoc2, sidPhoneNumberNext);
            if (sNumber == null)
                {
                return false;
                }

            // handle the type
            // get the selected item first
            string sTypeID = ArbWebControl.SGetSelectSelectedValue(oDoc2, sPhoneTypeNext);
            // convert the type into the name
            sType = ArbWebControl.SGetSelectValFromDoc(oDoc2, sPhoneTypeNext, sTypeID);

            rsteOut.SetPhoneNumber(iNextPhone, sNumber, sType);

            if (rsteNew != null && String.Compare(sType, sTypeNew) != 0)
                {
                // now set the type if we have a new number
                string sNewTypeID = ArbWebControl.SGetSelectIDFromDoc(oDoc2, sPhoneTypeNext, sTypeNew);
                ArbWebControl.FSetSelectControlValue(oDoc2, sPhoneTypeNext, sNewTypeID, false);
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
                string sReversed = String.Format("{0}, {1}", rste.Last, rste.First);

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
        private static void VisitRankCallbackDownload(IRoster irst, string sRank, Dictionary<string, int> mpRanked, Dictionary<string, string> mpRankedId, ArbWebControl awc,
            StatusBox.StatusRpt srpt)
        {
            // don't do anything with unranked
            // just add the rankings
            foreach (string s in mpRanked.Keys)
                irst.FAddRanking(s, sRank, mpRanked[s]);
        }

        private delegate void VisitRankCallback(IRoster irst, string sRank, Dictionary<string, int> mpRanked, Dictionary<string, string> mpRankedId, ArbWebControl awc,
            StatusBox.StatusRpt srpt);

        /*----------------------------------------------------------------------------
        	%%Function: HandleRankings
        	%%Qualified: ArbWeb.AwMainForm.HandleRankings
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void HandleRankings(IRoster irst, IRoster irstBuilding)
        {
            if (irst != null && irst.PlsRankings == null)
                return;

            NavigateArbiterRankings();

            IHTMLDocument2 oDoc2;

            oDoc2 = m_awc.Document2;

            Dictionary<string, string> mpRankFilter = ArbWebControl.MpGetSelectValues(m_srpt, oDoc2, WebCore._s_RanksEdit_Select_PosNames);
            List<string> plsRankings = PlsRankingsBuildFromRst(irst, irstBuilding, mpRankFilter);

            if (m_pr.SkipZ)
                {
                List<string> plsKeysToRemove = new List<string>();
                foreach (string sKey in mpRankFilter.Keys)
                    {
                    if (sKey.StartsWith("z"))
                        plsKeysToRemove.Add(sKey);
                    }

                foreach (string sKey in plsKeysToRemove)
                    {
                    mpRankFilter.Remove(sKey);
                    }

                int i = plsRankings.Count;

                while (--i >= 0)
                    {
                    if (plsRankings[i].StartsWith("z"))
                        plsRankings.RemoveAt(i);
                    }
                }

            if (irst == null)
                VisitRankings(plsRankings, mpRankFilter, VisitRankCallbackDownload, irstBuilding, false /*fVerbose*/);
            else
                VisitRankings(plsRankings, mpRankFilter, VisitRankCallbackUpload, irst, false /*fVerbose*/); // true
        }

        /* V I S I T  R A N K I N G S */
        /*----------------------------------------------------------------------------
	    	%%Function: VisitRankings
	    	%%Qualified: ArbWeb.AwMainForm.VisitRankings
	    	%%Contact: rlittle
         
            Visit a rankings page. Used for both upload and download, with the
            callback interface used to differentiate up/down.
	    ----------------------------------------------------------------------------*/
        private void VisitRankings(List<string> plsRankedPositions, IDictionary<string, string> mpRankFilter, VisitRankCallback pfnVrc, IRoster irstParam, bool fVerboseLog)
        {
            // now, navigate to every ranked positions' page and either fetch or sync every
            // official
            m_srpt.LogData("Visit Rankings", 1, StatusRpt.MSGT.Header);
            m_srpt.LogData("plsRankedPositions:", 2, StatusRpt.MSGT.Body, plsRankedPositions);

            foreach (string sRankPosition in plsRankedPositions)
                {
                m_srpt.AddMessage(String.Format("Processing ranks for {0}...", sRankPosition));

                if (!FNavigateToRankPosition(mpRankFilter, sRankPosition))
                    {
                    m_srpt.AddMessage("Ranks for position '{0}' do not exist on Arbiter!  Skipping...",
                                      StatusBox.StatusRpt.MSGT.Error);
                    continue;
                    }

                IHTMLDocument2 oDoc2 = m_awc.Document2;
                // m_awc.RefreshPage();

                Dictionary<string, int> mpRanked;
                Dictionary<string, string> mpRankedId;

                BuildRankingMapFromPage(oDoc2, sRankPosition, out mpRanked, out mpRankedId);

                m_srpt.LogData("Rankings built: mpRanked:", 4, StatusRpt.MSGT.Body, mpRanked);
                m_srpt.LogData("Rankings built: mpRankedId:", 4, StatusRpt.MSGT.Body, mpRankedId);

                pfnVrc(irstParam, sRankPosition, mpRanked, mpRankedId, m_awc, m_srpt);

                if (fVerboseLog)
                    {
                    m_awc.RefreshPage();

                    Dictionary<string, int> mpRankedCheck;
                    Dictionary<string, string> mpRankedIdCheck;

                    BuildRankingMapFromPage(oDoc2, sRankPosition, out mpRankedCheck, out mpRankedIdCheck);

                    List<string> plsUnrank;
                    Dictionary<int, List<string>> mpRank;
                    Dictionary<int, List<string>> mpRerank;
                    BuildRankingJobs(irstParam, sRankPosition, mpRankedCheck, out plsUnrank, out mpRank, out mpRerank);

                    if (plsUnrank.Count != 0)
                        m_srpt.LogData("plsUnrank not empty: ", 1, StatusRpt.MSGT.Error, plsUnrank);
                    else
                        m_srpt.LogData("plsUnrank empty after upload", 4, StatusRpt.MSGT.Header);

                    if (mpRank.Count != 0)
                        m_srpt.LogData("mpRank not empty: ", 1, StatusRpt.MSGT.Error, mpRank);
                    else
                        m_srpt.LogData("mpRank empty after upload", 4, StatusRpt.MSGT.Header);
                    if (mpRerank.Count != 0)
                        m_srpt.LogData("mpRerank not empty: ", 1, StatusRpt.MSGT.Error, mpRerank);
                    else
                        m_srpt.LogData("mpRerank empty after upload", 4, StatusRpt.MSGT.Header);


                    }
                }
        }

        /*----------------------------------------------------------------------------
        	%%Function: FNavigateToRankPosition
        	%%Qualified: ArbWeb.AwMainForm.FNavigateToRankPosition
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private bool FNavigateToRankPosition(IDictionary<string, string> mpRankFilter, string sRankPosition)
        {
// try to navigate to the page
            if (!mpRankFilter.ContainsKey(sRankPosition))
                return false;

            // make sure we have the right checkbox states 
            // (Show unranked only = false, Show Active only = false)
            ArbWebControl.FSetCheckboxControlVal(m_awc.Document2, false, WebCore._s_RanksEdit_Checkbox_Active);
            ArbWebControl.FSetCheckboxControlVal(m_awc.Document2, false, WebCore._s_RanksEdit_Checkbox_Rank);

            m_awc.ResetNav();
            m_awc.FSetSelectControlText(m_awc.Document2, WebCore._s_RanksEdit_Select_PosNames, WebCore._sid_RanksEdit_Select_PosNames, sRankPosition, false);
            m_awc.FWaitForNavFinish();
            return true;
        }

        /*----------------------------------------------------------------------------
        	%%Function: BuildRankingMapFromPage
        	%%Qualified: ArbWeb.AwMainForm.BuildRankingMapFromPage
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void BuildRankingMapFromPage(IHTMLDocument2 oDoc2, string sRankPosition, out Dictionary<string, int> mpRanked, out Dictionary<string, string> mpRankedId)
        {
            List<string> plsUnranked = new List<string>();
            mpRanked = new Dictionary<string, int>();
            mpRankedId = new Dictionary<string, string>();

            Dictionary<string, string> mpT;

            // unranked officials
            mpT = ArbWebControl.MpGetSelectValues(m_srpt, oDoc2, WebCore._s_RanksEdit_Select_NotRanked);

            foreach (string s in mpT.Keys)
                plsUnranked.Add(s);

            // ranked officials
            mpT = ArbWebControl.MpGetSelectValues(m_srpt, oDoc2, WebCore._s_RanksEdit_Select_Ranked);

            foreach (string s in mpT.Keys)
                {
                int iColon = s.IndexOf(":");
                if (iColon == -1)
                    throw new Exception("bad format for ranked official on arbiter!");

                int nRank = Int32.Parse(s.Substring(0, iColon));

                int iStart = iColon + 1;
                while (Char.IsWhiteSpace(s.Substring(iStart, 1)[0]))
                    iStart++;

                string sRankKey = s.Substring(iStart);
                if (!mpRanked.ContainsKey(sRankKey))
                    mpRanked.Add(sRankKey, nRank);
                else
                    {
                    m_srpt.AddMessage(
                        String.Format("Duplicate key {0} adding rank {1} to rank {2}", sRankKey, nRank, sRankPosition),
                        StatusRpt.MSGT.Error);
                    }

                if (!mpRankedId.ContainsKey(sRankKey))
                    mpRankedId.Add(sRankKey, mpT[s]);
                else
                    {
                    m_srpt.AddMessage(
                        String.Format("Duplicate key {0} adding rankid {1} to rank {2}", sRankKey, mpT[s], sRankPosition),
                        StatusRpt.MSGT.Error);
                    }
                }
        }

        /*----------------------------------------------------------------------------
        	%%Function: PlsRankingsBuildFromRst
        	%%Qualified: ArbWeb.AwMainForm.PlsRankingsBuildFromRst
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private static List<string> PlsRankingsBuildFromRst(IRoster irst, IRoster irstBuilding, Dictionary<string, string> mpRankFilter)
        {
            List<string> plsRankings;
            if (irst == null)
                {
                // now, build up our plsRankedPositions
                plsRankings = new List<string>();

                foreach (string s in mpRankFilter.Keys)
                    plsRankings.Add(s);

                irstBuilding.PlsRankings = plsRankings;
                }
            else
                plsRankings = irst.PlsRankings;

            return plsRankings;
        }

        /*----------------------------------------------------------------------------
        	%%Function: NavigateArbiterRankings
        	%%Qualified: ArbWeb.AwMainForm.NavigateArbiterRankings
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void NavigateArbiterRankings()
        {
            m_awc.ResetNav();
            if (!m_awc.FNavToPage(WebCore._s_RanksEdit))
                throw (new Exception("could not navigate to the bulk rankings page"));
            m_awc.FWaitForNavFinish();
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