﻿using System;
using System.Diagnostics;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Net;
using System.IO;
using Microsoft.Win32;
using AxSHDocVw;
using StatusBox;
using mshtml;
using System.Text.RegularExpressions;
using Microsoft.Office;
using System.Runtime.InteropServices;
using Outlook=Microsoft.Office.Interop.Outlook;
using Excel=Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;

namespace ArbWeb
{
    /// <summary>
    /// Summary description for AwMainForm.
    /// </summary>
    public partial class AwMainForm : System.Windows.Forms.Form
    {
	    private class PGL
	    {
	        public class OFI
	        {
	            public string sOfficialID;
	            public string sEmail;

	            public OFI()
	            {
	                sOfficialID = null;
	                sEmail = null;
	            }
	        };

	        public PGL()
	        {
	            plofi = new List<OFI>();
	        }
	        public List<OFI> plofi;
//	        public List<string> rgsLinks;
//	        public List<string> rgsData;
	        public int iCur;
	    };

        /* N A V I G A T E  O F F I C I A L S  P A G E  A L L  O F F I C I A L S */
        /*----------------------------------------------------------------------------
	    	%%Function: NavigateOfficialsPageAllOfficials
	    	%%Qualified: ArbWeb.AwMainForm.NavigateOfficialsPageAllOfficials
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
        private void NavigateOfficialsPageAllOfficials()
        {
            EnsureLoggedIn();

            ThrowIfNot(m_awc.FNavToPage(_s_Page_OfficialsView), "Couldn't nav to officials view!");
            m_awc.FWaitForNavFinish();

            // from the officials view, make sure we are looking at active officials
            m_awc.ResetNav();
            IHTMLDocument2 oDoc2 = m_awc.Document2;

            ArbWebControl.FSetSelectControlText(oDoc2, _s_OfficialsView_Select_Filter, "All Officials", true);
            m_awc.FWaitForNavFinish();
        }

        /* P O P U L A T E  P G L  F R O M  P A G E  C O R E */
        /*----------------------------------------------------------------------------
			%%Function: PopulatePglFromPageCore
			%%Qualified: ArbWeb.AwMainForm.PopulatePglFromPageCore
			%%Contact: rlittle

			Return a PGL (page of links) from the give sUrl.  

				rx3 is a match for either the link name or the link text
				rx4 is a match for the link name always (will supercede rx3)
				rxData, if set, is the match that will be used to populat the rgsData

			on exit, rgsLinkNames, rgsLinks, and (optionally) rgsData will be
			populated in the pglLinks

            we need to collect information from two separate places in the DOM --
            the Official Name (Last, First) will be in an anchor linking to the
            offical page (which we can get the official ID from).  Then we have the
            email address from which we can get the actual email address.
         
            because of this, we collect the email first, then note that we
            are looking for the official ID.  essentially, a state machine.. (albeit 2 
            states)
		----------------------------------------------------------------------------*/
        private void PopulatePglOfficialsFromPageCore(PGL pgl, IHTMLDocument2 oDoc)
        {
            IHTMLElementCollection links = oDoc.links;
            Regex rx3 = new Regex("OfficialEdit.aspx\\?userID=.*");
            Regex rxData = new Regex("mailto:.*");

            bool fLookingForEmail = false;

            // build up a list of probable index links
            foreach (HTMLAnchorElementClass link in links)
                {
                string sLinkName = link.nameProp;
                string sLinkText = link.innerText;
                string sLinkTarget = link.href;

                if (sLinkText == null)
                    sLinkText = "";

                if (sLinkName == null)
                    sLinkName = link.innerText.Substring(1, link.innerText.Length - 1);

                if (rxData != null && sLinkTarget != null && rxData.IsMatch(sLinkTarget))
                    {
                    if (fLookingForEmail)
                        {
                        // adjust the top item in plofi...
                        pgl.plofi[pgl.plofi.Count - 1].sEmail = sLinkTarget;
                        fLookingForEmail = false;
                        }
                    else
                        {
                        m_srpt.AddMessage("Found (" + sLinkTarget + ") when not looking for email!", StatusRpt.MSGT.Error);
                        }
                    }

                if (rx3.IsMatch(sLinkName))
                    {
                    PGL.OFI ofi = new PGL.OFI();

                    ofi.sEmail = "";
                    ofi.sOfficialID = link.href;
                    pgl.plofi.Add(ofi);

                    fLookingForEmail = true;
                    }

                }
            pgl.iCur = 0;
        }

        private bool FAppendToFile(string sFile, string sName, string sEmail, List<string> plsMisc)
        {
            StreamWriter sw = new StreamWriter(sFile, true, System.Text.Encoding.Default);

            sEmail = sEmail.Substring(sEmail.IndexOf(":") + 1);

            if (sw == null)
                return false;

            string sFirst, sLast;

            if (sName.IndexOf(" ") == -1)
                {
                sFirst = sName;
                sLast = "";
                }
            else
                {
                sFirst = sName.Substring(0, sName.IndexOf(" "));
                sLast = sName.Substring(sName.IndexOf(" ") + 1);
                }

            sw.Write("\"{0}\",\"{1}\",\"{2}\"", sFirst, sLast, sEmail);
            foreach (string s in plsMisc)
                sw.Write(",\"{0}\"", s);

            sw.WriteLine();
            sw.Close();
            return true;
        }

        /* U P D A T E  M I S C */
        /*----------------------------------------------------------------------------
			%%Function: UpdateMisc
			%%Qualified: ArbWeb.AwMainForm.UpdateMisc
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
        private void UpdateMisc(PGL pgl, Roster rst, ref RosterEntry rste)
        {
            // ok, nav to the page and scrape
//			m_awc.ResetNav();
//			ThrowIfNot(FClickControl(m_awc.Document2, "pgeOfficialEdit_cmnRelatedData_tskEditMiscFields"), "couldn't find misc fields link");
//          m_awc.FWaitForNavFinish();

            if (!m_awc.FNavToPage(_s_EditUser_MiscFields + pgl.plofi[pgl.iCur].sOfficialID))
                {
                throw (new Exception("could not navigate to the officials page"));
                }

            IHTMLDocument oDoc = m_awc.Document;
            IHTMLDocument2 oDoc2 = m_awc.Document2;
            IHTMLDocument3 oDoc3 = m_awc.Document3;

            IHTMLElementCollection hec;

            // misc field info.  every text input field is a misc field we want to save
            hec = (IHTMLElementCollection) oDoc2.all.tags("input");
            List<string> plsValue = new List<string>();
            string sValue = null;
            bool fNeedSave = false;

            foreach (IHTMLInputElement ihie in hec)
                {
                if (String.Compare(ihie.type, "text", true) == 0)
                    {
                    // cool, extract the value
                    sValue = ihie.value;
                    if (sValue == null)
                        sValue = "";
                    if (rst != null)
                        {
                        // check to see if it matches what we have
                        // find a match on email address first
                        List<string> plsMisc = rst.PlsLookupEmail(pgl.plofi[pgl.iCur].sEmail);

                        if (plsMisc != null)
                            {
                            if (plsMisc.Count <= plsValue.Count
                                && ihie.value != null
                                && ihie.value.Length > 0)
                                {
                                // null means empty which replaces non-empty
                                ihie.value = "";
                                fNeedSave = true;
                                }
                            else if (plsMisc.Count > plsValue.Count
                                     && String.Compare(plsMisc[plsValue.Count], sValue, true /*ignoreCase*/) != 0)
                                {
                                ihie.value = plsMisc[plsValue.Count];
                                fNeedSave = true;
                                }
                            }
                        }
                    plsValue.Add(sValue);
                    // don't break here -- just get the next misc value...
                    }
                }

            if (fNeedSave)
                {
                m_srpt.AddMessage(String.Format("Updating misc info...", pgl.plofi[pgl.iCur].sEmail));
                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_MiscFields_Button_Save), "Couldn't find save button");

                m_awc.FWaitForNavFinish();
                }
            else
                {
                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_MiscFields_Button_Cancel), "Couldn't find cancel button");
//				((IHTMLElement)(oDoc2.all.item("pgeMiscFieldsEdit_navMiscFieldsEdit_btnSave", 0))).click();
                m_awc.FWaitForNavFinish();
                }

            if (sValue == null || plsValue.Count == 0)
                throw new Exception("couldn't extract misc field for official");

            rste.m_plsMisc = plsValue;
        }

        private void MatchAssignText(IHTMLInputElement ihie, Roster rst, string sMatch, string sRsteMatched, ref string sAssign, ref bool fNeedSave, ref bool fFailAssign)
        {
            if (ihie.name.Contains(sMatch))
                {
                sAssign = ihie.value;

                if (sAssign == null)
                    sAssign = "";

                if (rst != null)
                    {
                    // check to see if it matches what we have
                    // find a match on email address first
                    if (sRsteMatched != null
                        && String.Compare(sRsteMatched, sAssign, true /*ignoreCase*/) != 0)
                        {
                        if (ihie.disabled)
                            {
                            fFailAssign = true;
                            }
                        else
                            {
                            ihie.value = sRsteMatched;
                            fNeedSave = true;
                            }
                        }
                    }
                }
        }

        /* U P D A T E  I N F O */
        /*----------------------------------------------------------------------------
			%%Function: UpdateInfo
			%%Qualified: ArbWeb.AwMainForm.UpdateInfo
			%%Contact: rlittle

			when we leave, if rst was null, then rste will have the values as we
			fetched from arbiter
		----------------------------------------------------------------------------*/
        private void UpdateInfo(PGL pgl, Roster rst, ref RosterEntry rste, bool fMarkOnly)
        {
            bool fNeedSave = false;
            bool fFailUpdate = false;

            RosterEntry rsteMatch = null;

            if (rst != null)
                rsteMatch = rst.RsteLookupEmail(pgl.plofi[pgl.iCur].sEmail);

            if (rsteMatch == null)
                rsteMatch = new RosterEntry(); // just to get nulls filled in to the member variables
            else
                rsteMatch.Marked = true;

            if (fMarkOnly)
                return;

            // ok, nav to the page and scrape
            if (!m_awc.FNavToPage(_s_EditUser + pgl.plofi[pgl.iCur].sOfficialID))
                {
                throw (new Exception("could not navigate to the officials page"));
                }

            IHTMLDocument oDoc = m_awc.Document;
            IHTMLDocument2 oDoc2 = m_awc.Document2;
            IHTMLDocument3 oDoc3 = m_awc.Document3;

            IHTMLElementCollection hec;

            hec = (IHTMLElementCollection) oDoc2.all.tags("input");

            foreach (IHTMLInputElement ihie in hec)
                {
                if (String.Compare(ihie.type, "checkbox", true) == 0)
                    {
                    // checkboxes are either ready or active
                    if (ihie.name.Contains("Active"))
                        rste.m_fActive = String.Compare(ihie.value, "on", true) == 0;
                    else if (ihie.name.Contains("Ready"))
                        rste.m_fReady = String.Compare(ihie.value, "on", true) == 0;
                    }

                if (String.Compare(ihie.type, "text", true) == 0)
                    {
                    if (ihie.name.Contains("Email"))
                        {
//						if (ihie.value == null && rste.m_sEmail != null && rste.m_sEmail != "")
                        // continue;

                        if (ihie.value != null && rste.m_sEmail != null)
                            {
                            if (String.Compare(ihie.value, rste.m_sEmail, true) != 0)
                                throw new Exception("email addresses don't match!");
                            }
                        else
                            {
                            m_srpt.AddMessage(String.Format("NULL Email address for {0},{1}", rste.m_sFirst, rste.m_sLast), StatusBox.StatusRpt.MSGT.Error);
                            }
                        }
                    MatchAssignText(ihie, rst, "FirstName", rsteMatch.m_sFirst, ref rste.m_sFirst, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "LastName", rsteMatch.m_sLast, ref rste.m_sLast, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "Address1", rsteMatch.m_sAddress1, ref rste.m_sAddress1, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "Address2", rsteMatch.m_sAddress2, ref rste.m_sAddress2, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "City", rsteMatch.m_sCity, ref rste.m_sCity, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "State", rsteMatch.m_sState, ref rste.m_sState, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "PostalCode", rsteMatch.m_sZip, ref rste.m_sZip, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "OfficialNumber", rsteMatch.m_sOfficialNumber, ref rste.m_sOfficialNumber, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "DateOfBirth", rsteMatch.m_sDateOfBirth, ref rste.m_sDateOfBirth, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "DateJoined", rsteMatch.m_sDateJoined, ref rste.m_sDateJoined, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "GamesPerDay", rsteMatch.m_sGamesPerDay, ref rste.m_sGamesPerDay, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "GamesPerWeek", rsteMatch.m_sGamesPerWeek, ref rste.m_sGamesPerWeek, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "GamesTotal", rsteMatch.m_sTotalGames, ref rste.m_sTotalGames, ref fNeedSave, ref fFailUpdate);
                    MatchAssignText(ihie, rst, "WaitMinutes", rsteMatch.m_sWaitMinutes, ref rste.m_sWaitMinutes, ref fNeedSave, ref fFailUpdate);
                    }
                }

            if (fFailUpdate)
                {
                m_srpt.AddMessage(String.Format("FAILED to update some general info!  '{0}' was read only", pgl.plofi[pgl.iCur].sEmail), StatusBox.StatusRpt.MSGT.Error);
                }
            if (fNeedSave)
                {
                m_srpt.AddMessage(String.Format("Updating general info...", pgl.plofi[pgl.iCur].sEmail));
                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_OfficialsEdit_Button_Save), "couldn't find save button");
                m_awc.FWaitForNavFinish();
                }
            else
                {
                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_OfficialsEdit_Button_Cancel), "Couldn't find cancel button!");
                m_awc.FWaitForNavFinish();
                }
            if (rste.m_sAddress1 == null || rste.m_sAddress2 == null || rste.m_sCity == null || rste.m_sDateJoined == null
                || rste.m_sDateOfBirth == null || rste.m_sEmail == null || rste.m_sFirst == null || rste.m_sGamesPerDay == null
                || rste.m_sGamesPerWeek == null || rste.m_sLast == null || rste.m_sOfficialNumber == null
                || rste.m_sState == null || rste.m_sTotalGames == null || rste.m_sWaitMinutes == null
                || rste.m_sZip == null)
                {
                throw new Exception("couldn't extract one more more fields from official info");
                }
        }

        private static void VisitRankCallbackUpload(Roster rst, string sRankPosition, Dictionary<string, int> mpRanked, Dictionary<string, string> mpRankedId, ArbWebControl2 awc, StatusBox.StatusRpt srpt)
        {
            IHTMLDocument2 oDoc2;
            oDoc2 = awc.Document2;

            List<string> plsUnrank;
            Dictionary<int, List<string>> mpRank;
            Dictionary<int, List<string>> mpRerank;

            BuildRankingJobs(rst, sRankPosition, mpRanked, out plsUnrank, out mpRank, out mpRerank);

            // at this point, we have a list of jobs to do.

            // first, unrank everyone that needs unranked
            if (plsUnrank.Count > 0)
                {
                ArbWebControl.FResetMultiSelectOptions(oDoc2, _s_RanksEdit_Select_Ranked);
                foreach (string s in plsUnrank)
                    {
                    if (!ArbWebControl.FSelectMultiSelectOption(oDoc2, _s_RanksEdit_Select_Ranked, mpRankedId[s], true))
                        throw new Exception("couldn't select an official for unranking!");
                    }

                // now, do the unrank
                awc.ResetNav();
                awc.FClickControl(oDoc2, _s_RanksEdit_Button_Unrank);
                awc.FWaitForNavFinish();
                oDoc2 = awc.Document2;
                }

            // now, let's rerank the folks that need to be re-ranked
            // we will do this once for every new rank we are setting
            foreach (int nRank in mpRerank.Keys)
                {
                ArbWebControl.FResetMultiSelectOptions(oDoc2, _s_RanksEdit_Select_Ranked);
                foreach (string s in mpRerank[nRank])
                    {
                    if (!ArbWebControl.FSelectMultiSelectOption(oDoc2, _s_RanksEdit_Select_Ranked, mpRankedId[s], true))
                        throw new Exception("couldn't select an official for reranking!");
                    }
                ArbWebControl.FSetInputControlText(oDoc2, _s_RanksEdit_Input_Rank, nRank.ToString(), false);

                // now, rank'em
                awc.ResetNav();
                awc.FClickControl(oDoc2, _s_RanksEdit_Button_ReRank);
                awc.FWaitForNavFinish();
                oDoc2 = awc.Document2;
                }

            // finally, let's rank the folks that weren't ranked before

            foreach (int nRank in mpRank.Keys)
                {
                ArbWebControl.FResetMultiSelectOptions(oDoc2, _s_RanksEdit_Select_NotRanked);
                foreach (string s in mpRank[nRank])
                    {
                    if (!ArbWebControl.FSelectMultiSelectOption(oDoc2, _s_RanksEdit_Select_NotRanked, s, false))
                        srpt.AddMessage(String.Format("Could not select an official for ranking: {0}", s),
                                        StatusRpt.MSGT.Error);
                    // throw new Exception("couldn't select an official for ranking!");
                    }

                ArbWebControl.FSetInputControlText(oDoc2, _s_RanksEdit_Input_Rank, nRank.ToString(), false);

                // now, rank'em
                awc.ResetNav();
                awc.FClickControl(oDoc2, _s_RanksEdit_Button_Rank);
                awc.FWaitForNavFinish();
                oDoc2 = awc.Document2;
                }
        }

        // make the rankings on the page match the rankings in our roster
        private static void BuildRankingJobs(
            Roster rst,
            string sRankPosition,
            Dictionary<string, int> mpRanked,
            out List<string> plsUnrank, // officials that need to be unranked
            out Dictionary<int, List<string>> mpRank, // officials that need to be ranked
            out Dictionary<int, List<string>> mpRerank) // officials that need to be re-ranked
        {
            List<RosterEntry> plrste = rst.Plrste;

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
            foreach (RosterEntry rste in plrste)
                {
                string sReversed = String.Format("{0}, {1}", rste.m_sLast, rste.m_sFirst);

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

        private static void VisitRankCallbackDownload(Roster rst, string sRank, Dictionary<string, int> mpRanked, Dictionary<string, string> mpRankedId, ArbWebControl2 awc, StatusBox.StatusRpt srpt)
        {
            // don't do anything with unranked
            // just add the rankings
            foreach (string s in mpRanked.Keys)
                rst.FAddRanking(s, sRank, mpRanked[s]);
        }

        private delegate void VisitRankCallback(Roster rst, string sRank, Dictionary<string, int> mpRanked, Dictionary<string, string> mpRankedId, ArbWebControl2 awc, StatusBox.StatusRpt srpt);

        private void HandleRankings(Roster rst, ref Roster rstBuilding)
        {
            if (rst != null && rst.PlsRankings == null)
                return;

            NavigateArbiterRankings();

            IHTMLDocument2 oDoc2;

            oDoc2 = m_awc.Document2;

            Dictionary<string, string> mpRankFilter = ArbWebControl.MpGetSelectValues(m_srpt, oDoc2, _s_RanksEdit_Select_PosNames);
            List<string> plsRankings = PlsRankingsBuildFromRst(rst, rstBuilding, mpRankFilter);

            if (rst == null)
                VisitRankings(plsRankings, mpRankFilter, VisitRankCallbackDownload, rstBuilding, false /*fVerbose*/);
            else
                VisitRankings(plsRankings, mpRankFilter, VisitRankCallbackUpload, rst, false /*fVerbose*/); // true
        }

        /* V I S I T  R A N K I N G S */
        /*----------------------------------------------------------------------------
	    	%%Function: VisitRankings
	    	%%Qualified: ArbWeb.AwMainForm.VisitRankings
	    	%%Contact: rlittle
         
            Visit a rankings page. Used for both upload and download, with the
            callback interface used to differentiate up/down.
	    ----------------------------------------------------------------------------*/
        private void VisitRankings(List<string> plsRankedPositions, IDictionary<string, string> mpRankFilter, VisitRankCallback pfnVrc, Roster rstParam, bool fVerboseLog)
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

                pfnVrc(rstParam, sRankPosition, mpRanked, mpRankedId, m_awc, m_srpt);

                if (fVerboseLog)
                    {
                    m_awc.RefreshPage();

                    Dictionary<string, int> mpRankedCheck;
                    Dictionary<string, string> mpRankedIdCheck;

                    BuildRankingMapFromPage(oDoc2, sRankPosition, out mpRankedCheck, out mpRankedIdCheck);

                    List<string> plsUnrank;
                    Dictionary<int, List<string>> mpRank;
                    Dictionary<int, List<string>> mpRerank;
                    BuildRankingJobs(rstParam, sRankPosition, mpRankedCheck, out plsUnrank, out mpRank, out mpRerank);

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

        private bool FNavigateToRankPosition(IDictionary<string, string> mpRankFilter, string sRankPosition)
        {
// try to navigate to the page
            if (!mpRankFilter.ContainsKey(sRankPosition))
                return false;

            // make sure we have the right checkbox states 
            // (Show unranked only = false, Show Active only = false)
            ArbWebControl.FSetCheckboxControlVal(m_awc.Document2, false, _s_RanksEdit_Checkbox_Active);
            ArbWebControl.FSetCheckboxControlVal(m_awc.Document2, false, _s_RanksEdit_Checkbox_Rank);

            m_awc.ResetNav();
            ArbWebControl.FSetSelectControlText(m_awc.Document2, _s_RanksEdit_Select_PosNames, sRankPosition, false);
            m_awc.FWaitForNavFinish();
            return true;
        }

        private void BuildRankingMapFromPage(IHTMLDocument2 oDoc2, string sRankPosition, out Dictionary<string, int> mpRanked, out Dictionary<string, string> mpRankedId)
        {
            List<string> plsUnranked = new List<string>();
            mpRanked = new Dictionary<string, int>();
            mpRankedId = new Dictionary<string, string>();

            Dictionary<string, string> mpT;

            // unranked officials
            mpT = ArbWebControl.MpGetSelectValues(m_srpt, oDoc2, _s_RanksEdit_Select_NotRanked);

            foreach (string s in mpT.Keys)
                plsUnranked.Add(s);

            // ranked officials
            mpT = ArbWebControl.MpGetSelectValues(m_srpt, oDoc2, _s_RanksEdit_Select_Ranked);

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

        private void VisitRosterRankUploaded(Roster rst, string sRank, Dictionary<string, int> mpRanked, IHTMLDocument2 oDoc2,
            Dictionary<string, string> mpRankedId)
        {
        }

        private static void VisitRosterRankDownload(Roster rstBuilding, Dictionary<string, int> mpRanked, string sRank)
        {
        }

        private static List<string> PlsRankingsBuildFromRst(Roster rst, Roster rstBuilding, Dictionary<string, string> mpRankFilter)
        {
            List<string> plsRankings;
            if (rst == null)
                {
                // now, build up our plsRankedPositions
                plsRankings = new List<string>();

                foreach (string s in mpRankFilter.Keys)
                    plsRankings.Add(s);

                rstBuilding.PlsRankings = plsRankings;
                }
            else
                plsRankings = rst.PlsRankings;
            return plsRankings;
        }

        private void NavigateArbiterRankings()
        {
            m_awc.ResetNav();
            if (!m_awc.FNavToPage(_s_RanksEdit))
                throw (new Exception("could not navigate to the bulk rankings page"));
            m_awc.FWaitForNavFinish();
        }

        private void AddOfficials(List<RosterEntry> plrsteNew)
        {
            foreach (RosterEntry rste in plrsteNew)
                {
                // add the official rste
                m_srpt.AddMessage(String.Format("Adding official '{0}', {1}", rste.Name, rste.m_sEmail), StatusBox.StatusRpt.MSGT.Body);
                m_srpt.PushLevel();

                // go to the add user page
                m_awc.ResetNav();
                if (!m_awc.FNavToPage(_s_AddUser))
                    {
                    throw (new Exception("could not navigate to the add user page"));
                    }
                m_awc.FWaitForNavFinish();

                IHTMLDocument2 oDoc2;

                oDoc2 = m_awc.Document2;

                ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, _s_AddUser_Input_FirstName, rste.m_sFirst, false /*fCheck*/), "Failed to find first name control");
                ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, _s_AddUser_Input_LastName, rste.m_sLast, false /*fCheck*/), "Failed to find last name control");
                ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, _s_AddUser_Input_Email, rste.m_sEmail, false /*fCheck*/), "Failed to find email control");

                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_AddUser_Button_Next), "Can't click next button on adduser");
                m_awc.FWaitForNavFinish();

                // we are either adding a new user, or a user that arbiter already knows
                // about...
                // 
                if (!ArbWebControl.FCheckForControl(oDoc2, _sid_AddUser_Input_Address1))
                    {
                    m_srpt.AddMessage(String.Format("Email {0} already in use", rste.m_sEmail), StatusBox.StatusRpt.MSGT.Warning);

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
                    if (String.Compare(sName, rste.Name, true /*ignoreCase*/) != 0)
                        {
                        if (MessageBox.Show(String.Format("Trying to add office {0} and found a mismatch with existing official {1}, with email {2}", rste.Name, sName, rste.m_sEmail), "ArbWeb", MessageBoxButtons.YesNo) != DialogResult.Yes)
                            {
                            // ok, then just cancel...
                            m_awc.ResetNav();
                            ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_AddUser_Button_Cancel), "Can't click cancel button on adduser");
                            m_awc.FWaitForNavFinish();
                            continue;
                            }
                        }
                    // cool, just go on...
                    m_awc.ResetNav();
                    ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_AddUser_Button_Next), "Can't click next button on adduser");
                    m_awc.FWaitForNavFinish();

                    // sigh, now we're being asked whether we want to add local personal info.  of course
                    // we don't since it will be thrown away when they choose to join our group!

                    // but make sure that we're really on that page...
                    sText = oDoc2.body.innerText;
                    ThrowIfNot(sText.IndexOf("as a fully integrated user") > 0, "Didn't find the confirmation text on 'personal info' portion of existing user sequence");

                    // cool, let's just move on again...
                    m_awc.ResetNav();
                    ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_AddUser_Button_Next), "Can't click next button on adduser");
                    m_awc.FWaitForNavFinish();

                    // now fallthrough to the "Official's info" page handling, which is common
                    }
                else
                    {
                    // if there's an address control, then this is a brand new official
                    ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, _s_AddUser_Input_Address1, rste.m_sAddress1, false /*fCheck*/), "Failed to find address1 control");
                    ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, _s_AddUser_Input_Address1, rste.m_sAddress2, false /*fCheck*/), "Failed to find address2 control");
                    ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, _s_AddUser_Input_City, rste.m_sCity, false /*fCheck*/), "Failed to find city control");
                    ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, _s_AddUser_Input_State, rste.m_sState, false /*fCheck*/), "Failed to find state control");
                    ThrowIfNot(ArbWebControl.FSetInputControlText(oDoc2, _s_AddUser_Input_Zip, rste.m_sZip, false /*fCheck*/), "Failed to find zip control");

                    m_awc.ResetNav();
                    ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_AddUser_Button_Next), "Can't click next button on adduser");
                    m_awc.FWaitForNavFinish();

                    // fallthrough to the common handling below
                    }

                // now we are on the last add official page
                // the only thing that *might* be interesting on this page is the active button which is
                // not checked by default...
                ThrowIfNot(ArbWebControl.FCheckForControl(oDoc2, _sid_AddUser_Input_IsActive), "bad hierarchy in add user.  expected screen with 'active' checkbox, didn't find it.");

                // don't worry about Active for now...Just click next again
                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_AddUser_Button_Next), "Can't click next button on adduser");
                m_awc.FWaitForNavFinish();

                // and now we're on the finish page.  oddly enough, the finish button has the "Cancel" ID
                ThrowIfNot(String.Compare("Finish", ArbWebControl.SGetControlValue(oDoc2, _sid_AddUser_Button_Cancel)) == 0, "Finish screen didn't have a finish button");

                m_awc.ResetNav();
                ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_AddUser_Button_Cancel), "Can't click finish/cancel button on adduser");
                m_awc.FWaitForNavFinish();
                m_srpt.PopLevel();
                // and now we're back somewhere (probably officials edit page)
                // continue to the next one...
                }
            // and that's it...simple n'est pas?
        }

        // Update the "last login" value.  since we are scraping the screen for this, we have to deal with pagination
        private static void VOPC_UpdateLastAccess(AwMainForm awf, IHTMLDocument2 oDoc2, Object o)
        {
            Roster rstBuilding = (Roster) o;

            awf.UpdateLastAccessFromDoc(rstBuilding, oDoc2);
        }

        // object could be RST or PGL
        private delegate void VisitOfficialsPageCallback(AwMainForm awf, IHTMLDocument2 oDoc2, Object o);

        private void ProcessAllOfficialPages(VisitOfficialsPageCallback vopc, Object o)
        {
            NavigateOfficialsPageAllOfficials();

            IHTMLDocument2 oDoc2 = m_awc.Document2;

            // first, get the first pages and callback

            vopc(this, oDoc2, o);

            // figure out how many pages we have
            // find all of the <a> tags with an href that targets a pagination postback
            IHTMLElementCollection ihec = (IHTMLElementCollection) oDoc2.all.tags("a");
            List<string> plsHrefs = new List<string>();

            foreach (IHTMLAnchorElement iha in ihec)
                {
                if (iha.href != null && iha.href.Contains(_s_OfficialsView_PaginationHrefPostbackSubstr))
                    {
                    // we can't just remember this element because we will be navigating around.  instead we will
                    // just remember the entire target so we can find it again
                    plsHrefs.Add(iha.href);
                    }
                }

            // now, we are going to navigate to each page by finding and clicking each pagination link in turn
            foreach (string sHref in plsHrefs)
                {
                ihec = (IHTMLElementCollection) oDoc2.all.tags("a");
                foreach (IHTMLAnchorElement iha in ihec)
                    {
                    if (String.Compare(iha.href, sHref, true /*ignoreCase*/) == 0)
                        {
                        // now we need to click on the navigation
                        ((IHTMLElement) iha).click();
                        m_awc.FWaitForNavFinish();
                        oDoc2 = m_awc.Document2;

                        vopc(this, oDoc2, o);
                        break; // done processing the element collection -- have to process the next one for the next doc
                        }
                    }

                }
        }


        // Assuming we are on the core officials page...
        private void UpdateLastAccessFromDoc(Roster rstBuilding, IHTMLDocument2 oDoc2)
        {
            IHTMLTable ihtbl;

            // misc field info.  every text input field is a misc field we want to save
            ihtbl = (IHTMLTable) oDoc2.all.item(_sid_OfficialsView_ContentTable, 0);

            foreach (IHTMLTableRow ihtr in ihtbl.rows)
                {
                IHTMLElement iheEmail = (IHTMLElement) ihtr.cells.item(3);
                IHTMLElement iheSignedIn = (IHTMLElement) ihtr.cells.item(4);

                if (iheEmail == null || iheSignedIn == null)
                    continue;

                string sEmail = iheEmail.innerText;
                string sSignedIn = iheSignedIn.innerText;

                RosterEntry rste = rstBuilding.RsteLookupEmail(sEmail);
                if (rste == null)
                    {
                    m_srpt.AddMessage(
                                      String.Format("Lookup failed during ProcessAllOfficialPages for official '{0}'({1})",
                                                    ((IHTMLElement) ihtr.cells.item(2)).innerText, sEmail), StatusBox.StatusRpt.MSGT.Error);
                    continue;
                    }

                m_srpt.AddMessage(String.Format("Updating last access for official '{0}', {1}", rste.Name, sSignedIn),
                                  StatusBox.StatusRpt.MSGT.Body);
                rste.m_sLastSignin = sSignedIn;
                }
        }


        /* D O  C O R E  R O S T E R  U P D A T E */
        /*----------------------------------------------------------------------------
			%%Function: DoCoreRosterUpdate
			%%Qualified: ArbWeb.AwMainForm.DoCoreRosterUpdate
			%%Contact: rlittle

			Do the core roster updating.  We are being given the list of links on
			the official's edit page, the roster that we are uploading (if any),
			and a list of officials to limit our handling to (this is used when 
			we just added new officials and we just want to update their info/misc
			fields...)
			
		----------------------------------------------------------------------------*/
        private void DoCoreRosterUpdate(PGL pgl, Roster rst, Roster rstBuilding, List<RosterEntry> plrsteLimit)
        {
            pgl.iCur = 0;
            Dictionary<string, bool> mpOfficials = new Dictionary<string, bool>();

            if (plrsteLimit != null)
                {
                foreach (RosterEntry rsteCheck in plrsteLimit)
                    mpOfficials.Add("MAILTO:" + rsteCheck.m_sEmail.ToUpper(), true);
                }

            while (pgl.iCur < pgl.plofi.Count && (rst == null || m_cbRankOnly.Checked == false) && pgl.iCur < pgl.plofi.Count)
                {
                if (rst == null
                    || (rst.PlsLookupEmail(pgl.plofi[pgl.iCur].sEmail) != null
                        && pgl.plofi[pgl.iCur].sEmail.Length != 0))
                    {
                    RosterEntry rste = new RosterEntry();
                    bool fMarkOnly = false;

                    rste.SetEmail((string) pgl.plofi[pgl.iCur].sEmail);
                    m_srpt.AddMessage(String.Format("Processing roster info for {0}...", pgl.plofi[pgl.iCur].sEmail));

                    if (m_cbAddOfficialsOnly.Checked && plrsteLimit == null)
                        fMarkOnly = true;

                    if (plrsteLimit != null)
                        {
                        if (!mpOfficials.ContainsKey(((string) pgl.plofi[pgl.iCur].sEmail.ToUpper())))
                            {
                            pgl.iCur++;
                            continue; // it doesn't match an official in the "limit-to" list.
                            }
                        fMarkOnly = false; // we want to process this one.
                        }

                    if (!fMarkOnly)
                        UpdateMisc(pgl, rst, ref rste);

                    // don't call UpdateInfo on a newly added official
                    if (plrsteLimit == null && (rst == null || !rst.IsQuick))
                        UpdateInfo(pgl, rst, ref rste, fMarkOnly);

                    if (rst == null)
                        {
                        rstBuilding.Add(rste);
//                        rste.AppendToFile(sOutFile, m_rgsRankings);
                        // at this point, we have the name and the affiliation
                        //						if (!FAppendToFile(sOutFile, sName, (string)pgl.rgsData[pgl.iCur], plsValue))
                        //							throw new Exception("couldn't append to the file!");
                        }
                    else
                        {
                        RosterEntry rsteT = rst.RsteLookupEmail(pgl.plofi[pgl.iCur].sEmail);

                        if (rsteT != null)
                            rsteT.Marked = true;
                        }

                    if (m_cbTestOnly.Checked)
                        {
                        break;
                        }
                    }

                pgl.iCur++;
                }
        }

        private static void VOPC_PopulatePgl(AwMainForm awf, IHTMLDocument2 oDoc2, Object o)
        {
            awf.PopulatePglOfficialsFromPageCore((PGL) o, oDoc2);
        }

        private PGL PglGetOfficialsFromWeb()
        {
            EnsureLoggedIn();
            int i;

            PGL pgl = new PGL();

            ProcessAllOfficialPages(VOPC_PopulatePgl, pgl);

//            NavigateOfficialsPageAllOfficials();

//            IHTMLDocument2 oDoc = m_awc.Document2;

//			PopulatePglFromPageCore(pgl, oDoc);
            // for now, assume that all the officials fit on the same screen!!

            // if there are no links, then we aren't logged in yet
            if (pgl.plofi.Count == 0)
                {
                throw (new Exception("Not logged in after EnsureLoggedIn()!!"));
                }

            // ok, now grab the userIDs and put those in the pgl
            i = 0;
            while (i < pgl.plofi.Count)
                {
                string s = (string) pgl.plofi[i].sOfficialID;

                string sID = s.Substring(s.IndexOf("=") + 1);
                pgl.plofi[i].sOfficialID = sID;
                i++;
                }

            return pgl;
        }

        /* H A N D L E  R O S T E R */
        /*----------------------------------------------------------------------------
			%%Function: HandleRoster
			%%Qualified: ArbWeb.AwMainForm.HandleRoster
			%%Contact: rlittle

			If rst == null, then we're downloading the roster.  Otherwise, we are
			uploading
		----------------------------------------------------------------------------*/
        private void HandleRoster(Roster rst, string sOutFile)
        {
            Roster rstBuilding = null;
            PGL pgl;

            // we're not going to write the roster out until the end now...

            if (rst == null)
                rstBuilding = new Roster();

            pgl = PglGetOfficialsFromWeb();
            DoCoreRosterUpdate(pgl, rst, rstBuilding, null /*plrsteLimit*/);

            if (rstBuilding != null)
                {
                // get the last login date from the officials main page
                NavigateOfficialsPageAllOfficials();
                ProcessAllOfficialPages(VOPC_UpdateLastAccess, rstBuilding);
                }

            if (rst != null)
                {
                List<RosterEntry> plrsteUnmarked = rst.PlrsteUnmarked();

                // we might have some officials left "unmarked".  These need to be added

                // at this point, all the officials have either been marked or need to 
                // be added

                if (plrsteUnmarked.Count > 0)
                    {
                    if (MessageBox.Show(String.Format("There are {0} new officials.  Add these officials?", plrsteUnmarked.Count), "ArbWeb", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                        AddOfficials(plrsteUnmarked);
                        // now we have to reload the page of links and do the whole thing again (updating info, etc)
                        // so we get the misc fields updated.  Then fall through to the rankings and do everyone at
                        // once
                        pgl = PglGetOfficialsFromWeb(); // refresh to get new officials
                        DoCoreRosterUpdate(pgl, rst, null /*rstBuilding*/, plrsteUnmarked);
                        // now we can fall through to our core ranking handling...
                        }
                    }
                }

            // now, do the rankings.  this is easiest done in the bulk rankings tool...
            HandleRankings(rst, ref rstBuilding);
            // lastly, if we're downloading, then output the roster

            if (rst == null)
                rstBuilding.WriteRoster(sOutFile);

            if (m_cbTestOnly.Checked)
                {
                MessageBox.Show("Stopping after 1 roster item");
                }
        }
		string SBuildRosterFilename()
		{
			string sOutFile;
            string sPrefix = "";
            
			if (m_ebRoster.Text.Length < 1)
				{
				sOutFile = String.Format("{0}", Environment.GetEnvironmentVariable("temp"));
				}
			else
				{
				sOutFile = System.IO.Path.GetDirectoryName(m_ebRoster.Text);
				string[] rgs;
				if (m_ebRoster.Text.Length > 5 && sOutFile.Length > 0)
					{
					rgs = CountsData.RexHelper.RgsMatch(m_ebRoster.Text.Substring(sOutFile.Length + 1), "([.*])roster");
					if (rgs != null && rgs.Length > 0 && rgs[0] != null)
						sPrefix = rgs[0];
					}
				}

			sOutFile = String.Format("{0}{2}\\roster_{1:MM}{1:dd}{1:yy}_{1:HH}{1:mm}.csv", sOutFile, DateTime.Now, sPrefix);
			return sOutFile;
		}

		private void DoDownloadQuickRoster(object sender, EventArgs e)
		{
			m_srpt.AddMessage("Starting Quick Roster download...");
			m_srpt.PushLevel();
			PushCursor(Cursors.WaitCursor);

		    string sOutFile = SBuildRosterFilename();

			m_ebRoster.Text = sOutFile;

			//string sTempFile = "C:\\Users\\rlittle\\AppData\\Local\\Temp\\temp3c92cb56-0b95-41c0-8eb5-37387bacf4f6.csv";
			string sTempFile = SRosterFileDownload();

//			m_awc.PopSaveToFile();
//			m_awc.PopNewWindow3Delegate();

			// now, get the rankings and update the last access date

			Roster rstBuilding = new Roster();

			rstBuilding.ReadRoster(sTempFile);

    		ProcessAllOfficialPages(VOPC_UpdateLastAccess, rstBuilding);
			HandleRankings(null, ref rstBuilding);
			rstBuilding.WriteRoster(m_ebRoster.Text);
//			System.IO.File.Copy(sTempFile, m_ebRoster.Text);
			System.IO.File.Delete(m_ebRosterWorking.Text);
			System.IO.File.Copy(sOutFile, m_ebRosterWorking.Text);

			PopCursor();
			m_srpt.PopLevel();

			m_srpt.AddMessage("Completed Quick Roster download.");
		}

	    private string SRosterFileDownload()
	    {
	        IHTMLDocument2 oDoc2;
// navigate to the officials page...
	        EnsureLoggedIn();

	        ThrowIfNot(m_awc.FNavToPage(_s_Page_OfficialsView), "Couldn't nav to officials view!");
	        m_awc.FWaitForNavFinish();

	        oDoc2 = m_awc.Document2;

	        // from the officials view, make sure we are looking at active officials
	        m_awc.ResetNav();
	        ArbWebControl.FSetSelectControlText(oDoc2, _s_OfficialsView_Select_Filter, "All Officials", true);
	        m_awc.FWaitForNavFinish();

	        oDoc2 = m_awc.Document2;
	        // now we have all officials showing.  download the report

	        // sometimes running the javascript takes a while, but the page isn't busy
	        int cTry = 3;
	        while (cTry > 0)
	            {
	            m_awc.ResetNav();
	            m_awc.ReportNavState("Before click on PrintRoster: ");
	            ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_OfficialsView_PrintRoster), "Can't click on roster control");
	            m_awc.FWaitForNavFinish();

	            oDoc2 = m_awc.Document2;
	            if (ArbWebControl.FCheckForControl(oDoc2, _sid_RosterPrint_MergeStyle))
	                break;

	            cTry--;
	            }

	        // now we are on the PrintRoster screen

	        // clicking on the Merge Style control will cause a page refresh
	        m_awc.ResetNav();
	        ThrowIfNot(m_awc.FClickControl(oDoc2, _sid_RosterPrint_MergeStyle), "Can't click on roster control");
	        m_awc.FWaitForNavFinish();

	        oDoc2 = m_awc.Document2;

	        ThrowIfNot(ArbWebControl.FCheckForControl(oDoc2, _sid_RosterPrint_DateJoined),
	                   "Couldn't find expected control on roster print config!");

	        // check a whole bunch of config checkboxes
	        ArbWebControl.FSetCheckboxControlVal(oDoc2, true, _s_RosterPrint_DateJoined);
	        ArbWebControl.FSetCheckboxControlVal(oDoc2, true, _s_RosterPrint_OfficialNumber);
	        ArbWebControl.FSetCheckboxControlVal(oDoc2, true, _s_RosterPrint_MiscFields);
	        ArbWebControl.FSetCheckboxControlVal(oDoc2, true, _s_RosterPrint_NonPublicPhone);
	        ArbWebControl.FSetCheckboxControlVal(oDoc2, true, _s_RosterPrint_NonPublicAddress);

	        m_awc.ResetNav();
//    		m_awc.PushNewWindow3Delegate(new DWebBrowserEvents2_NewWindow3EventHandler(DownloadQuickRosterNewWindowDelegate));
//          m_awc.PushSaveToFile(sOutFile);


	        ((IHTMLElement) (oDoc2.all.item(_sid_RosterPrint_BeginPrint, 0))).click();

	        string sTempFile = String.Format("{0}\\temp{1}.csv", Environment.GetEnvironmentVariable("Temp"),
	                                         System.Guid.NewGuid().ToString());
	        System.Windows.Forms.Clipboard.SetText(sTempFile);
	        MessageBox.Show(
	            String.Format(
	                "Please download the roster to {0}. This path is on the clipboard, so you can just past it into the file/save dialog when you click Save.\n\nWhen the download is complete, click OK.",
	                sTempFile), "ArbWeb", MessageBoxButtons.OK);
	        return sTempFile;
	    }
		private void InvalRoster()
		{
			m_rst = null;
		}

		private Roster RstEnsure(string sInFile)
		{
			if (m_rst != null)
				return m_rst;

			m_rst = new Roster();

			m_rst.ReadRoster(sInFile);
			return m_rst;
		}

    }
}