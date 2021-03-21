﻿using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using TCore.StatusBox;
using System.Runtime.InteropServices;
using Outlook=Microsoft.Office.Interop.Outlook;
using System.Threading.Tasks;
using TCore.Util;
using HtmlAgilityPack;
using OpenQA.Selenium;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace ArbWeb
{
    /// <summary>
    /// Summary description for AwMainForm.
    /// </summary>
    public partial class AwMainForm : System.Windows.Forms.Form
    {

        /* B U I L D  A N N  N A M E */
        /*----------------------------------------------------------------------------
        	%%Function: BuildAnnName
        	%%Qualified: ArbWeb.AwMainForm.BuildAnnName
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private static string BuildAnnName(string sPrefix, string sSuffix, string sCtl)
        {
            return $"{sPrefix}{sCtl}{sSuffix}";
        }

        /* D O  G E N  M A I L  M E R G E  A N D  A N N O U C E */
        /*----------------------------------------------------------------------------
        	%%Function: DoGenMailMergeAndAnnouce
        	%%Qualified: ArbWeb.AwMainForm.DoGenMailMergeAndAnnouce
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoGenMailMergeAndAnnouce()
        {
            CountsData gc = GcEnsure(m_pr.RosterWorking, m_pr.GameCopy, m_cbIncludeCanceled.Checked);
            Roster rst = RstEnsure(m_pr.RosterWorking);
            m_srpt.AddMessage("Generating mail merge documents...", MSGT.Header, false);

            m_srpt.LogData("GamesFromFilter...", 3, MSGT.Header);

            // first, generate the mailmerge source csv file.  this is either the entire roster, or just the folks 
            // rated for the sports we are filtered to
            GameData.GameSlots gms = gc.GamesFromFilter(
	            WebCore.RgsFromChlbx(m_cbFilterSport.Checked, m_chlbxSports),
	            WebCore.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels),
	            false,
	            m_saOpenSlots);

            m_srpt.LogData("Beginning to build rosterfiltered...", 3, MSGT.Body);
            Roster rstFiltered;

            if (m_cbFilterRank.Checked)
                rstFiltered = rst.FilterByRanks(gms.RequiredRanks());
            else
                rstFiltered = rst;

            string sCsvTemp = Filename.SBuildTempFilename("MailMergeRoster", "csv");
            m_srpt.LogData($"Writing filtered roster to {sCsvTemp}", 3, MSGT.Body);
            StreamWriter sw = new StreamWriter(sCsvTemp, false, System.Text.Encoding.Default);

            sw.WriteLine("email,firstname,lastname");
            foreach (RosterEntry rste in rstFiltered.Plrste)
                {
                if (m_ebFilter.Text == "" || rste.FMatchAnyMisc(m_ebFilter.Text))
                    sw.WriteLine("{0},{1},{2}", rste.Email, rste.First, rste.Last);
                }
            sw.Flush();
            sw.Close();

            // ok, now create the mailmerge .docx
            string sTempName;
            string sArbiterHelpNeeded;
            m_srpt.LogData("Filtered roster written", 3, MSGT.Body);

            string sApp = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string sFile = Path.Combine(Path.GetPathRoot(sApp),Path.GetDirectoryName(sApp), "mailmergedoc.docx");

            sTempName = Filename.SBuildTempFilename("mailmergedoc", "docx");
            m_srpt.LogData($"Writing mailmergedoc to {sTempName} using template at {sFile}", 3, MSGT.Body);
            OOXML.CreateMailMergeDoc(sFile, sTempName, sCsvTemp, gms, out sArbiterHelpNeeded);

            m_srpt.LogData($"ArbiterHelp HTML created: {sArbiterHelpNeeded}", 5, MSGT.Body);
            System.Windows.Forms.Clipboard.SetText(sArbiterHelpNeeded);
            if (m_cbLaunch.Checked)
                {
                m_srpt.AddMessage("Done, launching document...", MSGT.Header, false);
                System.Diagnostics.Process.Start(sTempName);
                }
            if (m_cbSetArbiterAnnounce.Checked)
                SetArbiterAnnounce(sArbiterHelpNeeded);

            DoPendingQueueUIOp();
        }

        /* S E T  A R B I T E R  A N N O U N C E */
        /*----------------------------------------------------------------------------
        	%%Function: SetArbiterAnnounce
        	%%Qualified: ArbWeb.AwMainForm.SetArbiterAnnounce
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void SetArbiterAnnounce(string sArbiterHelpNeeded)
        {
	        m_srpt.AddMessage("Starting Announcement Set...");
	        m_srpt.PushLevel();

	        EnsureLoggedIn();
	        ThrowIfNot(m_webControl.FNavToPage(WebCore._s_Announcements), "Couldn't nav to announcements page!");
	        m_webControl.WaitForPageLoad();

	        // now we need to find the URGENT HELP NEEDED row
	        string sHtml = m_webControl.Driver.FindElement(By.XPath("//body")).GetAttribute("innerHTML");
	        HtmlDocument html = new HtmlDocument();
	        html.LoadHtml(sHtml);

	        string sXpath = "//div[@id='D9UrgentHelpNeeded']";

            HtmlNode node = html.DocumentNode.SelectSingleNode(sXpath);

	        string sCtl = null;

	        m_srpt.LogData("Found D9UrgentHelpNeeded DIV, looking for parent TR element", 3, MSGT.Body);


	        // ok, go up to the parent TR.

	        HtmlNode nodeFind = node;

	        while (nodeFind.Name.ToLower() != "tr")
	        {
		        nodeFind = nodeFind.ParentNode;
		        ThrowIfNot(nodeFind != null, "Can't find HELP announcement");
	        }

	        m_srpt.LogData("Found D9UrgentHelpNeeded parent TR", 3, MSGT.Body);

	        // now find one of our controls and get its control number
	        string s = nodeFind.InnerHtml;
	        int ich = s.IndexOf(WebCore._s_Announcements_Button_Edit_Prefix);
	        if (ich > 0)
	        {
		        sCtl = s.Substring(ich + WebCore._s_Announcements_Button_Edit_Prefix.Length, 5);
	        }

	        m_srpt.LogData($"Extracted ID for announcment to set: {sCtl}", 3, MSGT.Body);
        
        ThrowIfNot(sCtl != null, "Can't find HELP announcement");

		string sidControl = BuildAnnName(WebCore._sid_Announcements_Button_Edit_Prefix, WebCore._sid_Announcements_Button_Edit_Suffix, sCtl);

            ThrowIfNot(m_webControl.FClickControlId(sidControl), "Couldn't find edit button");
            m_webControl.WaitForPageLoad();

            // now edit the text
            string sNameControl = BuildAnnName(WebCore._s_Announcements_Textarea_Text_Prefix, WebCore._s_Announcements_Textarea_Text_Suffix, sCtl);

            m_webControl.FSetTextAreaTextForControlName(sNameControl, sArbiterHelpNeeded, true);
            m_webControl.WaitForPageLoad();

            sidControl = BuildAnnName(WebCore._sid_Announcements_Button_Save_Prefix, WebCore._sid_Announcements_Button_Save_Suffix, sCtl);

            ThrowIfNot(m_webControl.FClickControlId(sidControl), "Couldn't find save button");
            m_webControl.WaitForPageLoad();

            // and now save it.

            m_srpt.PopLevel();
            m_srpt.AddMessage("Completed Announcement Set.");
        }

        /* D O  G E N  O P E N  S L O T S  M A I L */
        /*----------------------------------------------------------------------------
        	%%Function: DoGenOpenSlotsMail
        	%%Qualified: ArbWeb.AwMainForm.DoGenOpenSlotsMail
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoGenOpenSlotsMail()
        {
            CountsData gc = GcEnsure(m_pr.RosterWorking, m_pr.GameCopy, m_cbIncludeCanceled.Checked);
            string sTempFile = $"{Environment.GetEnvironmentVariable("Temp")}\\temp{System.Guid.NewGuid().ToString()}.htm";
            Roster rst = RstEnsure(m_pr.RosterWorking);

            string sBcc = m_cbTestEmail.Checked ? "" : rst.SBuildAddressLine(m_ebFilter.Text);
            ;

            Outlook.Application appOlk = (Outlook.Application) Marshal.GetActiveObject("Outlook.Application");

            if (appOlk == null)
                {
                MessageBox.Show("No running instance of outlook!");
                return;
                }

            Outlook.MailItem oNote = appOlk.CreateItem(Outlook.OlItemType.olMailItem);
            // Outlook.MailItem oNote = (Outlook.MailItem)appOlk.CreateItem(Outlook.OlItemType.olMailItem);

            oNote.To = "rlittle@thetasoft.com";
            oNote.BCC = sBcc;
            oNote.Subject = "This is a test";
            oNote.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            oNote.HTMLBody = "<html><style>\r\n*#myId {\ncolor:Blue;\n}\n</style><body><p>Put your preamble here...</p>";

            if (m_cbSplitSports.Checked)
                {
                string[] rgs;

                oNote.HTMLBody += "<h1>Baseball open slots</h1>";
                rgs = WebCore.RgsFromChlbxSport(m_cbFilterSport.Checked, m_chlbxSports, "Softball", false);
                gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
                                      rgs, WebCore.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
                oNote.HTMLBody += SHtmlReadFile(sTempFile) + "<h1>Softball Open Slots</h1>";
                rgs = WebCore.RgsFromChlbxSport(m_cbFilterSport.Checked, m_chlbxSports, "Softball", true);
                gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
                                      rgs, WebCore.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
                oNote.HTMLBody += SHtmlReadFile(sTempFile);
                }
            else
                {
                gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
                                      WebCore.RgsFromChlbx(m_cbFilterSport.Checked, m_chlbxSports),
                                      WebCore.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
                oNote.HTMLBody += SHtmlReadFile(sTempFile);
                }
            oNote.Display(true);

            appOlk = null;
            System.IO.File.Delete(sTempFile);
        }

        private SlotAggr m_saOpenSlots;

        /* D O  C A L C  O P E N  S L O T S */
        /*----------------------------------------------------------------------------
        	%%Function: DoCalcOpenSlots
        	%%Qualified: ArbWeb.AwMainForm.DoCalcOpenSlots
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private async void DoCalcOpenSlots()
        {
            Task<CountsData> taskCalc = new Task<CountsData>(CalcOpenSlotsWork);
            taskCalc.Start();

            CountsData cd = await taskCalc;
            
            m_srpt.PopLevel();
            m_srpt.AddMessage("Updating listboxes...", MSGT.Header, false);
            // update regenerate the listboxes...
            string[] rgsSports = WebCore.RgsFromChlbx(true, m_chlbxSports);
            string[] rgsSportLevels = WebCore.RgsFromChlbx(true, m_chlbxSportLevels);

            bool fCheckAllSports = false;
            bool fCheckAllSportLevels = false;

            if (rgsSports.Length == 0 && m_chlbxSports.Items.Count == 0)
                fCheckAllSports = true;

            if (rgsSports.Length == 0 && m_chlbxSportLevels.Items.Count == 0)
                fCheckAllSportLevels = true;

            WebCore.UpdateChlbxFromRgs(m_chlbxSports, cd.GetOpenSlotSports(m_saOpenSlots), rgsSports, null, fCheckAllSports);
            WebCore.UpdateChlbxFromRgs(m_chlbxSportLevels, cd.GetOpenSlotSportLevels(m_saOpenSlots), rgsSportLevels, fCheckAllSports ? null : rgsSports, fCheckAllSportLevels);
            string[] rgsRosterSites = WebCore.RgsFromChlbx(true, m_chlbxRoster);

            WebCore.UpdateChlbxFromRgs(m_chlbxRoster, cd.GetSiteRosterSites(m_saOpenSlots), rgsRosterSites, null, false);
            m_srpt.PopLevel();
            DoPendingQueueUIOp();
        }

        /* C A L C  O P E N  S L O T S  W O R K */
        /*----------------------------------------------------------------------------
	    	%%Function: CalcOpenSlotsWork
	    	%%Qualified: ArbWeb.AwMainForm.CalcOpenSlotsWork
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
        private CountsData CalcOpenSlotsWork()
        {
            m_srpt.AddMessage("Calculating slot data...", MSGT.Header, false);

            CountsData gc = GcEnsure(m_pr.RosterWorking, m_pr.GameCopy, m_cbIncludeCanceled.Checked);
            Roster rst = RstEnsure(m_pr.RosterWorking);

            m_srpt.PopLevel();
            m_srpt.AddMessage("Calculating open slots...", MSGT.Header, false);
            m_saOpenSlots = gc.CalcOpenSlots(m_dtpStart.Value, m_dtpEnd.Value);
            return gc;
        }

        private void DoGenSiteRosterReport()
        {
            CountsData gc = GcEnsure(m_pr.RosterWorking, m_pr.GameCopy, m_cbIncludeCanceled.Checked);
            string sTempFile = $"{Environment.GetEnvironmentVariable("Temp")}\\temp{System.Guid.NewGuid().ToString()}.doc";
            Roster rst = RstEnsure(m_pr.RosterWorking);

            gc.GenSiteRosterReport(sTempFile, rst, WebCore.RgsFromChlbx(true, m_chlbxRoster), m_dtpStart.Value, m_dtpEnd.Value);
            // launch word with the file
            Process.Start(sTempFile);
            // System.IO.File.Delete(sTempFile);
        }
    }
}