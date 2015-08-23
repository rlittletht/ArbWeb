using System;
using System.Diagnostics;
using System.Drawing;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using StatusBox;
using mshtml;
using System.Runtime.InteropServices;
using Outlook=Microsoft.Office.Interop.Outlook;
using System.Threading.Tasks;
using TCore.Settings;

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
            return String.Format("{0}{1}{2}", sPrefix, sCtl, sSuffix);
        }

        /* D O  G E N  M A I L  M E R G E  A N D  A N N O U C E */
        /*----------------------------------------------------------------------------
        	%%Function: DoGenMailMergeAndAnnouce
        	%%Qualified: ArbWeb.AwMainForm.DoGenMailMergeAndAnnouce
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private void DoGenMailMergeAndAnnouce()
        {
            CountsData gc = GcEnsure(m_ebRosterWorking.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
            Roster rst = RstEnsure(m_ebRosterWorking.Text);
            m_srpt.AddMessage("Generating mail merge documents...", StatusRpt.MSGT.Header, false);

            m_srpt.LogData("GamesFromFilter...", 3, StatusRpt.MSGT.Header);

            // first, generate the mailmerge source csv file.  this is either the entire roster, or just the folks 
            // rated for the sports we are filtered to
            GameData.GameSlots gms = gc.GamesFromFilter(ArbWebControl.RgsFromChlbx(m_cbFilterSport.Checked, m_chlbxSports),
                                                        ArbWebControl.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), false, m_saOpenSlots);

            m_srpt.LogData("Beginning to build rosterfiltered...", 3, StatusRpt.MSGT.Body);
            Roster rstFiltered;

            if (m_cbFilterRank.Checked)
                rstFiltered = rst.FilterByRanks(gms.RequiredRanks());
            else
                rstFiltered = rst;

            string sCsvTemp = SBuildTempFilename("MailMergeRoster", "csv");
            m_srpt.LogData(String.Format("Writing filtered roster to {0}", sCsvTemp), 3, StatusRpt.MSGT.Body);
            StreamWriter sw = new StreamWriter(sCsvTemp, false, System.Text.Encoding.Default);

            sw.WriteLine("email,firstname,lastname");
            foreach (RosterEntry rste in rstFiltered.Plrste)
                {
                sw.WriteLine("{0},{1},{2}", rste.Email, rste.First, rste.m_sLast);
                }
            sw.Flush();
            sw.Close();

            // ok, now create the mailmerge .docx
            string sTempName;
            string sArbiterHelpNeeded;
            m_srpt.LogData("Filtered roster written", 3, StatusRpt.MSGT.Body);

            string sApp = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string sFile = Path.Combine(Path.GetPathRoot(sApp),Path.GetDirectoryName(sApp), "mailmergedoc.docx");

            sTempName = SBuildTempFilename("mailmergedoc", "docx");
            m_srpt.LogData(String.Format("Writing mailmergedoc to {0} using template at {1}", sTempName, sFile), 3, StatusRpt.MSGT.Body);
            OOXML.CreateMailMergeDoc(sFile, sTempName, sCsvTemp, gms, out sArbiterHelpNeeded);

            m_srpt.LogData(String.Format("ArbiterHelp HTML created: {0}", sArbiterHelpNeeded), 5, StatusRpt.MSGT.Body);
            System.Windows.Forms.Clipboard.SetText(sArbiterHelpNeeded);
            if (m_cbLaunch.Checked)
                {
                m_srpt.AddMessage("Done, launching document...", StatusRpt.MSGT.Header, false);
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
            ThrowIfNot(m_awc.FNavToPage(_s_Announcements), "Couldn't nav to announcements page!");
            m_awc.FWaitForNavFinish();

            // now we need to find the URGENT HELP NEEDED row
            IHTMLDocument2 oDoc2 = m_awc.Document2;
            IHTMLElementCollection hec = (IHTMLElementCollection) oDoc2.all.tags("div");

            string sCtl = null;

            foreach (IHTMLElement he in hec)
                {
                if (he.id == "D9UrgentHelpNeeded")
                    {
                    m_srpt.LogData("Found D9UrgentHelpNeeded DIV, looking for parent TR element", 3, StatusRpt.MSGT.Body);

                    IHTMLElement heFind = he;
                    while (heFind.tagName.ToLower() != "tr")
                        {
                        heFind = heFind.parentElement;
                        ThrowIfNot(heFind != null, "Can't find HELP announcement");
                        }
                    m_srpt.LogData("Found D9UrgentHelpNeeded parent TR", 3, StatusRpt.MSGT.Body);
                    // ok, go up to the parent TR.
                    // now find one of our controls and get its control number
                    string s = heFind.innerHTML;
                    int ich = s.IndexOf(_s_Announcements_Button_Edit_Prefix);
                    if (ich > 0)
                        {
                        sCtl = s.Substring(ich + _s_Announcements_Button_Edit_Prefix.Length, 5);
                        }
                    m_srpt.LogData(String.Format("Extracted ID for announcment to set: {0}", sCtl), 3, StatusRpt.MSGT.Body);
                    break;
                    }
                }

            ThrowIfNot(sCtl != null, "Can't find HELP announcement");

            m_awc.ResetNav();
            string sControl = BuildAnnName(_sid_Announcements_Button_Edit_Prefix, _sid_Announcements_Button_Edit_Suffix, sCtl);

            ThrowIfNot(m_awc.FClickControl(oDoc2, sControl), "Couldn't find edit button");
            m_awc.FWaitForNavFinish();

            // now edit the text
            sControl = BuildAnnName(_s_Announcements_Textarea_Text_Prefix, _s_Announcements_Textarea_Text_Suffix, sCtl);

            m_awc.FSetTextareaControlText(oDoc2, sControl, sArbiterHelpNeeded, true);
            m_awc.FWaitForNavFinish();

            sControl = BuildAnnName(_sid_Announcements_Button_Save_Prefix, _sid_Announcements_Button_Save_Suffix, sCtl);

            ThrowIfNot(m_awc.FClickControl(oDoc2, sControl), "Couldn't find save button");
            m_awc.FWaitForNavFinish();

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
            CountsData gc = GcEnsure(m_ebRosterWorking.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
            string sTempFile = String.Format("{0}\\temp{1}.htm", Environment.GetEnvironmentVariable("Temp"), System.Guid.NewGuid().ToString());
            Roster rst = RstEnsure(m_ebRosterWorking.Text);

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
                rgs = ArbWebControl.RgsFromChlbxSport(m_cbFilterSport.Checked, m_chlbxSports, "Softball", false);
                gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
                                      rgs, ArbWebControl.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
                oNote.HTMLBody += SHtmlReadFile(sTempFile) + "<h1>Softball Open Slots</h1>";
                rgs = ArbWebControl.RgsFromChlbxSport(m_cbFilterSport.Checked, m_chlbxSports, "Softball", true);
                gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
                                      rgs, ArbWebControl.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
                oNote.HTMLBody += SHtmlReadFile(sTempFile);
                }
            else
                {
                gc.GenOpenSlotsReport(sTempFile, m_cbOpenSlotDetail.Checked, m_cbFuzzyTimes.Checked, m_cbDatePivot.Checked,
                                      ArbWebControl.RgsFromChlbx(m_cbFilterSport.Checked, m_chlbxSports),
                                      ArbWebControl.RgsFromChlbx(m_cbFilterLevel.Checked, m_chlbxSportLevels), m_saOpenSlots);
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
            m_srpt.AddMessage("Updating listboxes...", StatusRpt.MSGT.Header, false);
            // update regenerate the listboxes...
            string[] rgsSports = ArbWebControl.RgsFromChlbx(true, m_chlbxSports);
            string[] rgsSportLevels = ArbWebControl.RgsFromChlbx(true, m_chlbxSportLevels);

            bool fCheckAllSports = false;
            bool fCheckAllSportLevels = false;

            if (rgsSports.Length == 0 && m_chlbxSports.Items.Count == 0)
                fCheckAllSports = true;

            if (rgsSports.Length == 0 && m_chlbxSportLevels.Items.Count == 0)
                fCheckAllSportLevels = true;

            ArbWebControl.UpdateChlbxFromRgs(m_chlbxSports, cd.GetOpenSlotSports(m_saOpenSlots), rgsSports, null, fCheckAllSports);
            ArbWebControl.UpdateChlbxFromRgs(m_chlbxSportLevels, cd.GetOpenSlotSportLevels(m_saOpenSlots), rgsSportLevels, fCheckAllSports ? null : rgsSports, fCheckAllSportLevels);
            string[] rgsRosterSites = ArbWebControl.RgsFromChlbx(true, m_chlbxRoster);

            ArbWebControl.UpdateChlbxFromRgs(m_chlbxRoster, cd.GetSiteRosterSites(m_saOpenSlots), rgsRosterSites, null, false);
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
            m_srpt.AddMessage("Calculating slot data...", StatusRpt.MSGT.Header, false);

            CountsData gc = GcEnsure(m_ebRosterWorking.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
            Roster rst = RstEnsure(m_ebRosterWorking.Text);

            m_srpt.PopLevel();
            m_srpt.AddMessage("Calculating open slots...", StatusRpt.MSGT.Header, false);
            m_saOpenSlots = gc.CalcOpenSlots(m_dtpStart.Value, m_dtpEnd.Value);
            return gc;
        }

        private void DoGenSiteRosterReport()
        {
            CountsData gc = GcEnsure(m_ebRosterWorking.Text, m_ebGameCopy.Text, m_cbIncludeCanceled.Checked);
            string sTempFile = String.Format("{0}\\temp{1}.doc", Environment.GetEnvironmentVariable("Temp"),
                                             System.Guid.NewGuid().ToString());
            Roster rst = RstEnsure(m_ebRosterWorking.Text);

            gc.GenSiteRosterResport(sTempFile, rst, ArbWebControl.RgsFromChlbx(true, m_chlbxRoster), m_dtpStart.Value, m_dtpEnd.Value);
            // launch word with the file
            Process.Start(sTempFile);
            // System.IO.File.Delete(sTempFile);
        }
    }
}