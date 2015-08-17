using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;

using Microsoft.Win32;
using System.Collections;
using System.Net;
using AxSHDocVw;
using mshtml;
using Microsoft.Office;
using System.Runtime.InteropServices;
using StatusBox;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArbWeb
	{
	// ================================================================================
	//  C S V 
	// ================================================================================
	class Csv
	{
		public static string[] LineToArray(string line)
		{
			String pattern = ",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))";
			Regex r = new Regex(pattern);

			string[] rgs = r.Split(line);

			for (int i = 0; i < rgs.Length; i++)
				{
				if (rgs[i].Length > 0 && rgs[i][0] == '"')
					rgs[i] = rgs[i].Substring(1, rgs[i].Length - 2);
				}
			return rgs;
	   }
	};

	// ================================================================================
	//  C O U N T S  D A T A 
	// ================================================================================
	public partial class CountsData
		{
	    public static class RexHelper
	    {
	        public static string[] RgsMatch(string sInput, string sRex)
	        {
	            Regex rex = new Regex(sRex);
	            Match m = rex.Match(sInput);
	            int c = m.Groups.Count - 1;

	            string[] rgs = new string[m.Groups.Count - 1];
	            int i;
	            for (i = 1; i <= c; i++)
	                rgs[i - 1] = m.Groups[i].Value;

	            return rgs;
	        }
	    }

        StatusBox.StatusRpt m_srpt;

		/* G E N  C O U N T S */
		/*----------------------------------------------------------------------------
			%%Function: CountsData
			%%Qualified: GenCount.CountsData.CountsData
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
        public CountsData(StatusBox.StatusRpt srpt)
		{
			m_srpt = srpt;
		}

		GameData m_gmd;

		/* L O A D  D A T A */
		/*----------------------------------------------------------------------------
			%%Function: LoadData
			%%Qualified: ArbWeb.CountsData.LoadData
			%%Contact: rlittle

		----------------------------------------------------------------------------*/
		public void LoadData(string sRoster, string sSource, bool fIncludeCanceled, int iMiscAffiliation)
		{
//            srpt.UnitTest();
			m_gmd = new GameData(m_srpt);
		    m_srpt.AddMessage("Loading roster...", StatusRpt.MSGT.Header, false);

			m_gmd.FLoadRoster(sRoster, iMiscAffiliation);
            // m_srpt.AddMessage(String.Format("Using plsMisc[{0}] ({1}) for team affiliation", iMiscAffiliation, m_gmd.SMiscHeader(iMiscAffiliation)), StatusBox.StatusRpt.MSGT.Body);

		    m_srpt.PopLevel();
		    m_srpt.AddMessage("Loading games...", StatusRpt.MSGT.Header, false);
			m_gmd.FLoadGames(sSource, fIncludeCanceled);
		    m_srpt.PopLevel();
			// read in the roster of umpires...
		}

        public void GenSiteRosterResport(string sReportFile, Roster rst, string[] rgsRosterFilter, DateTime dttmStart, DateTime dttmEnd)
        {
            m_gmd.GenSiteRosterReport(sReportFile, rst, rgsRosterFilter, dttmStart, dttmEnd);
        }
		/* G E N  O P E N  S L O T S  R E P O R T */
		/*----------------------------------------------------------------------------
			%%Function: GenOpenSlotsReport
			%%Qualified: ArbWeb.CountsData.GenOpenSlotsReport
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
		public void GenOpenSlotsReport(string sReport, bool fDetail, bool fFuzzyTimes, bool fDatePivot, string[] rgsSportFilter, string[] rgsSportLevelFilter, SlotAggr sa)
		{
			m_gmd.GenOpenSlotsReport(sReport, fDetail, fFuzzyTimes, fDatePivot, rgsSportFilter, rgsSportLevelFilter, sa);
		}

        /* G A M E S  F R O M  F I L T E R */
        /*----------------------------------------------------------------------------
        	%%Function: GamesFromFilter
        	%%Qualified: ArbWeb.CountsData.GamesFromFilter
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public GameData.GameSlots GamesFromFilter(string[] rgsSportFilter, string[] rgsSportLevelFilter, bool fOpenOnly, SlotAggr sa)
        {
            return m_gmd.GamesFromFilter(rgsSportFilter, rgsSportLevelFilter, fOpenOnly, sa);
        }

		/* C A L C  O P E N  S L O T S */
		/*----------------------------------------------------------------------------
			%%Function: CalcOpenSlots
			%%Qualified: ArbWeb.CountsData.CalcOpenSlots
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
		public SlotAggr CalcOpenSlots(DateTime dttmStart, DateTime dttmEnd)
		{
			return m_gmd.GenOpenSlots(dttmStart, dttmEnd);
		}

		/* G E T  O P E N  S L O T  S P O R T S */
		/*----------------------------------------------------------------------------
			%%Function: GetOpenSlotSports
			%%Qualified: ArbWeb.CountsData.GetOpenSlotSports
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
		public string[] GetOpenSlotSports(SlotAggr sa)
		{
			return m_gmd.GetOpenSlotSports(sa);
		}

        public string[] GetSiteRosterSites(SlotAggr sa)
        {
            return m_gmd.GetSiteRosterSites(sa);
        }

		/* G E T  O P E N  S L O T  S P O R T  L E V E L S */
		/*----------------------------------------------------------------------------
			%%Function: GetOpenSlotSportLevels
			%%Qualified: ArbWeb.CountsData.GetOpenSlotSportLevels
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
		public string[] GetOpenSlotSportLevels(SlotAggr sa)
		{
			return m_gmd.GetOpenSlotSportLevels(sa);
		}

		/* G E N  R E P O R T */
		/*----------------------------------------------------------------------------
			%%Function: GenReport
			%%Qualified: ArbWeb.CountsData.GenReport
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
		public void GenAnalysis(string sReport)
		{
			m_gmd.GenAnalysis(sReport);
		}

		/* G E N  G A M E S  R E P O R T */
		/*----------------------------------------------------------------------------
			%%Function: GenGamesReport
			%%Qualified: ArbWeb.CountsData.GenGamesReport
			%%Contact: rlittle
			
		----------------------------------------------------------------------------*/
		public void GenGamesReport(string sReport)
		{
			m_gmd.GenGamesReport(sReport);
		}
		}
	}
