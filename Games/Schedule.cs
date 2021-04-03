using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using NUnit.Framework;
using TCore.StatusBox;

namespace ArbWeb.Games
{
    // ================================================================================
    //  G A M E  D A T A 
    // ================================================================================
    public class Schedule
    {
        private Roster m_rst;
        private ScheduleGames m_scheduleGames;
        private StatusBox m_srpt;

        public string SMiscHeader(int i)
        {
            return m_rst.SMiscHeader(i);
        }

        /* L O A D  R O S T E R */
        /*----------------------------------------------------------------------------
				%%Function: LoadRoster
				%%Qualified: GenCount.CountsData:GameData.LoadRoster
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public bool FLoadRoster(string sRoster, int iMiscAffiliation)
        {
            m_rst = new Roster(m_srpt);
            bool f = m_rst.LoadRoster(sRoster, iMiscAffiliation);

            if (f)
                {
                m_scheduleGames.SetMiscHeadings(m_rst.PlsMiscHeadings);
                }
            return f;
        }

        /* F  L O A D  G A M E S */
        /*----------------------------------------------------------------------------
				%%Function: FLoadGames
				%%Qualified: ArbWeb.CountsData:GameData.FLoadGames
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public bool FLoadGames(string sGames, bool fIncludeCanceled)
        {
	        GamesLoader_Arbiter loader = new GamesLoader_Arbiter(m_srpt);
	        
            return loader.FLoadGames(sGames, m_rst, fIncludeCanceled, m_scheduleGames);
        }

        /* G E N  R E P O R T */
        /*----------------------------------------------------------------------------
				%%Function: GenReport
				%%Qualified: ArbWeb.CountsData:GameData.GenReport
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public void GenAnalysis(string sReport)
        {
            m_scheduleGames.ReduceTeams();
            m_scheduleGames.GenReport(sReport);
        }

        public void GenGamesReport(string sReport)
        {
            m_scheduleGames.GenGamesReport(sReport);
        }

        /* G E N  O P E N  S L O T S */
        /*----------------------------------------------------------------------------
				%%Function: GenOpenSlots
				%%Qualified: ArbWeb.CountsData:GameData.GenOpenSlots
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public SlotAggr GenOpenSlots(DateTime dttmStart, DateTime dttmEnd)
        {
            return m_scheduleGames.GenOpenSlots(dttmStart, dttmEnd);
        }

        public void GenSiteRosterReport(string sReportFile, ArbWeb.Roster rst, string[] rgsRosterFilter, DateTime dttmStart, DateTime dttmEnd)
        {
            m_scheduleGames.GenSiteRosterReport(sReportFile, rst, rgsRosterFilter, dttmStart, dttmEnd);
        }
        /* G E N  O P E N  S L O T S */
        /*----------------------------------------------------------------------------
				%%Function: GenOpenSlots
				%%Qualified: ArbWeb.CountsData:GameData.GenOpenSlots
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public void GenOpenSlotsReport(string sReport, bool fDetail, bool fFuzzyTimes, bool fDatePivot, string[] rgsSportFilter, string[] rgsSportLevelFilter, SlotAggr sa)
        {
            m_scheduleGames.GenOpenSlotsReport(sReport, fDetail, fFuzzyTimes, fDatePivot, rgsSportFilter, rgsSportLevelFilter, sa);
        }

        /* G A M E S  F R O M  F I L T E R */
        /*----------------------------------------------------------------------------
        	%%Function: GamesFromFilter
        	%%Qualified: ArbWeb.GameData.GamesFromFilter
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public ScheduleGames GamesFromFilter(string[] rgsSportFilter, string[] rgsSportLevelFilter, bool fOpenOnly, SlotAggr sa)
        {
            return m_scheduleGames.GamesFromFilter(rgsSportFilter, rgsSportLevelFilter, fOpenOnly, sa);
        }

        /* G E T  O P E N  S L O T  S P O R T S */
        /*----------------------------------------------------------------------------
				%%Function: GetOpenSlotSports
				%%Qualified: ArbWeb.CountsData:GameData.GetOpenSlotSports
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public string[] GetOpenSlotSports(SlotAggr sa)
        {
            return m_scheduleGames.GetOpenSlotSports(sa);
        }

        public string[] GetSiteRosterSites(SlotAggr sa)
        {
            return m_scheduleGames.GetSiteRosterSites(sa);
        }

        /* G E T  O P E N  S L O T  S P O R T  L E V E L S */
        /*----------------------------------------------------------------------------
				%%Function: GetOpenSlotSportLevels
				%%Qualified: ArbWeb.CountsData:GameData.GetOpenSlotSportLevels
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public string[] GetOpenSlotSportLevels(SlotAggr sa)
        {
            return m_scheduleGames.GetOpenSlotSportLevels(sa);
        }

        /* G E N  G A M E  S T A T S */
        /*----------------------------------------------------------------------------
				%%Function: GameData
				%%Qualified: GenCount.CountsData:GameData.GameData
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public Schedule(StatusBox srpt)
        {
            //  m_sRoster = null;
            m_srpt = srpt;
            m_rst = new Roster(srpt);
            m_scheduleGames = new ScheduleGames(srpt);
        }
    } // END  GameData

}