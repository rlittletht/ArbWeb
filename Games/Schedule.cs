using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using ArbWeb.Reports;
using NUnit.Framework;
using TCore.StatusBox;

namespace ArbWeb.Games
{
    public class Schedule
    {
        private Roster m_rst;
        private ScheduleGames m_scheduleGames;
        private StatusBox m_srpt;

        public ScheduleGames Games => m_scheduleGames;

        /*----------------------------------------------------------------------------
			%%Function:FLoadRoster
			%%Qualified:ArbWeb.Games.Schedule.FLoadRoster
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

        /*----------------------------------------------------------------------------
			%%Function:FLoadGames
			%%Qualified:ArbWeb.Games.Schedule.FLoadGames
        ----------------------------------------------------------------------------*/
        public bool FLoadGames(string sGames, bool fIncludeCanceled)
        {
	        GamesLoader_Arbiter loader = new GamesLoader_Arbiter(m_srpt);
	        
            return loader.FLoadGames(sGames, m_rst, fIncludeCanceled, m_scheduleGames);
        }

        /*----------------------------------------------------------------------------
			%%Function:GenAnalysis
			%%Qualified:ArbWeb.Games.Schedule.GenAnalysis
        ----------------------------------------------------------------------------*/
        public void GenAnalysis(string sReport)
        {
            AnalysisReport.GenReport(m_scheduleGames, sReport);
        }

        /*----------------------------------------------------------------------------
			%%Function:GenOpenSlots
			%%Qualified:ArbWeb.Games.Schedule.GenOpenSlots
        ----------------------------------------------------------------------------*/
        public SlotAggr GenOpenSlots(DateTime dttmStart, DateTime dttmEnd)
        {
            return m_scheduleGames.GenOpenSlots(dttmStart, dttmEnd);
        }
        
        /*----------------------------------------------------------------------------
			%%Function:GamesFromFilter
			%%Qualified:ArbWeb.Games.Schedule.GamesFromFilter
        ----------------------------------------------------------------------------*/
        public ScheduleGames GamesFromFilter(string[] rgsSportFilter, string[] rgsSportLevelFilter, bool fOpenOnly, SlotAggr sa)
        {
            return m_scheduleGames.GamesFromFilter(rgsSportFilter, rgsSportLevelFilter, fOpenOnly, sa);
        }

        /*----------------------------------------------------------------------------
			%%Function:GetOpenSlotSports
			%%Qualified:ArbWeb.Games.Schedule.GetOpenSlotSports
        ----------------------------------------------------------------------------*/
        public string[] GetOpenSlotSports(SlotAggr sa)
        {
            return m_scheduleGames.GetOpenSlotSports(sa);
        }

        /*----------------------------------------------------------------------------
			%%Function:GetSiteRosterSites
			%%Qualified:ArbWeb.Games.Schedule.GetSiteRosterSites
        ----------------------------------------------------------------------------*/
        public string[] GetSiteRosterSites(SlotAggr sa)
        {
            return m_scheduleGames.GetSiteRosterSites(sa);
        }

        /*----------------------------------------------------------------------------
			%%Function:GetOpenSlotSportLevels
			%%Qualified:ArbWeb.Games.Schedule.GetOpenSlotSportLevels
        ----------------------------------------------------------------------------*/
        public string[] GetOpenSlotSportLevels(SlotAggr sa)
        {
            return m_scheduleGames.GetOpenSlotSportLevels(sa);
        }

        /*----------------------------------------------------------------------------
			%%Function:Schedule
			%%Qualified:ArbWeb.Games.Schedule.Schedule
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