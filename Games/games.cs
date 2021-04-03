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
    public class GameData
    {
        private Roster m_rst;
        private GameSlots m_gms;
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
                m_gms.SetMiscHeadings(m_rst.PlsMiscHeadings);
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
            return m_gms.FLoadGames(sGames, m_rst, fIncludeCanceled);
        }

        /* G E N  R E P O R T */
        /*----------------------------------------------------------------------------
				%%Function: GenReport
				%%Qualified: ArbWeb.CountsData:GameData.GenReport
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public void GenAnalysis(string sReport)
        {
            m_gms.ReduceTeams();
            m_gms.GenReport(sReport);
        }

        public void GenGamesReport(string sReport)
        {
            m_gms.GenGamesReport(sReport);
        }

        /* G E N  O P E N  S L O T S */
        /*----------------------------------------------------------------------------
				%%Function: GenOpenSlots
				%%Qualified: ArbWeb.CountsData:GameData.GenOpenSlots
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public SlotAggr GenOpenSlots(DateTime dttmStart, DateTime dttmEnd)
        {
            return m_gms.GenOpenSlots(dttmStart, dttmEnd);
        }

        public void GenSiteRosterReport(string sReportFile, ArbWeb.Roster rst, string[] rgsRosterFilter, DateTime dttmStart, DateTime dttmEnd)
        {
            m_gms.GenSiteRosterReport(sReportFile, rst, rgsRosterFilter, dttmStart, dttmEnd);
        }
        /* G E N  O P E N  S L O T S */
        /*----------------------------------------------------------------------------
				%%Function: GenOpenSlots
				%%Qualified: ArbWeb.CountsData:GameData.GenOpenSlots
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public void GenOpenSlotsReport(string sReport, bool fDetail, bool fFuzzyTimes, bool fDatePivot, string[] rgsSportFilter, string[] rgsSportLevelFilter, SlotAggr sa)
        {
            m_gms.GenOpenSlotsReport(sReport, fDetail, fFuzzyTimes, fDatePivot, rgsSportFilter, rgsSportLevelFilter, sa);
        }

        /* G A M E S  F R O M  F I L T E R */
        /*----------------------------------------------------------------------------
        	%%Function: GamesFromFilter
        	%%Qualified: ArbWeb.GameData.GamesFromFilter
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public GameSlots GamesFromFilter(string[] rgsSportFilter, string[] rgsSportLevelFilter, bool fOpenOnly, SlotAggr sa)
        {
            return m_gms.GamesFromFilter(rgsSportFilter, rgsSportLevelFilter, fOpenOnly, sa);
        }

        /* G E T  O P E N  S L O T  S P O R T S */
        /*----------------------------------------------------------------------------
				%%Function: GetOpenSlotSports
				%%Qualified: ArbWeb.CountsData:GameData.GetOpenSlotSports
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public string[] GetOpenSlotSports(SlotAggr sa)
        {
            return m_gms.GetOpenSlotSports(sa);
        }

        public string[] GetSiteRosterSites(SlotAggr sa)
        {
            return m_gms.GetSiteRosterSites(sa);
        }

        /* G E T  O P E N  S L O T  S P O R T  L E V E L S */
        /*----------------------------------------------------------------------------
				%%Function: GetOpenSlotSportLevels
				%%Qualified: ArbWeb.CountsData:GameData.GetOpenSlotSportLevels
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public string[] GetOpenSlotSportLevels(SlotAggr sa)
        {
            return m_gms.GetOpenSlotSportLevels(sa);
        }

        /* G E N  G A M E  S T A T S */
        /*----------------------------------------------------------------------------
				%%Function: GameData
				%%Qualified: GenCount.CountsData:GameData.GameData
				%%Contact: rlittle

			----------------------------------------------------------------------------*/
        public GameData(StatusBox srpt)
        {
            //  m_sRoster = null;
            m_srpt = srpt;
            m_rst = new Roster(srpt);
            m_gms = new GameSlots(srpt);
        }
    } // END  GameData

}