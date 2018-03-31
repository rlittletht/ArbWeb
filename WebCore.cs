using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using mshtml;
using StatusBox;
using TCore.Util;
using Win32Win;

namespace ArbWeb
{
    class WebCore
    {
        #region ArbiterStrings

        // ============================================================================
        // T O P  L E V E L    
        // ============================================================================
        public const string _s_Home = "https://www.arbitersports.com/Shared/SignIn/Signin.aspx"; // ok2010
        public const string _s_Assigning = "https://www.arbitersports.com/Assigner/Games/NewGamesView.aspx"; // ok2010
        public const string _s_RanksEdit = "https://www.arbitersports.com/Assigner/RanksEdit.aspx"; // ok2010
        public const string _s_AddUser = "https://www.arbitersports.com/Assigner/UserAdd.aspx?userTypeID=3"; // ok2010u
        private const string _s_OfficialsView = "https://www.arbitersports.com/Assigner/OfficialsView.aspx"; // ok2010u
        public const string _s_Announcements = "https://www.arbitersports.com/Shared/AnnouncementsEdit.aspx"; // ok2015
        public const string _s_ContactsView = "https://www.arbitersports.com/Assigner/ContactsView.aspx"; // ok2018

        // ============================================================================
        // D I R E C T  A C C E S S
        // ============================================================================
        public const string _s_Assigning_PrintAddress = "https://www.arbitersports.com/Assigner/Games/Print.aspx?filterID="; // ok2010
        public const string _s_EditUser_MiscFields = "https://www.arbitersports.com/Official/MiscFieldsEdit.aspx?userID="; // ok2010
        public const string _s_EditUser = "https://www.arbitersports.com/Official/OfficialEdit.aspx?userID="; // ok2010u

        // ============================================================================
        // H O M E
        // ============================================================================
        private const string _s_Home_Anchor_Login = "SignInButton"; // ctl00$ucMiniLogin$SignInButton"; // ok2010
        public const string _s_Home_Input_Email = "ctl00$ContentHolder$pgeSignIn$conSignIn$txtEmail"; // ctl00$ucMiniLogin$UsernameTextBox"; // ok2016
        public const string _s_Home_Input_Password = "txtPassword"; // ctl00$ucMiniLogin$PasswordTextBox"; // ok2016
        public const string _s_Home_Button_SignIn = "ctl00$ContentHolder$pgeSignIn$conSignIn$btnSignIn"; // ctl00$ucMiniLogin$SignInButton"; // ok2016

        public const string _sid_Home_Div_PnlAccounts = "ctl00_ContentHolder_pgeDefault_conDefault_pnlAccounts"; // ok2010
        public const string _sid_Home_Anchor_NeedHelpLink = "ctl00_PageHelpTextLink"; // ok2017

        // ============================================================================
        // A S S I G N I N G
        // ============================================================================
        // (games view) links
        public const string _s_Assigning_Select_Filters = "ctl00$ContentHolder$pgeGamesView$conGamesView$ddlSavedFilters"; // ok2010
        public const string _s_Assigning_Reports_Select_Format = "ctl00$ContentHolder$pgePrint$conPrint$ddlFormat"; // ok2010
        public const string _s_Assigning_Reports_Submit_Print = "ctl00$ContentHolder$pgePrint$navPrint$btnBeginPrint"; // ok2010

        public const string _sid_Assigning_Reports_Select_Format = "ctl00_ContentHolder_pgePrint_conPrint_ddlFormat"; // ok2010
        

        // ============================================================================
        // A D D  U S E R
        // ============================================================================
        public const string _s_AddUser_Input_FirstName = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$txtFirstName"; // ok2010
        public const string _s_AddUser_Input_LastName = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$txtLastName"; // ok2010
        public const string _s_AddUser_Input_Email = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$txtEmail"; // ok2010
        public const string _s_AddUser_Input_Address1 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$uclAddress$address_txtAddress1"; // ok2010
        private const string _s_AddUser_Input_Address2 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$uclAddress$address_txtAddress2"; // ok2010
        public const string _s_AddUser_Input_City = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$uclAddress$address_txtCity"; // ok2010
        public const string _s_AddUser_Input_State = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$uclAddress$address_txtState"; // ok2010
        public const string _s_AddUser_Input_Zip = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$uclAddress$address_txtPostalCode"; // ok2010
        public const string _sid_AddUser_Input_Zip = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_uclAddress_address_txtPostalCode"; // ok2010

        public const string _s_AddUser_Input_PhoneNum1 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl00$txtPhone"; // ok2018
        public const string _s_AddUser_Input_PhoneNum2 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl01$txtPhone"; // ok2018
        public const string _s_AddUser_Input_PhoneNum3 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl02$txtPhone"; // ok2018
        public const string _s_AddUser_Input_PhoneType1 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl00$ddlPhone"; // ok2018
        public const string _s_AddUser_Input_PhoneType2 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl01$ddlPhone"; // ok2018
        public const string _s_AddUser_Input_PhoneType3 = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$rptPhones$ctl02$ddlPhone"; // ok2018
        private const string _sid_AddUser_Input_PhoneType1 = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_rptPhones_ctl00_ddlPhone"; // ok2010a
        private const string _sid_AddUser_Input_PhoneType2 = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_rptPhones_ctl01_ddlPhone"; // ok2010a
        private const string _sid_AddUser_Input_PhoneType3 = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_rptPhones_ctl02_ddlPhone"; // ok2010a

        public const string _s_AddUser_Input_Country = "ctl00$ContentHolder$pgeUserAdd$conUserAdd$uclAddress$address_ddlCountryType"; // ok2018
        public const string _sid_AddUser_Input_Country = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_uclAddress_address_ddlCountryType"; // ok2018

        // ============================================================================
        // U S E R  I N F O
        // ============================================================================
        public const string _s_EditUser_Email = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclEmails$rptEmail$ctl01$txtEmail"; // ok2018
        public const string _s_EditUser_FirstName = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtFirstName"; // ok2018
        public const string _s_EditUser_LastName = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtLastName"; // ok2018
        public const string _s_EditUser_Address1 = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclAddress$address_txtAddress1"; // ok2018
        public const string _s_EditUser_Address2 = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclAddress$address_txtAddress2"; // ok2018
        public const string _s_EditUser_City = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtCity"; // ok2016
        public const string _s_EditUser_State = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtState"; // ok2016
        public const string _s_EditUser_PostalCode = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclAddress$address_txtPostalCode"; // ok2016
        public const string _s_EditUser_OfficialNumber = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtOfficialNumber"; // ok2016
        public const string _s_EditUser_DateOfBirth = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtDateOfBirth"; // ok2016
        public const string _s_EditUser_DateJoined = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtDateJoined"; // ok2016
        public const string _s_EditUser_GamesPerDay = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtGamesPerDay"; // ok2016
        public const string _s_EditUser_GamesPerWeek = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtGamesPerWeek"; // ok2016
        public const string _s_EditUser_GamesTotal = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtGamesTotal"; // ok2016
        public const string _s_EditUser_WaitMinutes = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtWaitMinutes"; // ok2016

        // ============================================================================
        // P H O N E  C O N T R O L S
        // ============================================================================
        // format of control id is "ctl##" where ## is the line number starting at 1
        public const string _s_EditUser_PhoneNumber_Prefix = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclPhones$rptPhone$"; // ok2016
        public const string _s_EditUser_PhoneNumber_Suffix = "$txtPhone"; // ok2016
        public const string _sid_EditUser_PhoneNumber_Prefix = "ctl00_ContentHolder_pgeOfficialEdit_conOfficialEdit_uclPhones_rptPhone_"; // ok 2016
        public const string _sid_EditUser_PhoneNumber_Suffix = "_txtPhone"; // ok 2016
        public const string _s_EditUser_PhoneType_Prefix = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclPhones$rptPhone$"; // ok2016
        public const string _s_EditUser_PhoneType_Suffix = "$ddlPhoneType"; // ok2016
        private const string _sid_EditUser_PhoneType_Prefix = "ctl00_ContentHolder_pgeOfficialEdit_conOfficialEdit_uclPhones_rptPhone_"; // ok2016
        private const string _sid_EditUser_PhoneType_Suffix = "_ddlPhoneType"; // ok 2016
        public const string _s_EditUser_PhoneCarrier_Prefix = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclPhones$rptPhone$"; // ok2016
        public const string _s_EditUser_PhoneCarrier_Suffix = "$ddlCarrier"; // ok2016
        private const string _sid_EditUser_PhoneCarrier_Prefix = "ctl00_ContentHolder_pgeOfficialEdit_conOfficialEdit_uclPhones_rptPhone_"; // ok 2016
        private const string _sid_EditUser_PhoneCarrier_Suffix = "_ddlCarrier";

        public const string _s_EditUser_PhonePublic_Prefix = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclPhones$rptPhone$";
        public const string _s_EditUser_PhonePublic_Suffix = "$chkPublic";
        private const string _sid_EditUser_PhonePublic_Prefix = "ctl00_ContentHolder_pgeOfficialEdit_conOfficialEdit_uclPhones_rptPhone_";
        private const string _sid_EditUser_PhonePublic_Suffix = "_chkPublic";

        public const string _sid_EditUser_PhoneNumber_AddNew = "ctl00_ContentHolder_pgeOfficialEdit_conOfficialEdit_uclPhones_rptPhone_ctl00_lnkAdd"; // ok2016

        public const string _sid_MiscFields_Button_Save = "ctl00_ContentHolder_pgeMiscFieldsEdit_navMiscFieldsEdit_btnSave"; // ok2010u
        public const string _sid_MiscFields_Button_Cancel = "ctl00_ContentHolder_pgeMiscFieldsEdit_navMiscFieldsEdit_lnkCancel"; // ok2010u

        // phone types are Home, Work, Fax, Cellular, Pager, Security, Other
        public const string _sid_AddUser_Button_Next = "ctl00_ContentHolder_pgeUserAdd_navUserAdd_btnNext"; // ok2010
        public const string _sid_AddUser_Input_Address1 = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_uclAddress_address_txtAddress1"; // ok2018
        public const string _sid_AddUser_Input_IsActive = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_chkIsActive"; // ok2010
        public const string _sid_AddUser_Button_Cancel = "ctl00_ContentHolder_pgeUserAdd_navUserAdd_lnkCancel"; // ok2010

        // ============================================================================
        // O F F I C I A L S  V I E W
        // ============================================================================
        public const string _s_Page_OfficialsView = "https://www.arbitersports.com/Assigner/OfficialsView.aspx"; // ok2010
        public const string _s_OfficialsView_Select_Filter = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$ddlFilter"; // ok2010
        public const string _sid_OfficialsView_Select_Filter = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_ddlFilter"; // ok2010
        
        public const string _sid_OfficialsView_PrintRoster = "ctl00_ContentHolder_pgeOfficialsView_sbrReports_tskPrint"; // ok2010u
        public const string _sid_OfficialsView_ContentTable = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_dgOfficials"; // ok2013

        public const string _s_OfficialsView_PaginationHrefPostbackSubstr = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$dgOfficials$ctl204$ctl"; // ok2014

        // ============================================================================
        // O F F I C I A L S  E D I T
        // ============================================================================
        public const string _sid_OfficialsEdit_Button_Save = "ctl00_ContentHolder_pgeOfficialEdit_navOfficialEdit_btnSave"; // ok2010u
        public const string _sid_OfficialsEdit_Button_Cancel = "ctl00_ContentHolder_pgeOfficialEdit_navOfficialEdit_lnkCancel"; // ok2010u

        // ============================================================================
        // R O S T E R  P R I N T
        // ============================================================================
        public const string _sid_RosterPrint_MergeStyle = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkMerge"; // ok2010u 

        public const string _s_RosterPrint_OfficialNumber = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$chkOfficialNumber"; // ok2010
        public const string _s_RosterPrint_DateJoined = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$chkDateJoined"; // ok2010
        public const string _s_RosterPrint_MiscFields = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$chkMiscFields"; // ok2010
        public const string _s_RosterPrint_NonPublicPhone = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$chkIncludeNonPublic"; // ok2010
        public const string _s_RosterPrint_NonPublicAddress = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$chkIncludeNonPublicAddress"; // ok2010

        private const string _sid_RosterPrint_OfficialNumber = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkOfficialNumber"; // ok2010
        public const string _sid_RosterPrint_DateJoined = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkDateJoined"; // ok2010
        private const string _sid_RosterPrint_MiscFields = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkMiscFields"; // ok2010
        private const string _sid_RosterPrint_NonPublicPhone = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkIncludeNonPublic"; // ok2010
        private const string _sid_RosterPrint_NonPublicAddress = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_chkIncludeNonPublicAddress"; // ok2010

        public const string _sid_RosterPrint_BeginPrint = "ctl00_ContentHolder_pgeOfficialsView_navOfficialsView_btnBeginPrint"; // ok2010

        // ============================================================================
        // R A N K S
        // ============================================================================
        public const string _s_RanksEdit_Select_PosNames = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$ddlPosNames"; // ok2010
        public const string _sid_RanksEdit_Select_PosNames = "ctl00_ContentHolder_pgeRanksEdit_conRanksEdit_ddlPosNames"; // ok2010

        
        public const string _s_RanksEdit_Checkbox_Active = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$chkActive"; // ok2010
        public const string _s_RanksEdit_Checkbox_Rank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$chkRank"; // ok2010
        public const string _s_RanksEdit_Select_NotRanked = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$lstNotRanked"; // ok2010
        public const string _s_RanksEdit_Select_Ranked = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$lstRanked"; // ok2010

        public const string _s_RanksEdit_Button_Unrank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$btnUnrank"; // ok2010
        public const string _s_RanksEdit_Button_ReRank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$btnReRank"; // ok2010
        public const string _s_RanksEdit_Button_Rank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$btnRank"; // ok2010
        public const string _s_RanksEdit_Input_Rank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$txtRank"; // ok2010

        // ============================================================================
        // A N N O U N C E M E N T S
        // ============================================================================
        public const string _s_Announcements_Button_Edit_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$"; // the control will be "ctl##"
        private const string _s_Announcements_Button_Edit_Suffix = "$btnEdit";

        private const string _s_Announcements_Checkbox_Assigners_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$";
        // the control will be "ctl##"

        private const string _s_Announcements_Checkbox_Assigners_Suffix = "$chkToAssigners_Edit";

        private const string _s_Announcements_Checkbox_Contacts_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$";
        // the control will be "ctl##"

        private const string _s_Announcements_Checkbox_Contacts_Suffix = "$chkToContacts_Edit";

        private const string _s_Announcements_Select_Filters_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$";
        // the control will be "ctl##"

        private const string _s_Announcements_Select_Filters_Suffix = "$ddlFilters";
        private const string _s_Announcements_Button_Save_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$"; // the control will be "ctl##"
        private const string _s_Announcements_Button_Save_Suffix = "$btnSave";

        public const string _s_Announcements_Textarea_Text_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$";
        public const string _s_Announcements_Textarea_Text_Suffix = "$txtAnnouncement";

        public const string _sid_Announcements_Button_Edit_Prefix = "ctl00_ContentHolder_pgeAnnouncementsEdit_conAnnouncementsEdit_dgAnnouncements_";
        public const string _sid_Announcements_Button_Edit_Suffix = "_btnEdit";

        public const string _sid_Announcements_Button_Save_Prefix = "ctl00_ContentHolder_pgeAnnouncementsEdit_conAnnouncementsEdit_dgAnnouncements_";
        public const string _sid_Announcements_Button_Save_Suffix = "_btnSave";

        public const string _sid_Login_Span_Type_Prefix = "ctl00_ContentHolder_pgeDefault_conDefault_dgAccounts_";
        private const string _sid_Login_Span_Type_Suffix = "_lblType2";

        public const string _sid_Login_Anchor_TypeLink_Prefix = "ctl00_ContentHolder_pgeDefault_conDefault_dgAccounts_";
        public const string _sid_Login_Anchor_TypeLink_Suffix = "_UserTypeLink";

        public const string s_MiscField_EditControlSubstring = "txtMiscFieldValue";

        // ============================================================================
        // C O N T A C T S
        // ============================================================================
        public const string _sid_Contacts_Anchor_ContactsReport = "ctl00_ContentHolder_pgeContactsView_cmnReports_tskPrint"; //ok2018

        // contacts roster page
        public const string _s_Contacts_Roster_Select_Format = "ctl00$ContentHolder$pgeContactsView$conContactsView$ddlFormat"; // ok2018
        public const string _sid_Contacts_Roster_Select_Format = "ctl00_ContentHolder_pgeContactsView_conContactsView_ddlFormat"; // ok2018

        public const string _s_Contacts_Roster_Check_Address = "ctl00$ContentHolder$pgeContactsView$conContactsView$chkAddress"; // ok2018
        public const string _sid_Contacts_Roster_Check_Address = "ctl00_ContentHolder_pgeContactsView_conContactsView_chkAddress"; // ok2018
        public const string _s_Contacts_Roster_Check_Phone = "ctl00$ContentHolder$pgeContactsView$conContactsView$chkPhoneNumber"; // ok2018
        public const string _sid_Contacts_Roster_Check_Phone = "ctl00_ContentHolder_pgeContactsView_conContactsView_chkPhoneNumber"; // ok2018
        public const string _s_Contacts_Roster_Check_Email = "ctl00$ContentHolder$pgeContactsView$conContactsView$chkEmail"; // ok2018
        public const string _sid_Contacts_Roster_Check_Email = "ctl00_ContentHolder_pgeContactsView_conContactsView_chkEmail"; // ok2018
        public const string _s_Contacts_Roster_Check_PageHeader = "ctl00$ContentHolder$pgeContactsView$conContactsView$chkPageHeader"; // ok2018
        public const string _sid_Contacts_Roster_Check_PageHeader = "ctl00_ContentHolder_pgeContactsView_conContactsView_chkPageHeader"; // ok2018
        public const string _s_Contacts_Roster_Check_Site = "ctl00$ContentHolder$pgeContactsView$conContactsView$chkSite"; // ok2018
        public const string _sid_Contacts_Roster_Check_Site = "ctl00_ContentHolder_pgeContactsView_conContactsView_chkSite"; // ok2018
        public const string _s_Contacts_Roster_Check_Team = "ctl00$ContentHolder$pgeContactsView$conContactsView$chkTeam"; // ok2018
        public const string _sid_Contacts_Roster_Check_Team = "ctl00_ContentHolder_pgeContactsView_conContactsView_chkTeam"; // ok2018

        public const string _s_Contacts_Roster_Submit_Print = "ctl00$ContentHolder$pgeContactsView$navContactsView$btnPrint"; // ok2018
        public const string _sid_Contacts_Roster_Submit_Print = "ctl00_ContentHolder_pgeContactsView_navContactsView_btnPrint"; // ok2018

        #endregion
    }

    public class DownloadGenericExcelReport
    {
        private readonly string m_sFilterReq;
        private readonly string m_sDescription;
        private readonly IAppContext m_iac;
        private readonly string m_sReportPage;
        private readonly string m_sSelectFilterControlName;
        private readonly string m_sReportPrintPagePrefix;
        private readonly string m_sReportPrintSubmitPrintControlName;
        private readonly string m_sFullExpectedName;
        private readonly string m_sExpectedName;
        private readonly string m_sidReportPageLink;
        private string m_sGameFile;
        private readonly string m_sGameCopy;

        public class ControlSetting<T>
        {
            private string m_sControlName;
            private string m_sidControlExtra;   // this is usually something like the Choice element ID
            private T m_tControlValue;

            public string ControlName => m_sControlName;
            public string IdControlExtra => m_sidControlExtra;
            public T ControlValue => m_tControlValue;

            public ControlSetting(string sSelectControlName, string sidChoiceControl, T sSelectValue)
            {
                m_sControlName = sSelectControlName;
                m_sidControlExtra = sidChoiceControl;
                m_tControlValue = sSelectValue;
            }

            public ControlSetting(string sControlName, T sControlValue)
            {
                m_sControlName = sControlName;
                m_tControlValue = sControlValue;
            }
        }

        private readonly ControlSetting<bool>[] m_rgCheckedSettings;
        private readonly ControlSetting<string>[] m_rgSelectSettings;

        // this version selects a filter
        public DownloadGenericExcelReport(
            string sFilterReq, 
            string sDescription, 
            string sReportPage, 
            string sSelectFilterControlName, 
            string sReportPrintPagePrefix, 
            string sReportPrintSubmitPrintControlName, 
            string sFullExpectedName, 
            string sExpectedName, 
            ControlSetting<string>[] rgSelectSettings,
            string sGameFile,
            string sGameCopy,
            IAppContext iac)
        {
            m_sFilterReq = sFilterReq;
            m_sDescription = sDescription;
            m_iac = iac;
            m_sReportPage = sReportPage;
            m_sSelectFilterControlName = sSelectFilterControlName;
            m_sReportPrintPagePrefix = sReportPrintPagePrefix;
            m_sReportPrintSubmitPrintControlName = sReportPrintSubmitPrintControlName;
            m_sFullExpectedName = sFullExpectedName;
            m_sExpectedName = sExpectedName;
            m_rgSelectSettings = rgSelectSettings;
            m_sGameFile = sGameFile;
            m_sGameCopy = sGameCopy;
        }

        // this version does not set a filter, it just goes to the start page, clicks a link, then sets the report params
        // this version selects a filter
        public DownloadGenericExcelReport(
            string sDescription,
            string sReportPage,
            string sidReportPageLink,
            string sReportPrintSubmitPrintControlName,
            string sFullExpectedName,
            string sExpectedName,
            ControlSetting<string>[] rgSelectSettings,
            ControlSetting<bool>[] rgCheckedSettings,
            string sGameFile,
            string sGameCopy,
            IAppContext iac)
        {
            m_sDescription = sDescription;
            m_iac = iac;
            m_sReportPage = sReportPage;
            m_sidReportPageLink = sidReportPageLink;
            m_sReportPrintSubmitPrintControlName = sReportPrintSubmitPrintControlName;
            m_sFullExpectedName = sFullExpectedName;
            m_sExpectedName = sExpectedName;
            m_rgSelectSettings = rgSelectSettings;
            m_rgCheckedSettings = rgCheckedSettings;
            m_sGameFile = sGameFile;
            m_sGameCopy = sGameCopy;
        }

        /*----------------------------------------------------------------------------
        	%%Function: DownloadGames
        	%%Qualified: ArbWeb.AwMainForm.DownloadGames
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public void DownloadGeneric(out string sGameFileNew)
        {
            m_iac.StatusReport.AddMessage($"Starting {m_sDescription} download...");
            m_iac.StatusReport.PushLevel();
            string sTempFile = Filename.SBuildTempFilename("temp", "xls");

            sTempFile = DownloadGenericToFile(sTempFile);
            HandleDownloadGames(sTempFile);

            System.IO.File.Delete(sTempFile);

            // ok, now we have all games selected...
            // time to try to download a report
            m_iac.StatusReport.PopLevel();
            m_iac.StatusReport.AddMessage($"Completed downloading {m_sDescription}.");
            sGameFileNew = m_sGameFile;
        }

        /*----------------------------------------------------------------------------
            %%Function: HandleDownloadGames
            %%Qualified: ArbWeb.AwMainForm.HandleDownloadGames
            %%Contact: rlittle
        
        ----------------------------------------------------------------------------*/
        private void HandleDownloadGames(string sFile)
        {
            object missing = System.Type.Missing;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook wkb;

            wkb = app.Workbooks.Open(sFile, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);

            string sOutFile = "";
            string sPrefix = "";

            if (m_sGameFile.Length < 1)
            {
                sOutFile = String.Format("{0}", Environment.GetEnvironmentVariable("temp"));
            }
            else
            {
                sOutFile = System.IO.Path.GetDirectoryName(m_sGameFile);
                string[] rgs;
                if (m_sGameFile.Length > 5 && sOutFile.Length > 0)
                {
                    rgs = CountsData.RexHelper.RgsMatch(m_sGameFile.Substring(sOutFile.Length + 1), "([.*])games");
                    if (rgs != null && rgs.Length > 0 && rgs[0] != null)
                        sPrefix = rgs[0];
                }
            }


            sOutFile = String.Format("{0}\\{2}games_{1:MM}{1:dd}{1:yy}_{1:HH}{1:mm}.csv", sOutFile, DateTime.Now, sPrefix);

            if (wkb != null)
            {
                wkb.SaveAs(sOutFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                wkb.Close(0, missing, missing);
            }
            app.Quit();
            app = null;
            m_sGameFile = sOutFile;
            System.IO.File.Delete(m_sGameCopy);
            System.IO.File.Copy(sOutFile, m_sGameCopy);
        }

        /*----------------------------------------------------------------------------
        	%%Function: DownloadGamesToFile
        	%%Qualified: ArbWeb.AwMainForm.DownloadGamesToFile
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        private string DownloadGenericToFile(string sTempFile)
        {
            m_iac.EnsureLoggedIn();

            m_iac.StatusReport.LogData("LaunchDownloadGeneric async task launched", 3, StatusRpt.MSGT.Body);
            var evtDownload = LaunchDownloadGeneric(sTempFile);
            m_iac.StatusReport.LogData("Before evtDownload.Wait()", 3, StatusRpt.MSGT.Body);
            evtDownload.WaitOne();
            m_iac.StatusReport.LogData("evtDownload.WaitOne() complete", 3, StatusRpt.MSGT.Body);

            return sTempFile;
        }

        private delegate AutoResetEvent LaunchDownloadGenericDel(string sTempFile);

        /*----------------------------------------------------------------------------
        	%%Function: LaunchDownloadGames
        	%%Qualified: ArbWeb.AwMainForm.LaunchDownloadGames
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        AutoResetEvent LaunchDownloadGeneric(string sTempFile)
        {
            if (m_iac.WebControl.InvokeRequired)
            {
            m_iac.StatusReport.LogData("InvokeRequired true for LaunchDownloadGeneric", 3, StatusRpt.MSGT.Body);

                IAsyncResult rsl = m_iac.WebControl.BeginInvoke(new LaunchDownloadGenericDel(DoLaunchDownloadGeneric), sTempFile);
                return (AutoResetEvent)m_iac.WebControl.EndInvoke(rsl);
            }
            else
            {
            m_iac.StatusReport.LogData("InvokeRequired false for DoLaunchDownloadGames", 3, StatusRpt.MSGT.Body);
                return DoLaunchDownloadGeneric(sTempFile);
            }
        }

        bool FNeedSelectReportFilter()
        {
            return m_sSelectFilterControlName != null;
        }

        /*----------------------------------------------------------------------------
        	%%Function: DoLaunchDownloadGames
        	%%Qualified: ArbWeb.AwMainForm.DoLaunchDownloadGames
        	%%Contact: rlittle

            This will optionally select a filter dropdown item from the starting 
            page. if not requested, this will just navigate to the start page and
            invoke the link to the report page.
        ----------------------------------------------------------------------------*/
        private AutoResetEvent DoLaunchDownloadGeneric(string sTempFile)
        {
            IHTMLDocument2 oDoc2 = m_iac.WebControl.Document2;
            int count = 0;
            string sFilter = null;

            while (count < 2)
                {
                // ok, now we're at the main assigner page...
                if (!m_iac.WebControl.FNavToPage(m_sReportPage))
                    throw (new Exception("could not navigate to games view"));

                oDoc2 = m_iac.WebControl.Document2;
                if (FNeedSelectReportFilter())
                    {
                    sFilter = m_iac.WebControl.SGetFilterID(oDoc2, m_sSelectFilterControlName, m_sFilterReq);
                    if (sFilter != null)
                        break;
                    }
                else
                    {
                    break;
                    }

                count++;
                }

            if (FNeedSelectReportFilter())
                {
                if (sFilter == null)
                    throw (new Exception($"there is no '{m_sFilterReq}' filter"));

                // now set that filter

                m_iac.WebControl.ResetNav();
                m_iac.WebControl.FSetSelectControlText(oDoc2, m_sSelectFilterControlName, null, m_sFilterReq, false);
                m_iac.WebControl.FWaitForNavFinish();

                if (!m_iac.WebControl.FNavToPage(m_sReportPrintPagePrefix + sFilter))
                    throw (new Exception("could not navigate to the reports page!"));
                }
            else
                {
                m_iac.WebControl.ResetNav();
                m_iac.ThrowIfNot(m_iac.WebControl.FClickControl(oDoc2, m_sidReportPageLink), "could not click on report link");
                m_iac.WebControl.FWaitForNavFinish();
                }

            oDoc2 = m_iac.WebControl.Document2;

            // loop through the Select controls we have to set (typically, this will include the file format)
            if (m_rgSelectSettings != null)
                {
                foreach (ControlSetting<string> cs in m_rgSelectSettings)
                    {
                    m_iac.WebControl.FSetSelectControlText(oDoc2, cs.ControlName, cs.IdControlExtra, cs.ControlValue, false);
                    }
                }

            if (m_rgCheckedSettings != null)
                {
                foreach (ControlSetting<bool> cs in m_rgCheckedSettings)
                    {
                    ArbWebControl.FSetCheckboxControlVal(oDoc2, cs.ControlValue, cs.ControlName);
                    }
                }

            m_iac.StatusReport.LogData(String.Format("Setting clipboard data: {0}", sTempFile), 3, StatusRpt.MSGT.Body);
            System.Windows.Forms.Clipboard.SetText(sTempFile);

            m_iac.WebControl.ResetNav();
            //          m_awc.PushNewWindow3Delegate(new DWebBrowserEvents2_NewWindow3EventHandler(DownloadGamesNewWindowDelegate));

            AutoResetEvent evtDownload = new AutoResetEvent(false);

            m_iac.StatusReport.LogData("Setting up TrapFileDownload", 3, StatusRpt.MSGT.Body);
            Win32Win.TrapFileDownload aww = new TrapFileDownload(m_iac.StatusReport, m_sFullExpectedName, m_sExpectedName, sTempFile, null, evtDownload);
            m_iac.WebControl.FClickControlNoWait(oDoc2, m_sReportPrintSubmitPrintControlName);
            return evtDownload;
        }

    }

    public class HandleGenericRoster
    {
        private IAppContext m_iac;
        private bool m_fAddOfficialsOnly;
        private delDoPass1Visit m_delDoPass1Visit;
        private delAddOfficials m_delAddOfficials;
        private delDoPostHandleRoster m_delDoPostHandleRoster;

        public delegate void delDoPass1Visit(string sEmail, string sOfficialID, IRoster irst, IRoster irstServer, IRosterEntry irste, IRoster irstBuilding, bool fJustAdded, bool fMarkOnly);
        public delegate void delAddOfficials(List<IRosterEntry> plirste);
        public delegate void delDoPostHandleRoster(IRoster irstUpload, IRoster irstBuilding);

        public HandleGenericRoster(IAppContext iac, bool fNeedPass1OnUpload, bool fAddOfficialsOnly, delDoPass1Visit doPass1Visit, delAddOfficials doAddOfficials, delDoPostHandleRoster doPostHandleRoster)
        {
            m_iac = iac;
            m_fNeedPass1OnUpload = fNeedPass1OnUpload;
            m_fAddOfficialsOnly = fAddOfficialsOnly;
            m_delDoPass1Visit = doPass1Visit;
            m_delAddOfficials = doAddOfficials;
            m_delDoPostHandleRoster = doPostHandleRoster;
        }

        public class PGL
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
                        m_iac.StatusReport.AddMessage("Found (" + sLinkTarget + ") when not looking for email!", StatusRpt.MSGT.Error);
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

        private bool m_fNeedPass1OnUpload;
        /*----------------------------------------------------------------------------
            %%Function: DoCoreRosterSync
            %%Qualified: ArbWeb.AwMainForm.DoCoreRosterSync
            %%Contact: rlittle

            Do the core roster syncing. 

            We are either syncing server->local (download) 
            or local->server (upload).

            We are being given the list of links on
            the official's edit page, the roster that we are uploading (if any),
            and a list of officials to limit our handling to (this is used when 
            we just added new officials and we just want to update their info/misc
            fields...)

            rstServer is the latest roster from the server -- useful for quickly
            determining what we need to update (without having to check the 
            server again)
        ----------------------------------------------------------------------------*/
        private void DoCoreRosterSync(PGL pgl, IRoster irst, IRoster irstBuilding, IRoster irstServer, List<IRosterEntry> plirsteLimit)
        {
            pgl.iCur = 0;
            Dictionary<string, bool> mpOfficials = new Dictionary<string, bool>();

            if (plirsteLimit != null)
                {
                foreach (IRosterEntry irsteCheck in plirsteLimit)
                    mpOfficials.Add("MAILTO:" + irsteCheck.Email.ToUpper(), true);
                }

            while (pgl.iCur < pgl.plofi.Count // we have links left to visit
                   && (irst == null
                       || m_fNeedPass1OnUpload) // why is this condition part of the while?! rst and cbRankOnly never changes in the loop
                   && pgl.iCur < pgl.plofi.Count)
                {
                if (irst == null
                    || (irst.PlsMiscLookupEmail(pgl.plofi[pgl.iCur].sEmail) != null
                        && pgl.plofi[pgl.iCur].sEmail.Length != 0))
                    {
                    IRosterEntry irste = irstBuilding.CreateRosterEntry();
                    bool fMarkOnly = false;

                    irste.SetEmail((string) pgl.plofi[pgl.iCur].sEmail);
                    m_iac.StatusReport.AddMessage(String.Format("Processing roster info for {0}...", pgl.plofi[pgl.iCur].sEmail));

                    if (m_fAddOfficialsOnly && plirsteLimit == null)
                        fMarkOnly = true;

                    if (plirsteLimit != null)
                        {
                        if (!mpOfficials.ContainsKey(((string) pgl.plofi[pgl.iCur].sEmail.ToUpper())))
                            {
                            pgl.iCur++;
                            continue; // it doesn't match an official in the "limit-to" list.
                            }

                        fMarkOnly = false; // we want to process this one.
                        }

                    bool fJustAdded = plirsteLimit == null && (irst == null || !irst.IsQuick || irst.IsUploadableQuickroster);
                    m_delDoPass1Visit(pgl.plofi[pgl.iCur].sEmail, pgl.plofi[pgl.iCur].sOfficialID, irst, irstServer, irste, irstBuilding, fJustAdded, fMarkOnly);

                    if (irst == null && !String.IsNullOrEmpty(irste.Email))
                        {
                        irstBuilding.Add(irste);
                        //                        rste.AppendToFile(sOutFile, m_rgsRankings);
                        // at this point, we have the name and the affiliation
                        //						if (!FAppendToFile(sOutFile, sName, (string)pgl.rgsData[pgl.iCur], plsValue))
                        //							throw new Exception("couldn't append to the file!");
                        }
                    else
                        {
                        if (!String.IsNullOrEmpty(pgl.plofi[pgl.iCur].sEmail))
                            {
                            IRosterEntry irsteT = irst.IrsteLookupEmail(pgl.plofi[pgl.iCur].sEmail);

                            if (irsteT != null)
                                irsteT.Marked = true;
                            }
                        }

                    if (m_iac.Profile.TestOnly)
                        {
                        break;
                        }
                    }

                pgl.iCur++;
                }
        }

        private void VOPC_PopulatePgl(IHTMLDocument2 oDoc2, Object o)
        {
            PopulatePglOfficialsFromPageCore((PGL)o, oDoc2);
        }


        /* N A V I G A T E  O F F I C I A L S  P A G E  A L L  O F F I C I A L S */
        /*----------------------------------------------------------------------------
	    	%%Function: NavigateOfficialsPageAllOfficials
	    	%%Qualified: ArbWeb.AwMainForm.NavigateOfficialsPageAllOfficials
	    	%%Contact: rlittle
	    	
	    ----------------------------------------------------------------------------*/
        public void NavigateOfficialsPageAllOfficials()
        {
            m_iac.EnsureLoggedIn();

            m_iac.ThrowIfNot(m_iac.WebControl.FNavToPage(WebCore._s_Page_OfficialsView), "Couldn't nav to officials view!");
            m_iac.WebControl.FWaitForNavFinish();

            // from the officials view, make sure we are looking at active officials
            m_iac.WebControl.ResetNav();
            IHTMLDocument2 oDoc2 = m_iac.WebControl.Document2;

            m_iac.WebControl.FSetSelectControlText(oDoc2, WebCore._s_OfficialsView_Select_Filter, WebCore._sid_OfficialsView_Select_Filter, "All Officials", true);
            m_iac.WebControl.FWaitForNavFinish();
        }

        // object could be RST or PGL
        public delegate void VisitOfficialsPageCallback(IHTMLDocument2 oDoc2, Object o);

        /*----------------------------------------------------------------------------
        	%%Function: ProcessAllOfficialPages
        	%%Qualified: ArbWeb.HandleGenericRoster.ProcessAllOfficialPages
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public void ProcessAllOfficialPages(VisitOfficialsPageCallback vopc, Object o)
        {
            NavigateOfficialsPageAllOfficials();

            IHTMLDocument2 oDoc2 = m_iac.WebControl.Document2;

            // first, get the first pages and callback

            vopc(oDoc2, o);

            // figure out how many pages we have
            // find all of the <a> tags with an href that targets a pagination postback
            IHTMLElementCollection ihec = (IHTMLElementCollection) oDoc2.all.tags("a");
            List<string> plsHrefs = new List<string>();

            foreach (IHTMLAnchorElement iha in ihec)
                {
                if (iha.href != null && iha.href.Contains(WebCore._s_OfficialsView_PaginationHrefPostbackSubstr))
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
                        m_iac.WebControl.FWaitForNavFinish();
                        oDoc2 = m_iac.WebControl.Document2;

                        vopc(oDoc2, o);
                        break; // done processing the element collection -- have to process the next one for the next doc
                        }
                    }

                }
        }

        private PGL PglGetOfficialsFromWeb()
        {
            m_iac.EnsureLoggedIn();
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
                string s = (string)pgl.plofi[i].sOfficialID;

                string sID = s.Substring(s.IndexOf("=") + 1);
                pgl.plofi[i].sOfficialID = sID;
                i++;
                }

            return pgl;
        }


        public delegate void HandleRosterPostUpdateDelegate(HandleGenericRoster gr, IRoster irst);

        /*----------------------------------------------------------------------------
        	%%Function: GenericVisitRoster
        	%%Qualified: ArbWeb.HandleGenericRoster.GenericVisitRoster
        	%%Contact: rlittle
        	
			If rst == null, then we're downloading the roster.  Otherwise, we are
			uploading

            FUTURE: Make this a generic "VisitRoster" with callbacks or methods
            specific to upload or download (i.e. core out the code shared by 
            upload and download, then make separate upload and download functions
            with no duplication)
        ----------------------------------------------------------------------------*/
        public void GenericVisitRoster(IRoster irstUpload, IRoster irstBuilding, string sOutFile, IRoster irstServer, HandleRosterPostUpdateDelegate handleRosterPostUpdate)
        {
            //Roster rstBuilding = null;
            PGL pgl;

            if (irstUpload != null && irstBuilding != null)
                throw new Exception("cannot upload AND download at the same time");

            // we're not going to write the roster out until the end now...

            //if (rstUpload == null)
                //rstBuilding = new Roster();

            pgl = PglGetOfficialsFromWeb();
            DoCoreRosterSync(pgl, irstUpload, irstBuilding, irstServer, null /*plrsteLimit*/);

            handleRosterPostUpdate?.Invoke(this, irstBuilding);

            if (irstUpload != null)
            {
                List<IRosterEntry> plirsteUnmarked = irstUpload.PlirsteUnmarked();

                // we might have some officials left "unmarked".  These need to be added

                // at this point, all the officials have either been marked or need to 
                // be added

                if (plirsteUnmarked.Count > 0)
                {
                    if (MessageBox.Show(String.Format("There are {0} new officials.  Add these officials?", plirsteUnmarked.Count), "ArbWeb", MessageBoxButtons.YesNo) ==
                        DialogResult.Yes)
                    {
                        m_delAddOfficials?.Invoke(plirsteUnmarked);
                        // now we have to reload the page of links and do the whole thing again (updating info, etc)
                        // so we get the misc fields updated.  Then fall through to the rankings and do everyone at
                        // once
                        pgl = PglGetOfficialsFromWeb(); // refresh to get new officials
                        DoCoreRosterSync(pgl, irstUpload, null /*rstBuilding*/, irstServer, plirsteUnmarked);
                        // now we can fall through to our core ranking handling...
                    }
                }
            }

            // now, do the rankings.  this is easiest done in the bulk rankings tool...
            m_delDoPostHandleRoster?.Invoke(irstUpload, irstBuilding);
            // lastly, if we're downloading, then output the roster

            if (irstUpload == null)
                irstBuilding.WriteRoster(sOutFile);

            if (m_iac.Profile.TestOnly)
            {
                MessageBox.Show("Stopping after 1 roster item");
            }
        }

        /*----------------------------------------------------------------------------
        	%%Function: SBuildRosterFilename
        	%%Qualified: ArbWeb.AwMainForm.SBuildRosterFilename
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public static string SBuildRosterFilename(string sRosterName)
        {
            string sOutFile;
            string sPrefix = "";

            if (sRosterName.Length < 1)
                {
                sOutFile = String.Format("{0}", Environment.GetEnvironmentVariable("temp"));
                }
            else
                {
                sOutFile = System.IO.Path.GetDirectoryName(sRosterName);
                string[] rgs;
                if (sRosterName.Length > 5 && sOutFile.Length > 0)
                    {
                    rgs = CountsData.RexHelper.RgsMatch(sRosterName.Substring(sOutFile.Length + 1), "([.*])roster");
                    if (rgs != null && rgs.Length > 0 && rgs[0] != null)
                        sPrefix = rgs[0];
                    }
                }

            sOutFile = String.Format("{0}{2}\\roster_{1:MM}{1:dd}{1:yy}_{1:HH}{1:mm}.csv", sOutFile, DateTime.Now, sPrefix);
            return sOutFile;
        }
    }
}
