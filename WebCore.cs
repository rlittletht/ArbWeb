using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
}
