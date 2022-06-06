using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OpenQA.Selenium;
using TCore.StatusBox;
using TCore.Util;
using HtmlAgilityPack;
using TCore.WebControl;

namespace ArbWeb
{
    class WebCore
    {
        #region ArbiterStrings

        // ============================================================================
        // T O P  L E V E L    
        // ============================================================================
        public const string _s_Home = "https://www1.arbitersports.com/Shared/SignIn/Signin.aspx"; // ok2010
        public const string _s_Assigning = "https://www1.arbitersports.com/Assigner/Games/NewGamesView.aspx"; // ok2010
        public const string _s_RanksEdit = "https://www1.arbitersports.com/Assigner/RanksEdit.aspx"; // ok2010
        public const string _s_AddUser = "https://www1.arbitersports.com/Assigner/UserAdd.aspx?userTypeID=3"; // ok2010u
        private const string _s_OfficialsView = "https://www1.arbitersports.com/Assigner/OfficialsView.aspx"; // ok2010u
        public const string _s_Announcements = "https://www1.arbitersports.com/Shared/AnnouncementsEdit.aspx"; // ok2015
        public const string _s_ContactsView = "https://www1.arbitersports.com/Assigner/ContactsView.aspx"; // ok2018

        // ============================================================================
        // D I R E C T  A C C E S S
        // ============================================================================
        public const string _s_Assigning_PrintAddress = "https://www1.arbitersports.com/Assigner/Games/Print.aspx?filterID="; // ok2010
        public const string _s_EditUser_MiscFields = "https://www1.arbitersports.com/Official/MiscFieldsEdit.aspx?userID="; // ok2010
        public const string _s_EditUser = "https://www1.arbitersports.com/Official/OfficialEdit.aspx?userID="; // ok2010u

        // ============================================================================
        // H O M E
        // ============================================================================
        private const string _s_Home_Anchor_Login = "SignInButton"; // ctl00$ucMiniLogin$SignInButton"; // ok2010
        public const string _s_Home_Input_Email = "ctl00$ContentHolder$pgeSignIn$conSignIn$txtEmail"; // ctl00$ucMiniLogin$UsernameTextBox"; // ok2016
        public const string _s_Home_Input_Password = "txtPassword"; // ctl00$ucMiniLogin$PasswordTextBox"; // ok2016
        public const string _s_Home_Button_SignIn = "ctl00$ContentHolder$pgeSignIn$conSignIn$btnSignIn"; // ctl00$ucMiniLogin$SignInButton"; // ok2016

        public const string _sid_Home_Div_PnlAccounts = "ctl00_ContentHolder_pgeDefault_conDefault_pnlAccounts"; // ok2010
        public const string _sid_Home_MessagingText = "ctl00_MessagingText"; // ok2022
        public const string _sid_Home_Anchor_NeedHelpLink = "ctl00_PageHelpTextLink"; // not ok 2022

        // ============================================================================
        // A S S I G N I N G
        // ============================================================================
        // (games view) links
        public const string _s_Assigning_Select_Filters = "ctl00$ContentHolder$pgeGamesView$conGamesView$ddlSavedFilters"; // ok2021
        public const string _sid_Assigning_Select_Filters = "ddlSavedFilters"; // ok2021
        public const string _s_Assigning_Reports_Select_Format = "ctl00$ContentHolder$pgePrint$conPrint$ddlFormat"; // ok2010
        public const string _s_Assigning_Reports_Submit_Print = "ctl00$ContentHolder$pgePrint$navPrint$btnBeginPrint"; // ok2010
        public const string _sid_Assigning_GamesReport = "ctl00_ContentHolder_pgeGamesView_cmnReports_tskPrint"; // ok2021
        public const string _sid_Assigning_Reports_Done = "ctl00_ContentHolder_pgePrint_navPrint_lnkDone"; // ok2021
        public const string _s_Assigning_CheckAll = "ctl00$ContentHolder$pgeGamesView$conGamesView$dgGames$ctl01$chkAll"; // ok2021

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

        public const string _sid_AddUser_Input_State = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_uclAddress_address_ddlState"; // ok2021
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
        public const string _s_EditUser_MiddleName = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtMiddleName"; // ok2018
        public const string _s_EditUser_LastName = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtLastName"; // ok2018
        public const string _s_EditUser_Address1 = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclAddress$address_txtAddress1"; // ok2018
        public const string _s_EditUser_Address2 = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclAddress$address_txtAddress2"; // ok2018
        public const string _s_EditUser_City = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclAddress$address_txtCity"; // ok2021
        public const string _s_EditUser_State = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$txtState"; // ok2016
        public const string _s_EditUser_PostalCode = "ctl00$ContentHolder$pgeOfficialEdit$conOfficialEdit$uclAddress$address_txtPostalCode"; // ok2021
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

        public const string _sid_MiscFields_MainBodyContentTable = "MainBodyContentTable"; // ok2021
        
        // phone types are Home, Work, Fax, Cellular, Pager, Security, Other
        public const string _sid_AddUser_Button_Next = "ctl00_ContentHolder_pgeUserAdd_navUserAdd_btnNext"; // ok2010
        public const string _sid_AddUser_Input_Address1 = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_uclAddress_address_txtAddress1"; // ok2018
        public const string _sid_AddUser_Input_IsActive = "ctl00_ContentHolder_pgeUserAdd_conUserAdd_chkIsActive"; // ok2010
        public const string _sid_AddUser_Button_Cancel = "ctl00_ContentHolder_pgeUserAdd_navUserAdd_lnkCancel"; // ok2010

        // ============================================================================
        // O F F I C I A L S  V I E W
        // ============================================================================
        public const string _s_Page_OfficialsView = "https://www1.arbitersports.com/Assigner/OfficialsView.aspx"; // ok2021
        public const string _s_OfficialsView_Select_Filter = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$ddlFilter"; // ok2010
        public const string _sid_OfficialsView_Select_Filter = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_ddlFilter"; // ok2021
        
        public const string _sid_OfficialsView_PrintRoster = "ctl00_ContentHolder_pgeOfficialsView_sbrReports_tskPrint"; // ok2010u
        public const string _sid_OfficialsView_PrintCustomRoster = "ctl00_ContentHolder_pgeOfficialsView_sbrReports_showCustomRoster"; // ok2021
        
        public const string _sid_OfficialsView_ContentTable = "ctl00_ContentHolder_pgeOfficialsView_conOfficialsView_dgOfficials"; // ok2013

//        public const string _s_OfficialsView_PaginationHrefPostbackSubstr = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$dgOfficials$ctl204$ctl"; // ok2014
        public const string _s_OfficialsView_PaginationHrefPostbackSubstr = "ctl00$ContentHolder$pgeOfficialsView$conOfficialsView$dgOfficials$ctl"; // ok2021

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

        public const string _sid_CustomRosterPrint_UserFilter = "OfficialFilterId"; // ok2021
        public const string _s_CustomRosterPrint_UserFilter = "OfficialFilterId"; // ok2021
        public const string _sid_CustomRosterPrint_DateOfBirth = "DateOfBirth"; // ok2021
        public const string _sid_CustomRosterPrint_OfficialNumber = "OfficialNumber"; // ok2021
        public const string _sid_CustomRosterPrint_DateJoined = "DateJoined"; // ok2021
        public const string _sid_CustomRosterPrint_UserID = "UserId"; // ok2021
        public const string _sid_CustomRosterPrint_MiddleName = "MiddleName"; // ok2021

        public const string _sid_CustomRosterPrint_SelectAllCustomFields = "selectAllCustomFields"; // ok2021
        public const string _sid_CustomRosterPrint_GenerateRosterReport = "generateRosterReport"; // ok2021

        public const string _sid_CustomRosterPrint_CustomFieldListDropdown = "CustomFieldList"; // ok2021

        // ============================================================================
        // R A N K S
        // ============================================================================
        public const string _s_RanksEdit_Select_PosNames = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$ddlPosNames"; // ok2010
        public const string _sid_RanksEdit_Select_PosNames = "ctl00_ContentHolder_pgeRanksEdit_conRanksEdit_ddlPosNames"; // ok2021

        
        public const string _s_RanksEdit_Checkbox_Active = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$chkActive"; // ok2010
        public const string _s_RanksEdit_Checkbox_Rank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$chkRank"; // ok2010
        public const string _s_RanksEdit_Select_NotRanked = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$lstNotRanked"; // ok2010
        public const string _s_RanksEdit_Select_Ranked = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$lstRanked"; // ok2010

        public const string _sid_RanksEdit_Checkbox_Active = "ctl00_ContentHolder_pgeRanksEdit_conRanksEdit_chkActive"; // ok2021
        public const string _sid_RanksEdit_Checkbox_Rank = "ctl00_ContentHolder_pgeRanksEdit_conRanksEdit_chkRank"; // ok2021
        public const string _sid_RanksEdit_Select_NotRanked = "ctl00_ContentHolder_pgeRanksEdit_conRanksEdit_lstNotRanked"; // ok2021
        public const string _sid_RanksEdit_Select_Ranked = "ctl00_ContentHolder_pgeRanksEdit_conRanksEdit_lstRanked"; // ok2021

        public const string _s_RanksEdit_Button_Unrank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$btnUnrank"; // ok2010
        public const string _s_RanksEdit_Button_ReRank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$btnReRank"; // ok2010
        public const string _s_RanksEdit_Button_Rank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$btnRank"; // ok2010
        public const string _s_RanksEdit_Input_Rank = "ctl00$ContentHolder$pgeRanksEdit$conRanksEdit$txtRank"; // ok2010

        public const string _sid_RanksEdit_Button_Unrank = "ctl00_ContentHolder_pgeRanksEdit_conRanksEdit_btnUnrank"; // ok2021
        public const string _sid_RanksEdit_Button_ReRank = "ctl00_ContentHolder_pgeRanksEdit_conRanksEdit_btnReRank"; // ok2021
        public const string _sid_RanksEdit_Button_Rank = "ctl00_ContentHolder_pgeRanksEdit_conRanksEdit_btnRank"; // ok2021
        public const string _sid_RanksEdit_Input_Rank = "ctl00_ContentHolder_pgeRanksEdit_conRanksEdit_txtRank"; // ok2021

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
        
        #region Support Functions
        
	    /* R G S  F R O M  C H L B X */
	    /*----------------------------------------------------------------------------
		    %%Function: RgsFromChlbx
		    %%Qualified: ArbWeb.ArbWebControl.RgsFromChlbx
		    %%Contact: rlittle
		    
	    ----------------------------------------------------------------------------*/
	    public static string[] RgsFromChlbx(bool fUse, CheckedListBox chlbx)
	    {
		    return RgsFromChlbx(fUse, chlbx, -1, false, null, false);
	    }

	    /* R G S  F R O M  C H L B X  S P O R T */
	    /*----------------------------------------------------------------------------
		    %%Function: RgsFromChlbxSport
		    %%Qualified: ArbWeb.ArbWebControl.RgsFromChlbxSport
		    %%Contact: rlittle
		    
	    ----------------------------------------------------------------------------*/
	    public static string[] RgsFromChlbxSport(bool fUse, CheckedListBox chlbx, string sSport, bool fMatch)
	    {
		    return RgsFromChlbx(fUse, chlbx, -1, false, sSport, fMatch);
	    }

	    /* R G S  F R O M  C H L B X */
	    /*----------------------------------------------------------------------------
		    %%Function: RgsFromChlbx
		    %%Qualified: ArbWeb.ArbWebControl.RgsFromChlbx
		    %%Contact: rlittle
		    
	    ----------------------------------------------------------------------------*/
	    public static string[] RgsFromChlbx(
		    bool fUse,
		    CheckedListBox chlbx,
		    int iForceToggle,
		    bool fForceOn,
		    string sSport,
		    bool fMatch)
	    {
		    string sSport2 = sSport == "Softball" ? "SB" : sSport;

		    if (!fUse && sSport == null)
			    return null;

		    int c = chlbx.CheckedItems.Count;

		    if (!fUse)
			    c = chlbx.Items.Count;

		    if (iForceToggle != -1)
		    {
			    if (fForceOn)
				    c++;
			    else
				    c--;
		    }

		    string[] rgs = new string[c];
		    int i = 0;

		    if (!fUse)
		    {
			    int iT = 0;

			    for (i = 0; i < c; i++)
			    {
				    rgs[iT] = (string) chlbx.Items[i];
				    if (sSport != null)
				    {
					    if ((rgs[iT].IndexOf(sSport) >= 0 && fMatch)
					        || (rgs[iT].IndexOf(sSport) == -1 && !fMatch)
					        || (rgs[iT].IndexOf(sSport2) >= 0 && fMatch)
					        || (rgs[iT].IndexOf(sSport2) == -1 && !fMatch))
					    {
						    iT++;
					    }
				    }
				    else
				    {
					    iT++;
				    }
			    }

			    if (iT < c)
				    Array.Resize(ref rgs, iT);

			    return rgs;
		    }

		    i = 0;
		    foreach (int iChecked in chlbx.CheckedIndices)
		    {
			    if (iChecked == iForceToggle)
				    continue;
			    rgs[i] = (string) chlbx.Items[iChecked];
			    if (sSport != null)
			    {
				    if ((rgs[i].IndexOf(sSport) >= 0 && fMatch)
				        || (rgs[i].IndexOf(sSport) == -1 && !fMatch))
				    {
					    i++;
				    }
			    }
			    else
			    {
				    i++;
			    }
		    }

		    if (fForceOn && iForceToggle != -1)
			    rgs[i++] = (string) chlbx.Items[iForceToggle];

		    if (i < c)
			    Array.Resize(ref rgs, i);

		    return rgs;
	    }

	    /* U P D A T E  C H L B X  F R O M  R G S */
	    /*----------------------------------------------------------------------------
		    %%Function: UpdateChlbxFromRgs
		    %%Qualified: ArbWeb.ArbWebControl.UpdateChlbxFromRgs
		    %%Contact: rlittle
		    
	    ----------------------------------------------------------------------------*/
	    public static void UpdateChlbxFromRgs(
		    CheckedListBox chlbx,
		    string[] rgsSource,
		    string[] rgsChecked,
		    string[] rgsFilterPrefix,
		    bool fCheckAll)
	    {
		    chlbx.Items.Clear();
		    SortedList<string, int> mp = Utils.PlsUniqueFromRgs(rgsChecked);

		    foreach (string s in rgsSource)
		    {
			    bool fSkip = false;

			    if (rgsFilterPrefix != null)
			    {
				    fSkip = true;
				    foreach (string sPrefix in rgsFilterPrefix)
				    {
					    if (s.Length > sPrefix.Length && String.Compare(s.Substring(0, sPrefix.Length), sPrefix, true /*ignoreCase*/) == 0)
					    {
						    fSkip = false;
						    break;
					    }
				    }
			    }

			    if (fSkip)
				    continue;

			    CheckState cs;

			    if (fCheckAll || mp.ContainsKey(s))
				    cs = CheckState.Checked;
			    else
				    cs = CheckState.Unchecked;

			    int i = chlbx.Items.Add(s, cs);
		    }
	    }
	    
	    #endregion
    }
    
    

    public class DownloadGenericExcelReport
    {
        private readonly string m_sFilterOptionTextReq;
        private readonly string m_sDescription;
        private readonly IAppContext m_appContext;
        private readonly string m_sReportPage;
        private readonly string m_sSelectFilterControlName;
        private readonly string m_sReportPrintPagePrefix;
        private readonly string m_sReportPrintSubmitPrintControlName;
        private readonly string[] m_expectedFullNameTemplates;
        private readonly string m_sExpectedName;
        private readonly string m_sidReportPageLink;
        private string m_sGameFile;
        private readonly string m_sGameCopy;

        public class ControlSetting<T>
        {
            private string m_sControlName;
            private string m_sidControlExtra;   // this is usually something like the Choice element ID
            private T m_tControlValue; // for select controls, this is the option text value

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

        private string m_sidSelectFilterControl;

        // this version selects a filter
        public DownloadGenericExcelReport(
            string sFilterOptionTextReq, 
            string sDescription, 
            string sReportPage, 
            string sSelectFilterControlName, 
            string sidSelectFilterControl,
            string sReportPrintPagePrefix, 
            string sReportPrintSubmitPrintControlName,
            string[] expectedFullNameTemplates,
            string sExpectedName, 
            ControlSetting<string>[] rgSelectSettings,
            string sGameFile,
            string sGameCopy,
            IAppContext appContext)
        {
            m_sFilterOptionTextReq = sFilterOptionTextReq;
            m_sDescription = sDescription;
            m_appContext = appContext;
            m_sReportPage = sReportPage;
            m_sSelectFilterControlName = sSelectFilterControlName;
            m_sidSelectFilterControl = sidSelectFilterControl;
            m_sReportPrintPagePrefix = sReportPrintPagePrefix;
            m_sReportPrintSubmitPrintControlName = sReportPrintSubmitPrintControlName;
            m_sExpectedName = sExpectedName;
            m_expectedFullNameTemplates = expectedFullNameTemplates;
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
            IAppContext appContext)
        {
            m_sDescription = sDescription;
            m_appContext = appContext;
            m_sReportPage = sReportPage;
            m_sidReportPageLink = sidReportPageLink;
            m_sReportPrintSubmitPrintControlName = sReportPrintSubmitPrintControlName;
            m_expectedFullNameTemplates = new string[] {sFullExpectedName};
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
            m_appContext.StatusReport.AddMessage($"Starting {m_sDescription} download...");
            m_appContext.StatusReport.PushLevel();
            string sTempFile = Filename.SBuildTempFilename("temp", "xls");

            sTempFile = DownloadGenericToFile(sTempFile);
            HandleDownloadGames(sTempFile);

            System.IO.File.Delete(sTempFile);

            // ok, now we have all games selected...
            // time to try to download a report
            m_appContext.StatusReport.PopLevel();
            m_appContext.StatusReport.AddMessage($"Completed downloading {m_sDescription}.");
            sGameFileNew = m_sGameFile;
        }

        /*----------------------------------------------------------------------------
			%%Function:ConvertExcelFileToCsv
			%%Qualified:ArbWeb.DownloadGenericExcelReport.ConvertExcelFileToCsv

        ----------------------------------------------------------------------------*/
        public static void ConvertExcelFileToCsv(string sExcelFile, string sTargetCsvFile)
        {
	        object missing = System.Type.Missing;
	        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

	        Microsoft.Office.Interop.Excel.Workbook wkb;

	        wkb = app.Workbooks.Open(sExcelFile, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);

	        if (wkb != null)
	        {
		        Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet) wkb.Worksheets[1];
		        sheet.Select(Type.Missing);
                wkb.SaveAs(sTargetCsvFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
		        wkb.Close(0, missing, missing);
	        }
	        app.Quit();
	        app = null;
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
                sOutFile = $"{Environment.GetEnvironmentVariable("temp")}";
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
            m_appContext.EnsureLoggedIn();

            DoLaunchDownloadGeneric(sTempFile);

            return sTempFile;
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
        private void DoLaunchDownloadGeneric(string sTempFile)
        {
            int count = 0;
            string sFilterOptionValue = null;

            while (count < 2)
                {
                // ok, now we're at the main assigner page...
                if (!m_appContext.WebControl.FNavToPage(m_sReportPage))
                    throw (new Exception("could not navigate to games view"));

                if (FNeedSelectReportFilter())
                    {
                    sFilterOptionValue = m_appContext.WebControl.GetOptionValueFromFilterOptionTextForControlName(m_sSelectFilterControlName, m_sFilterOptionTextReq);
                    if (sFilterOptionValue != null)
                        break;
                    }
                else
                    {
                    break;
                    }

                // throw new Exception("needed more than one iteration?");
                count++;
                }

            if (FNeedSelectReportFilter())
                {
                if (sFilterOptionValue == null)
                    throw (new Exception($"there is no '{m_sFilterOptionTextReq}' filter"));

                // now set that filter

                m_appContext.WebControl.FSetSelectedOptionTextForControlId(m_sidSelectFilterControl, m_sFilterOptionTextReq);

                m_appContext.WebControl.FSetCheckboxControlNameVal(true, WebCore._s_Assigning_CheckAll);

                if (!(m_appContext.WebControl.FClickControlId(WebCore._sid_Assigning_GamesReport, WebCore._sid_Assigning_Reports_Done)))
                    throw (new Exception("could not navigate to the reports page!"));
                }
            else
                {
                Utils.ThrowIfNot(m_appContext.WebControl.FClickControlId(m_sidReportPageLink), "could not click on report link");
                }

            // loop through the Select controls we have to set (typically, this will include the file format)
            if (m_rgSelectSettings != null)
                {
                foreach (ControlSetting<string> cs in m_rgSelectSettings)
                    {
                    m_appContext.WebControl.FSetSelectedOptionTextForControlId(cs.IdControlExtra, cs.ControlValue);
                    }
                }

            if (m_rgCheckedSettings != null)
                {
                foreach (ControlSetting<bool> cs in m_rgCheckedSettings)
                    m_appContext.WebControl.FSetCheckboxControlNameVal(cs.ControlValue, cs.ControlName);
                }

            // m_iac.StatusReport.LogData($"Setting clipboard data: {sTempFile}", 3, StatusRpt.MSGT.Body);
            // System.Windows.Forms.Clipboard.SetText(sTempFile);

            WebControl.FileDownloader downloader = new WebControl.FileDownloader(
	            m_appContext.WebControl,
	            m_expectedFullNameTemplates,
	            sTempFile,
	            () => m_appContext.WebControl.FClickControlName(m_sReportPrintSubmitPrintControlName));
            
            downloader.GetDownloadedFile();
        }
    }

    public class HandleGenericRoster
    {
        private IAppContext m_iac;
        private bool m_fAddOfficialsOnly;
        private delDoPass1Visit m_delDoPass1Visit;
        private delAddOfficials m_delAddOfficials;
        private delDoPostHandleRoster m_delDoPostHandleRoster;

        public delegate void delDoPass1Visit(string sEmail, string sOfficialID, IRoster irst, IRoster irstServer, IRosterEntry irste, IRoster irstBuilding, bool fNotJustAdded, bool fMarkOnly);
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
        private void PopulatePglOfficialsFromPageCore(PGL pgl)
        {
	        // grab the info from the current navigated page
	        IWebElement table = m_iac.WebControl.Driver.FindElement(By.XPath("//body"));

	        string sHtml = table.GetAttribute("outerHTML");
	        HtmlAgilityPack.HtmlDocument html = new HtmlAgilityPack.HtmlDocument();

	        html.LoadHtml(sHtml);
	        string sHrefSubstringMatch = "OfficialEdit.aspx\\?userID=";
	        
            Regex rx3 = new Regex($"{sHrefSubstringMatch}.*");
            Regex rxData = new Regex("mailto:.*");

            bool fLookingForEmail = false;

            string sXpath = $"//a"; // [contains(@href, '{sHrefSubstringMatch}')]

            // build up a list of probable index links
            HtmlNodeCollection links = html.DocumentNode.SelectSingleNode(".")
	            .SelectNodes(sXpath);
            
            foreach (HtmlNode link in links)
            {
                string sLinkTarget = link.GetAttributeValue("href", "");

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
                        m_iac.StatusReport.AddMessage("Found (" + sLinkTarget + ") when not looking for email!", MSGT.Error);
                    }
                }

                if (rx3.IsMatch(sLinkTarget))
                {
                    PGL.OFI ofi = new PGL.OFI();

                    ofi.sEmail = "";
                    ofi.sOfficialID = sLinkTarget;
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
        private void DoCoreRosterSync(PGL pgl, IRoster irstUpload, IRoster irstBuilding, IRoster irstServer, List<IRosterEntry> plirsteLimit)
        {
            pgl.iCur = 0;
            Dictionary<string, bool> mpOfficials = new Dictionary<string, bool>();

            if (plirsteLimit != null)
                {
                foreach (IRosterEntry irsteCheck in plirsteLimit)
                    mpOfficials.Add("MAILTO:" + irsteCheck.Email.ToUpper(), true);
                }

            while (pgl.iCur < pgl.plofi.Count // we have links left to visit
                   && (irstUpload == null
                       || m_fNeedPass1OnUpload) // why is this condition part of the while?! rst and cbRankOnly never changes in the loop
                   && pgl.iCur < pgl.plofi.Count)
                {
                // if we aren't uploading, or if we are uploading and we values for this email address AND the current link has an email address
                if (irstUpload == null
                    || (!String.IsNullOrEmpty(pgl.plofi[pgl.iCur].sEmail)
                        && irstUpload.IrsteLookupEmail(pgl.plofi[pgl.iCur].sEmail) != null))
                    {
                    IRosterEntry irste;

                    // we need to build a RosterEntry for this. if we are downloading, then we will add it to the building roster
                    // otherwise, we will use it to compare with the one we are uploading to see if we need to update anything

                    if (irstBuilding != null)
                        irste = irstBuilding.CreateRosterEntry();
                    else
                        irste = irstUpload.CreateRosterEntry();

                    bool fMarkOnly = false;

                    irste.SetEmail((string) pgl.plofi[pgl.iCur].sEmail);
                    m_iac.StatusReport.AddMessage($"Processing roster info for {pgl.plofi[pgl.iCur].sEmail}...");

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

                    // if we don't have a limit list, then we aren't in pass right after adding officials
                    // (we build the limit list when we actually add officials)
                    bool fNotJustAdded = plirsteLimit == null && (irstUpload == null || irstUpload.IsUploadableRoster);
                    m_delDoPass1Visit(pgl.plofi[pgl.iCur].sEmail, pgl.plofi[pgl.iCur].sOfficialID, irstUpload, irstServer, irste, irstBuilding, fNotJustAdded, fMarkOnly);

                    if (irstUpload == null && !String.IsNullOrEmpty(irste.Email))
                        {
                        irstBuilding.Add(irste);
                        }
                    else
                        {
                        if (!String.IsNullOrEmpty(pgl.plofi[pgl.iCur].sEmail))
                            {
                            IRosterEntry irsteT = irstUpload.IrsteLookupEmail(pgl.plofi[pgl.iCur].sEmail);

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

        private void VOPC_PopulatePgl(Object o)
        {
            PopulatePglOfficialsFromPageCore((PGL)o);
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

            Utils.ThrowIfNot(m_iac.WebControl.FNavToPage(WebCore._s_Page_OfficialsView), "Couldn't nav to officials view!");

            m_iac.WebControl.FSetSelectedOptionTextForControlId(WebCore._sid_OfficialsView_Select_Filter, "All Officials");
        }

        // object could be RST or PGL
        public delegate void VisitOfficialsPageCallback(Object o);

        public static string ToXPath(string value)
        {
	        const string apostrophe = "'";
	        const string quote = "\"";

	        if (value.Contains(quote))
	        {
		        if (value.Contains(apostrophe))
		        {
			        throw new Exception("Illegal XPath string literal.");
		        }
		        else
		        {
			        return apostrophe + value + apostrophe;
		        }
	        }
	        else
	        {
		        return quote + value + quote;
	        }
        }
        
        /*----------------------------------------------------------------------------
        	%%Function: ProcessAllOfficialPages
        	%%Qualified: ArbWeb.HandleGenericRoster.ProcessAllOfficialPages
        	%%Contact: rlittle
        	
        ----------------------------------------------------------------------------*/
        public void ProcessAllOfficialPages(VisitOfficialsPageCallback visit, Object o)
        {
	        MicroTimer timer = new MicroTimer();

            NavigateOfficialsPageAllOfficials();

            // first, get the first pages and callback
            timer.Stop();
            m_iac.StatusReport.LogData($"Process All Officials(NavigateAtStart) elapsedTime: {timer.MsecFloat}", 1, MSGT.Body);

            visit(o);

	        timer.Reset();
	        timer.Start();
	        
	        // figure out how many pages we have
	        // find all of the <a> tags with an href that targets a pagination postback
	        IList<IWebElement> anchors = m_iac.WebControl.Driver.FindElements(By.XPath($"//tr[@class='numericPaging']//a[contains(@href, '{WebCore._s_OfficialsView_PaginationHrefPostbackSubstr}')]"));
	        List<string> plsHrefs = new List<string>();

	        foreach (IWebElement anchor in anchors)
	        {
		        string href = anchor.GetAttribute("href");

		        if (href != null && href.Contains(WebCore._s_OfficialsView_PaginationHrefPostbackSubstr))
		        {
			        // we can't just remember this element because we will be navigating around.  instead we will
			        // just remember the entire target so we can find it again
			        plsHrefs.Add(href);
		        }
	        }

	        timer.Stop();
	        m_iac.StatusReport.LogData($"Process All Officials(buildAnchorList) elapsedTime: {timer.MsecFloat}", 1, MSGT.Body);
	        
	        // now, we are going to navigate to each page by finding and clicking each pagination link in turn
	        foreach (string sHref in plsHrefs)
	        {
		        timer.Reset();
		        timer.Start();
		        
		        string sXpath = $"//a[@href={ToXPath(sHref)}]";

                IWebElement anchor;

                try
                {
	                anchor = m_iac.WebControl.Driver.FindElement(By.XPath(sXpath));
                }
                catch
                {
                    anchor = null;
                }
                
                timer.Stop();
                m_iac.StatusReport.LogData($"Process All Officials(find item by xpath) elapsedTime: {timer.MsecFloat}", 1, MSGT.Body);

                timer.Reset();
                timer.Start();
                
                anchor.Click();

                timer.Stop();
                m_iac.StatusReport.LogData($"Process All Officials(anchor click) elapsedTime: {timer.MsecFloat}", 1, MSGT.Body);

                visit(o);
	        }
        }

        private PGL PglGetOfficialsFromWeb()
        {
            m_iac.EnsureLoggedIn();
            int i;

            PGL pgl = new PGL();

            ProcessAllOfficialPages(VOPC_PopulatePgl, pgl);

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
                    if (MessageBox.Show($"There are {plirsteUnmarked.Count} new officials.  Add these officials?", "ArbWeb", MessageBoxButtons.YesNo) ==
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
                sOutFile = $"{Environment.GetEnvironmentVariable("temp")}";
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
