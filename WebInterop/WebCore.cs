using System;
using System.Collections.Generic;
using System.Windows.Forms;
using TCore.Util;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
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
        public const string _s_OfficialsView = "https://www1.arbitersports.com/Assigner/OfficialsView.aspx"; // ok2010u
        public const string _s_Announcements = "https://www1.arbitersports.com/Shared/AnnouncementsEdit.aspx"; // ok2015
        public const string _s_ContactsView = "https://www1.arbitersports.com/Assigner/ContactsView.aspx"; // ok2018

        public const string s_RightsEdit = "https://www1.arbitersports.com/Assigner/RightsEdit.aspx"; // ok2024

        public const string _s_TeamsView = "https://www1.arbitersports.com/assigner/TeamsView.aspx"; // ok2023

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

        public const string
            _s_Home_Input_Email = "Email"; // "ctl00$ContentHolder$pgeSignIn$conSignIn$txtEmail"; // ctl00$ucMiniLogin$UsernameTextBox"; // ok2023

        public const string _s_Home_Input_Password = "Password"; // txtPassword"; // ctl00$ucMiniLogin$PasswordTextBox"; // ok2023

        public const string _sid_Home_Button_SignIn = "next"; // ctl00$ucMiniLogin$SignInButton"; // ok2023
        // public const string _s_Home_Button_SignIn = "ctl00$ContentHolder$pgeSignIn$conSignIn$btnSignIn"; // ctl00$ucMiniLogin$SignInButton"; // ok2016

        public const string _sid_Home_MultiFactor_Label = "mfaEnroll_label"; // ok2023
        public const string _sid_Home_MfaEnroll_False = "mfaEnroll_false"; // ok2023
        public const string _sid_Home_Mfa_continue = "continue"; // ok2023

        public const string _sid_Home_Div_PnlAccounts = "ctl00_ContentHolder_pgeDefault_conDefault_pnlAccounts"; // ok2010
        public const string _sid_Home_MessagingText = "ctl00_MessagingText"; // ok2022
        public const string _sid_Home_Anchor_NeedHelpLink = "ctl00_PageHelpTextLink"; // not ok 2022

        public const string _sid_Home_LoggedInUserId = "LoggedInUserId"; // ok2023

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

        public const string _sid_OfficialsView_IsReadyStatusPrefix = "IsReadyStatusText";
        public const string _sid_OfficialsView_EditAccount_Email0 = "emails_0_email1";
        public const string _sid_OfficialsView_EditAccount_FirstName = "firstName";
        public const string _sid_OfficialsView_EditAccount_NickName = "nickName";
        public const string _sid_OfficialsView_EditAccount_LastName = "lastName";
        public const string _sid_OfficialsView_EditAccount_MiddleName = "middleName";
        public const string _sid_OfficialsView_EditAccount_Address1 = "address1";
        public const string _sid_OfficialsView_EditAccount_Address2 = "address2";
        public const string _sid_OfficialsView_EditAccount_ZipCode = "postalCode";
        public const string _sid_OfficialsView_EditAccount_CityName = "cityName";
        public const string _sid_OfficialsView_EditAccount_State = "stateId";
        public const string _sid_OfficialsView_EditAccount_ButtonSave = "save-changes-button";
        public const string _sid_OfficialsView_EditAccount_ButtonCancel = "cancel-changes-button";
        public const string _sid_OfficialsView_EditAccount_BirthDate = "dateOfBirth";

        public const string _xpath_modalDialogRoot = "//div[@class='ant-modal-root']";
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
        public const string
            _s_Announcements_Button_Edit_Prefix =
                "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$"; // the control will be "ctl##"

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

        private const string
            _s_Announcements_Button_Save_Prefix =
                "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$"; // the control will be "ctl##"

        private const string _s_Announcements_Button_Save_Suffix = "$btnSave";

        public const string _sid_cke_Prefix = "cke_";
        public const string _sid_Announcements_Textarea_Text_Prefix = "ctl00_ContentHolder_pgeAnnouncementsEdit_conAnnouncementsEdit_dgAnnouncements_";
        public const string _sid_Announcements_Textarea_Text_Suffix = "_txtAnnouncement";

        public const string _s_Announcements_Textarea_Text_Prefix = "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$";
        public const string _s_Announcements_Textarea_Text_Suffix = "$txtAnnouncement";

        public const string _sid_Announcements_Button_Edit_Prefix = "ctl00_ContentHolder_pgeAnnouncementsEdit_conAnnouncementsEdit_dgAnnouncements_";
        public const string _sid_Announcements_Button_Edit_Suffix = "_btnEdit";

        public const string _sid_Announcements_Button_ToAssigners_Prefix =
            "ctl00_ContentHolder_pgeAnnouncementsEdit_conAnnouncementsEdit_dgAnnouncements_"; // ok2024

        public const string _sid_Announcements_Button_ToAssigners_Suffix = "_chkToAssigners_Edit"; // ok2024

        public const string _sid_Announcements_Button_ToOfficials_Prefix =
            "ctl00_ContentHolder_pgeAnnouncementsEdit_conAnnouncementsEdit_dgAnnouncements_"; // ok2024

        public const string _sid_Announcements_Button_ToOfficials_Suffix = "_ddlFilters"; // ok2024

        public const string _s_Announcements_Button_ToOfficials_Prefix =
            "ctl00$ContentHolder$pgeAnnouncementsEdit$conAnnouncementsEdit$dgAnnouncements$"; // ok2024

        public const string _s_Announcements_Button_ToOfficials_Suffix = "$ddlFilters"; // ok2024

        public const string _sid_Announcements_Button_ToContacts_Prefix =
            "ctl00_ContentHolder_pgeAnnouncementsEdit_conAnnouncementsEdit_dgAnnouncements_"; // ok2024

        public const string _sid_Announcements_Button_ToContacts_Suffix = "_chkToContacts_Edit"; // ok2024

        public const string _sid_Announcements_Button_Save_Prefix = "ctl00_ContentHolder_pgeAnnouncementsEdit_conAnnouncementsEdit_dgAnnouncements_";
        public const string _sid_Announcements_Button_Save_Suffix = "_btnSave";

        public const string _sid_Announcements_Button_Cancel_Prefix = "ctl00_ContentHolder_pgeAnnouncementsEdit_conAnnouncementsEdit_dgAnnouncements_";
        public const string _sid_Announcements_Button_Cancel_Suffix = "_btnCancel";

        public const string _sid_Login_Span_Type_Prefix = "ctl00_ContentHolder_pgeDefault_conDefault_dgAccounts_";
        private const string _sid_Login_Span_Type_Suffix = "_lblType2";

        public const string _sid_Login_Anchor_TypeLink_Prefix = "ctl00_ContentHolder_pgeDefault_conDefault_dgAccounts_";
        public const string _sid_Login_Anchor_TypeLink_Suffix = "_UserTypeLink";

        public const string s_MiscField_EditControlSubstring = "txtMiscFieldValue";

        public const string s_RightsEdit_PermissionsTarget = "ctl00$ContentHolder$pgeRightsEdit$conRightsEdit$ddlUserTypes"; // ok2024
        public const string sid_RightsEdit_PermissionsTarget = "ctl00_ContentHolder_pgeRightsEdit_conRightsEdit_ddlUserTypes"; // ok2024

        public const string s_RightsEdit_AllowUsersTo = "ctl00$ContentHolder$pgeRightsEdit$conRightsEdit$ddlRights"; // ok2024
        public const string sid_RightsEdit_AllowUsersTo = "ctl00_ContentHolder_pgeRightsEdit_conRightsEdit_ddlRights"; // ok2024

        public const string s_RightsEdit_WithoutPermissions = "ctl00$ContentHolder$pgeRightsEdit$conRightsEdit$lstNotAllowed"; // ok2024
        public const string sid_RightsEdit_WithoutPermissions = "ctl00_ContentHolder_pgeRightsEdit_conRightsEdit_lstNotAllowed"; // ok2024

        public const string s_RightsEdit_WithPermissions = "ctl00$ContentHolder$pgeRightsEdit$conRightsEdit$lstAllowed"; // ok2024
        public const string sid_RightsEdit_WithPermissions = "ctl00_ContentHolder_pgeRightsEdit_conRightsEdit_lstAllowed"; // ok2024

        public const string s_RightsEdit_RemovePermission = "ctl00$ContentHolder$pgeRightsEdit$conRightsEdit$btnDeny"; // ok2024
        public const string sid_RightsEdit_RemovePermission = "ctl00_ContentHolder_pgeRightsEdit_conRightsEdit_btnDeny"; // ok2024

        public const string s_RightsEdit_AddPermission = "ctl00$ContentHolder$pgeRightsEdit$conRightsEdit$btnAllow"; // ok2024
        public const string sid_RightsEdit_AddPermission = "ctl00_ContentHolder_pgeRightsEdit_conRightsEdit_btnAllow"; // ok2024

//        public const string s_ = ""; // ok2024
//        public const string sid_ = ""; // ok2024

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
                    rgs[iT] = (string)chlbx.Items[i];
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
                rgs[i] = (string)chlbx.Items[iChecked];
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
                rgs[i++] = (string)chlbx.Items[iForceToggle];

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
            private string m_sidControlExtra; // this is usually something like the Choice element ID
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
            m_expectedFullNameTemplates = new string[] { sFullExpectedName };
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

            wkb = app.Workbooks.Open(
                sExcelFile,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing);

            if (wkb != null)
            {
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)wkb.Worksheets[1];
                sheet.Select(Type.Missing);
                wkb.SaveAs(
                    sTargetCsvFile,
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV,
                    missing,
                    missing,
                    missing,
                    missing,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing);
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

            wkb = app.Workbooks.Open(
                sFile,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing,
                missing);

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
                wkb.SaveAs(
                    sOutFile,
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV,
                    missing,
                    missing,
                    missing,
                    missing,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing);
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
                    sFilterOptionValue =
                        m_appContext.WebControl.GetOptionValueFromFilterOptionTextForControlName(m_sSelectFilterControlName, m_sFilterOptionTextReq);
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

    public class OfficialLinkInfo
    {
        public string sOfficialEditLink = null;
        public string sOfficialID = null;
        public string sEmail = null;

        // PageNavLink is the pagination link (from the main page) that is required
        // to get to the page this official is on
        public string PageNavLink = "";
    }
}
