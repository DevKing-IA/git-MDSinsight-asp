<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Routing.asp"-->
<%
UserNo = Request.Form("txtUserNo")
ActiveTab = Request.Form("txtTab")
'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM tblUsers where Userno = " & UserNo 
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_userFirstName = rs("userFirstName")
	Orig_userLastName = rs("userLastName")
	Orig_userEmail = rs("userEmail")
	Orig_userPassword = rs("userPassword")
	Orig_userLastLogin = rs("userLastLogin")
	Orig_userEnabled = rs("userEnabled")
	Orig_userDownloadEmail = rs("userDownloadEmail")
	Orig_userUpdateCalendar = rs("userUpdateCalendar")	
	Orig_userCanAuthSwaps = rs("userCanAuthSwaps")
	Orig_userReceivePartsRequestEmails = rs("userReceivePartsRequestEmails")
	Orig_userDisplayName = rs("userDisplayName")
	Orig_userCellNumber = rs("userCellNumber")
	Orig_userType = rs("userType")
	Orig_LoginLandingPage = rs("LoginLandingPageURL")
	Orig_userTruckNumber = rs("userTruckNumber")
	Orig_userFilterRoutes = rs("userFilterRoutes")
	Orig_userCRMAccessType  = rs("userCRMAccessType")
	Orig_userOrderAPIAccessType = rs("userOrderAPIAccessType")
	Orig_userCRMDeleteAccess  = rs("userCRMDeleteAccess")
	Orig_userProspectingAddEditAccess = rs("userProspectingAddEditAccess")
	Orig_userEmailSystemID = rs("userEmailSystemID")
	Orig_userEmailServer = rs("userEmailServer")
	Orig_userEmailSystemPass = rs("userEmailSystemPass")
	Orig_userSystemVMSID = rs("userVMS_ID")
	Orig_userSalesPersonNumber = rs("userSalesPersonNumber")
	Orig_userSalesPersonNumber2 = rs("userSalesPersonNumber2")
	Orig_userForceNextStopSelectionOverride = rs("userForceNextStopSelectionOverride")
	Orig_userNextStopNagMessageOverride = rs("userNextStopNagMessageOverride")
	Orig_userNextStopNagMinutes = rs("userNextStopNagMinutes")
	Orig_userNextStopNagIntervalMinutes = rs("userNextStopNagIntervalMinutes")
	Orig_userNextStopNagMessageMaxToSendPerStop = rs("userNextStopNagMessageMaxToSendPerStop")
	Orig_userNextStopNagMessageMaxToSendThisDriverPerDay = rs("userNextStopNagMessageMaxToSendThisDriverPerDay")
	Orig_userNextStopNagMessageSendMethod = rs("userNextStopNagMessageSendMethod")
	Orig_userNoActivityNagMessageOverride = rs("userNoActivityNagMessageOverride")
	Orig_userNoActivityNagMinutes = rs("userNoActivityNagMinutes")
	Orig_userNoActivityNagIntervalMinutes = rs("userNoActivityNagIntervalMinutes")
	Orig_userNoActivityNagMessageMaxToSendPerStop = rs("userNoActivityNagMessageMaxToSendPerStop")
	Orig_userNoActivityNagMessageMaxToSendPerDriverPerDay = rs("userNoActivityNagMessageMaxToSendPerDriverPerDay")
	Orig_userNoActivityNagMessageSendMethod = rs("userNoActivityNagMessageSendMethod")
	Orig_userNoActivityNagTimeOfDay = rs("userNoActivityNagTimeOfDay")

	
	Orig_userNoActivityNagMessageOverride_FS = rs("userNoActivityNagMessageOverride_FS")
	Orig_userNoActivityNagMinutes_FS = rs("userNoActivityNagMinutes_FS")
	Orig_userNoActivityNagIntervalMinutes_FS = rs("userNoActivityNagIntervalMinutes_FS")
	Orig_userNoActivityNagMessageMaxToSendPerStop_FS = rs("userNoActivityNagMessageMaxToSendPerStop_FS")
	Orig_userNoActivityNagMessageMaxToSendPerDriverPerDay_FS = rs("userNoActivityNagMessageMaxToSendPerDriverPerDay_FS")
	Orig_userNoActivityNagMessageSendMethod_FS = rs("userNoActivityNagMessageSendMethod_FS")
	Orig_userNoActivityNagTimeOfDay_FS = rs("userNoActivityNagTimeOfDay_FS")
	
	
	Orig_userLoginDisableAccessHolidays	= rs("userLoginDisableAccessHolidays")
	Orig_userInventoryControlAccessType = rs("userInventoryControlAccessType")
	Orig_userMobileInventoryControlAccess = rs("userMobileInventoryControlAccess")
	Orig_userCanEditEqpTablesOnFly = rs("userEditEqpOnTheFly")
	Orig_userEditCRMOnTheFly = rs("userEditCRMOnTheFly")
	Orig_userCreateEquipmentSymptomCodesOnTheFly = rs("userCreateEquipmentSymptomCodesOnTheFly")
	Orig_userCreateEquipmentProblemCodesOnTheFly = rs("userCreateEquipmentProblemCodesOnTheFly")
	Orig_userCreateEquipmentResolutionCodesOnTheFly = rs("userCreateEquipmentResolutionCodesOnTheFly")
	
	Orig_userLeftNavAPIModule = rs("userLeftNavAPIModule")
	Orig_userLeftNavBIModule = rs("userLeftNavBIModule")
	Orig_userLeftNavProspectingModule = rs("userLeftNavProspectingModule")
	Orig_userLeftNavCustomerServiceModule = rs("userLeftNavCustomerServiceModule")
	Orig_userLeftNavEquipmentModule = rs("userLeftNavEquipmentModule")
	Orig_userLeftNavInventoryControlModule = rs("userLeftNavInventoryControlModule")
	Orig_userLeftNavAccountsReceivableModule = rs("userLeftNavAccountsReceivableModule")
	Orig_userLeftNavAccountsPayableModule = rs("userLeftNavAccountsPayableModule")
	Orig_userLeftNavServiceModule = rs("userLeftNavServiceModule")
	Orig_userLeftNavRoutingModule = rs("userLeftNavRoutingModule")
	Orig_userLeftNavQuickbooksModule = rs("userLeftNavQuickbooksModule")
	Orig_userLeftNavFiltertraxModule = rs("userLeftNavFiltertraxModule")
	Orig_userLeftNavSystem = rs("userLeftNavSystem")

	Orig_userCreateNewServiceTicket = rs("userCreateNewServiceTicket")
	Orig_userAccessServiceDispatchCenter = rs("userAccessServiceDispatchCenter")
	Orig_userAccessServiceActionsModalButton = rs("userAccessServiceActionsModalButton")
	Orig_userAccessServiceDispatchButton = rs("userAccessServiceDispatchButton")
	Orig_userAccessServiceCloseCancelButton= rs("userAccessServiceCloseCancelButton")
	
End If


set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

FirstName = Request.Form("txtFirstName")
LastName = Request.Form("txtLastName")
LastName = Replace(LastName,"'","''")
DisplayName = Request.Form("txtDisplayName")
CellNumber = Request.Form("txtCellNumber")
Email= Request.Form("txtEmail")
Password = Request.Form("txtPassword")
Enabled = Request.Form("chkEnabled")
DownloadEmail = Request.Form("chkDownloadEmail")
UpdateCalendar = Request.Form("chkUpdateCalendar")
userCanAuthSwaps = Request.Form ("chkuserCanAuthSwaps")
userReceivePartsRequestEmails = Request.Form ("chkuserReceivePartsRequestEmails")
UserType = Request.Form("selUserType")
LoginLandingPage = Request.Form("selLoginLandingPage")
UserTruckNumber = Request.Form("selRouteNumber")
userFilterRoutes = Replace(Request.Form("txtRoutes")," ","")
userCRMAccessType = Request.Form("optCRMAccessType")
userOrderAPIAccessType = Request.Form("optOrderAPIAccessType")
chkCRMAddEditAccess = Request.Form("chkCRMAddEditAccess")
chkCRMDeleteAccess = Request.Form("chkCRMDeleteAccess")
userEmailSystemID = Request.Form("txtEmailSystemID")
userEmailSystemPass = Request.Form("txtEmailSystemPass")
userEmailServer = Request.Form("txtEmailServer")
userSystemVMSID = Request.Form("txtUserSystemVMSID")

If MUV_READ("biModuleOn") = "Enabled" Then 
	userSalesPersonNumber = Request.Form("seluserSalesPersonNumber")
	userSalesPersonNumber2 = Request.Form("seluserSalesPersonNumber2")
End If

userForceNextStopSelectionOverride = Request.Form("seluserForceNextStopSelectionOverride")
userNextStopNagMessageOverride = Request.Form("seluserNextStopNagMessageOverride")
userNextStopNagMinutes = Request.Form("seluserNextStopNagMinutes")
userNextStopNagIntervalMinutes = Request.Form("seluserNextStopNagIntervalMinutes")
userNextStopNagMessageMaxToSendPerStop = Request.Form("seluserNextStopNagMessageMaxToSendPerStop")
userNextStopNagMessageMaxToSendThisDriverPerDay = Request.Form("seluserNextStopNagMessageMaxToSendThisDriverPerDay")
userNextStopNagMessageSendMethod = Request.Form("seluserNextStopNagMessageSendMethod")

userNoActivityNagMessageOverride = Request.Form("seluserNoActivityNagMessageOverride")
userNoActivityNagMinutes = Request.Form("seluserNoActivityNagMinutes")
userNoActivityNagIntervalMinutes = Request.Form("seluserNoActivityNagIntervalMinutes")
userNoActivityNagMessageMaxToSendPerStop = Request.Form("seluserNoActivityNagMessageMaxToSendPerStop")
userNoActivityNagMessageMaxToSendPerDriverPerDay = Request.Form("seluserNoActivityNagMessageMaxToSendPerDriverPerDay")
userNoActivityNagMessageSendMethod = Request.Form("seluserNoActivityNagMessageSendMethod")
userNoActivityNagTimeOfDay = Request.Form("seluserNoActivityNagTimeOfDay")
chkDisableHolidayLogins = Request.Form("chkDisableHolidayLogins")


userNoActivityNagMessageOverride_FS = Request.Form("seluserNoActivityNagMessageOverride_FS")
userNoActivityNagMinutes_FS = Request.Form("seluserNoActivityNagMinutes_FS")
userNoActivityNagIntervalMinutes_FS = Request.Form("seluserNoActivityNagIntervalMinutes_FS")
userNoActivityNagMessageMaxToSendPerStop_FS = Request.Form("seluserNoActivityNagMessageMaxToSendPerStop_FS")
userNoActivityNagMessageMaxToSendPerDriverPerDay_FS = Request.Form("seluserNoActivityNagMessageMaxToSendPerDriverPerDay_FS")
userNoActivityNagMessageSendMethod_FS = Request.Form("seluserNoActivityNagMessageSendMethod_FS")
userNoActivityNagTimeOfDay_FS= Request.Form("seluserNoActivityNagTimeOfDay_FS")


userInventoryControlAccessType = Request.Form("optInvControlAccessType")
userMobileInventoryControlAccess = Request.Form("chkInvControlMobileAccess")
If userMobileInventoryControlAccess = "on" then userMobileInventoryControlAccess = 1 Else userMobileInventoryControlAccess = 0

userCanEditEqpTablesOnFly = Request.Form("chkUserCanEditEqpTablesOnFly")
If userCanEditEqpTablesOnFly = "on" then userCanEditEqpTablesOnFly = 1 Else userCanEditEqpTablesOnFly = 0

userCreateEquipmentSymptomCodesOnTheFly = Request.Form("chkUserCreateEquipmentSymptomCodesOnTheFly")
If userCreateEquipmentSymptomCodesOnTheFly = "on" then userCreateEquipmentSymptomCodesOnTheFly = 1 Else userCreateEquipmentSymptomCodesOnTheFly = 0

userCreateEquipmentProblemCodesOnTheFly = Request.Form("chkUserCreateEquipmentProblemCodesOnTheFly")
If userCreateEquipmentProblemCodesOnTheFly = "on" then userCreateEquipmentProblemCodesOnTheFly = 1 Else userCreateEquipmentProblemCodesOnTheFly = 0

userCreateEquipmentResolutionCodesOnTheFly = Request.Form("chkUserCreateEquipmentResolutionCodesOnTheFly")
If userCreateEquipmentResolutionCodesOnTheFly = "on" then userCreateEquipmentResolutionCodesOnTheFly = 1 Else userCreateEquipmentResolutionCodesOnTheFly = 0

userCreateNewServiceTicket = Request.Form("chkUserCreateNewServiceTicket")
If userCreateNewServiceTicket = "on" then userCreateNewServiceTicket = 1 Else userCreateNewServiceTicket = 0

userAccessServiceDispatchCenter = Request.Form("chkUserAccessServiceDispatchCenter")
If userAccessServiceDispatchCenter = "on" then userAccessServiceDispatchCenter = 1 Else userAccessServiceDispatchCenter = 0

userAccessServiceActionsModalButton = Request.Form("chkUserAccessServiceActionsModalButton")
If userAccessServiceActionsModalButton = "on" then userAccessServiceActionsModalButton = 1 Else userAccessServiceActionsModalButton = 0

userAccessServiceDispatchButton = Request.Form("chkUserAccessServiceDispatchButton")
If userAccessServiceDispatchButton = "on" then userAccessServiceDispatchButton = 1 Else userAccessServiceDispatchButton = 0

userAccessServiceCloseCancelButton = Request.Form("chkUserAccessServiceCloseCancelButton")
If userAccessServiceCloseCancelButton = "on" then userAccessServiceCloseCancelButton = 1 Else userAccessServiceCloseCancelButton = 0


If Enabled = "on" then Enabled =1 Else Enabled = 0

If userCanAuthSwaps = "on" then userCanAuthSwaps = 1 Else userCanAuthSwaps = 0
If userReceivePartsRequestEmails = "on" then userReceivePartsRequestEmails = 1 Else userReceivePartsRequestEmails = 0

userEditCRMOnTheFly = Request.Form("chkUserEditCRMOnTheFly")
If userEditCRMOnTheFly = "on" then userEditCRMOnTheFly = 1 Else userEditCRMOnTheFly = 0


If DownloadEmail = "on" then DownloadEmail = 1 Else DownloadEmail = 0
If UpdateCalendar = "on" then UpdateCalendar = 1 Else UpdateCalendar = 0
If chkCRMAddEditAccess = "on" then chkCRMAddEditAccess = 1 Else chkCRMAddEditAccess = 0
If chkCRMDeleteAccess = "on" then chkCRMDeleteAccess = 1 Else chkCRMDeleteAccess = 0
If chkDisableHolidayLogins = "on" then chkDisableHolidayLogins = 1 Else chkDisableHolidayLogins = 0

userLeftNavAPIModule = Request.Form("chkUserLeftNavAPIModule")
userLeftNavBIModule = Request.Form("chkUserLeftNavBIModule")
userLeftNavProspectingModule = Request.Form("chkUserLeftNavProspectingModule")
userLeftNavCustomerServiceModule = Request.Form("chkUserLeftNavCustomerServiceModule")
userLeftNavEquipmentModule = Request.Form("chkUserLeftNavEquipmentModule")
userLeftNavInventoryControlModule = Request.Form("chkUserLeftNavInventoryControlModule")
userLeftNavAccountsReceivableModule = Request.Form("chkUserLeftNavAccountsReceivableModule")
userLeftNavAccountsPayableModule = Request.Form("chkUserLeftNavAccountsPayableModule")
userLeftNavServiceModule = Request.Form("chkUserLeftNavServiceModule")
userLeftNavRoutingModule = Request.Form("chkUserLeftNavRoutingModule")
userLeftNavQuickbooksModule = Request.Form("chkUserLeftNavQuickbooksModule")
userLeftNavFiltertraxModule = Request.Form("chkUserLeftNavFiltertraxModule")
userLeftNavSystem = Request.Form("chkUserLeftNavSystem")

If userLeftNavAPIModule = "on" then userLeftNavAPIModule = 1 Else userLeftNavAPIModule = 0
If userLeftNavBIModule = "on" then userLeftNavBIModule = 1 Else userLeftNavBIModule = 0
If userLeftNavProspectingModule = "on" then userLeftNavProspectingModule = 1 Else userLeftNavProspectingModule = 0
If userLeftNavCustomerServiceModule = "on" then userLeftNavCustomerServiceModule = 1 Else userLeftNavCustomerServiceModule = 0
If userLeftNavEquipmentModule = "on" then userLeftNavEquipmentModule = 1 Else userLeftNavEquipmentModule = 0
If userLeftNavInventoryControlModule = "on" then userLeftNavInventoryControlModule = 1 Else userLeftNavInventoryControlModule = 0
If userLeftNavAccountsReceivableModule = "on" then userLeftNavAccountsReceivableModule = 1 Else userLeftNavAccountsReceivableModule = 0
If userLeftNavAccountsPayableModule = "on" then userLeftNavAccountsPayableModule = 1 Else userLeftNavAccountsPayableModule = 0
If userLeftNavServiceModule = "on" then userLeftNavServiceModule = 1 Else userLeftNavServiceModule = 0
If userLeftNavRoutingModule = "on" then userLeftNavRoutingModule = 1 Else userLeftNavRoutingModule = 0
If userLeftNavQuickbooksModule = "on" then userLeftNavQuickbooksModule = 1 Else userLeftNavQuickbooksModule = 0
If userLeftNavFiltertraxModule = "on" then userLeftNavFiltertraxModule = 1 Else userLeftNavFiltertraxModule = 0
If userLeftNavSystem = "on" then userLeftNavSystem = 1 Else userLeftNavSystem = 0


SQL = "UPDATE tblUsers SET "
SQL = SQL &  "userFirstName = '" & FirstName & "',"
SQL = SQL &  "userLastName = '" & LastName & "',"
SQL = SQL &  "userDisplayName = '" & DisplayName & "',"
SQL = SQL &  "userCellNumber = '" & CellNumber & "',"
SQL = SQL &  "userEmail = '" & Email & "',"
SQL = SQL &  "userPassword = '" & Password & "',"
SQL = SQL &  "userEmailSystemID = '" & userEmailSystemID & "',"
SQL = SQL &  "userEmailSystemPass = '" & userEmailSystemPass & "',"
SQL = SQL &  "userEmailServer = '" & userEmailServer & "',"
SQL = SQL &  "userVMS_ID = '" & userSystemVMSID & "',"
SQL = SQL &  "userType = '" & UserType & "',"
SQL = SQL &  "LoginLandingPageURL = '" & LoginLandingPage & "',"
SQL = SQL &  "userEnabled = " & Enabled & ","
SQL = SQL &  "userDownloadEmail = " & DownloadEmail & ","
SQL = SQL &  "userUpdateCalendar = " & UpdateCalendar & ","
SQL = SQL &  "userCanAuthSwaps = " & userCanAuthSwaps & ", "
SQL = SQL &  "userReceivePartsRequestEmails = " & userReceivePartsRequestEmails & ", "
SQL = SQL &  "userTruckNumber = '" & userTruckNumber & "', "
SQL = SQL &  "userFilterRoutes = '" & userFilterRoutes & "', "
SQL = SQL &  "userProspectingAddEditAccess = " & chkCRMAddEditAccess & ","
SQL = SQL &  "userCRMDeleteAccess = " & chkCRMDeleteAccess & ","
SQL = SQL &  "userCRMAccessType = '" & userCRMAccessType & "', "
SQL = SQL &  "userEditCRMOnTheFly = " & userEditCRMOnTheFly & ","
SQL = SQL &  "userOrderAPIAccessType = '" & userOrderAPIAccessType & "', "

SQL = SQL &  "userForceNextStopSelectionOverride = '" & userForceNextStopSelectionOverride & "' "
SQL = SQL &  "," & "userNextStopNagMessageOverride = '" & userNextStopNagMessageOverride & "' "
SQL = SQL &  "," & "userNextStopNagMinutes = " & userNextStopNagMinutes
SQL = SQL &  "," & "userNextStopNagIntervalMinutes = " & userNextStopNagIntervalMinutes
SQL = SQL &  "," & "userNextStopNagMessageMaxToSendPerStop = " & userNextStopNagMessageMaxToSendPerStop
SQL = SQL &  "," & "userNextStopNagMessageMaxToSendThisDriverPerDay = " & userNextStopNagMessageMaxToSendThisDriverPerDay
SQL = SQL &  "," & "userNextStopNagMessageSendMethod = '" & userNextStopNagMessageSendMethod & "' "

SQL = SQL &  "," & "userNoActivityNagMessageOverride = '" & userNoActivityNagMessageOverride & "' "
SQL = SQL &  "," & "userNoActivityNagMinutes = " & userNoActivityNagMinutes
SQL = SQL &  "," & "userNoActivityNagIntervalMinutes = " & userNoActivityNagIntervalMinutes 
SQL = SQL &  "," & "userNoActivityNagMessageMaxToSendPerStop = " & userNoActivityNagMessageMaxToSendPerStop 
SQL = SQL &  "," & "userNoActivityNagMessageMaxToSendPerDriverPerDay = " & userNoActivityNagMessageMaxToSendPerDriverPerDay
SQL = SQL &  "," &"userNoActivityNagMessageSendMethod = '" & userNoActivityNagMessageSendMethod & "' "
SQL = SQL &  "," &"userNoActivityNagTimeOfDay = '" & userNoActivityNagTimeOfDay & "' "

SQL = SQL &  "," & "userNoActivityNagMessageOverride_FS = '" & userNoActivityNagMessageOverride_FS & "' "
SQL = SQL &  "," & "userNoActivityNagMinutes_FS = " & userNoActivityNagMinutes_FS
SQL = SQL &  "," & "userNoActivityNagIntervalMinutes_FS = " & userNoActivityNagIntervalMinutes_FS 
SQL = SQL &  "," & "userNoActivityNagMessageMaxToSendPerStop_FS = " & userNoActivityNagMessageMaxToSendPerStop_FS
SQL = SQL &  "," & "userNoActivityNagMessageMaxToSendPerDriverPerDay_FS = " & userNoActivityNagMessageMaxToSendPerDriverPerDay_FS
SQL = SQL &  "," &"userNoActivityNagMessageSendMethod_FS = '" & userNoActivityNagMessageSendMethod_FS & "' "
SQL = SQL &  "," &"userNoActivityNagTimeOfDay_FS = '" & userNoActivityNagTimeOfDay_FS & "' "

SQL = SQL &  "," & "userLoginDisableAccessHolidays= " & chkDisableHolidayLogins & " "
SQL = SQL &  "," & "userInventoryControlAccessType= '" & userInventoryControlAccessType & "' "
SQL = SQL &  "," & "userMobileInventoryControlAccess= " & userMobileInventoryControlAccess & " "
SQL = SQL &  "," & "userEditEqpOnTheFly= " & userCanEditEqpTablesOnFly & " "

SQL = SQL &  "," & "userCreateEquipmentSymptomCodesOnTheFly = " & userCreateEquipmentSymptomCodesOnTheFly & " "
SQL = SQL &  "," & "userCreateEquipmentProblemCodesOnTheFly = " & userCreateEquipmentProblemCodesOnTheFly & " "
SQL = SQL &  "," & "userCreateEquipmentResolutionCodesOnTheFly = " & userCreateEquipmentResolutionCodesOnTheFly & " "

SQL = SQL &  "," & "userCreateNewServiceTicket = " & userCreateNewServiceTicket & " "
SQL = SQL &  "," & "userAccessServiceDispatchCenter = " & userAccessServiceDispatchCenter & " "
SQL = SQL &  "," & "userAccessServiceActionsModalButton = " & userAccessServiceActionsModalButton & " "
SQL = SQL &  "," & "userAccessServiceDispatchButton = " & userAccessServiceDispatchButton & " "
SQL = SQL &  "," & "userAccessServiceCloseCancelButton = " & userAccessServiceCloseCancelButton & " "

SQL = SQL &  "," & "userLeftNavAPIModule = " & userLeftNavAPIModule & " "
SQL = SQL &  "," & "userLeftNavBIModule = " & userLeftNavBIModule & " "
SQL = SQL &  "," & "userLeftNavProspectingModule = " & userLeftNavProspectingModule & " "
SQL = SQL &  "," & "userLeftNavCustomerServiceModule = " & userLeftNavCustomerServiceModule & " "
SQL = SQL &  "," & "userLeftNavEquipmentModule = " & userLeftNavEquipmentModule & " "
SQL = SQL &  "," & "userLeftNavInventoryControlModule = " & userLeftNavInventoryControlModule & " "
SQL = SQL &  "," & "userLeftNavAccountsReceivableModule = " & userLeftNavAccountsReceivableModule & " "
SQL = SQL &  "," & "userLeftNavAccountsPayableModule = " & userLeftNavAccountsPayableModule & " "
SQL = SQL &  "," & "userLeftNavServiceModule = " & userLeftNavServiceModule & " "
SQL = SQL &  "," & "userLeftNavRoutingModule = " & userLeftNavRoutingModule & " "
SQL = SQL &  "," & "userLeftNavQuickbooksModule = " & userLeftNavQuickbooksModule & " "
SQL = SQL &  "," & "userLeftNavFiltertraxModule = " & userLeftNavFiltertraxModule & " "
SQL = SQL &  "," & "userLeftNavSystem = " & userLeftNavSystem & " "

If MUV_READ("biModuleOn") = "Enabled" Then 
	SQL = SQL &  "," & "userSalesPersonNumber = " & userSalesPersonNumber & " "
	SQL = SQL &  "," & "userSalesPersonNumber2 = " & userSalesPersonNumber2 & " "
End If


SQL = SQL &  " WHERE userNo = " & UserNo
	
'Response.write(SQL)
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 

'Response.Write(SQL&"<br>")
	
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""

Description = "User " & Orig_userDisplayName & ": "

If Orig_userFirstName <> FirstName Then
	Description = Description & "First name changed from " & Orig_userFirstName & " to " & FirstName
End If
If Orig_userLastName <> LastName Then
	Description = Description & "  Last name changed from " & Orig_userLastName & " to " & LastName
End If
If Orig_userDisplayName <> DisplayName Then
	Description = Description & "  Display name changed from " & Orig_userDisplayName & " to " & DisplayName
End If
If Orig_userCellNumber <> CellNumber Then
	Description = Description & "  Cell number changed from " & Orig_userCellNumber & " to " & CellNumber
End If
If Orig_userEmail <> Email Then
	Description = Description & "  Email changed from " & Orig_userEmail & " to " & Email
End If
If Orig_userPassword  <> Password Then
	Description = Description & "  Password changed"
End If
If Orig_userEmailSystemID  <> userEmailSystemID Then
	Description = Description & "  System Email ID changed from " & Orig_userEmailSystemID & " to " & userEmailSystemID
End If
If Orig_userEmailSystemPass  <> userEmailSystemPass Then
	Description = Description & " System Email Password changed"
End If
If Orig_userSystemVMSID <> userSystemVMSID Then
	Description = Description & " System VMS ID changed"
End If
If Orig_userEmailServer <> userEmailServer Then
	Description = Description & "  System Email Server changed from " & Orig_userEmailServer & " to " & userEmailServer
End If
If Orig_userType <> UserType Then
	Description = Description & "  User type changed from " & GetTerm(Orig_userType) & " to " & GetTerm(userType)
End If
If Orig_LoginLandingPage <> LoginLandingPage Then
	Description = Description & "  Login landing page changed from " & Orig_LoginLandingPage & " to " & LoginLandingPage
End If
If Orig_userFilterRoutes <> userFilterRoutes Then
	Description = Description & "  Filter change routes changed from " & userFilterRoutes & " to " & userFilterRoutes
End If
If Orig_userTruckNumber <> UserTruckNumber Then
	Description = Description & "  The Route number field was changed changed from " & Orig_UserTruckNumber & " (" & GetDriverNameByTruckID(Orig_UserTruckNumber) & ") to " & UserTruckNumber & " (" & GetDriverNameByTruckID(UserTruckNumber)  & ") for the user: " & DisplayName
End If
If Orig_userCRMAccessType <> userCRMAccessType Then
	Description = Description & "  The " & GetTerm("Prospecting") & " access type was changed from " & Orig_userCRMAccessType & " to " & userCRMAccessType & " for the user: " & DisplayName
End If

If Orig_userOrderAPIAccessType <> userOrderAPIAccessType Then
	Description = Description & "  The order API access type was changed from " & Orig_userOrderAPIAccessType & " to " & userOrderAPIAccessType & " for the user: " & DisplayName
End If

If Orig_userSalesPersonNumber <> userSalesPersonNumber Then
	Description = Description & "  The " & GetTerm("Primary Salesman") & " number was changed from " & Orig_userSalesPersonNumber & " to " & userSalesPersonNumber & " for the user: " & DisplayName
End If

If Orig_userSalesPersonNumber2 <> userSalesPersonNumber2 Then
	Description = Description & "  The " & GetTerm("Secondary Salesman") & " number was changed from " & Orig_userSalesPersonNumber2 & " to " & userSalesPersonNumber2 & " for the user: " & DisplayName
End If

If Orig_userProspectingAddEditAccess <> userProspectingAddEditAccess Then
	Description = Description & "  The " & GetTerm("Prospecting") & " add/edit menu access permission was changed from " & Orig_userProspectingAddEditAccess & " to " & userProspectingAddEditAccess & " for the user: " & DisplayName
End If
If Orig_userCRMDeleteAccess <> userCRMDeleteAccess Then
	Description = Description & "  The " & GetTerm("Prospecting") & " delete " & GetTerm("Prospect") & " button access permission was changed from " & Orig_userCRMDeleteAccess & " to " & userCRMDeleteAccess & " for the user: " & DisplayName
End If
'Need to change the value of the bit fields for the purposes of the audit trail
If Orig_userEnabled = True Then Orig_userEnabled = "True" else Orig_userEnabled = "False"

If Orig_userDownloadEmail = True Then Orig_userDownloadEmail = "True" else Orig_userDownloadEmail = "False"
If Orig_userUpdateCalendar = True Then Orig_userUpdateCalendar = "True" else Orig_userUpdateCalendar = "False"



If Orig_userCanAuthSwaps = True Then Orig_userCanAuthSwaps = "True" else Orig_userCanAuthSwaps = "False"
If Orig_userReceivePartsRequestEmails = True Then Orig_userReceivePartsRequestEmails = "True" else Orig_userReceivePartsRequestEmails = "False"


If Enabled = 1 Then Enabled = "True" else Enabled = "False"
If Orig_userEnabled <> Enabled Then
	Description = Description & "  Enabled changed from " & Orig_userEnabled & " to " & Enabled
End If


If userCanAuthSwaps = 1 Then userCanAuthSwaps  = "True" else userCanAuthSwaps = "False"
If Orig_userCanAuthSwaps <> userCanAuthSwaps Then
	Description = Description & "  User can authorize equipment swaps changed from " & Orig_userCanAuthSwaps & " to " & userCanAuthSwaps 
End If

If userEditCRMOnTheFly = 1 Then userEditCRMOnTheFly = "True" else userEditCRMOnTheFly = "False"
If Orig_userEditCRMOnTheFly <> userEditCRMOnTheFly Then
	Description = Description & "  User can edit CRM on the fly changed from " & Orig_userEditCRMOnTheFly & " to " & userEditCRMOnTheFly
End If


If userReceivePartsRequestEmails = 1 Then userReceivePartsRequestEmails = "True" else userReceivePartsRequestEmails = "False"
If Orig_userReceivePartsRequestEmails <> userReceivePartsRequestEmails Then
	Description = Description & "  User receives parts request emails changed from " & Orig_userReceivePartsRequestEmails & " to " & userReceivePartsRequestEmails
End If


If DownloadEmail = 1 Then DownloadEmail = "True" else DownloadEmail = "False"
If Orig_userDownloadEmail  <> DownloadEmail Then
	Description = Description & "  Download email changed from " & Orig_userDownloadEmail  & " to " & DownloadEmail 
End If
If UpdateCalendar = 1 Then UpdateCalendar = "True" else UpdateCalendar = "False"
If Orig_userUpdateCalendar  <> UpdateCalendar Then
	Description = Description & "  Update calendar changed from " & Orig_userUpdateCalendar  & " to " & UpdateCalendar
End If

'Audit trail for nag alerts


If Orig_userNextStopNagMessageOverride <> userNextStopNagMessageOverride Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>Turn on 'No Next Stop' nag messages</strong> was changed from " & Orig_userNextStopNagMessageOverride & " to " & userNextStopNagMessageOverride & " for the user: " & DisplayName
End If
If Orig_userNextStopNagMinutes<> userNextStopNagMinutes Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>Send when the Next Stop has not been set for X minutes</strong> was changed from " & Orig_userNextStopNagMinutes & " to " & userNextStopNagMinutes & " for the user: " & DisplayName
End If
If Orig_userNextStopNagIntervalMinutes <> userNextStopNagIntervalMinutes Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>Continue to send 'No Next Stop' nag messages every X minutes</strong> was changed from " & Orig_userNextStopNagIntervalMinutes & " to " & userNextStopNagIntervalMinutes & " for the user: " & DisplayName
End If
If Orig_userNextStopNagMessageMaxToSendPerStop <> userNextStopNagMessageMaxToSendPerStop Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>Send a maximum of X 'No Next Stop' nag messages each time a 'No Next Stop' event occurs</strong> was changed from " & Orig_userNextStopNagMessageMaxToSendPerStop & " to " & userNextStopNagMessageMaxToSendPerStop & " for the user: " & DisplayName
End If
If Orig_userNextStopNagMessageMaxToSendThisDriverPerDay <> userNextStopNagMessageMaxToSendThisDriverPerDay Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>Send a maxium of X 'No Next Stop' nag messages to this driver on any given day</strong> was changed from " & Orig_userNextStopNagMessageMaxToSendThisDriverPerDay & " to " & userNextStopNagMessageMaxToSendThisDriverPerDay & " for the user: " & DisplayName
End If
If Orig_userNextStopNagMessageSendMethod <> userNextStopNagMessageSendMethod Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>'No Next Stop' nag messages send method</strong> was changed from " & Orig_userNextStopNagMessageSendMethod & " to " & userNextStopNagMessageSendMethod & " for the user: " & DisplayName
End If

If Orig_userNoActivityNagMessageOverride <> userNoActivityNagMessageOverride Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>Turn on 'No Activity' nag messages</strong> was changed from " & Orig_userNoActivityNagMessageOverride& " to " & userNoActivityNagMessageOverride& " for the user: " & DisplayName
End If
If Orig_userNoActivityNagMinutes <> userNoActivityNagMinutes Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>Send when there has been 'No Activity' for X minutes</strong> was changed from " & Orig_userNoActivityNagMinutes & " to " & userNoActivityNagMinutes& " for the user: " & DisplayName
End If
If Orig_userNoActivityNagIntervalMinutes <> userNoActivityNagIntervalMinutes Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>Continue to send 'No Activity' nag messages every X minutes</strong> was changed from " & Orig_userNoActivityNagIntervalMinutes & " to " & userNoActivityNagIntervalMinutes & " for the user: " & DisplayName
End If
If Orig_userNoActivityNagMessageMaxToSendPerStop <> userNoActivityNagMessageMaxToSendPerStop Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>Send a maximum of X 'No Activity' nag messages each time a 'No Next Stop' event occurs</strong> was changed from " & Orig_userNoActivityNagMessageMaxToSendPerStop & " to " & userNoActivityNagMessageMaxToSendPerStop & " for the user: " & DisplayName
End If
If Orig_userNoActivityNagMessageMaxToSendPerDriverPerDay <> userNoActivityNagMessageMaxToSendPerDriverPerDay Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>Send a maxium of X 'No Activity' nag messages to this driver on any given day</strong> was changed from " & Orig_userNoActivityNagMessageMaxToSendPerDriverPerDay & " to " & userNoActivityNagMessageMaxToSendPerDriverPerDay & " for the user: " & DisplayName
End If
If Orig_userNoActivityNagTimeOfDay <> userNoActivityNagTimeOfDay Then
	Description = Description & "  The " & GetTerm("routing") & "setting </strong>Start sending messages if there has been 'No Activity' by {time} </strong> was changed from " & Orig_userNoActivityNagTimeOfDay & " to " & userNoActivityNagTimeOfDay & " for the user: " & DisplayName
End If


If Orig_userNoActivityNagMessageOverride_FS <> userNoActivityNagMessageOverride_FS Then
	Description = Description & "  The " & GetTerm("Field Service") & "setting </strong>Turn on 'No Activity' nag messages</strong> was changed from " & Orig_userNoActivityNagMessageOverride_FS & " to " & userNoActivityNagMessageOverride_FS & " for the user: " & DisplayName
End If
If Orig_userNoActivityNagMinutes_FS <> userNoActivityNagMinutes_FS Then
	Description = Description & "  The " & GetTerm("Field Service") & "setting </strong>Send when there has been 'No Activity' for X minutes</strong> was changed from " & Orig_userNoActivityNagMinutes_FS & " to " & userNoActivityNagMinutes_FS & " for the user: " & DisplayName
End If
If Orig_userNoActivityNagIntervalMinutes_FS <> userNoActivityNagIntervalMinutes_FS Then
	Description = Description & "  The " & GetTerm("Field Service") & "setting </strong>Continue to send 'No Activity' nag messages every X minutes</strong> was changed from " & Orig_userNoActivityNagIntervalMinutes_FS & " to " & userNoActivityNagIntervalMinutes_FS & " for the user: " & DisplayName
End If
If Orig_userNoActivityNagMessageMaxToSendPerStop_FS <> userNoActivityNagMessageMaxToSendPerStop_FS Then
	Description = Description & "  The " & GetTerm("Field Service") & "setting </strong>Send a maximum of X 'No Activity' nag messages each time a 'No Next Stop' event occurs</strong> was changed from " & Orig_userNoActivityNagMessageMaxToSendPerStop_FS & " to " & userNoActivityNagMessageMaxToSendPerStop_FS & " for the user: " & DisplayName
End If
If Orig_userNoActivityNagMessageMaxToSendPerDriverPerDay_FS <> userNoActivityNagMessageMaxToSendPerDriverPerDay_FS Then
	Description = Description & "  The " & GetTerm("Field Service") & "setting </strong>Send a maxium of X 'No Activity' nag messages to this driver on any given day</strong> was changed from " & Orig_userNoActivityNagMessageMaxToSendPerDriverPerDay_FS & " to " & userNoActivityNagMessageMaxToSendPerDriverPerDay_FS & " for the user: " & DisplayName
End If
If Orig_userNoActivityNagTimeOfDay_FS <> userNoActivityNagTimeOfDay_FS Then
	Description = Description & "  The " & GetTerm("Field Service") & "setting </strong>Start sending messages if there has been 'No Activity' by {time} </strong> was changed from " & Orig_userNoActivityNagTimeOfDay_FS & " to " & userNoActivityNagTimeOfDay_FS & " for the user: " & DisplayName
End If




If userLoginDisableAccessHolidays = 1 Then userLoginDisableAccessHolidays = "True" else userLoginDisableAccessHolidays = "False"
If Orig_userLoginDisableAccessHolidays <> userLoginDisableAccessHolidays Then
	Description = Description & " Disable login access on holidays was changed from " & Orig_userLoginDisableAccessHolidays & " to " & userLoginDisableAccessHolidays
End If

If Orig_userInventoryControlAccessType <> userInventoryControlAccessType Then
	Description = Description & "  The " & GetTerm("Inventory Control") & " access type was changed from " & Orig_userInventoryControlAccessType & " to " & userInventoryControlAccessType & " for the user: " & DisplayName
End If
If userMobileInventoryControlAccess = 1 Then userMobileInventoryControlAccess = "True" else userMobileInventoryControlAccess = "False"
If Orig_userInventoryControlAccessType <> userMobileInventoryControlAccess Then
	Description = Description & " Mobile " & GetTerm("Inventory Control") & " access changed from " & Orig_userInventoryControlAccessType & " to " & userMobileInventoryControlAccess
End If

If userCanEditEqpTablesOnFly = 1 Then userCanEditEqpTablesOnFly = "True" else userCanEditEqpTablesOnFly= "False"
If Orig_userCanEditEqpTablesOnFly <> userCanEditEqpTablesOnFly Then
	Description = Description & " " & GetTerm("Equipment") & " User Can Edit Equipment Tables on the Fly access changed from " & Orig_userCanEditEqpTablesOnFly & " to " & userCanEditEqpTablesOnFly
End If

If userCreateEquipmentSymptomCodesOnTheFly = 1 Then userCreateEquipmentSymptomCodesOnTheFly = "True" else userCreateEquipmentSymptomCodesOnTheFly = "False"
If Orig_userCreateEquipmentSymptomCodesOnTheFly <> userCreateEquipmentSymptomCodesOnTheFly Then
	Description = Description & " " & GetTerm("Service") & " User Can Edit " & GetTerm("Service") & " " & GetTerm("Equipment") & " Symptom Codes On The Fly Access changed from " & Orig_userCreateEquipmentSymptomCodesOnTheFly & " to " & userCreateEquipmentSymptomCodesOnTheFly
End If

If userCreateEquipmentProblemCodesOnTheFly = 1 Then userCreateEquipmentProblemCodesOnTheFly = "True" else userCreateEquipmentProblemCodesOnTheFly = "False"
If Orig_userCreateEquipmentProblemCodesOnTheFly <> userCreateEquipmentProblemCodesOnTheFly Then
	Description = Description & " " & GetTerm("Service") & " User Can Edit " & GetTerm("Service") & " " & GetTerm("Equipment") & " Problem Codes On The Fly Access changed from " & Orig_userCreateEquipmentProblemCodesOnTheFly & " to " & userCreateEquipmentProblemCodesOnTheFly
End If


If userCreateEquipmentResolutionCodesOnTheFly = 1 Then userCreateEquipmentResolutionCodesOnTheFly = "True" else userCreateEquipmentResolutionCodesOnTheFly = "False"
If Orig_userCreateEquipmentResolutionCodesOnTheFly <> userCreateEquipmentResolutionCodesOnTheFly Then
	Description = Description & " " & GetTerm("Service") & " User Can Edit " & GetTerm("Service") & " " & GetTerm("Equipment") & " Resolution Codes On The Fly Access changed from " & Orig_userCreateEquipmentResolutionCodesOnTheFly & " to " & userCreateEquipmentResolutionCodesOnTheFly
End If




If Orig_userLeftNavAPIModule = True Then Orig_userLeftNavAPIModule = "True" else Orig_userLeftNavAPIModule = "False"
If Orig_userLeftNavBIModule = True Then Orig_userLeftNavBIModule = "True" else Orig_userLeftNavBIModule = "False"
If Orig_userLeftNavProspectingModule = True Then Orig_userLeftNavProspectingModule = "True" else Orig_userLeftNavProspectingModule = "False"
If Orig_userLeftNavCustomerServiceModule = True Then Orig_userLeftNavCustomerServiceModule = "True" else Orig_userLeftNavCustomerServiceModule = "False"
If Orig_userLeftNavEquipmentModule = True Then Orig_userLeftNavEquipmentModule = "True" else Orig_userLeftNavEquipmentModule = "False"
If Orig_userLeftNavInventoryControlModule = True Then Orig_userLeftNavInventoryControlModule = "True" else Orig_userLeftNavInventoryControlModule = "False"
If Orig_userLeftNavAccountsReceivableModule = True Then Orig_userLeftNavAccountsReceivableModule = "True" else Orig_userLeftNavAccountsReceivableModule = "False"
If Orig_userLeftNavAccountsPayableModule = True Then Orig_userLeftNavAccountsPayableModule = "True" else Orig_userLeftNavAccountsPayableModule = "False"
If Orig_userLeftNavServiceModule = True Then Orig_userLeftNavServiceModule = "True" else Orig_userLeftNavServiceModule = "False"
If Orig_userLeftNavRoutingModule = True Then Orig_userLeftNavRoutingModule = "True" else Orig_userLeftNavRoutingModule = "False"
If Orig_userLeftNavQuickbooksModule = True Then Orig_userLeftNavQuickbooksModule = "True" else Orig_userLeftNavQuickbooksModule = "False"
If Orig_userLeftNavFiltertraxModule = True Then Orig_userLeftNavFiltertraxModule = "True" else Orig_userLeftNavFiltertraxModule = "False"
If Orig_userLeftNavSystem = True Then Orig_userLeftNavSystem = "True" else Orig_userLeftNavSystem = "False"


If userLeftNavAPIModule = 1 Then userLeftNavAPIModule = "True" else userLeftNavAPIModule = "False"
If Orig_userLeftNavAPIModule <> userLeftNavAPIModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("API") & " Module Menu Link changed from " & Orig_userLeftNavAPIModule & " to " & userLeftNavAPIModule
End If

If userLeftNavBIModule = 1 Then userLeftNavBIModule = "True" else userLeftNavBIModule = "False"
If Orig_userLeftNavBIModule <> userLeftNavBIModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("Business Intelligence") & " Module Menu Link changed from " & Orig_userLeftNavBIModule & " to " & userLeftNavBIModule
End If

If userLeftNavProspectingModule = 1 Then userLeftNavProspectingModule = "True" else userLeftNavProspectingModule = "False"
If Orig_userLeftNavProspectingModule <> userLeftNavProspectingModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("Prospecting") & " Module Menu Link changed from " & Orig_userLeftNavProspectingModule & " to " & userLeftNavProspectingModule 
End If

If userLeftNavCustomerServiceModule = 1 Then userLeftNavCustomerServiceModule = "True" else userLeftNavCustomerServiceModule = "False"
If Orig_userLeftNavCustomerServiceModule <> userLeftNavCustomerServiceModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("Customer Service") & " Module Menu Link changed from " & Orig_userLeftNavCustomerServiceModule & " to " & userLeftNavCustomerServiceModule 
End If

If userLeftNavEquipmentModule = 1 Then userLeftNavEquipmentModule = "True" else userLeftNavEquipmentModule = "False"
If Orig_userLeftNavEquipmentModule <> userLeftNavEquipmentModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("Equipment") & " Module Menu Link changed from " & Orig_userLeftNavEquipmentModule & " to " & userLeftNavEquipmentModule 
End If

If userLeftNavInventoryControlModule = 1 Then userLeftNavInventoryControlModule = "True" else userLeftNavInventoryControlModule = "False"
If Orig_userLeftNavInventoryControlModule <> userLeftNavInventoryControlModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("Inventory Control") & " Module Menu Link changed from " & Orig_userLeftNavInventoryControlModule & " to " & userLeftNavInventoryControlModule 
End If

If userLeftNavAccountsReceivableModule = 1 Then userLeftNavAccountsReceivableModule = "True" else userLeftNavAccountsReceivableModule = "False"
If Orig_userLeftNavAccountsReceivableModule <> userLeftNavAccountsReceivableModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("Accounts Receivable") & " Module Menu Link changed from " & Orig_userLeftNavAccountsReceivableModule & " to " & userLeftNavAccountsReceivableModule 
End If

If userLeftNavAccountsPayableModule = 1 Then userLeftNavAccountsPayableModule = "True" else userLeftNavAccountsPayableModule = "False"
If Orig_userLeftNavAccountsPayableModule <> userLeftNavAccountsPayableModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("Accounts Payable") & " Module Menu Link changed from " & Orig_userLeftNavAccountsPayableModule & " to " & userLeftNavAccountsPayableModule 
End If

If userLeftNavServiceModule = 1 Then userLeftNavServiceModule = "True" else userLeftNavServiceModule = "False"
If Orig_userLeftNavServiceModule <> userLeftNavServiceModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("Service") & " Module Menu Link changed from " & Orig_userLeftNavServiceModule & " to " & userLeftNavServiceModule 
End If

If userLeftNavRoutingModule = 1 Then userLeftNavRoutingModule = "True" else userLeftNavRoutingModule = "False"
If Orig_userLeftNavRoutingModule <> userLeftNavRoutingModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("Routing") & " Module Menu Link changed from " & Orig_userLeftNavRoutingModule & " to " & userLeftNavRoutingModule 
End If

If userLeftNavQuickbooksModule = 1 Then userLeftNavQuickbooksModule = "True" else userLeftNavQuickbooksModule = "False"
If Orig_userLeftNavQuickbooksModule <> userLeftNavQuickbooksModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("QuickBooks") & " Module Menu Link changed from " & Orig_userLeftNavQuickbooksModule & " to " & userLeftNavQuickbooksModule 
End If

If userLeftNavFiltertraxModule = 1 Then userLeftNavFiltertraxModule = "True" else userLeftNavFiltertraxModule = "False"
If Orig_userLeftNavFiltertraxModule <> userLeftNavFiltertraxModule Then
	Description = Description & " User Can View Left Navigation " & GetTerm("FilterTrax") & " Module Menu Link changed from " & Orig_userLeftNavFiltertraxModule & " to " & userLeftNavFiltertraxModule 
End If

If userLeftNavSystem = 1 Then userLeftNavSystem = "True" else userLeftNavSystem = "False"
If Orig_userLeftNavSystem <> userLeftNavSystem Then
	Description = Description & " User Can View Left Navigation " & GetTerm("System") & " Menu Link changed from " & Orig_userLeftNavSystem & " to " & userLeftNavSystem 
End If


If Orig_userCreateNewServiceTicket = True Then Orig_userCreateNewServiceTicket = "True" else Orig_userCreateNewServiceTicket = "False"
If Orig_userAccessServiceDispatchCenter = True Then Orig_userAccessServiceDispatchCenter = "True" else Orig_userAccessServiceDispatchCenter = "False"
If Orig_userAccessServiceActionsModalButton = True Then Orig_userAccessServiceActionsModalButton = "True" else Orig_userAccessServiceActionsModalButton = "False"
If Orig_userAccessServiceDispatchButton = True Then Orig_userAccessServiceDispatchButton = "True" else Orig_userAccessServiceDispatchButton = "False"
If Orig_userAccessServiceCloseCancelButton = True Then Orig_userAccessServiceCloseCancelButton = "True" else Orig_userAccessServiceCloseCancelButton = "False"

If userCreateNewServiceTicket = 1 Then userCreateNewServiceTicket = "True" else userCreateNewServiceTicket = "False"
If Orig_userCreateNewServiceTicket <> userCreateNewServiceTicket Then
	Description = Description & " User Can Create New " & GetTerm("Service") & " Ticket changed from " & Orig_userCreateNewServiceTicket & " to " & userCreateNewServiceTicket 
End If

If userAccessServiceDispatchCenter = 1 Then userAccessServiceDispatchCenter = "True" else userAccessServiceDispatchCenter = "False"
If Orig_userAccessServiceDispatchCenter <> userAccessServiceDispatchCenter Then
	Description = Description & " User Can Access " & GetTerm("Service") & " Dispatch Center changed from " & Orig_userAccessServiceDispatchCenter & " to " & userAccessServiceDispatchCenter 
End If

If userAccessServiceActionsModalButton = 1 Then userAccessServiceActionsModalButton = "True" else userAccessServiceActionsModalButton = "False"
If Orig_userAccessServiceActionsModalButton <> userAccessServiceActionsModalButton Then
	Description = Description & " User Can Access " & GetTerm("Service") & " Actions Button changed from " & Orig_userAccessServiceActionsModalButton & " to " & userAccessServiceActionsModalButton 
End If

If userAccessServiceDispatchButton = 1 Then userAccessServiceDispatchButton = "True" else userAccessServiceDispatchButton = "False"
If Orig_userAccessServiceDispatchButton <> userAccessServiceDispatchButton Then
	Description = Description & " User Can Access " & GetTerm("Service") & " Dispatch Button changed from " & Orig_userAccessServiceDispatchButton & " to " & userAccessServiceDispatchButton 
End If

If userAccessServiceCloseCancelButton = 1 Then userAccessServiceCloseCancelButton = "True" else userAccessServiceCloseCancelButton = "False"
If Orig_userAccessServiceCloseCancelButton <> userAccessServiceCloseCancelButton Then
	Description = Description & " User Can Access " & GetTerm("Service") & " Close/Cancel Button changed from " & Orig_userAccessServiceCloseCancelButton & " to " & userAccessServiceCloseCancelButton 
End If


CreateAuditLogEntry "User Edited","User Edited","Major",0,Description

Response.Redirect("main.asp#" & ActiveTab)

%>















