<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->

<%
ActiveTab = Request.Form("txtTab")
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
userCanAuthSwaps = Request.Form("chkuserCanAuthSwaps")
userReceivePartsRequestEmails = Request.Form("chkuserReceivePartsRequestEmails")
UserType = Request.Form("selUserType")
LoginLandingPage = Request.Form("selLoginLandingPage")
UserTruckNumber = Request.Form("selRouteNumber")
FilterRoutes = Replace(Request.Form("txtRoutes")," ","")
userCRMAccessType = Request.Form("optCRMAccessType")
userOrderAPIAccessType = Request.Form("optOrderAPIAccessType")
chkCRMAddEditAccess = Request.Form("chkCRMAddEditAccess")
chkCRMDeleteAccess = Request.Form("chkCRMDeleteAccess")
If MUV_READ("biModuleOn") = "Enabled" Then 
	userSalesPersonNumber = Request.Form("seluserSalesPersonNumber")
	userSalesPersonNumber2 = Request.Form("seluserSalesPersonNumber2")
End If
userInventoryControlAccessType = Request.Form("optInvControlAccessType")
userMobileInventoryControlAccess = Request.Form("chkInvControlMobileAccess")
If userMobileInventoryControlAccess = "on" then userMobileInventoryControlAccess = 1 Else userMobileInventoryControlAccess = 0

userCanEditEqpTablesOnFly = Request.Form("chkUserCanEditEqpTablesOnFly")
If userCanEditEqpTablesOnFly = "on" then userCanEditEqpTablesOnFly = 1 Else userCanEditEqpTablesOnFly = 0

userEditCRMOnTheFly = Request.Form("chkUserEditCRMOnTheFly")
If userEditCRMOnTheFly = "on" then userEditCRMOnTheFly = 1 Else userEditCRMOnTheFly = 0

userCreateEquipmentSymptomCodesOnTheFly = Request.Form("chkUserCreateEquipmentSymptomCodesOnTheFly")
If userCreateEquipmentSymptomCodesOnTheFly = "on" then userCreateEquipmentSymptomCodesOnTheFly = 1 Else userCreateEquipmentSymptomCodesOnTheFly = 0

userCreateEquipmentResolutionCodesOnTheFly = Request.Form("chkUserCreateEquipmentResolutionCodesOnTheFly")
If userCreateEquipmentResolutionCodesOnTheFly = "on" then userCreateEquipmentResolutionCodesOnTheFly = 1 Else userCreateEquipmentResolutionCodesOnTheFly = 0

userCreateEquipmentProblemCodesOnTheFly = Request.Form("chkUserCreateEquipmentProblemCodesOnTheFly")
If userCreateEquipmentProblemCodesOnTheFly = "on" then userCreateEquipmentProblemCodesOnTheFly = 1 Else userCreateEquipmentProblemCodesOnTheFly = 0


userEmailSystemID = Request.Form("txtEmailSystemID")
userEmailSystemPass = Request.Form("txtEmailSystemPass")
userEmailServer = Request.Form("txtEmailServer")
userSystemVMSID = Request.Form("txtUserSystemVMSID")
If Enabled = "on" then Enabled =1 Else Enabled = 0

If userCanAuthSwaps = "on" then userCanAuthSwaps = 1 Else userCanAuthSwaps = 0
If userReceivePartsRequestEmails = "on" then userReceivePartsRequestEmails = 1 Else userReceivePartsRequestEmails = 0

If DownloadEmail = "on" then DownloadEmail = 1 Else DownloadEmail = 0
If UpdateCalendar = "on" then UpdateCalendar = 1 Else UpdateCalendar = 0
If chkCRMAddEditAccess = "on" then chkCRMAddEditAccess = 1 Else chkCRMAddEditAccess = 0
If chkCRMDeleteAccess = "on" then chkCRMDeleteAccess = 1 Else chkCRMDeleteAccess = 0

ForceNextStop = Request.Form("selForceNextStop")
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


userNoActivityNagMessageOverride_FS = Request.Form("seluserNoActivityNagMessageOverride_FS")
userNoActivityNagMinutes_FS = Request.Form("seluserNoActivityNagMinutes_FS")
userNoActivityNagIntervalMinutes_FS = Request.Form("seluserNoActivityNagIntervalMinutes_FS")
userNoActivityNagMessageMaxToSendPerStop_FS = Request.Form("seluserNoActivityNagMessageMaxToSendPerStop_FS")
userNoActivityNagMessageMaxToSendPerDriverPerDay_FS = Request.Form("seluserNoActivityNagMessageMaxToSendPerDriverPerDay_FS")
userNoActivityNagMessageSendMethod_FS = Request.Form("seluserNoActivityNagMessageSendMethod_FS")
userNoActivityNagTimeOfDay_FS= Request.Form("seluserNoActivityNagTimeOfDay_FS")

chkDisableHolidayLogins = Request.Form("chkDisableHolidayLogins")
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
userCreateNewServiceTicket = Request.Form("chkUserCreateNewServiceTicket")
userAccessServiceDispatchCenter = Request.Form("chkUserAccessServiceDispatchCenter")
userAccessServiceActionsModalButton = Request.Form("chkUserAccessServiceActionsModalButton")
userAccessServiceDispatchButton = Request.Form("chkUserAccessServiceDispatchButton")
userAccessServiceCloseCancelButton = Request.Form("chkUserAccessServiceCloseCancelButton")

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
If userCreateNewServiceTicket = "on" then userCreateNewServiceTicket = 1 Else userCreateNewServiceTicket = 0
If userAccessServiceDispatchCenter = "on" then userAccessServiceDispatchCenter = 1 Else userAccessServiceDispatchCenter = 0
If userAccessServiceActionsModalButton = "on" then userAccessServiceActionsModalButton = 1 Else userAccessServiceActionsModalButton = 0
If userAccessServiceDispatchButton = "on" then userAccessServiceDispatchButton = 1 Else userAccessServiceDispatchButton = 0
If userAccessServiceCloseCancelButton = "on" then userAccessServiceCloseCancelButton = 1 Else userAccessServiceCloseCancelButton = 0


SQL = "INSERT INTO tblUsers (userFirstName,userLastName,userEmail,userPassword, "
SQL = SQL & "userEnabled,userDownloadEmail,userUpdateCalendar,userDisplayName,userCellNumber,userType,userCanAuthSwaps,userFilterRoutes,userCRMAccessType, "
SQL = SQL &  " userOrderAPIAccessType,userProspectingAddEditAccess,userCRMDeleteAccess,userEditCRMOnTheFly, UserTruckNumber,userLicense,userLicenseExpiration, "
SQL = SQL &  " userEmailSystemID ,userEmailSystemPass, userEmailServer, userVMS_ID, userForceNextStopSelectionOverride,  "
SQL = SQL &  " userNextStopNagMessageOverride ,userNextStopNagMinutes,userNextStopNagIntervalMinutes,  "
SQL = SQL &  " userNextStopNagMessageMaxToSendPerStop,userNextStopNagMessageMaxToSendThisDriverPerDay,userNextStopNagMessageSendMethod,  "
SQL = SQL &  " userNoActivityNagMessageOverride,userNoActivityNagMinutes,userNoActivityNagIntervalMinutes,  "
SQL = SQL &  " userNoActivityNagMessageMaxToSendPerStop,userNoActivityNagMessageMaxToSendPerDriverPerDay,userNoActivityNagMessageSendMethod,userNoActivityNagTimeOfDay, "
SQL = SQL &  " userNoActivityNagMessageOverride_FS,userNoActivityNagMinutes_FS,userNoActivityNagIntervalMinutes_FS,  "
SQL = SQL &  " userNoActivityNagMessageMaxToSendPerStop_FS,userNoActivityNagMessageMaxToSendPerDriverPerDay_FS,userNoActivityNagMessageSendMethod_FS,userNoActivityNagTimeOfDay_FS, "
SQL = SQL &  " userLoginDisableAccessHolidays,userInventoryControlAccessType,userMobileInventoryControlAccess,userEditEqpOnTheFly, "
SQL = SQL &  " userCreateEquipmentSymptomCodesOnTheFly,userCreateEquipmentProblemCodesOnTheFly, userCreateEquipmentResolutionCodesOnTheFly, "
SQL = SQL &  " userLeftNavAPIModule,userLeftNavBIModule, userLeftNavProspectingModule, "
SQL = SQL &  " userLeftNavCustomerServiceModule,userLeftNavEquipmentModule, userLeftNavInventoryControlModule, "
SQL = SQL &  " userLeftNavAccountsReceivableModule,userLeftNavAccountsPayableModule, userLeftNavServiceModule, "
SQL = SQL &  " userLeftNavRoutingModule,userLeftNavQuickbooksModule, userLeftNavFiltertraxModule, userLeftNavSystem, "
SQL = SQL &  " userCreateNewServiceTicket,userAccessServiceDispatchCenter, userAccessServiceActionsModalButton, "
SQL = SQL &  " userAccessServiceDispatchButton,userAccessServiceCloseCancelButton "


If MUV_READ("biModuleOn") = "Enabled" Then 
	SQL = SQL &  " ,userSalesPersonNumber "
	SQL = SQL &  " ,userSalesPersonNumber2 "
End If
SQL = SQL &  " 	,userReceivePartsRequestEmails "
SQL = SQL &  " ,LoginLandingPageURL)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & FirstName & "'"
SQL = SQL & ",'"  & LastName & "'"
SQL = SQL & ",'"  & Email & "'"
SQL = SQL & ",'"  & Password & "'"		
SQL = SQL & ","  & Enabled 
SQL = SQL & ","  & DownloadEmail 
SQL = SQL & ","  & UpdateCalendar
SQL = SQL & ",'"  & DisplayName & "'"
SQL = SQL & ",'"  & CellNumber & "'"
SQL = SQL & ",'" & UserType  & "'"	
SQL = SQL & "," & userCanAuthSwaps 
SQL = SQL & ",'" & FilterRoutes & "'"	
SQL = SQL & ",'" & userCRMAccessType & "'"
SQL = SQL & ",'" & userOrderAPIAccessType & "'"
SQL = SQL & ","  & chkCRMAddEditAccess 
SQL = SQL & ","  & chkCRMDeleteAccess 
SQL = SQL & ","  & userEditCRMOnTheFly
SQL = SQL & ",'" & UserTruckNumber & "'"
SQL = SQL & ",'Free'"
If MUV_READ("SERNO")="1071" Then 
	SQL = SQL & ", dateadd(year,5,getdate())"
Else
	SQL = SQL & ", dateadd(d,30,getdate())"
End If
SQL = SQL & ",'" & userEmailSystemID & "'"
SQL = SQL & ",'" & userEmailSystemPass & "'"
SQL = SQL & ",'" & userEmailServer & "'"
SQL = SQL & ",'" & userSystemVMSID & "'"
SQL = SQL & ",'" & ForceNextStop & "'"
SQL = SQL & ",'" & userNextStopNagMessageOverride & "'"
SQL = SQL & "," & userNextStopNagMinutes
SQL = SQL & "," & userNextStopNagIntervalMinutes
SQL = SQL & "," & userNextStopNagMessageMaxToSendPerStop
SQL = SQL & "," & userNextStopNagMessageMaxToSendThisDriverPerDay
SQL = SQL & ",'" & userNextStopNagMessageSendMethod & "'"
SQL = SQL & ",'" & userNoActivityNagMessageOverride & "'"
SQL = SQL & "," & userNoActivityNagMinutes
SQL = SQL & "," & userNoActivityNagIntervalMinutes
SQL = SQL & "," & userNoActivityNagMessageMaxToSendPerStop
SQL = SQL & "," & userNoActivityNagMessageMaxToSendPerDriverPerDay
SQL = SQL & ",'" & userNoActivityNagMessageSendMethod & "'"
SQL = SQL & ",'" & userNoActivityNagTimeOfDay & "'"
SQL = SQL & ",'" & userNoActivityNagMessageOverride_FS & "'"
SQL = SQL & "," & userNoActivityNagMinutes_FS
SQL = SQL & "," & userNoActivityNagIntervalMinutes_FS
SQL = SQL & "," & userNoActivityNagMessageMaxToSendPerStop_FS
SQL = SQL & "," & userNoActivityNagMessageMaxToSendPerDriverPerDay_FS
SQL = SQL & ",'" & userNoActivityNagMessageSendMethod_FS & "'"
SQL = SQL & ",'" & userNoActivityNagTimeOfDay_FS & "'"
SQL = SQL & ",'" & chkDisableHolidayLogins & "'"
SQL = SQL & ",'" & userInventoryControlAccessType& "'"
SQL = SQL & "," & userMobileInventoryControlAccess
SQL = SQL & "," & userCanEditEqpTablesOnFly
SQL = SQL & "," & userCreateEquipmentSymptomCodesOnTheFly
SQL = SQL & "," & userCreateEquipmentProblemCodesOnTheFly
SQL = SQL & "," & userCreateEquipmentResolutionCodesOnTheFly
SQL = SQL & "," & userLeftNavAPIModule
SQL = SQL & "," & userLeftNavBIModule
SQL = SQL & "," & userLeftNavProspectingModule
SQL = SQL & "," & userLeftNavCustomerServiceModule
SQL = SQL & "," & userLeftNavEquipmentModule
SQL = SQL & "," & userLeftNavInventoryControlModule
SQL = SQL & "," & userLeftNavAccountsReceivableModule
SQL = SQL & "," & userLeftNavAccountsPayableModule
SQL = SQL & "," & userLeftNavServiceModule
SQL = SQL & "," & userLeftNavRoutingModule
SQL = SQL & "," & userLeftNavQuickbooksModule
SQL = SQL & "," & userLeftNavFiltertraxModule
SQL = SQL & "," & userLeftNavSystem
SQL = SQL & "," & userCreateNewServiceTicket
SQL = SQL & "," & userAccessServiceDispatchCenter
SQL = SQL & "," & userAccessServiceActionsModalButton
SQL = SQL & "," & userAccessServiceDispatchButton
SQL = SQL & "," & userAccessServiceCloseCancelButton


If MUV_READ("biModuleOn") = "Enabled" Then 
	SQL = SQL & "," & userSalesPersonNumber
	SQL = SQL & "," & userSalesPersonNumber2
End If	
SQL = SQL & "," & userReceivePartsRequestEmails
SQL = SQL & ",'" & LoginLandingPageURL
SQL = SQL & "')"


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = ""
Description = Description & "Email: "  & Email
Description = Description & "     Display Name: "  & DisplayName 
Description = Description & "     Admin: "  
If Admin = 1 then Description = Description & "True" else Description = Description & "False" 
Description = Description & "     Enabled: "  
If Enabled = 1 then Description = Description & "True" else Description = Description & "False" 

Description = Description & "     DownloadEmail: "  
If DownloadEmail = 1 then Description = Description & "True" else Description = Description & "False" 
Description = Description & "     UpdateCalendar: "  
If UpdateCalendar = 1 then Description = Description & "True" else Description = Description & "False" 



If advancedDispatchIsOn() Then
	Description = Description & "     Can Auth Swaps: "  
	If userCanAuthSwaps = 1 then Description = Description & "True" else Description = Description & "False" 
	Description = Description & "     Receive part req emails: "  
	If userReceivePartsRequestEmails = 1 then Description = Description & "True" else Description = Description & "False" 
End If




Description = Description & "     User Type: "  & GetTerm(UserType)
CreateAuditLogEntry "User Added","User Added","Major",0,Description



'***********************************************************************************
'***********************************************************************************
'UPDATE SC_UserRestrictedLoginSchedule, REPLACING TEMP USER RECORDS, THOSE WITH A
'USERNO OF -1 TO THE USERNO OF THE NEWLY CREATED USER
'***********************************************************************************

SQL = "SELECT * FROM tblUsers WHERE userEmail = '" & Email & "'"
Set rs8 = cnn8.Execute(SQL)

If NOT rs8.EOF Then
	newUserNo = rs8("userNo")

	Set cnn9 = Server.CreateObject("ADODB.Connection")
	cnn9.open (Session("ClientCnnString"))
	Set rs9 = Server.CreateObject("ADODB.Recordset")
	rs9.CursorLocation = 3 

	SQL9 = "UPDATE SC_UserRestrictedLoginSchedule SET userNo = " & newUserNo & " WHERE userNo = -1"
	Set rs9 = cnn9.Execute(SQL9)
	set rs9 = Nothing
End If

set rs8 = Nothing

'***********************************************************************************
'***********************************************************************************



Response.Redirect("main.asp#" & ActiveTab)

%>















