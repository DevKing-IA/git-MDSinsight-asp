<%
'********************************
'List of all the functions & subs
'********************************
'Func EzTextingUserID()
'Func EzTextingPassword()
'Func userTrafficReconcileUser()
'Func userIsArchived(passedUserNo)
'Func userIsEnabled(passedUserNo)
'Func getUserType(passedUserNo)
'Func userIsFinance(passedUserNo)
'Func userIsCSR(passedUserNo)
'Func userIsCSRManager(passedUserNo)
'Func userIsRouteManager(passedUserNo)
'Func userIsInsideSalesManager(passedUserNo)
'Func userIsOutsideSalesManager(passedUserNo)
'Func userIsAdmin(passedUserNo)
'Func userIsDriver(passedUserNo)
'Func userIsCSRManager(passedUserNo)
'Func userIsServiceManager(passedUserNo)
'Func userIsFinanceManager(passedUserNo)
'Func userIsSalesGroup(passedUserNo)
'Func GetUserDisplayNameByUserNo(passedUserNo)
'Func GetUserFirstAndLastNameByUserNo(passedUserNo)
'Func GetUserNoByUserDisplayName(passedUserDisplayName)
'Func GetUserNoByEmailAddress(passedEmailAddress)
'Func GetUserEmailByUserNo(passedUserNo)
'Func userViewLeftNavAPIModule(passedUserNo)
'Func userViewLeftNavBIModule(passedUserNo)
'Func userViewLeftNavProspectingModule(passedUserNo)
'Func userViewLeftNavCustomerServiceModule(passedUserNo)
'Func userViewLeftNavEquipmentModule(passedUserNo)
'Func userViewLeftNavInventoryControlModule(passedUserNo)
'Func userViewLeftNavAccountsReceivableModule(passedUserNo)
'Func userViewLeftNavAccountsPayableModule(passedUserNo)
'Func userViewLeftNavServiceModule(passedUserNo)
'Func userViewLeftNavRoutingModule(passedUserNo)
'Func userViewLeftNavQuickbooksModule(passedUserNo)
'Func userViewLeftNavFiltertraxModule(passedUserNo)
'Func userViewLeftNavSystem(passedUserNo)
'Func userCanCreateNewServiceTicket(passedUserNo)
'Func userCanAccessServiceDispatchCenter(passedUserNo)
'Func userCanAccessServiceActionsModalButton(passedUserNo)
'Func userCanAccessServiceDispatchButton(passedUserNo)
'Func userCanAccessServiceCloseCancelButton(passedUserNo)
'Func GetTotalNumberOfTeamMembers(passedTeamIntRecID)
'Func GetTeamNameByTeamIntRecID(passedTeamIntRecID)
'Func GetTeamUserNosByTeamIntRecID(passedTeamIntRecID)
'Func SetUserGeoLocation(passedUserNo,passedGeoLoc)
'Func GetCRMPermissionLevel(passedUserNo)
'Func GetCRMAddEditMenuPermissionLevel(passedUserNo)
'Func GetInventoryControlAddEditMenuPermissionLevel(passedUserNo)
'Func GetCRMDeleteProspectPermissionLevel(passedUserNo)
'Func NoteNewForUser(passedCustNum,passedEntryDateTime)
'Func MARKNoteNewForUser(passedCustNum)
'Func getUserEmailAddress(passedUserNo)
'Func getUserCellNumber(passedUserNo)
'Func GetUserEmailSystemIDByUserNo(passedUserNo)
'Func GetUserEmailSystemPassByUserNo(passedUserNo)
'Func GetUserEmailServerByUserNo(passedUserNo)
'Func AllowUpdatesToUsersCalendar(passedUserNo)
'Func GetUserOrderAPIPermissionLevel(passedUserNo)
'Func userCanEditEqpOnFly(passedUserNo)

'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function EzTextingUserID()

	resultEzTextingUserID = ""

	Set cnnEzTextingUserID = Server.CreateObject("ADODB.Connection")
	cnnEzTextingUserID.open (Session("ClientCnnString"))
	Set rsEzTextingUserID = Server.CreateObject("ADODB.Recordset")
	rsEzTextingUserID.CursorLocation = 3 

	SQLEzTextingUserID = "SELECT EZTextingID FROM Settings_Global"

	Set rsEzTextingUserID = cnnEzTextingUserID.Execute(SQLEzTextingUserID)

	If not rsEzTextingUserID.EOF Then resultEzTextingUserID = rsEzTextingUserID("EZTextingID")
	set rsEzTextingUserID = Nothing
	cnnEzTextingUserID.close
	set cnnEzTextingUserID = Nothing
	
	EzTextingUserID = resultEzTextingUserID

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function EzTextingPassword()

	resultEzTextingPassword = ""

	Set cnnEzTextingPassword = Server.CreateObject("ADODB.Connection")
	cnnEzTextingPassword.open (Session("ClientCnnString"))
	Set rsEzTextingPassword = Server.CreateObject("ADODB.Recordset")
	rsEzTextingPassword.CursorLocation = 3 

	SQLEzTextingPassword = "SELECT EZTextingPassword FROM Settings_Global"

	Set rsEzTextingPassword = cnnEzTextingPassword.Execute(SQLEzTextingPassword)

	If not rsEzTextingPassword.EOF Then resultEzTextingPassword = rsEzTextingPassword("EZTextingPassword")
	set rsEzTextingPassword = Nothing
	cnnEzTextingPassword.close
	set cnnEzTextingPassword = Nothing
	
	EzTextingPassword = resultEzTextingPassword

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsSalesGroup(passedUserNo)

	result = false
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userType") = "Inside Sales" or rsBoost1("userType") = "Outside Sales" or rsBoost1("userType") = "Telemarketing" then result = True
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	userIsSalesGroup = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsFinance(passedUserNo)

	resultuserIsFinance = false
	
	Set cnnuserIsFinance = Server.CreateObject("ADODB.Connection")
	cnnuserIsFinance.open Session("ClientCnnString")

	SQLuserIsFinance = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsuserIsFinance = Server.CreateObject("ADODB.Recordset")
	rsuserIsFinance.CursorLocation = 3 
	Set rsuserIsFinance= cnnuserIsFinance.Execute(SQLuserIsFinance)
	
	If not rsuserIsFinance.eof then 
		If rsuserIsFinance("userType") = "Finance" then resultuserIsFinance = True
	End IF	
	set rsuserIsFinance= Nothing
	set cnnuserIsFinance= Nothing
	
	userIsFinance = resultuserIsFinance

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsTelemarketing(passedUserNo)

	resultuserIsTelemarketing = false
	
	Set cnnuserIsTelemarketing = Server.CreateObject("ADODB.Connection")
	cnnuserIsTelemarketing.open Session("ClientCnnString")

	SQLuserIsTelemarketing = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsuserIsTelemarketing = Server.CreateObject("ADODB.Recordset")
	rsuserIsTelemarketing.CursorLocation = 3 
	Set rsuserIsTelemarketing= cnnuserIsTelemarketing.Execute(SQLuserIsTelemarketing)
	
	If not rsuserIsTelemarketing.eof then 
		If rsuserIsTelemarketing("userType") = "Telemarketing" then resultuserIsTelemarketing = True
	End IF	
	set rsuserIsTelemarketing= Nothing
	set cnnuserIsTelemarketing= Nothing
	
	userIsTelemarketing = resultuserIsTelemarketing

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsCSR(passedUserNo)

	resultuserIsCSR = false
	
	Set cnnuserIsCSR = Server.CreateObject("ADODB.Connection")
	cnnuserIsCSR.open Session("ClientCnnString")

	SQLuserIsCSR = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsuserIsCSR = Server.CreateObject("ADODB.Recordset")
	rsuserIsCSR.CursorLocation = 3 
	Set rsuserIsCSR= cnnuserIsCSR.Execute(SQLuserIsCSR)
	
	If not rsuserIsCSR.eof then 
		If rsuserIsCSR("userType") = "CSR" then resultuserIsCSR = True
	End IF	
	set rsuserIsCSR= Nothing
	set cnnuserIsCSR= Nothing
	
	userIsCSR = resultuserIsCSR

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsCSRManager(passedUserNo)

	resultuserIsCSRManager = false
	
	Set cnnuserIsCSRManager = Server.CreateObject("ADODB.Connection")
	cnnuserIsCSRManager.open Session("ClientCnnString")

	SQLuserIsCSRManager = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsuserIsCSRManager = Server.CreateObject("ADODB.Recordset")
	rsuserIsCSRManager.CursorLocation = 3 
	Set rsuserIsCSRManager = cnnuserIsCSRManager.Execute(SQLuserIsCSRManager)
	
	If not rsuserIsCSRManager.eof then 
		If rsuserIsCSRManager("userType") = "CSR Manager" then resultuserIsCSRManager = True
	End IF	
	set rsuserIsCSRManager = Nothing
	set cnnuserIsCSRManager = Nothing
	
	userIsCSRManager = resultuserIsCSRManager 

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsAdmin(passedUserNo)

	userIsAdmin= false
	
	Set cnnuserIsAdmin = Server.CreateObject("ADODB.Connection")
	cnnuserIsAdmin.open Session("ClientCnnString")

	SQLuserIsAdmin = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsuserIsAdmin = Server.CreateObject("ADODB.Recordset")
	rsuserIsAdmin.CursorLocation = 3 
	Set rsuserIsAdmin = cnnuserIsAdmin.Execute(SQLuserIsAdmin)
	
	If not rsuserIsAdmin.eof then 
		If rsuserIsAdmin("userType") = "Admin" then resultuserIsAdmin = True
	End IF	
	set rsuserIsAdmin = Nothing
	set cnnuserIsAdmin = Nothing
	
	userIsAdmin = resultuserIsAdmin 

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsDriver(passedUserNo)

	userIsDriver= false
	
	Set cnnuserIsDriver = Server.CreateObject("ADODB.Connection")
	cnnuserIsDriver.open Session("ClientCnnString")

	SQLuserIsDriver = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsuserIsDriver = Server.CreateObject("ADODB.Recordset")
	rsuserIsDriver.CursorLocation = 3 
	Set rsuserIsDriver = cnnuserIsDriver.Execute(SQLuserIsDriver)
	
	If not rsuserIsDriver.eof then 
		If rsuserIsDriver("userType") = "Driver" then resultuserIsDriver = True
	End IF	
	set rsuserIsDriver = Nothing
	set cnnuserIsDriver = Nothing
	
	userIsDriver = resultuserIsDriver 

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsArchived(passedUserNo)

	result = false
	
	If passedUserNo <> "*Not Found*" Then
	
		Set cnn = Server.CreateObject("ADODB.Connection")
		cnn.open Session("ClientCnnString")
	
		SQLUserArchived = "Select userArchived from tblUsers where UserNo = " & passedUserNo
		
		'Response.write(SQLUserArchived & "<br><br>")
		 
		Set rsUserArchived = Server.CreateObject("ADODB.Recordset")
		rsUserArchived.CursorLocation = 3 
		Set rsUserArchived= cnn.Execute(SQLUserArchived)
		
		If not rsUserArchived.eof then 
			If rsUserArchived("userArchived") = 1 then result = True
		End IF	
		set rsUserArchived= Nothing
		set cnn= Nothing
		
	End If
	
	userIsArchived = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsEnabled(passedUserNo)

	result = true
	
	If passedUserNo <> "*Not Found*" Then
		
		Set cnn = Server.CreateObject("ADODB.Connection")
		cnn.open Session("ClientCnnString")
	
		SQLUserEnabled = "Select userEnabled from tblUsers where UserNo = " & passedUserNo
		 
		Set rsUserEnabled = Server.CreateObject("ADODB.Recordset")
		rsUserEnabled.CursorLocation = 3 
		Set rsUserEnabled= cnn.Execute(SQLUserEnabled)
		
		If not rsUserEnabled.eof then 
			If rsUserEnabled("userEnabled") = 0 then result = false
		End IF	
		set rsUserEnabled= Nothing
		set cnn= Nothing
		
	End If
	
	userIsEnabled = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function getUserType(passedUserNo)

	result = false
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "SELECT * FROM tblUsers WHERE UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		result= rsBoost1("userType")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	getUserType = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsCSRManager(passedUserNo)

	result = false
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userType") = "CSR Manager" then result = True
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	userIsCSRManager = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsRouteManager(passedUserNo)

	result = false
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userType") = "Route Manager" then result = True
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	userIsRouteManager = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsInsideSalesManager(passedUserNo)

	result = false
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userType") = "Inside Sales Manager" then result = True
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	userIsInsideSalesManager = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsOutsideSales(passedUserNo)

	result = false
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userType") = "Outside Sales" then result = True
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	userIsOutsideSales = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsInsideSales(passedUserNo)

	result = false
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userType") = "Inside Sales" then result = True
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	userIsInsideSales = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsOutsideSalesManager(passedUserNo)

	result = false
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userType") = "Outside Sales Manager" then result = True
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	userIsOutsideSalesManager = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsServiceManager(passedUserNo)

	result = false
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userType") = "Service Manager" then result = True
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing

	'Fixit
	' cheap fix to let adam henchel see service stuff wihtout being a service manager
	If passedUserNo = 56 Then result = True
	
	userIsServiceManager = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userIsFinanceManager(passedUserNo)

	result = false
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userType") = "Finance Manager" then result = True
	End IF	 
	set rsBoost1= Nothing
	set cnn= Nothing
	
	userIsFinanceManager = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavAPIModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else

		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavAPIModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavAPIModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
	
	userViewLeftNavAPIModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavBIModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
	
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavBIModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavBIModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
	
	End If
	
	userViewLeftNavBIModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavProspectingModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
	
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavProspectingModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavProspectingModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
	
	userViewLeftNavProspectingModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavCustomerServiceModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
	
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavCustomerServiceModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavCustomerServiceModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
	
	End If
	
	userViewLeftNavCustomerServiceModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavEquipmentModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
	
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavEquipmentModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavEquipmentModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
	
	userViewLeftNavEquipmentModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavInventoryControlModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
	
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavInventoryControlModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavInventoryControlModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
	
	userViewLeftNavInventoryControlModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavAccountsReceivableModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
	
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavAccountsReceivableModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavAccountsReceivableModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
	
	userViewLeftNavAccountsReceivableModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavAccountsPayableModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
	
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavAccountsPayableModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavAccountsPayableModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
	
	End If
	
	userViewLeftNavAccountsPayableModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavServiceModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
		
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavServiceModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavServiceModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
	
	End If
	
	userViewLeftNavServiceModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavRoutingModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else

		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavRoutingModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavRoutingModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
	
	End If
	
	userViewLeftNavRoutingModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavQuickbooksModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
	
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavQuickbooksModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavQuickbooksModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
	
	End If
	
	userViewLeftNavQuickbooksModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavQuickbooksModule(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
	
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavQuickbooksModule FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavQuickbooksModule") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
	
	userViewLeftNavQuickbooksModule = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userViewLeftNavSystem(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
		
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userLeftNavSystem FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userLeftNavSystem") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
	
	userViewLeftNavSystem = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userCanCreateNewServiceTicket(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
	
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userCreateNewServiceTicket FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userCreateNewServiceTicket") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
	
	userCanCreateNewServiceTicket = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userCanAccessServiceDispatchCenter(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
		
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userAccessServiceDispatchCenter FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userAccessServiceDispatchCenter") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
	
	userCanAccessServiceDispatchCenter = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userCanAccessServiceActionsModalButton(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
	
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userAccessServiceActionsModalButton FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userAccessServiceActionsModalButton") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
		
	userCanAccessServiceActionsModalButton = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userCanAccessServiceDispatchButton(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else
		
		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
	
		SQL = "SELECT userAccessServiceDispatchButton FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userAccessServiceDispatchButton") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
	
	userCanAccessServiceDispatchButton = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userCanAccessServiceCloseCancelButton(passedUserNo)

	result = false
	
	If userIsAdmin(passedUserNo) Then
	
		result = true
		
	Else

		Set cnnUser = Server.CreateObject("ADODB.Connection")
		cnnUser.open Session("ClientCnnString")
		
		SQL = "SELECT userAccessServiceCloseCancelButton FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUser = Server.CreateObject("ADODB.Recordset")
		rsUser.CursorLocation = 3 
		Set rsUser= cnnUser.Execute(SQL)
		
		If NOT rsUser.EOF then 
			If rsUser("userAccessServiceCloseCancelButton") = 1 then result = true
		End If	
		
		set rsUser= Nothing
		set cnnUser= Nothing
		
	End If
		
	userCanAccessServiceCloseCancelButton = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfTeamMembers(passedTeamIntRecID)

	Set cnnTotalNumTeamMembers = Server.CreateObject("ADODB.Connection")
	cnnTotalNumTeamMembers.open Session("ClientCnnString")

	resultTotalNumTeamMembers = 0
		
	SQLTotalNumTeamMembers = "SELECT * FROM USER_TEAMS WHERE InternalRecordIdentifier = " & passedTeamIntRecID
	 
	Set rsTotalNumTeamMembers = Server.CreateObject("ADODB.Recordset")
	rsTotalNumTeamMembers.CursorLocation = 3 
	
	rsTotalNumTeamMembers.Open SQLTotalNumTeamMembers,cnnTotalNumTeamMembers 
	
	TeamUserNos = rsTotalNumTeamMembers("TeamUserNos")
	TeamUserNosArray = Split(TeamUserNos,",")
			
	resultTotalNumTeamMembers = UBound(TeamUserNosArray) + 1
	
	rsTotalNumTeamMembers.Close
	set rsTotalNumTeamMembers = Nothing
	cnnTotalNumTeamMembers.Close	
	set cnnTotalNumTeamMembers = Nothing
	
	GetTotalNumberOfTeamMembers = resultTotalNumTeamMembers
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTeamNameByTeamIntRecID(passedTeamIntRecID)

	Set cnnTeamNameByTeamIntRecID = Server.CreateObject("ADODB.Connection")
	cnnTeamNameByTeamIntRecID.open Session("ClientCnnString")

	resultTeamNameByTeamIntRecID = 0
		
	SQLTeamNameByTeamIntRecID = "SELECT * FROM USER_TEAMS WHERE InternalRecordIdentifier = " & passedTeamIntRecID
	 
	Set rsTeamNameByTeamIntRecID = Server.CreateObject("ADODB.Recordset")
	rsTeamNameByTeamIntRecID.CursorLocation = 3 
	
	rsTeamNameByTeamIntRecID.Open SQLTeamNameByTeamIntRecID,cnnTeamNameByTeamIntRecID 
				
	resultTeamNameByTeamIntRecID = rsTeamNameByTeamIntRecID("TeamName")
	
	rsTeamNameByTeamIntRecID.Close
	set rsTeamNameByTeamIntRecID = Nothing
	cnnTeamNameByTeamIntRecID.Close	
	set cnnTeamNameByTeamIntRecID = Nothing
	
	GetTeamNameByTeamIntRecID = resultTeamNameByTeamIntRecID
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTeamNameByTeamIntRecID(passedTeamIntRecID)

	Set cnnTeamNameByTeamIntRecID = Server.CreateObject("ADODB.Connection")
	cnnTeamNameByTeamIntRecID.open Session("ClientCnnString")

	resultTeamNameByTeamIntRecID = 0
		
	SQLTeamNameByTeamIntRecID = "SELECT * FROM USER_TEAMS WHERE InternalRecordIdentifier = " & passedTeamIntRecID
	 
	Set rsTeamNameByTeamIntRecID = Server.CreateObject("ADODB.Recordset")
	rsTeamNameByTeamIntRecID.CursorLocation = 3 
	
	rsTeamNameByTeamIntRecID.Open SQLTeamNameByTeamIntRecID,cnnTeamNameByTeamIntRecID 
				
	resultTeamNameByTeamIntRecID = rsTeamNameByTeamIntRecID("TeamName")
	
	rsTeamNameByTeamIntRecID.Close
	set rsTeamNameByTeamIntRecID = Nothing
	cnnTeamNameByTeamIntRecID.Close	
	set cnnTeamNameByTeamIntRecID = Nothing
	
	GetTeamNameByTeamIntRecID = resultTeamNameByTeamIntRecID
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************

Function GetTeamUserNosByTeamIntRecID(passedTeamIntRecID)

	Set cnnTeamUserNosByTeamIntRecID = Server.CreateObject("ADODB.Connection")
	cnnTeamUserNosByTeamIntRecID.open Session("ClientCnnString")

	resultTeamUserNosByTeamIntRecID = 0
		
	SQLTeamUserNosByTeamIntRecID = "SELECT * FROM USER_TEAMS WHERE InternalRecordIdentifier = " & passedTeamIntRecID
	 
	Set rsTeamUserNosByTeamIntRecID = Server.CreateObject("ADODB.Recordset")
	rsTeamUserNosByTeamIntRecID.CursorLocation = 3 
	
	rsTeamUserNosByTeamIntRecID.Open SQLTeamUserNosByTeamIntRecID,cnnTeamUserNosByTeamIntRecID 
				
	resultTeamUserNosByTeamIntRecID = rsTeamUserNosByTeamIntRecID("TeamUserNos")
	
	rsTeamUserNosByTeamIntRecID.Close
	set rsTeamUserNosByTeamIntRecID = Nothing
	cnnTeamUserNosByTeamIntRecID.Close	
	set cnnTeamUserNosByTeamIntRecID = Nothing
	
	GetTeamUserNosByTeamIntRecID = resultTeamUserNosByTeamIntRecID
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetCRMPermissionLevel(passedUserNo)
	
	GetCRMPermissionLevelresult = "NONE"
	
	Set cnnGetCRMPermissionLevel = Server.CreateObject("ADODB.Connection")
	cnnGetCRMPermissionLevel.open Session("ClientCnnString")

	SQLGetCRMPermissionLevel = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsGetCRMPermissionLevel = Server.CreateObject("ADODB.Recordset")
	rsGetCRMPermissionLevel.CursorLocation = 3 
	Set rsGetCRMPermissionLevel= cnnGetCRMPermissionLevel.Execute(SQLGetCRMPermissionLevel)
	
	If not rsGetCRMPermissionLevel.eof then 
		resultGetCRMPermissionLevel = rsGetCRMPermissionLevel("userCRMAccessType")
	End IF	
	
	set rsGetCRMPermissionLevel = Nothing
	cnnGetCRMPermissionLevel.Close
	set cnnGetCRMPermissionLevel = Nothing
	
	GetCRMPermissionLevel = resultGetCRMPermissionLevel

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetCRMAddEditMenuPermissionLevel(passedUserNo)
	
	GetCRMAddEditMenuPermissionLevelresult = 0
	
	Set cnnGetCRMAddEditMenuPermissionLevel = Server.CreateObject("ADODB.Connection")
	cnnGetCRMAddEditMenuPermissionLevel.open Session("ClientCnnString")

	SQLGetCRMAddEditMenuPermissionLevel = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsGetCRMAddEditMenuPermissionLevel = Server.CreateObject("ADODB.Recordset")
	rsGetCRMAddEditMenuPermissionLevel.CursorLocation = 3 
	Set rsGetCRMAddEditMenuPermissionLevel= cnnGetCRMAddEditMenuPermissionLevel.Execute(SQLGetCRMAddEditMenuPermissionLevel)
	
	If not rsGetCRMAddEditMenuPermissionLevel.eof then 
		resultGetCRMAddEditMenuPermissionLevel = rsGetCRMAddEditMenuPermissionLevel("userProspectingAddEditAccess")
	End IF	
	
	set rsGetCRMAddEditMenuPermissionLevel = Nothing
	cnnGetCRMAddEditMenuPermissionLevel.Close
	set cnnGetCRMAddEditMenuPermissionLevel = Nothing
	
	GetCRMAddEditMenuPermissionLevel = resultGetCRMAddEditMenuPermissionLevel

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetInventoryControlAddEditMenuPermissionLevel(passedUserNo)
	
	resultGetInventoryControlAddEditMenuPermissionLevel = 0
	
	Set cnnGetInventoryControlAddEditMenuPermissionLevel  = Server.CreateObject("ADODB.Connection")
	cnnGetInventoryControlAddEditMenuPermissionLevel.open Session("ClientCnnString")

	SQLGetInventoryControlAddEditMenuPermissionLevel  = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsGetInventoryControlAddEditMenuPermissionLevel  = Server.CreateObject("ADODB.Recordset")
	rsGetInventoryControlAddEditMenuPermissionLevel.CursorLocation = 3 
	Set rsGetInventoryControlAddEditMenuPermissionLevel = cnnGetInventoryControlAddEditMenuPermissionLevel.Execute(SQLGetInventoryControlAddEditMenuPermissionLevel)
	
	If not rsGetInventoryControlAddEditMenuPermissionLevel.eof then 
		resultGetInventoryControlAddEditMenuPermissionLevel = rsGetInventoryControlAddEditMenuPermissionLevel("userInventoryControlAccess")
	End IF	
	
	set rsGetInventoryControlAddEditMenuPermissionLevel  = Nothing
	cnnGetInventoryControlAddEditMenuPermissionLevel.Close
	set cnnGetInventoryControlAddEditMenuPermissionLevel  = Nothing
	
	GetInventoryControlAddEditMenuPermissionLevel  = resultGetInventoryControlAddEditMenuPermissionLevel 

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetUserOrderAPIPermissionLevel(passedUserNo)
	
	GetUserOrderAPIPermissionLevelresult = "NONE"
	
	Set cnnGetUserOrderAPIPermissionLevel = Server.CreateObject("ADODB.Connection")
	cnnGetUserOrderAPIPermissionLevel.open Session("ClientCnnString")

	SQLGetUserOrderAPIPermissionLevel = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsGetUserOrderAPIPermissionLevel = Server.CreateObject("ADODB.Recordset")
	rsGetUserOrderAPIPermissionLevel.CursorLocation = 3 
	Set rsGetUserOrderAPIPermissionLevel = cnnGetUserOrderAPIPermissionLevel.Execute(SQLGetUserOrderAPIPermissionLevel)
	
	If not rsGetUserOrderAPIPermissionLevel.eof then 
		resultGetUserOrderAPIPermissionLevel = rsGetUserOrderAPIPermissionLevel("userOrderAPIAccessType")
	End IF	
	
	set rsGetUserOrderAPIPermissionLevel = Nothing
	cnnGetUserOrderAPIPermissionLevel.Close
	set cnnGetUserOrderAPIPermissionLevel = Nothing
	
	GetUserOrderAPIPermissionLevel = resultGetUserOrderAPIPermissionLevel

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetCRMDeleteProspectPermissionLevel(passedUserNo)
	
	GetCRMDeleteProspectPermissionLevelresult = 0
	
	Set cnnGetCRMDeleteProspectPermissionLevel = Server.CreateObject("ADODB.Connection")
	cnnGetCRMDeleteProspectPermissionLevel.open Session("ClientCnnString")

	SQLGetCRMDeleteProspectPermissionLevel = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsGetCRMDeleteProspectPermissionLevel = Server.CreateObject("ADODB.Recordset")
	rsGetCRMDeleteProspectPermissionLevel.CursorLocation = 3 
	Set rsGetCRMDeleteProspectPermissionLevel= cnnGetCRMDeleteProspectPermissionLevel.Execute(SQLGetCRMDeleteProspectPermissionLevel)
	
	If not rsGetCRMDeleteProspectPermissionLevel.eof then 
		resultGetCRMDeleteProspectPermissionLevel = rsGetCRMDeleteProspectPermissionLevel("userCRMDeleteAccess")
	End IF	
	
	set rsGetCRMDeleteProspectPermissionLevel = Nothing
	cnnGetCRMDeleteProspectPermissionLevel.Close
	set cnnGetCRMDeleteProspectPermissionLevel = Nothing
	
	GetCRMDeleteProspectPermissionLevel = resultGetCRMDeleteProspectPermissionLevel

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetUserDisplayNameByUserNo(passedUserNo)

	result = ""
	
	If IsNumeric(passedUserNo) Then	
		Set cnnGetUserDisplayNameByUserNo = Server.CreateObject("ADODB.Connection")
		cnnGetUserDisplayNameByUserNo.open Session("ClientCnnString")
	
		SQLGetUserDisplayNameByUserNo = "Select * from tblUsers where UserNo = " & passedUserNo
	
		Set rsGetUserDisplayNameByUserNo = Server.CreateObject("ADODB.Recordset")
		rsGetUserDisplayNameByUserNo.CursorLocation = 3 
		Set rsGetUserDisplayNameByUserNo = cnnGetUserDisplayNameByUserNo.Execute(SQLGetUserDisplayNameByUserNo)
		
		If not rsGetUserDisplayNameByUserNo.eof then 
			result = rsGetUserDisplayNameByUserNo("userDisplayName")
		End IF	
		set rsGetUserDisplayNameByUserNo= Nothing
		set cnGetUserDisplayNameByUserNon= Nothing
		
		If passedUserNo = 0 Then result="System" ' User 0 is special
	End If	
	GetUserDisplayNameByUserNo = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetUserFirstAndLastNameByUserNo(passedUserNo)

	result = ""

	If passedUserNo <> "" Then	
		Set cnnGetUserFirstAndLastNameByUserNo = Server.CreateObject("ADODB.Connection")
		cnnGetUserFirstAndLastNameByUserNo.open Session("ClientCnnString")
	
		SQLGetUserFirstAndLastNameByUserNo = "Select * from tblUsers where UserNo = " & passedUserNo
	
		Set rsGetUserFirstAndLastNameByUserNo = Server.CreateObject("ADODB.Recordset")
		rsGetUserFirstAndLastNameByUserNo.CursorLocation = 3 
		Set rsGetUserFirstAndLastNameByUserNo = cnnGetUserFirstAndLastNameByUserNo.Execute(SQLGetUserFirstAndLastNameByUserNo)
		
		If not rsGetUserFirstAndLastNameByUserNo.eof then 
			result = rsGetUserFirstAndLastNameByUserNo("userFirstName") & " " & rsGetUserFirstAndLastNameByUserNo("userLastName")
		End IF	
		set rsGetUserFirstAndLastNameByUserNo= Nothing
		set cnGetUserFirstAndLastNameByUserNon= Nothing
		
		If passedUserNo = 0 Then result="System" ' User 0 is special
	End If
	
	GetUserFirstAndLastNameByUserNo = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetUserNoByUserDisplayName(passedUserDisplayName)

	result = false
	
	Set cnnGetUserNoByUserDisplayName = Server.CreateObject("ADODB.Connection")
	cnnGetUserNoByUserDisplayName.open Session("ClientCnnString")

	SQLGetUserNoByUserDisplayName = "Select * from tblUsers where UserDisplayName = '" & passedUserDisplayName & "'"

	Set rsGetUserNoByUserDisplayName = Server.CreateObject("ADODB.Recordset")
	rsGetUserNoByUserDisplayName.CursorLocation = 3 
	Set rsGetUserNoByUserDisplayName= cnnGetUserNoByUserDisplayName.Execute(SQLGetUserNoByUserDisplayName)
	
	If not rsGetUserNoByUserDisplayName.eof then 
		result = rsGetUserNoByUserDisplayName("userNo")
	End IF	
	set rsGetUserNoByUserDisplayName= Nothing
	set cnnGetUserNoByUserDisplayName= Nothing
	
	GetUserNoByUserDisplayName = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetUserNoByEmailAddress(passedEmailAddress)

	result = ""
	
	Set cnnGetUserNoByEmailAddress = Server.CreateObject("ADODB.Connection")
	cnnGetUserNoByEmailAddress.open Session("ClientCnnString")

	SQLGetUserNoByEmailAddress = "Select * from tblUsers where UserEmail= '" & passedEmailAddress & "'"

	Set rsGetUserNoByEmailAddress = Server.CreateObject("ADODB.Recordset")
	rsGetUserNoByEmailAddress.CursorLocation = 3 
	Set rsGetUserNoByEmailAddress= cnnGetUserNoByEmailAddress.Execute(SQLGetUserNoByEmailAddress)
	
	If not rsGetUserNoByEmailAddress.eof then 
		result = rsGetUserNoByEmailAddress("userNo")
	End IF	
	set rsGetUserNoByEmailAddress= Nothing
	set cnnGetUserNoByEmailAddress= Nothing
	
	GetUserNoByEmailAddress = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetUserEmailByUserNo(passedUserNo)

	result = false
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		result = rsBoost1("userEmail")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetUserEmailByUserNo = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function SetUserGeoLocation(passedUserNo,passedGeoLoc)

	result = False
	RecFound = False
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	'See if it is there a record so we can update or insert
	SQL = "Select * from tblUsersGeoLocation where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then RecFound = True

	GeoArray = Split(passedGeoLoc,",")
	If Ubound(GeoArray) = 1 Then
		LatVar = GeoArray(0) 
		LongVar = GeoArray(1) 
		
		'Now Insert or Update
		If RecFound = True Then
			SQL = "UPDATE tblUsersGeoLocation Set Longitude = '" & LongVar & "', "
			SQL = SQL & "Latitude = '" & LatVar & "', LastLocationDateTime = getdate() "
			SQL = SQl & "WHERE Userno =" & passedUserNo	
		Else
			SQL = "INSERT INTO tblUsersGeoLocation (UserNo,Longitude,Latitude,LastLocationDateTime) "
			SQL = SQL &  " VALUES (" & passedUserNo & ", "
			SQL = SQL & "'" & LongVar & "', "
			SQL = SQL & "'" & LatVar & "', "
			SQL = SQL & " GetDate())"
		End If
		
		Set rsBoost1= cnn.Execute(SQL)
	End If
	set rsBoost1= Nothing
	set cnn= Nothing
	
	SetUserGeoLocation = result
	
'	Response.Redirect("http://maps.google.com/maps?q=" & LatVar & "," & LongVar)

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NoteNewForUser(passedCustNum,passedEntryDateTime)

	resultNoteNewForUser = False
	
	SQLNoteNewForUser = "SELECT * FROM tblCustomerNotesUserViewed Where CustNum ='" & passedCustNum & "' AND UserNo = " & Session("Userno")
	
	Set cnnNoteNewForUser = Server.CreateObject("ADODB.Connection")
	cnnNoteNewForUser.open (Session("ClientCnnString"))
	Set rNoteNewForUser = Server.CreateObject("ADODB.Recordset")
	rNoteNewForUser.CursorLocation = 3 
	Set rNoteNewForUser = cnnNoteNewForUser.Execute(SQLNoteNewForUser)

	If not rNoteNewForUser.EOF Then
		If cDate(rNoteNewForUser ("DateLastViewed")) < cDate(passedEntryDateTime) Then resultNoteNewForUser = True
	Else
		resultNoteNewForUser = True 'Also true if they have never seen any of them
	End If
	cnnNoteNewForUser.close
	set rNoteNewForUser = nothing
	set cnnNoteNewForUser= nothing	

	NoteNewForUser = resultNoteNewForUser

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function MARKNoteNewForUser(passedCustNum)

	SQLMARKNoteNewForUser = "SELECT * FROM tblCustomerNotesUserViewed Where CustNum ='" & passedCustNum & "' AND UserNo = " & Session("Userno")
	
	Set cnnMARKNoteNewForUser = Server.CreateObject("ADODB.Connection")
	cnnMARKNoteNewForUser.open (Session("ClientCnnString"))
	Set rMARKNoteNewForUser = Server.CreateObject("ADODB.Recordset")
	rMARKNoteNewForUser.CursorLocation = 3 
	Set rMARKNoteNewForUser = cnnMARKNoteNewForUser.Execute(SQLMARKNoteNewForUser)

	If rMARKNoteNewForUser.EOF Then ' Nothing there so we need to insert
		SQLMARKNoteNewForUser = "INSERT INTO tblCustomerNotesUserViewed (CustNum ,UserNo) VALUES ('" & passedCustNum & "',"  & Session("UserNo") & ")"
	Else
		SQLMARKNoteNewForUser = "UPDATE tblCustomerNotesUserViewed Set DateLastViewed = getdate() Where CustNum ='" & passedCustNum & "' AND UserNo = " & Session("Userno")
	End If
	
	Set rMARKNoteNewForUser = cnnMARKNoteNewForUser.Execute(SQLMARKNoteNewForUser)
		
	cnnMARKNoteNewForUser.close
	set rMARKNoteNewForUser = nothing
	set cnnMARKNoteNewForUser= nothing	

	MARKNoteNewForUser = 0

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function getUserEmailAddress(passedUserNo)

	result = ""
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userEmail") <> "" Then result = rsBoost1("UserEmail")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	getUserEmailAddress = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function getUserCellNumber(passedUserNo)

	result = ""
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userCellNumber") <> "" Then result = rsBoost1("userCellNumber")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	getUserCellNumber= result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetUserEmailSystemIDByUserNo(passedUserNo)

	resultGetUserEmailSystemIDByUserNo = ""

	Set cnnGetUserEmailSystemIDByUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetUserEmailSystemIDByUserNo.open Session("ClientCnnString")
		
	SQLGetUserEmailSystemIDByUserNo = "Select * from tblUsers Where UserNo = " & passedUserNo
 
	Set rsGetUserEmailSystemIDByUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetUserEmailSystemIDByUserNo.CursorLocation = 3 
	Set rsGetUserEmailSystemIDByUserNo = cnnGetUserEmailSystemIDByUserNo.Execute(SQLGetUserEmailSystemIDByUserNo)
			 
	If not rsGetUserEmailSystemIDByUserNo.EOF Then 
		resultGetUserEmailSystemIDByUserNo =  rsGetUserEmailSystemIDByUserNo("userEmailSystemID")
	End If
	
	rsGetUserEmailSystemIDByUserNo.Close
	set rsGetUserEmailSystemIDByUserNo= Nothing
	cnnGetUserEmailSystemIDByUserNo.Close	
	set cnnGetUserEmailSystemIDByUserNo= Nothing
	
	If IsNull(resultGetUserEmailSystemIDByUserNo) Then resultGetUserEmailSystemIDByUserNo = ""
		
	GetUserEmailSystemIDByUserNo = resultGetUserEmailSystemIDByUserNo
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetUserEmailSystemPassByUserNo(passedUserNo)

	resultGetUserEmailSystemPassByUserNo = ""

	Set cnnGetUserEmailSystemPassByUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetUserEmailSystemPassByUserNo.open Session("ClientCnnString")
		
	SQLGetUserEmailSystemPassByUserNo = "Select * from tblUsers Where UserNo = " & passedUserNo
 
	Set rsGetUserEmailSystemPassByUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetUserEmailSystemPassByUserNo.CursorLocation = 3 
	Set rsGetUserEmailSystemPassByUserNo = cnnGetUserEmailSystemPassByUserNo.Execute(SQLGetUserEmailSystemPassByUserNo)
			 
	If not rsGetUserEmailSystemPassByUserNo.EOF Then 
		resultGetUserEmailSystemPassByUserNo =  rsGetUserEmailSystemPassByUserNo("userEmailSystemPass")
	End If
	
	rsGetUserEmailSystemPassByUserNo.Close
	set rsGetUserEmailSystemPassByUserNo= Nothing
	cnnGetUserEmailSystemPassByUserNo.Close	
	set cnnGetUserEmailSystemPassByUserNo= Nothing
	
	If IsNull(resultGetUserEmailSystemPassByUserNo) Then resultGetUserEmailSystemPassByUserNo = ""
	
	GetUserEmailSystemPassByUserNo = resultGetUserEmailSystemPassByUserNo
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetUserEmailServerByUserNo(passedUserNo)

	resultGetUserEmailServerByUserNo = ""

	Set cnnGetUserEmailServerByUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetUserEmailServerByUserNo.open Session("ClientCnnString")
		
	SQLGetUserEmailServerByUserNo = "Select * from tblUsers Where UserNo = " & passedUserNo
 
	Set rsGetUserEmailServerByUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetUserEmailServerByUserNo.CursorLocation = 3 
	Set rsGetUserEmailServerByUserNo = cnnGetUserEmailServerByUserNo.Execute(SQLGetUserEmailServerByUserNo)
			 
	If not rsGetUserEmailServerByUserNo.EOF Then 
		resultGetUserEmailServerByUserNo =  rsGetUserEmailServerByUserNo("userEmailServer")
	End If
	
	rsGetUserEmailServerByUserNo.Close
	set rsGetUserEmailServerByUserNo= Nothing
	cnnGetUserEmailServerByUserNo.Close	
	set cnnGetUserEmailServerByUserNo= Nothing
	
	If IsNull(resultGetUserEmailServerByUserNo) Then resultGetUserEmailServerByUserNo = ""
		
	GetUserEmailServerByUserNo = resultGetUserEmailServerByUserNo
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function AllowUpdatesToUsersCalendar(passedUserNo)

	resultAllowUpdatesToUsersCalendar = False

	Set cnnAllowUpdatesToUsersCalendar = Server.CreateObject("ADODB.Connection")
	cnnAllowUpdatesToUsersCalendar.open Session("ClientCnnString")
		
	SQLAllowUpdatesToUsersCalendar = "Select * from tblUsers Where UserNo = " & passedUserNo
 
	Set rsAllowUpdatesToUsersCalendar = Server.CreateObject("ADODB.Recordset")
	rsAllowUpdatesToUsersCalendar.CursorLocation = 3 
	Set rsAllowUpdatesToUsersCalendar = cnnAllowUpdatesToUsersCalendar.Execute(SQLAllowUpdatesToUsersCalendar)
			 
	If not rsAllowUpdatesToUsersCalendar.EOF Then 
		If rsAllowUpdatesToUsersCalendar("userUpdateCalendar") = vbTrue Then resultAllowUpdatesToUsersCalendar = True
	End If
	
	rsAllowUpdatesToUsersCalendar.Close
	set rsAllowUpdatesToUsersCalendar= Nothing
	cnnAllowUpdatesToUsersCalendar.Close	
	set cnnAllowUpdatesToUsersCalendar= Nothing
	
	AllowUpdatesToUsersCalendar = resultAllowUpdatesToUsersCalendar
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function userCanEditEqpOnFly(passedUserNo)

	result = true
	
	If passedUserNo <> "*Not Found*" Then
		
		Set cnn = Server.CreateObject("ADODB.Connection")
		cnn.open Session("ClientCnnString")
	
		SQLUserEditEqpOnFly = "SELECT userEditEqpOnTheFly FROM tblUsers WHERE UserNo = " & passedUserNo
		 
		Set rsUserEditEqpOnFly = Server.CreateObject("ADODB.Recordset")
		rsUserEditEqpOnFly.CursorLocation = 3 
		Set rsUserEditEqpOnFly= cnn.Execute(SQLUserEditEqpOnFly)
		
		If NOT rsUserEditEqpOnFly.EOF then 
			If rsUserEditEqpOnFly("userEditEqpOnTheFly") = 0 then result = false
		End If
		set rsUserEditEqpOnFly= Nothing
		set cnn= Nothing
		
	End If
	
	userCanEditEqpOnFly = result

End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




%>
