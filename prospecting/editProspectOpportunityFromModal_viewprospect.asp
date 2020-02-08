<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%

txtInternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
ProspectName = GetProspectNameByNumber(txtInternalRecordIdentifier)

txtProjectedGPSpend = Request.Form("txtProjectedGPSpend")
txtNumEmployees = Request.Form("txtNumEmployees")
txtNumPantries = Request.Form("txtNumPantries")
txtLeaseExpirationDate = Request.Form("txtLeaseExpirationDate")
txtContractExpirationDate = Request.Form("txtContractExpirationDate")

'*******************************************************************************************************************
'GET ORIGINAL VALUES FOR OPPORTUNITY FIELDS FOR AUDIT TRAIL CHANGES
'*******************************************************************************************************************

	SQLProspect = "SELECT * FROM PR_Prospects WHERE InternalRecordIdentifier = " & txtInternalRecordIdentifier 

	Set cnnProspect = Server.CreateObject("ADODB.Connection")
	cnnProspect.open (Session("ClientCnnString"))
	Set rsProspect = Server.CreateObject("ADODB.Recordset")
	rsProspect.CursorLocation = 3 
	Set rsProspect = cnnProspect.Execute(SQLProspect)

	If not rsProspect.EOF Then
		ORIG_ProjectedGPSpend = rsProspect("ProjectedGPSpend")
		ORIG_NumberOfPantries = rsProspect("NumberOfPantries")
		ORIG_EmployeeRangeNumber = rsProspect("EmployeeRangeNumber")
		ORIG_LeaseExpirationDate = rsProspect("LeaseExpirationDate")	
		ORIG_ContractExpirationDate = rsProspect("ContractExpirationDate")											
	End If
	set rsProspect = Nothing
	cnnProspect.close
	set cnnProspect = Nothing
	

'*******************************************************************************************************************
'SET DEFAULT VALUES FOR ANY NON REQUIRED FIELDS LEFT BLANK DURING THE EDIT PROCESS
'*******************************************************************************************************************

If txtProjectedGPSpend = "" Then txtProjectedGPSpend = "0"
If txtNumEmployees = "" Then txtNumEmployees = "0"
If txtNumPantries = "" Then txtNumPantries = "1"
If txtLeaseExpirationDate = "" Then txtLeaseExpirationDate = ""
If txtContractExpirationDate = "" Then txtContractExpirationDate = ""

If ORIG_ProjectedGPSpend = "" Then ORIG_ProjectedGPSpend = "0"
If ORIG_EmployeeRangeNumber = "" Then ORIG_EmployeeRangeNumber = "0"
If ORIG_NumberOfPantries = "" Then ORIG_NumberOfPantries = "1"
If ORIG_LeaseExpirationDate = "" Then ORIG_LeaseExpirationDate = ""
If ORIG_ContractExpirationDate = "" Then ORIG_ContractExpirationDate = ""




'*******************************************************************************************************************
'PERFORM SQL UPDATE INTO PR_PROSPECTS AND PR_PROSPECTCONTACTS
'*******************************************************************************************************************

	'******************************************
	'Update PR_Prospects
	'******************************************
	
	SQLProspectUpdate = "UPDATE PR_Prospects SET EmployeeRangeNumber = " & cInt(txtNumEmployees) & ", NumberOfPantries = " & cInt(txtNumPantries) & ", ProjectedGPSpend= " & txtProjectedGPSpend & ", "
	SQLProspectUpdate = SQLProspectUpdate & "LeaseExpirationDate = '" & txtLeaseExpirationDate & "', ContractExpirationDate = '" & txtContractExpirationDate & "' "
	SQLProspectUpdate = SQLProspectUpdate & "WHERE InternalRecordIdentifier = " & txtInternalRecordIdentifier 
	
	Response.write(SQLProspectUpdate & "<br><br>")
	
	Set cnnProspectUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectUpdate.open (Session("ClientCnnString"))
	Set rsProspectUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectUpdate.CursorLocation = 3 
	Set rsProspectUpdate = cnnProspectUpdate.Execute(SQLProspectUpdate)
	
	Set rsProspectUpdate = Nothing
	cnnProspectUpdate.Close
	Set cnnProspectUpdate = Nothing


'*******************************************************************************************************************


'*******************************************************************************************************************
'PERFORM AUDIT LOG UPDATE ENTRIES
'*******************************************************************************************************************


	If cDbl(ORIG_ProjectedGPSpend) <> cDbl(txtProjectedGPSpend) Then
		If ORIG_ProjectedGPSpend = "" OR ORIG_ProjectedGPSpend = "0" Then ORIG_ProjectedGPSpend = "NONE ENTERED"
		If txtProjectedGPSpend = "" OR txtProjectedGPSpend = "0" Then txtProjectedGPSpend = "NONE ENTERED"
		Description = "The projected GP spend for prospect " & ProspectName  & " was changed to <strong><em>" & txtProjectedGPSpend & "</em></strong> from <strong><em>" & ORIG_ProjectedGPSpend & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect projected GP spend changed",GetTerm("Prospecting") & " prospect projected GP spend changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If
	

	If cInt(ORIG_EmployeeRangeNumber) = 0 AND cInt(txtNumEmployees) <> 0 AND (cInt(ORIG_EmployeeRangeNumber) <> cInt(txtNumEmployees)) Then
	
		Description = "The employee range for prospect " & ProspectName  & " was changed to <strong><em>" & GetEmployeeRangeByNum(txtNumEmployees) & "</em></strong> from <strong><em>No Employee Range Selected</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect employee range changed",GetTerm("Prospecting") & " prospect employee range changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
				
	ElseIf cInt(ORIG_EmployeeRangeNumber) <> 0 AND cInt(txtNumEmployees) = 0 AND (cInt(ORIG_EmployeeRangeNumber) <> cInt(txtNumEmployees)) Then
		Description = "The employee range for prospect " & ProspectName  & " was changed to <strong><em>No Employee Range Selected</em></strong> from <strong><em>" & GetEmployeeRangeByNum(ORIG_EmployeeRangeNumber) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect employee range changed",GetTerm("Prospecting") & " prospect employee range changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
				
	ElseIf cInt(ORIG_EmployeeRangeNumber) <> cInt(txtNumEmployees) Then
		Description = "The employee range for prospect " & ProspectName  & " was changed to <strong><em>" & GetEmployeeRangeByNum(txtNumEmployees) & "</em></strong> from <strong><em>" & GetEmployeeRangeByNum(ORIG_EmployeeRangeNumber) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect employee range changed",GetTerm("Prospecting") & " prospect employee range changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")		
	End If

	
	If cInt(ORIG_NumberOfPantries) <> cInt(txtNumPantries) Then
		If ORIG_NumberOfPantries = "" Then ORIG_NumberOfPantries = "NONE ENTERED"
		If txtNumPantries = "" Then txtNumPantries = "NONE ENTERED"
		Description = "The number of pantries for prospect " & ProspectName  & " was changed to <strong><em>" & txtNumPantries & "</em></strong> from <strong><em>" & ORIG_NumberOfPantries & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect number of pantries changed",GetTerm("Prospecting") & " prospect number of pantries changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If



	If ORIG_LeaseExpirationDate = "" OR IsNull(ORIG_LeaseExpirationDate) Then ORIG_LeaseExpirationDate = "NONE ENTERED"
	If txtLeaseExpirationDate = "" Then txtLeaseExpirationDate = "NONE ENTERED"

	If ORIG_LeaseExpirationDate <> txtLeaseExpirationDate Then
		
		If ORIG_LeaseExpirationDate <> "NONE ENTERED" AND txtLeaseExpirationDate <> "NONE ENTERED" Then
	
			If DateDiff("d",cDate(ORIG_LeaseExpirationDate),cDate(txtLeaseExpirationDate)) <> 0 Then
				Description = "The lease expiration date for prospect " & ProspectName  & " was changed to <strong><em>" & formatDateTime(txtLeaseExpirationDate,2) & "</em></strong> from <strong><em>" & formatDateTime(ORIG_LeaseExpirationDate,2) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
				CreateAuditLogEntry GetTerm("Prospecting") & " prospect lease expiration date changed",GetTerm("Prospecting") & " prospect lease expiration date changed","Major",0,Description
				Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
			End If
			
		Else
			Description = "The lease expiration date for prospect " & ProspectName  & " was changed to <strong><em>" & txtLeaseExpirationDate & "</em></strong> from <strong><em>" & ORIG_LeaseExpirationDate & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
			CreateAuditLogEntry GetTerm("Prospecting") & " prospect lease expiration date changed",GetTerm("Prospecting") & " prospect lease expiration date changed","Major",0,Description
			Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
		End If
		
	End If

	

	If ORIG_ContractExpirationDate = "" OR IsNull(ORIG_ContractExpirationDate) Then ORIG_ContractExpirationDate = "NONE ENTERED"
	If txtContractExpirationDate = "" Then txtContractExpirationDate = "NONE ENTERED"

	If ORIG_ContractExpirationDate <> txtContractExpirationDate Then
		
		If ORIG_ContractExpirationDate <> "NONE ENTERED" AND txtContractExpirationDate <> "NONE ENTERED" Then
		
			If DateDiff("d",cDate(ORIG_ContractExpirationDate),cDate(txtContractExpirationDate)) <> 0 Then
				Description = "The contract expiration date for prospect " & ProspectName  & " was changed to <strong><em>" & formatDateTime(txtContractExpirationDate,2) & "</em></strong> from <strong><em>" & formatDateTime(ORIG_ContractExpirationDate,2) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
				CreateAuditLogEntry GetTerm("Prospecting") & " prospect contract expiration date changed",GetTerm("Prospecting") & " prospect contract expiration date changed","Major",0,Description
				Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
			End If
			
		Else
			Description = "The contract expiration date for prospect " & ProspectName  & " was changed to <strong><em>" & txtContractExpirationDate & "</em></strong> from <strong><em>" & ORIG_ContractExpirationDate & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
			CreateAuditLogEntry GetTerm("Prospecting") & " prospect contract expiration date changed",GetTerm("Prospecting") & " prospect contract expiration date changed","Major",0,Description
			Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")

		End If
		
	End If

'*******************************************************************************************************************


Response.Redirect ("viewProspectDetail.asp?i=" & txtInternalRecordIdentifier)

%>