<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%

txtInternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
ProspectName = GetProspectNameByNumber(txtInternalRecordIdentifier)

txtSuffix = Request.Form("txtSuffix")
txtFirstName = Request.Form("txtFirstName")
txtLastName = Request.Form("txtLastName")
txtTitle = Request.Form("txtTitle")
txtCompanyName = Request.Form("txtCompanyName")
txtAddressLine1 = Request.Form("txtAddressLine1")
txtAddressLine2 = Request.Form("txtAddressLine2")
txtCity = Request.Form("txtCity")
txtState = Request.Form("txtState")
txtZipCode = Request.Form("txtZipCode")
txtCountry = Request.Form("txtCountry")
txtEmailAddress = Request.Form("txtEmailAddress")
txtWebsiteURL = Request.Form("txtWebsiteURL")
txtPhoneNumber = Request.Form("txtPhoneNumber")
txtPhoneNumberExt = Request.Form("txtPhoneNumberExt")
txtCellPhoneNumber = Request.Form("txtCellPhoneNumber")
txtFaxNumber = Request.Form("txtFaxNumber")
txtIndustry = Request.Form("txtIndustry")


'*******************************************************************************************************************
'GET ORIGINAL VALUES FOR BUSINESS CARDS FIELDS FOR AUDIT TRAIL CHANGES
'*******************************************************************************************************************

	SQLProspect = "SELECT * FROM PR_Prospects WHERE InternalRecordIdentifier = " & txtInternalRecordIdentifier 

	Set cnnProspect = Server.CreateObject("ADODB.Connection")
	cnnProspect.open (Session("ClientCnnString"))
	Set rsProspect = Server.CreateObject("ADODB.Recordset")
	rsProspect.CursorLocation = 3 
	Set rsProspect = cnnProspect.Execute(SQLProspect)

	If not rsProspect.EOF Then
		ORIG_Company = rsProspect("Company")
		ORIG_Street = rsProspect("Street")
		ORIG_Suite = rsProspect("Floor_Suite_Room__c")
		ORIG_City = rsProspect("City")
		ORIG_State = rsProspect("State")
		ORIG_PostalCode = rsProspect("PostalCode")
		ORIG_Country = rsProspect("Country")
		ORIG_Website = rsProspect("Website")	
		ORIG_IndustryNumber = rsProspect("IndustryNumber")
	End If
	set rsProspect = Nothing
	cnnProspect.close
	set cnnProspect = Nothing

'*******************************************************************************************************************
'GET ORIGINAL VALUES FOR BUSINESS CARDS FIELDS FOR AUDIT TRAIL CHANGES FROM PRIMARY CONTACT
'*******************************************************************************************************************

	SQLProspect = "SELECT * FROM PR_ProspectContacts WHERE ProspectIntRecID = " & txtInternalRecordIdentifier & " AND PrimaryContact = 1"

	Set cnnProspect = Server.CreateObject("ADODB.Connection")
	cnnProspect.open (Session("ClientCnnString"))
	Set rsProspectContact = Server.CreateObject("ADODB.Recordset")
	rsProspectContact.CursorLocation = 3 
	Set rsProspectContact = cnnProspect.Execute(SQLProspect)

	If not rsProspectContact.EOF Then
		ORIG_primarySuffix = rsProspectContact("Suffix")
		ORIG_primaryFirstName = rsProspectContact("FirstName")
		ORIG_primaryLastName = rsProspectContact("LastName")
		ORIG_primaryTitleNumber = rsProspectContact("ContactTitleNumber")
		ORIG_primaryEmail = rsProspectContact("Email")
		ORIG_primaryPhone = rsProspectContact("Phone")
		ORIG_primaryPhoneExt = rsProspectContact("PhoneExt")
		ORIG_primaryCell = rsProspectContact("Cell")
		ORIG_primaryFax = rsProspectContact("Fax")
	End If
	set rsProspectContact = Nothing
	cnnProspect.close
	set cnnProspect = Nothing




'*******************************************************************************************************************
'SET DEFAULT VALUES FOR ANY NON REQUIRED FIELDS LEFT BLANK DURING THE EDIT PROCESS
'*******************************************************************************************************************

If txtTitle = "" Then txtTitle = "0"
If txtIndustry = "" Then txtIndustry = "0"
If ORIG_primaryTitleNumber = "" Then ORIG_primaryTitleNumber = "0"
If ORIG_IndustryNumber = "" Then ORIG_IndustryNumber = "0"

'*******************************************************************************************************************

'*******************************************************************************************************************
'PERFORM SQL UPDATE INTO PR_PROSPECTS AND PR_PROSPECTCONTACTS
'*******************************************************************************************************************

	'******************************************
	'Update PR_Prospects
	'******************************************

	SQLProspectUpdate = "UPDATE PR_Prospects SET Company = '" & txtCompanyName & "', Street = '" & txtAddressLine1 & "', City = '" & txtCity & "', State = '" & txtState & "', "
	SQLProspectUpdate = SQLProspectUpdate & "PostalCode = '" & txtZipCode & "', Country = '" & txtCountry & "', Floor_Suite_Room__c = '" & txtAddressLine2 & "', "
	SQLProspectUpdate = SQLProspectUpdate & "Website = '" & txtWebsiteURL & "', IndustryNumber = " & txtIndustry & " "
	SQLProspectUpdate = SQLProspectUpdate & "WHERE InternalRecordIdentifier = " & txtInternalRecordIdentifier 
	
	'Response.write(SQLProspectUpdate & "<br><br>")
	
	Set cnnProspectUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectUpdate.open (Session("ClientCnnString"))
	Set rsProspectUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectUpdate.CursorLocation = 3 
	Set rsProspectUpdate = cnnProspectUpdate.Execute(SQLProspectUpdate)
	
	Set rsProspectUpdate = Nothing
	cnnProspectUpdate.Close
	Set cnnProspectUpdate = Nothing
		

	
	'******************************************
	'Update/Insert into PR_ProspectContacts
	'******************************************
	
	SQLProspectContactCheck = "SELECT * FROM PR_ProspectContacts WHERE ProspectIntRecID = " & txtInternalRecordIdentifier & " AND PrimaryContact = 1"
	Set cnnProspectContactCheck = Server.CreateObject("ADODB.Connection")
	cnnProspectContactCheck.open (Session("ClientCnnString"))
	Set rsProspectContactCheck = Server.CreateObject("ADODB.Recordset")
	rsProspectContactCheck.CursorLocation = 3 
	Set rsProspectContactCheck = cnnProspectContactCheck.Execute(SQLProspectContactCheck)
	
	If NOT rsProspectContactCheck.EOF Then
		SQLProspectContactUpdate = "UPDATE PR_ProspectContacts SET Suffix = '" & txtSuffix & "', FirstName = '" & txtFirstName & "', LastName = '" & txtLastName & "', "
		SQLProspectContactUpdate = SQLProspectContactUpdate & "ContactTitleNumber = " & txtTitle & ", Email = '" & txtEmailAddress & "', Phone = '" & txtPhoneNumber & "', "
		SQLProspectContactUpdate = SQLProspectContactUpdate & " PhoneExt = '" & txtPhoneNumberExt & "', Cell = '" & txtCellPhoneNumber & "', Fax = '" & txtFaxNumber & "' "
		SQLProspectContactUpdate = SQLProspectContactUpdate & " WHERE ProspectIntRecID = " & txtInternalRecordIdentifier & " AND PrimaryContact = 1"
	Else
		SQLProspectContactUpdate = "INSERT INTO PR_ProspectContacts (ProspectIntRecID, Suffix, FirstName, LastName, ContactTitleNumber, "
		SQLProspectContactUpdate = SQLProspectContactUpdate & " Email, Phone, PhoneExt, Cell, Fax, PrimaryContact) VALUES ( "
		SQLProspectContactUpdate = SQLProspectContactUpdate & txtInternalRecordIdentifier & ",'"  & txtSuffix & "','"  & txtFirstName & "','"  & txtLastName & "', "
		SQLProspectContactUpdate = SQLProspectContactUpdate & txtTitle & ",'" & txtEmailAddress & "','" & txtPhoneNumber & "','" & txtPhoneNumberExt & "','" & txtCellPhoneNumber & "',"
		SQLProspectContactUpdate = SQLProspectContactUpdate & "'" & txtFaxNumber & "',1)"
	End If
	
	Set rsProspectContactCheck = Nothing
	cnnProspectContactCheck.Close
	Set cnnProspectContactCheck = Nothing
	
	'Response.write(SQLProspectContactUpdate)
	
	Set cnnProspectContactUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectContactUpdate.open (Session("ClientCnnString"))
	Set rsProspectContactUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectContactUpdate.CursorLocation = 3 
	Set rsProspectContactUpdate = cnnProspectContactUpdate.Execute(SQLProspectContactUpdate)
	
	Set rsProspectContactUpdate = Nothing
	cnnProspectContactUpdate.Close
	Set cnnProspectContactUpdate = Nothing


'*******************************************************************************************************************

'*******************************************************************************************************************
'PERFORM AUDIT LOG UPDATE ENTRIES
'*******************************************************************************************************************	

	If ORIG_Company <> txtCompanyName Then
		If ORIG_Company = "" Then ORIG_COMPANY = "NONE ENTERED"
		If txtCompanyName  = "" Then txtCompanyName = "NONE ENTERED"
		Description = "The company name for prospect " & ProspectName  & " was changed to <strong><em>" & txtCompanyName & "</em></strong> from <strong><em>" & ORIG_Company & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect company name changed",GetTerm("Prospecting") & " prospect company name changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
	End If
	
	If ORIG_Street <> txtAddressLine1 Then
		If ORIG_Street = "" Then ORIG_Street = "NONE ENTERED"
		If txtAddressLine1 = "" Then txtAddressLine1 = "NONE ENTERED"	
		Description = "The street address for prospect " & ProspectName  & " was changed to <strong><em>" & txtAddressLine1 & "</em></strong> from <strong><em>" & ORIG_Street & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect street address changed",GetTerm("Prospecting") & " prospect street address changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If

	If ORIG_Suite <> txtAddressLine2 Then
		If ORIG_Suite = "" Then ORIG_Suite = "NONE ENTERED"
		If txtAddressLine2 = "" Then txtAddressLine2 = "NONE ENTERED"	
		Description = "The suite/floor number for prospect " & ProspectName  & " was changed to <strong><em>" & txtAddressLine2 & "</em></strong> from <strong><em>" & ORIG_Suite & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect suite/floor number changed",GetTerm("Prospecting") & " prospect suite/floor number changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If
	
	If ORIG_City <> txtCity Then
		If ORIG_City = "" Then ORIG_City = "NONE ENTERED"
		If txtCity = "" Then txtCity = "NONE ENTERED"
		Description = "The city for prospect " & ProspectName  & " was changed to <strong><em>" & txtCity & "</em></strong> from <strong><em>" & ORIG_City & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect city changed",GetTerm("Prospecting") & " prospect city changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If

	If ORIG_State <> txtState Then
		If ORIG_State = "" Then ORIG_State = "NONE ENTERED"
		If txtState = "" Then txtState = "NONE ENTERED"
		Description = "The state for prospect " & ProspectName  & " was changed to <strong><em>" & txtState & "</em></strong> from <strong><em>" & ORIG_State & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect state changed",GetTerm("Prospecting") & " prospect state changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If
	
	If ORIG_PostalCode <> txtZipCode Then
		If ORIG_PostalCode = "" Then ORIG_PostalCode = "NONE ENTERED"
		If txtZipCode = "" Then txtZipCode = "NONE ENTERED"
		Description = "The zip code for prospect " & ProspectName  & " was changed to <strong><em>" & txtZipCode & "</em></strong> from <strong><em>" & ORIG_PostalCode & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect zip code changed",GetTerm("Prospecting") & " prospect zip code changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If

	If ORIG_Country <> txtCountry Then
		If ORIG_Country = "" Then ORIG_Country = "NONE ENTERED"
		If txtCountry = "" Then txtCountry = "NONE ENTERED"
		Description = "The country for prospect " & ProspectName  & " was changed to <strong><em>" & txtCountry & "</em></strong> from <strong><em>" & ORIG_Country & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect country changed",GetTerm("Prospecting") & " prospect country changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If
		
	If ORIG_Website <> 	txtWebsiteURL Then
		If ORIG_Website = "" Then ORIG_Website = "NONE ENTERED"
		If txtWebsiteURL = "" Then txtWebsiteURL = "NONE ENTERED"
		Description = "The website URL for prospect " & ProspectName  & " was changed to <strong><em>" & txtWebsiteURL & "</em></strong> from <strong><em>" & ORIG_Website & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect  website URL changed",GetTerm("Prospecting") & " prospect  website URL changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If
					

	
	If cInt(ORIG_IndustryNumber) = 0 AND cInt(txtIndustry) <> 0 AND (cInt(ORIG_IndustryNumber) <> cInt(txtIndustry)) Then
	
		Description = "The industry for prospect " & ProspectName  & " was changed to <strong><em>" & GetIndustryByNum(txtIndustry) & "</em></strong> from <strong><em>No Industry Selected</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect industry changed",GetTerm("Prospecting") & " prospect industry changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
				
	ElseIf cInt(ORIG_IndustryNumber) <> 0 AND cInt(txtIndustry) = 0 AND (cInt(ORIG_IndustryNumber) <> cInt(txtIndustry)) Then
		Description = "The industry for prospect " & ProspectName  & " was changed to <strong><em>No Industry Set</em></strong> from <strong><em>" & GetIndustryByNum(ORIG_IndustryNumber) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect industry changed",GetTerm("Prospecting") & " prospect industry changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
				
	ElseIf cInt(ORIG_IndustryNumber) <> cInt(txtIndustry) Then
		Description = "The industry for prospect " & ProspectName  & " was changed to <strong><em>" & GetIndustryByNum(txtIndustry) & "</em></strong> from <strong><em>" & GetIndustryByNum(ORIG_IndustryNumber) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect industry changed",GetTerm("Prospecting") & " prospect industry changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")		
	End If
	
	
	
  	If ORIG_primarySuffix <> txtSuffix Then
  		If ORIG_primarySuffix = "" Then ORIG_primarySuffix = "NONE ENTERED"
		If txtSuffix = "" Then txtSuffix = "NONE ENTERED"
		Description = "The contact salutation for prospect " & ProspectName  & " was changed to <strong><em>" & txtSuffix & "</em></strong> from <strong><em>" & ORIG_primarySuffix & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect contact salutation changed",GetTerm("Prospecting") & " prospect contact salutation changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If

  	If ORIG_primaryFirstName <> txtFirstName Then
  		If ORIG_primaryFirstName = "" Then ORIG_primaryFirstName = "NONE ENTERED"
		If txtFirstName = "" Then txtFirstName = "NONE ENTERED"
		Description = "The contact first name for prospect " & ProspectName  & " was changed to <strong><em>" & txtFirstName & "</em></strong> from <strong><em>" & ORIG_primaryFirstName & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect contact first name changed",GetTerm("Prospecting") & " prospect contact first name changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If

	
	If ORIG_primaryLastName <> txtLastName Then
		If ORIG_primaryLastName = "" Then ORIG_primaryLastName = "NONE ENTERED"
		If txtLastName = "" Then txtLastName = "NONE ENTERED"
		Description = "The contact last name for prospect " & ProspectName  & " was changed to <strong><em>" & txtLastName & "</em></strong> from <strong><em>" & ORIG_primaryLastName & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect contact last name changed",GetTerm("Prospecting") & " prospect contact last name changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If
	
	
	'Response.write("ORIG_primaryTitleNumber: " & cInt(ORIG_primaryTitleNumber) & "<br>")
	'Response.write("txtTitle: " & cInt(txtTitle) & "<br>")
		
	If cInt(ORIG_primaryTitleNumber) = 0 AND cInt(txtTitle) <> 0 AND (cInt(ORIG_primaryTitleNumber) <> cInt(txtTitle)) Then
	
		Description = "The contact title for prospect " & ProspectName  & " was changed to <strong><em>" & GetContactTitleByNum(txtTitle) & "</em></strong> from <strong><em>No Title Selected</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect contact title changed",GetTerm("Prospecting") & " prospect contact title changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
				
	ElseIf cInt(ORIG_primaryTitleNumber) <> 0 AND cInt(txtTitle) = 0 AND (cInt(ORIG_primaryTitleNumber) <> cInt(txtTitle)) Then
		Description = "The contact title for prospect " & ProspectName  & " was changed to <strong><em>No Title Selected</em></strong> from <strong><em>" & GetContactTitleByNum(ORIG_primaryTitleNumber) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect contact title changed",GetTerm("Prospecting") & " prospect contact title changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
				
	ElseIf cInt(ORIG_primaryTitleNumber) <> cInt(txtTitle) Then
		Description = "The contact title for prospect " & ProspectName  & " was changed to <strong><em>" & GetContactTitleByNum(txtTitle) & "</em></strong> from <strong><em>" & GetContactTitleByNum(ORIG_primaryTitleNumber) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect contact title changed",GetTerm("Prospecting") & " prospect contact title changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
	
	End If
	
	
	
	If ORIG_primaryEmail <> txtEmailAddress Then
		If ORIG_primaryEmail = "" Then ORIG_primaryEmail = "NONE ENTERED"
		If txtEmailAddress = "" Then txtEmailAddress = "NONE ENTERED"
		Description = "The contact email address for prospect " & ProspectName  & " was changed to <strong><em>" & txtEmailAddress & "</em></strong> from <strong><em>" & ORIG_primaryEmail & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect contact email address changed",GetTerm("Prospecting") & " prospect contact email address changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If
	
	If ORIG_primaryPhone <> txtPhoneNumber Then
		If ORIG_primaryPhone = "" Then ORIG_primaryPhone = "NONE ENTERED"
		If txtPhoneNumber = "" Then txtPhoneNumber = "NONE ENTERED"
		Description = "The contact phone number for prospect " & ProspectName  & " was changed to <strong><em>" & txtPhoneNumber & "</em></strong> from <strong><em>" & ORIG_primaryPhone & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect contact phone number changed",GetTerm("Prospecting") & " prospect contact phone number changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If
	
	If ORIG_primaryPhoneExt <> txtPhoneNumberExt Then
		If ORIG_primaryPhoneExt = "" Then ORIG_primaryPhoneExt = "NONE ENTERED"
		If txtPhoneNumberExt = "" Then txtPhoneNumberExt = "NONE ENTERED"
		Description = "The contact phone number extension for prospect " & ProspectName  & " was changed to <strong><em>" & txtPhoneNumberExt & "</em></strong> from <strong><em>" & ORIG_primaryPhoneExt & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect contact phone number extension changed",GetTerm("Prospecting") & " prospect contact phone number changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If

	If ORIG_primaryCell <> txtCellPhoneNumber Then
		If ORIG_primaryCell = "" Then ORIG_primaryCell = "NONE ENTERED"
		If txtCellPhoneNumber = "" Then txtCellPhoneNumber = "NONE ENTERED"
		Description = "The contact cell phone number for prospect " & ProspectName  & " was changed to <strong><em>" & txtCellPhoneNumber & "</em></strong> from <strong><em>" & ORIG_primaryCell & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect contact cell phone number changed",GetTerm("Prospecting") & " prospect contact cell phone number changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If

	If ORIG_primaryFax <> txtFaxNumber Then
		If ORIG_primaryFax = "" Then ORIG_primaryFax = "NONE ENTERED"
		If txtFaxNumber = "" Then txtFaxNumber = "NONE ENTERED"
		Description = "The contact fax number for prospect " & ProspectName  & " was changed to <strong><em>" & txtFaxNumber & "</em></strong> from <strong><em>" & ORIG_primaryFax & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect contact fax number changed",GetTerm("Prospecting") & " prospect contact fax number changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
	End If

'*******************************************************************************************************************

Response.Redirect ("viewProspectDetail.asp?i=" & txtInternalRecordIdentifier)

%>