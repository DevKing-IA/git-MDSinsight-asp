<% @ Language = VBScript %>

<!--#include file="../inc/SubsAndFuncs.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<%

userResponse = Request.QueryString("resp")
userNo = Request.QueryString("u")
userClientID = Request.QueryString("c")
userProspectIntRecID = Request.QueryString("p")
Session("UserNo") = userNo 

SQLClientID = "SELECT * FROM tblServerInfo where clientKey='"& userClientID &"'"

Set ConnectionClientID  = Server.CreateObject("ADODB.Connection")
Set RecordsetClientID  = Server.CreateObject("ADODB.Recordset")

ConnectionClientID.Open InsightCnnString

'Open the recordset object executing the SQL statement and return records
RecordsetClientID.Open SQLClientID,ConnectionClientID,3,3

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If RecordsetClientID.recordcount <= 0 then
	RecordsetClientID.close
	ConnectionClientID.close
	set RecordsetClientID =nothing
	set ConnectionClientID =nothing
	info = "<font color='yellow'>Invaild Client Key. " & SQLClientID & "</font>"
	
Else
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & RecordsetClientID.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & RecordsetClientID.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & RecordsetClientID.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & RecordsetClientID.Fields("dbPassword") & ";"
	userQuickLoginURL = RecordsetClientID.Fields("QuickLoginURL")
End If



'**************************************************************************
'Set Mail Domain for Prospect Email
'**************************************************************************

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)


'**************************************************************************
'Obtain Prospect Information for Prospect Email
'**************************************************************************

Set cnnProspectInfo = Server.CreateObject("ADODB.Connection")
cnnProspectInfo.open Session("ClientCnnString")

'declare the SQL statement that will query the database
SQLProspectInfo = "SELECT * FROM PR_Prospects WHERE InternalRecordIdentifier = " & userProspectIntRecID 

'Open the recordset object executing the SQL statement and return records
Set rsProspectInfo = Server.CreateObject("ADODB.Recordset")
rsProspectInfo.CursorLocation = 3 
Set rsProspectInfo = cnnProspectInfo.Execute(SQLProspectInfo)

OwnerUserNo = rsProspectInfo("OwnerUserNo")
Company = rsProspectInfo("Company")
Street= rsProspectInfo("Street")
City= rsProspectInfo("City")
State= rsProspectInfo("State")
PostalCode = rsProspectInfo("PostalCode")
LeadSourceNumber = rsProspectInfo("LeadSourceNumber")
LeadSource = GetLeadSourceByNum(LeadSourceNumber)				
StageNumber = GetProspectCurrentStageByProspectNumber(userProspectIntRecID)
IndustryNumber = rsProspectInfo("IndustryNumber")	
Industry = GetIndustryByNum(IndustryNumber)											
OwnerUserNo = rsProspectInfo("OwnerUserNo")				
CreatedDate= rsProspectInfo("CreatedDate")
CreatedByUserNo= rsProspectInfo("CreatedByUserNo")				
TelemarketerUserNo = rsProspectInfo("TelemarketerUserNo")
Telemarketer = GetUserDisplayNameByUserNo(TelemarketerUserNo)
ProjectedGPSpend= rsProspectInfo("ProjectedGPSpend")
NumberOfPantries = rsProspectInfo("NumberOfPantries")
EmployeeRangeNumber = rsProspectInfo("EmployeeRangeNumber")
NumEmployees = GetEmployeeRangeByNum(EmployeeRangeNumber)
CreatedDate = rsProspectInfo("CreatedDate")
FormerCustNum = rsProspectInfo("FormerCustNum")
CancelDate = rsProspectInfo("CancelDate")
LeaseExpirationDate = rsProspectInfo("LeaseExpirationDate")	
ContractExpirationDate = rsProspectInfo("ContractExpirationDate")
Comments = rsProspectInfo("Comments")
CurrentOffering = rsProspectInfo("CurrentOffering")

set rsProspectInfo = Nothing
cnnProspectInfo.Close
set cnnProspectInfo = Nothing

'**************************************************************************
'Obtain Primary Competitor for the Prospect Email
'**************************************************************************

PrimaryCompetitorID = GetPrimaryCompetitorIDByProspectNumber(userProspectIntRecID)

If PrimaryCompetitorID <> "" Then
	PrimaryCompetitorName = GetCompetitorByNum(PrimaryCompetitorID)
Else
	PrimaryCompetitorName = "None Entered"
End If


'**************************************************************************
'Get First and Last Name of the Primary Contact for the Prospect Email
'**************************************************************************

Set cnnProspectContacts = Server.CreateObject("ADODB.Connection")
cnnProspectContacts.open Session("ClientCnnString")

'declare the SQL statement that will query the database
SQLProspectContacts = "SELECT * FROM PR_ProspectContacts WHERE ProspectIntRecID = " & userProspectIntRecID 

'Open the recordset object executing the SQL statement and return records
Set rsProspectContacts = Server.CreateObject("ADODB.Recordset")
rsProspectContacts.CursorLocation = 3 
Set rsProspectContacts = cnnProspectContacts.Execute(SQLProspectContacts)

If NOT rsProspectContacts.EOF Then
  	FirstName = rsProspectContacts("FirstName")
	LastName = rsProspectContacts("LastName")	
Else
 	FirstName = ""
	LastName = ""
End If

set rsProspectContacts = Nothing
cnnProspectContacts.Close
set cnnProspectContacts = Nothing


'**************************************************************************
'Get Next Activity, Appt or Meeting Information for the Prospect Email
'**************************************************************************


SQLNextActivity = "SELECT * FROM PR_ProspectActivities where ProspectRecID = " & userProspectIntRecID & " AND Status IS NULL"

Set cnnNextActivity = Server.CreateObject("ADODB.Connection")
cnnNextActivity.open (Session("ClientCnnString"))
Set rsNextActivity = Server.CreateObject("ADODB.Recordset")
rsNextActivity.CursorLocation = 3 
Set rsNextActivity = cnnNextActivity.Execute(SQLNextActivity)

If not rsNextActivity.EOF Then

  	NextActivityRecID = rsNextActivity("ActivityRecID")
  	NextActivity = GetActivityByNum(rsNextActivity("ActivityRecID"))
	NextActivityDueDate = FormatDateTime(rsNextActivity("ActivityDueDate"),2) & " " & FormatDateTime(rsNextActivity("ActivityDueDate"),3)
	daysOld = DateDiff("d",rsNextActivity("RecordCreationDateTime"),Now())
	daysOverdue = DateDiff("d",rsNextActivity("ActivityDueDate"),Now())
	
	ProspectApptOrMeeting = GetActivityApptOrMeetingByNum(GetCurrentProspectActivityNumberByProspectNumber(userProspectIntRecID)) 
	
	If ProspectApptOrMeeting <> "" Then
	
		If ProspectApptOrMeeting = "Appointment" Then
		
			Duration = rsNextActivity("ActivityAppointmentDuration")
			Location = ""

		ElseIf ProspectApptOrMeeting = "Meeting" Then
		
			Duration = rsNextActivity("ActivityMeetingDuration")
			Location = rsNextActivity("ActivityMeetingLocation")
			
		Else
		
			Duration = ""
			Location = ""
		
		End If
	End If				
End If
Set rsNextActivity = Nothing
cnnNextActivity.Close
Set cnnNextActivity = Nothing



If UCASE(userResponse) = "ACCEPT" Then

	'**************************************************************************
	'Send email to previous owner, letting them know the prospect was accepted
	'**************************************************************************

	Set cnnUsers = Server.CreateObject("ADODB.Connection")
	cnnUsers.open Session("ClientCnnString")
	
	'declare the SQL statement that will query the database
	SQLUsers = "SELECT * FROM tblUsers WHERE userNo= " & OwnerUserNo

	'Open the recordset object executing the SQL statement and return records
	Set rsUsers = Server.CreateObject("ADODB.Recordset")
	rsUsers.CursorLocation = 3 
	Set rsUsers = cnnUsers.Execute(SQLUsers)
	
	'If there is no record with the entered username, close connection
	If rsUsers.EOF then
		set rsUsers = Nothing
		cnnUsers.Close
		set cnnUsers = Nothing
	Else		
		userEmail = rsUsers("userEmail")
		userFirstName = rsUsers("userFirstName")
		userLastName = rsUsers("userLastName")
		userOwnerNo = GetProspectOwnerNoByNumber(userProspectIntRecID)
		userOwnerName = GetUserDisplayNameByUserNo(userOwnerNo)

		%><!--#include file="../emails/prospecting_owner_accepted_email.asp"--><%
		
		SendMail "mailsender@" & maildomain,userEmail,emailSubject,emailBody, GetTerm("Prospecting"), GetTerm("Prospecting") & " Prospect Owner Accepted"

		set rsUsers = Nothing
		cnnUsers.Close
		set cnnUsers = Nothing	
	End If	

	'***************************************
	'Update owner in PR_Prospects
	'***************************************

	Set cnnProspectUpdateOwner = Server.CreateObject("ADODB.Connection")
	cnnProspectUpdateOwner.open Session("ClientCnnString")
	
	SQLProspectUpdateOwner = "UPDATE PR_Prospects Set OwnerUserNo = " & userNo & " WHERE InternalRecordIdentifier = " & userProspectIntRecID 
	
	Set rsProspectUpdateOwner = Server.CreateObject("ADODB.Recordset")
	rsProspectUpdateOwner.CursorLocation = 3 
	Set rsProspectUpdateOwner = cnnProspectUpdateOwner.Execute(SQLProspectUpdateOwner)
	
	Description = GetUserDisplayNameByUserNo(userNo) & " accepted ownership for prospect " & GetProspectNameByNumber(userProspectIntRecID) 
	CreateAuditLogEntry GetTerm("Prospecting") & " ownership accepted via email",GetTerm("Prospecting") & " ownership accepted via email","Major",0,Description
	Record_PR_Activity userProspectIntRecID,Description,userNo
		
		
	set rsProspectUpdateOwner = Nothing
	cnnProspectUpdateOwner.Close
	set cnnProspectUpdateOwner = Nothing

	'***********************************************
	'Update next activity in PR_ProspectActivities
	'***********************************************

	Set cnnProspectNextActivityUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectNextActivityUpdate.open Session("ClientCnnString")
	
	SQLProspectNextActivityUpdate = "UPDATE PR_ProspectActivities Set ActivityCreatedByUserNo = " & userNo & " WHERE ProspectRecID = " & userProspectIntRecID & " AND Status IS NULL"
	
	Set rsProspectNextActivityUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectNextActivityUpdate.CursorLocation = 3 
	Set rsProspectNextActivityUpdate = cnnProspectNextActivityUpdate.Execute(SQLProspectNextActivityUpdate)
	
	Description = "The next activity for this prospect " & GetProspectNameByNumber(userProspectIntRecID) & " was reassigned to " & GetUserDisplayNameByUserNo(userNo) & " as a result of clicking the acceptance link in the email."
	CreateAuditLogEntry GetTerm("Prospecting") & " next activity reassigned",GetTerm("Prospecting") & " next activity reassigned","Major",0,Description
	Record_PR_Activity userProspectIntRecID,Description,userNo
	
	set rsProspectNextActivityUpdate = Nothing
	cnnProspectNextActivityUpdate.Close
	set cnnProspectNextActivityUpdate = Nothing
	
	dummy =	Prospect_Email_Accept(userProspectIntRecID,userNo)
	
	%><h1>Prospect has been successfully accepted. An email notification has been sent to the requesting user.</h1><%

	
Else


	Set cnnUsers = Server.CreateObject("ADODB.Connection")
	cnnUsers.open Session("ClientCnnString")
	
	'declare the SQL statement that will query the database
	SQLUsers = "SELECT * FROM tblUsers WHERE userNo= " & OwnerUserNo

	'Open the recordset object executing the SQL statement and return records
	Set rsUsers = Server.CreateObject("ADODB.Recordset")
	rsUsers.CursorLocation = 3 
	Set rsUsers = cnnUsers.Execute(SQLUsers)
	
	'If there is no record with the entered username, close connection
	If rsUsers.EOF then
		set rsUsers = Nothing
		cnnUsers.Close
		set cnnUsers = Nothing
	Else		
		userEmail = rsUsers("userEmail")
		userFirstName = rsUsers("userFirstName")
		userLastName = rsUsers("userLastName")
		userOwnerNo = GetProspectOwnerNoByNumber(userProspectIntRecID)
		userOwnerName = GetUserDisplayNameByUserNo(userOwnerNo)
				
		%><!--#include file="../emails/prospecting_owner_rejected_email.asp"--><%
		
		SendMail "mailsender@" & maildomain,userEmail,emailSubject,emailBody, GetTerm("Prospecting"), GetTerm("Prospecting") & " Prospect Owner Rejected"

		Description = GetTerm("Prospecting") & " Prospect Owner Request Rejected email sent to " & userEmail & " for user " & userFirstName & " " & userLastName
	
		CreateAuditLogEntry "Prospect Owner Request Rejected Emailed","Prospect Owner Request Rejected Emailed","Minor",0,Description 

		set rsUsers = Nothing
		cnnUsers.Close
		set cnnUsers = Nothing		
	End If	

	Description = GetUserDisplayNameByUserNo(userNo) & " rejected ownership for prospect " & GetProspectNameByNumber(userProspectIntRecID) 
	CreateAuditLogEntry GetTerm("Prospecting") & " owernship rejected via email",GetTerm("Prospecting") & " owernship rejected via email","Major",0,Description
	
	Description = GetUserDisplayNameByUserNo(userNo) & " rejected ownership for prospect " & GetProspectNameByNumber(userProspectIntRecID) 
	Record_PR_Activity userProspectIntRecID,Description,userNo
	
	%><h1>Prospect has been successfully rejected. An email notification has been sent back to the current owner.</h1><%


End If


%>