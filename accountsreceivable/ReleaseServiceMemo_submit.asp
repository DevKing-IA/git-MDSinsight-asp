<!--#include file="../inc/header-field-service.asp"-->

<!--#include file="../inc/mail.asp"-->
<%
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

Account = Request.Form("txtCustID")
sURL = Request.ServerVariables("SERVER_NAME")
MemoNumber = Request.Form("txtMemoNumber")
ReleaseNotes = Request.Form("releasenotes")
SaveOrRelease = Request.Form("btnSaveOrRelease")

'Quick and dirty bug correction
If Instr(Account,",") <> 0 Then
	Account = trim(Left(Account,Instr(Account,",")-1))
End If


'Trying a new way
MultiPost = 0
If GetServiceTicketStatus(MemoNumber) <> "HOLD" Then
		MultiPost = 1
		emailBody = "Memo Nunber: " & MemoNumber
		SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com","Bad Release -" & GetPOSTParams("Mode"),emailBody
Else
	If GetServiceTicketCurrentStage(MemoNumber) = "Released" Then
		MultiPost = 1
		emailBody = "Memo Nunber: " & MemoNumber
		SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com","Bad Release - Already Released" & GetPOSTParams("Mode"),emailBody
	End If
End If



If SaveOrRelease = "Save" Then
	'If we are only saving the notes then we do the stuff in this IF statement
	'If we are releasing, it will skip this & do everything below
	Set Connection2 = Server.CreateObject("ADODB.Connection")
	Set Recordset2 = Server.CreateObject("ADODB.Recordset")
	Recordset2.CursorLocation = 3 
	Connection2.Open Session("ClientCnnString")
	
	SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
	SQL = SQL & "SubmissionDateTime, USerNoSubmittingRecord,Remarks)"
	SQL = SQL &  " VALUES (" 
	SQL = SQL & "'"  & MemoNumber & "'"
	SQL = SQL & ",'"  & OpenAccountNumber & "'"
	SQL = SQL & ",'Under Review'"
	SQL = SQL & ",getdate() "
	SQL = SQL & ","  & Session("UserNo")
	SQL = SQL & ",'"  & ReleaseNotes & "')"
	Set Connection2 = Server.CreateObject("ADODB.Connection")
	Set Recordset2 = Server.CreateObject("ADODB.Recordset")
	Recordset2.CursorLocation = 3 
	Connection2.Open Session("ClientCnnString")
	Set Recordset2 = Connection2.Execute(SQL)
	Connection2.Close
	Set Recordset2 = Nothing
	Set Connection2 = Nothing
	
	Response.Redirect("TicketsOnHold.asp") ' Must do this so it doesn't release at this point
	
End If

If MultiPost = 0 Then ' Stop multiple posting

	'********************
	'Advanced dispatching   
	'*********************
	'If advanced dispatching is on we must mark tickets as released
	'First we emight need to lookup which dispatch method is used
	
	
	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Recordset.CursorLocation = 3 
	Connection.Open InsightCnnString
	

	' If it was only received, this would be the first deail record
	SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
	SQL = SQL & "SubmissionDateTime, USerNoSubmittingRecord,Remarks)"
	SQL = SQL &  " VALUES (" 
	SQL = SQL & "'"  & MemoNumber & "'"
	SQL = SQL & ",'"  & Account & "'"
	SQL = SQL & ",'Released'"
	SQL = SQL & ",getdate() "
	SQL = SQL & ","  & Session("UserNo")
	SQL = SQL & ",'"  & ReleaseNotes & "')"
	Set Connection2 = Server.CreateObject("ADODB.Connection")
	Set Recordset2 = Server.CreateObject("ADODB.Recordset")
	Recordset2.CursorLocation = 3 
	Connection2.Open Session("ClientCnnString")
	Set Recordset2 = Connection2.Execute(SQL)
	Connection2.Close
	Set Recordset2 = Nothing
	Set Connection2 = Nothing


	Connection.Close
	Set Recordset = Nothing
	Set Connection = Nothing

	
	SQL = "UPDATE FS_ServiceMemos SET CurrentStatus = 'OPEN', "
	SQL = SQL & "ReleasedByUserNo = " & Session("UserNo") 
	SQL = SQL & ",ReleasedDateTime = getdate() "
	If ReleaseNotes <> "" Then SQL = SQL & ", ReleasedNotes = '" & ReleaseNotes & "'"
	SQL = SQL & " WHERE MemoNumber = '" & MemoNumber & "'"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	'Now we must basically Clone this to make a record with an OPEN status
	SQL = "INSERT INTO FS_ServiceMemos (MemoNumber, CurrentStatus, RecordSubType, SubmittedByName, AccountNumber, Company, ProblemLocation, SubmittedByPhone, SubmittedByEmail, "
	SQL = SQL & " ProblemDescription, Mode, SubmissionSource, UserNoOfServiceTech, Dispatched, AlertEmailSent, EscalationAlertEmailSent, HoldAlertEmailSent, "
	SQL = SQL & " HoldEscalationAlertEmailSent, ReleasedDateTime, ReleasedByUserNo, ReleasedNotes, FilterChange,PMCall) "
	SQL = SQL & " Select MemoNumber, 'OPEN', 'OPEN', SubmittedByName, AccountNumber, Company, ProblemLocation, SubmittedByPhone, SubmittedByEmail, "
	SQL = SQL & " ProblemDescription, Mode, SubmissionSource, UserNoOfServiceTech, Dispatched, AlertEmailSent, EscalationAlertEmailSent, HoldAlertEmailSent, "
	SQL = SQL & " HoldEscalationAlertEmailSent, ReleasedDateTime, ReleasedByUserNo, ReleasedNotes, FilterChange, PMCall "
	SQL = SQL & " FROM FS_ServiceMemos Where MemoNumber = '" & MemoNumber  & "'"
	Set rs = cnn8.Execute(SQL)
	
	Set rs = Nothing
	cnn8.Close
	Set cnn8 = Nothing
	
	
	If ServiceNotes = "" Then ServiceNotes = "No Service Notes Provided"
	ServiceNotes = Replace(ServiceNotes,"&","%26") 
	ServiceNotes = Replace(ServiceNotes," ","%20") & "%0A%0D"
	
	data = "create_service_request&account="
	
	data = data & Account
	data = data & "&st=RELEASE"
	data = data & "&tnum=" & MemoNumber
	data = data & "&prob=" 
	data = data & ServiceNotes & "%0A%0D"
	data = data & "Released from hold by: "  & 	GetUserDisplayNameByUserNo(Session("UserNo")) & " - " & GetUserEmailByUserNo(Session("UserNo")) & "%0A%0D"
	
	data = data & "Submitted via MDS Insight" & "%0A%0D"
	
	
	data = data & "&md=" &  GetPOSTParams("Mode")
	data = data & "&serno="  & GetPOSTParams("Serno")
	data = data & "&src=MDS Insight"
	
	Description = "Service ticket #: " & MemoNumber & "-  "
	Description = Description & "Released from hold by: "  & 	GetUserDisplayNameByUserNo(Session("UserNo")) & " - " & GetUserEmailByUserNo(Session("UserNo")) 
	
	
	Description = Description & ",     Account: "  & Account & " - " & Company 
	Description = Description & ",    Submitted via MDS Insight" & "%0A%0D" 
	Description = Replace(Description ,"%20"," ")
	
	CreateAuditLogEntry "Service Ticket Released From Hold","Service Ticket Released From Hold","Major",0,Description

	For x = 1 to 2

		DO_Post = 0
		
		If x = 1 Then Do_Post = GetPOSTParams("ServiceMemoURL1ONOFF") 
		If x = 2 Then Do_Post = GetPOSTParams("ServiceMemoURL2ONOFF") 
		
		If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0
	
		If cint(Do_Post) = 1 Then
		
				CreateINSIGHTAuditLogEntry sURL,"Post Loop "& x,GetPOSTParams("Mode")
				
				If x = 1 Then
					Description = "Post to " & GetPOSTParams("ServiceMemoURL1")
				Else
					Description = "Post to " & GetPOSTParams("ServiceMemoURL2")
				End If

				CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
				Description = "data:" & data 
				CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
				
				Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
				
				If x = 1 Then
					httpRequest.Open "POST", GetPOSTParams("ServiceMemoURL1"), False
				Else
					httpRequest.Open "POST", GetPOSTParams("ServiceMemoURL2"), False				
				End If

				httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				httpRequest.Send data
					
				IF httpRequest.status = 200 THEN 
					Description = "httpRequest.responseText:" & httpRequest.responseText
					CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
					postResponse = httpRequest.responseText
				ELSE
					'In here it must email us if there are problems
					Description = "httpRequest.responseText:" & httpRequest.responseText
					CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
					postResponse= "Could not get data to " & GetTerm("Backend")
				END IF
					
				If postResponse <> "success" then 
					postResponse = httpRequest.responseText
					'In here it must email us if there are problems
				End If
	
				Set httpRequest = Nothing
		End IF
	Next
	
End If
Response.Redirect("TicketsOnHold.asp")
 
%><!--#include file="../inc/footer-field-service-noTimeout.asp"-->