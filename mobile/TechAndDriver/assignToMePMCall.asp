<%'Create a service memo & then dispatches it to ME automatically%>
<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/mail.asp"-->

<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

AssetNumber = Request.Form("txtAssetNumber")
'Lookup everything else we need
SQL = "SELECT * FROM Assets "
SQL = SQL & "INNER JOIN EQ_ScheduledServiceDates ON Assets.assetNumber = EQ_ScheduledServiceDates.assetNumber "
SQL = SQL & "WHERE EQ_ScheduledServiceDates.assetNumber = '" & AssetNumber & "'"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.eof then
	AccountNumber = rs("custAcctNum")
	Company = GetCustNameByCustNum(rs("custAcctNum"))
	ProblemDescription = "PREVENTATIVE MAINTENANCE"
	SumissionSource = "MDS Insight"
	PMCalldate = rs("nextDate1")
	PMCall = rs("Comment1")
end if
set rs = nothing
cnn8.close
set cnn8 = nothing

' Write Audit trail first, then post
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = "OPEN,    "
Description = Description & "Account: "  & AccountNumber & " - " & Company 
Description = Description & ",    Description: "  & ProblemDescription 
CreateAuditLogEntry "Service Memo Added","Service Memo Added","Major",0,Description

'Post to MDS goes here
data = "create_service_request&account="
data = data & AccountNumber
data = data & "&prob=PREVENTATIVE MAINTENANCE - " & PMCall & "%0A%0D"
data = data & "Opened by: " & 	MUV_Read("DisplayName") & " - " & Session("UserEmail")
data = data & "&probloc=" & ProblemLocation 
data = data & "&sbn=" & ContactName 
data = data & "&sbe=" & ContactEmail
data = data & "&sbp=" & ContactPhone 
data = data & "&contnm=" & ContactName
data = data & "&st=" & "OPEN"
data = data & "&md=" & GetPOSTParams("Mode")
data = data & "&serno=" & GetPOSTParams("Serno")
data = data & "&src=" & SumissionSource
data = data & "&usr=" & Session("userNo")
data = data & "&pm=" & Server.URLencode("Preventative Maintenance")
If GetPOSTParams("NeverPutOnHold") = 0  Then data = data & "&hld=1" Else data = data & "&hld=0" 'I know it's wierd. It's the opposite of how it is stored in the table

For x = 1 to 2

		DO_Post = 0
		
		If x = 1 Then Do_Post = GetPOSTParams("ServiceMemoURL1ONOFF") 
		If x = 2 Then Do_Post = GetPOSTParams("ServiceMemoURL2ONOFF") 
		
		If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0
	
		If cint(Do_Post) = 1 Then
		
				CreateSystemAuditLogEntry sURL,"Post Loop "& x,GetPOSTParams("Mode")

				If x = 1 Then
					Description = "Post to " & GetPOSTParams("ServiceMemoURL1")
				Else
					Description = "Post to " & GetPOSTParams("ServiceMemoURL2")
				End If

				CreateSystemAuditLogEntry sURL,Description,GetPOSTParams("Mode")
				Description = "data:" & data 
				CreateSystemAuditLogEntry sURL,Description,GetPOSTParams("Mode")

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
					CreateSystemAuditLogEntry sURL,Description,GetPOSTParams("Mode")
					If Instr(httpRequest.responseText,"success") = 0 Then
						Response.Write("POST RESPONSE:------X" & httpRequest.responseText & "<---------------<br>")
						emailBody = httpRequest.responseText
						emailBody = emailBody & "    PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
						emailBody = emailBody & "    POSTED DATA:" & data
						SendMail "mailsender@" & maildomain ,"projects@metroplexdata.com","BAD POST",emailBody,GetTerm("Field Service"),"Post Failure"
					End If
				ELSE
					'In here it must email us if there are problems
					Description = "httpRequest.responseText:" & httpRequest.responseText
					CreateSystemAuditLogEntry sURL,Description,GetPOSTParams("Mode")
					postResponse = httpRequest.responseText
					emailBody = postResponse
					emailBody = emailBody & "    Unable to communicate with Tomcat at" & GetPOSTParams("ServiceMemoURL1")
					emailBody = emailBody & "    POSTED DATA:" & data
					SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com","TOMCAT ERROR",emailBody,GetTerm("Field Service"),"Post Failure"
				END IF
				
		End If
Next
'******************************************
'Now create an entry in tblAssetPMSubmitted
'******************************************
SQL = "INSERT INTO tblAssetPMSubmitted (assetNumber, PMdate,filterOrPM) "
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & AssetNumber & "', "
SQL = SQL & "'" &  PMCalldate & "','P')"	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)

set rs = nothing
cnn8.close
set cnn8 = nothing

Response.Write("Debug:1<br>")
'*******************************************************************************
'*******************************************************************************
'*******************************************************************************
'This is the tricky part, now we need to find that memo & dispatch it to ourself
'*******************************************************************************
'*******************************************************************************
'*******************************************************************************
SQLDispatch = "Select * From FS_ServiceMemos WHERE "
SQLDispatch = SQLDispatch & "AccountNumber = " & AccountNumber & " AND "
SQLDispatch = SQLDispatch & "PMCall = 1 AND "
SQLDispatch = SQLDispatch & "CurrentStatus = 'OPEN' AND "
SQLDispatch = SQLDispatch & "UserNoOfServiceTech = " & Session("UserNo")
Set cnnDispatch = Server.CreateObject("ADODB.Connection")
cnnDispatch.open (Session("ClientCnnString"))
Set rsDispatch = Server.CreateObject("ADODB.Recordset")
Set rsDispatch = cnnDispatch.Execute(SQLDispatch)
If Not rsDispatch.Eof Then ServiceTicketNumber = rsDispatch("MemoNumber") Else ServiceTicketNumber = ""
Set rsDispatch = Nothing
cnnDispatch.close
Set cnnDispatch  = Nothing
Response.Write("Debug:2<br>")
'Failsafe
If ServiceTicketNumber = "" Then Response.Redirect("filterchanges.asp")
	
UserToDispatch = Session("UserNo") ' Self dispatch
CustNum = AccountNumber 

Response.Write("Debug:3<br>")
'Now do all the normal dispatch stuff
SQLDispatch = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
SQLDispatch = SQLDispatch & "UserNoOfServiceTech, SubmissionDateTime, USerNoSubmittingRecord,EmailAddressSentTo,TextNumberSentTo,OriginalDispatchDateTime)"
SQLDispatch = SQLDispatch &  " VALUES (" 
SQLDispatch = SQLDispatch & "'"  & ServiceTicketNumber & "'"
SQLDispatch = SQLDispatch & ",'"  & CustNum & "'"
SQLDispatch = SQLDispatch & ",'Dispatched'"
SQLDispatch = SQLDispatch & ","  & UserToDispatch 
SQLDispatch = SQLDispatch & ",getdate() "
SQLDispatch = SQLDispatch & ","  & Session("UserNo")
SQLDispatch = SQLDispatch & ",'"  & getUserEmailAddress(UserToDispatch) & "'"
SQLDispatch = SQLDispatch & ",'" & getUserCellNumber(UserToDispatch) & "' "
SQLDispatch = SQLDispatch & ", getDate())"	

Set cnnDispatch = Server.CreateObject("ADODB.Connection")
cnnDispatch.open (Session("ClientCnnString"))
Set rsDispatch = Server.CreateObject("ADODB.Recordset")
Set rsDispatch = cnnDispatch.Execute(SQLDispatch)


'Write audit trail for dispatch
'*******************************
Description = GetUserDisplayNameByUserNo(UserToDispatch) & " self dispatched to service ticket number " & ServiceTicketNumber & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
CreateAuditLogEntry "Service Ticket System","Self Dispatch","Minor",0,Description 

'Also set dispatched flag for simple dispatch model
SQLDispatch = "Update FS_ServiceMemos Set Dispatched = -1 Where MemoNumber = '"  & ServiceTicketNumber & "'"
Set rsDispatch = cnnDispatch.Execute(SQLDispatch)


Set rsDispatch = Nothing
cnnDispatch.Close
Set cnnDispatch = Nothing

'************************************************************
'Send emails to the service managers tell them this happenned
'************************************************************
'If SendEmail="on" then
'	If getUserEmailAddress(UserToDispatch) <> "" Then
'		Send_To = getUserEmailAddress(UserToDispatch)
'		
'		<!--#include file="../../emails/ADVdispatch_dispatch.asp"-->
'			
'		'Failsafe for dev
'		If Instr(ucase(sURL),"DEV") <> 0 Then Send_To = "rich@ocsaccess.com"
'		SendMail "mailsender@" & maildomain ,Send_To,emailSubject,emailBody,GetTerm("Field Service"),"Self Dispatch"
'		Description = "A dispatch email was sent to " & GetUserDisplayNameByUserNo(Session("UserNo")) & " (" & Send_To & ") at " & NOW()
'		CreateAuditLogEntry "Service Ticket System","Dispatch email sent","Minor",0,Description
'	Else
'		' Could not send dispatch email, no address on file
'		emailBody = "Insight was unable to send a dispatch email to " & GetUserDisplayNameByUserNo(UserToDispatch) & ". No email address on file"
'		If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com" else SEND_TO = "rich@ocsaccess.com"
'		SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send dispatch email",emailBody,GetTerm("Field Service"),"Missing Email"
'		Description = "Insight was unable to send a dispatch email to " & GetUserDisplayNameByUserNo(UserToDispatch) & ". No email address on file"
'		CreateAuditLogEntry "Service Ticket System","Unable to send dispatch email","Major",0,Description
'	End If
'End If

'*********************************************************
'*********************************************************
'*********************************************************
'Now another part, mark the dispatch as being acknowledged
'*********************************************************
'*********************************************************
'*********************************************************
Set cnnDispatch = Server.CreateObject("ADODB.Connection")
cnnDispatch.open (Session("ClientCnnString"))
Set rsDispatch = Server.CreateObject("ADODB.Recordset")
Response.Write("Debug:5<br>")
SQLDispatch = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
SQLDispatch = SQLDispatch & " SubmissionDateTime, USerNoSubmittingRecord,Remarks,UserNoOfServiceTech)"
SQLDispatch = SQLDispatch &  " VALUES (" 
SQLDispatch = SQLDispatch & "'"  & ServiceTicketNumber & "'"
SQLDispatch = SQLDispatch & ",'"  & CustNum & "'"
SQLDispatch = SQLDispatch & ",'Dispatch Acknowledged'"
SQLDispatch = SQLDispatch & ",getdate() "
SQLDispatch = SQLDispatch & ","  & Session("UserNo") & ", "
SQLDispatch = SQLDispatch & "'Automatically acknowled via self dispatch', "
SQLDispatch = SQLDispatch & Session("UserNo") & ")"

Set cnnDispatch = Server.CreateObject("ADODB.Connection")
cnnDispatch.open (Session("ClientCnnString"))
Set rsDispatch = Server.CreateObject("ADODB.Recordset")
Set rsDispatch = cnnDispatch.Execute(SQLDispatch)
'Write audit trail for dispatch
'*******************************
Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " automatically acknowldged dispatch for service ticket number " & ServiceTicketNumber & " due to self dispatch at " & NOW()
CreateAuditLogEntry "Service Ticket System","Automatic Dispatch Acknowledgement","Minor",0,Description 
Response.Write("Debug:7<br>")
Set rsDispatch = Nothing
cnnDispatch.Close
Set cnnDispatch = Nothing

Response.Write("This PM Call has been assigned to you. It now appears in your list of service tickets.")
%>