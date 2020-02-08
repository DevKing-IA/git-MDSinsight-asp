<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->

<%

emailReceivedAsArray = false

'****************************************************************************************
'IF WE RECEIVE A RECORD NUMBER FROM THE QUERYSTRING, THEN THE USER HAS REQUESTED TO 
'FORWARD A SINGLE EMAIL FROM THE VIEW FULL EMAIL MODAL WINDOW
'****************************************************************************************

EmailAddressToForwardTo = Request.Form("txtForwardEmailAddresses")
currentEmailCategory1ViewedID = Request.Form("txtCategory1Active")
currentEmailCategory2ViewedIDTab =  Request.Form("txtCategory2Active")
ClientID = Request.Form("txtClientID")

'****************************************************************************************
'PREPARE EMAIL SETTINGS
'****************************************************************************************

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

SQL = "SELECT * FROM tblServerInfo WHERE clientKey='"& ClientID &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")

'Response.Write("InsightCnnString:" & InsightCnnString & "<br>")

Connection.Open InsightCnnString

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
Else
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & Recordset.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & Recordset.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	Recordset.close
	Connection.close	
End If	
  				    		
'Must include down here after the session connection string has been set
%><!--#include file="../../inc/mail.asp"--><%

'******************************************************************************************
'IF WE DO NOT RECEIVE A RECORD NUMBER FROM THE QUERYSTRING, THEN THE USER HAS REQUESTED TO 
'FORWARD A SINGLE EMAIL OR MULTIPLE EMAILS FROM THE VIEW ALL EMAILS OUTBOX
'WE NEED TO SPLIT ON A COMMA TO SEE IF WE HAVE AN ARRAY/MULTIPLE EMAILS TO FORWARD
'******************************************************************************************

InternalRecordNumber = Request.Form("txtInternalRecordNumber")
emailReceivedAsArray = false
InternalRecordNumberArray = Split(InternalRecordNumber,",")

If Ubound(InternalRecordNumberArray) = 0 Then
	emailReceivedAsArray = false
	InternalRecordNumber = InternalRecordNumberArray(0)
Else
	emailReceivedAsArray = true
End If




'*******************************************************************************
'FORWARD SINGLE EMAIL
'*******************************************************************************

'Response.write("<br><br><br><br><br><br><br>emailReceivedAsArray : " & emailReceivedAsArray & "<br>")
'Response.write("InternalRecordNumber : " & InternalRecordNumber & "<br>")
'Response.write("EmailAddressToForwardTo : " & EmailAddressToForwardTo & "<br>")


If emailReceivedAsArray = false AND InternalRecordNumber <> "" AND EmailAddressToForwardTo <> "" Then

	
	SQL8 = "SELECT * FROM SC_EmailLog WHERE InternalRecordNumber = " & InternalRecordNumber 
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs8 = Server.CreateObject("ADODB.Recordset")
	rs8.CursorLocation = 3 
	Set rs8 = cnn8.Execute(SQL8)

	If NOT rs8.EOF AND Session("ClientCnnString") <> "" then
	
		EmailSendFrom = rs8("EmailFrom")
		EmailSendFromName = rs8("EmailFromName")
		Subject = rs8("Subject")
		Body = rs8("Body")
		Attachment = rs8("Attachment")   
		
		'The IF statement below makes sure that when run from DEV it only deos client keys with a d
		'and when run from LIVE it only does client keys without a d
		'Pretty smart, huh
		
		If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientID),"D") = 0)_
		or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientID),"D") <> 0) Then 
		
		   	If EmailAddressToForwardTo <> "" Then
				EmailAddressToForwardToArray = split(EmailAddressToForwardTo,";")
				
				If Ubound(EmailAddressToForwardToArray) = 0 then
					SendToList = SendToList & EmailAddressToForwardTo & ";"
				Else
					For z = 0 To Ubound(EmailAddressToForwardToArray)
						SendToList = SendToList & EmailAddressToForwardToArray(z) & ";"
					Next
				End If
			End If
		
			'Got all the addresses so now break them up
			SendToList_Array = Split(SendToList,";")

			'HERE WE ACTUALLY SEND THE EMAIL
			
			For x = 0 to Ubound(SendToList_Array) -1
				SendToList = SendToList_Array(x)
				
				'Failsafe for dev
				sURL = Request.ServerVariables("SERVER_NAME")
				If Instr(ucase(sURL),"DEV.") <> 0 Then SendToList = "rich@ocsaccess.com"
				
				If EmailSendFrom = "" Then
					EmailSendFrom = "mailsender@mdsinsight.com"
				End If
				
				If Attachment <> "" Then
				   SendMailWatt EmailSendFrom,SendToList,Subject,Body,Attachment,"Admin Email Log","Forwarded Emails"
				Else
					SendMail EmailSendFrom,SendToList,Subject,Body,"Admin Email Log","Forwarded Emails"
				End If	

				Description = "Email with subject, " & Subject & ", was forwarded to " & EmailAddressToForwardTo & " from " & EmailSendFromName & " at " & EmailSendFrom
				CreateAuditLogEntry "Email Forwarded From Admin","Email Forwarded From Admin","Minor",0,Description 
			Next

		End If
		
		set rs = Nothing
		set cnn8  = Nothing
		
		Response.Redirect ("allSentEmails.asp?cat1ID=" & currentEmailCategory1ViewedID & "&tab=" & currentEmailCategory2ViewedIDTab)

	Else
	
		%><div><br><br><br><br><br>
		Unable to send, unique email identifier not found: <%= InternalRecordNumber %>.
		</div>
		<%
		
	End If
	
'*******************************************************************************
'FORWARD MULTIPLE EMAILS AT ONCE
'*******************************************************************************
ElseIf emailReceivedAsArray = true AND emailAddressToForwardTo <> "" AND InternalRecordNumberArray(0)<>""  Then


	For q = 0 to ubound(InternalRecordNumberArray)
		
		SQL8 = "SELECT * FROM SC_EmailLog WHERE InternalRecordNumber = " & InternalRecordNumberArray(q)
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs8 = Server.CreateObject("ADODB.Recordset")
		rs8.CursorLocation = 3 
		Set rs8 = cnn8.Execute(SQL8)
	
		If NOT rs8.EOF AND Session("ClientCnnString") <> "" then
		
			EmailSendFrom = rs8("EmailFrom")
			EmailSendFromName = rs8("EmailFromName")
			Subject = rs8("Subject")
			Body = rs8("Body")
			Attachment = rs8("Attachment")   
			
			'The IF statement below makes sure that when run from DEV it only deos client keys with a d
			'and when run from LIVE it only does client keys without a d
			'Pretty smart, huh
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientID),"D") = 0)_
			or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientID),"D") <> 0) Then 
	    		
			   	If EmailAddressToForwardTo <> "" Then
					EmailAddressToForwardToArray = split(EmailAddressToForwardTo,";")
					
					If Ubound(EmailAddressToForwardToArray) = 0 then
						SendToList = SendToList & EmailAddressToForwardTo & ";"
					Else
						For z = 0 To Ubound(EmailAddressToForwardToArray)
							SendToList = SendToList & EmailAddressToForwardToArray(z) & ";"
						Next
					End If
				End If
		
				'Got all the addresses so now break them up
				SendToList_Array = Split(SendToList,";")

				'HERE WE ACTUALLY SEND THE EMAIL
				
				For x = 0 to Ubound(SendToList_Array) -1
					SendToList = SendToList_Array(x)
					
					'Failsafe for dev
					sURL = Request.ServerVariables("SERVER_NAME")
					If Instr(ucase(sURL),"DEV.") <> 0 Then SendToList = "rich@ocsaccess.com"
					
					If EmailSendFrom = "" Then
						EmailSendFrom = "mailsender@mdsinsight.com"
					End If
					
					If Attachment <> "" Then
					   SendMailWatt EmailSendFrom,SendToList,Subject,Body,Attachment,"Admin Email Log","Forwarded Emails"
					Else
						SendMail EmailSendFrom,SendToList,Subject,Body,"Admin Email Log","Forwarded Emails"
					End If	

					Description = "Email with subject, " & Subject & ", was forwarded to " & EmailAddressToForwardTo & " from " & EmailSendFromName & " at " & EmailSendFrom
					CreateAuditLogEntry "Email Forwarded From Admin","Email Forwarded From Admin","Minor",0,Description 
				Next
				
			End If
		End If
	Next


	set rs = Nothing
	set cnn8  = Nothing
		
	Response.Redirect ("allSentEmails.asp?cat1ID=" & currentEmailCategory1ViewedID & "&tab=" & currentEmailCategory2ViewedIDTab)
	
Else

	%><div><br><br><br><br><br>
	Unable to send, could not parse querystring for unqiue email identifier.
	</div>
	<%
	
End If

%><!--#include file="../../inc/footer-main.asp"-->