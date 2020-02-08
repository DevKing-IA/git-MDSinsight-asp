<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/mailDirectLaunch.asp"-->

<script type="text/javascript">
    function closeme() {
window.open('', '_parent', '');
window.close();  }
</script>
<%
Server.ScriptTimeout = 5000
'System Alert processing page
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/alerts/SystemAlertsCheck.asp?runlevel=run_now

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)

If Request.QueryString("runlevel") <> "run_now" then response.end

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
If left(maildomain,4) = "www." Then maildomain = right(maildomain,Len(maildomain)-4)


'This single page loops through and handles alerts for ALL clients
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString

	
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and exit
If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF
	
		PassPhrase = TopRecordset.Fields("directLaunchPassphrase")
		ClientKey = TopRecordset.Fields("clientkey")

		Response.Write("******** Processing " & ClientKey  & "************<br>")
		
		
		Call SetClientCnnString
		
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
		
		Response.Write("******** Setting Stopmail Vars for this client: " & ClientKey  & "************<br>")
		
		If Session("ClientCnnString") <> ""Then
			'SEE IF MAIL IS ON OR OFF
			SQLtoggle = "Select STOPALLEMAIL from " & MUV_Read("SQL_Owner") & ".Settings_Global"
			
			Response.write("ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ     " & SQLtoggle & "     ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ")
			Set cnntoggle = Server.CreateObject("ADODB.Connection")
			cnntoggle.open (Session("ClientCnnString"))
			Set rstoggle = Server.CreateObject("ADODB.Recordset")
			rstoggle.CursorLocation = 3 
			Set rstoggle = cnntoggle.Execute(SQLtoggle)
			If rstoggle.Eof Then 
				Session("MAILOFF") = 1 ' If eof then set email to off
				Response.Write("<br><font color='red'>MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-</font>")
			Else
				Session("MAILOFF") = rstoggle("STOPALLEMAIL")
				If Session("MAILOFF") = 1 Then
					Response.Write("<br><font color='red'>MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-</font>")				
				Else
				Response.Write("<br><font color='green'>MAIL ON-MAIL ON-MAIL ON-MAIL ON-MAIL ON-MAIL ON-</font>")				
				End IF
			End If
			set rstoggle = Nothing
			cnntoggle.close
			set cnntoggle = Nothing
		Else
			Session("MAILOFF") = 0 ' There was no valid ccn string, so assume it is on
		End If

		
		
		
		If MUV_READ("cnnStatus") = "OK" Then ' else it loops
		
			'Each alert is handled individually since we never know what's going to be added
			
			'****************************************
			'Begin System Alerts
			'****************************************
			Response.Write("Begin System Alerts<br>")
			
			'**********************************************
			'Delete alerts sent more than 30 days
			'ONLY delete system alerts, others will do their own
			'Don't really need to do this, but will delete stuff older
			'Than 30 days just to keep the file from getting
			'too big
			Set cnnAlertsSent = Server.CreateObject("ADODB.Connection")
			cnnAlertsSent.open (MUV_Read("ClientCnnString"))
			Set rsSCAlertsSent = Server.CreateObject("ADODB.Recordset")
			rsSCAlertsSent.CursorLocation = 3 
			
			rsSCAlertsSent = "SELECT * FROM SC_AlertsSent WHERE DateTimeSent < dateadd(d,-30,getdate())"
			Set rsSCAlertsSent = cnnAlertsSent.Execute(rsSCAlertsSent)
			on error goto 0
			
			If Not rsSCAlertsSent.EOF Then
			
				Set rsForDeletes = Server.CreateObject("ADODB.Recordset")
				rsForDeletes.CursorLocation = 3 

				Do While Not rsSCAlertsSent.EOF
					If GetAlertType(rsSCAlertsSent("Alert_InternalAlertRecNumber")) = "System" Then
						'OK, it is an old system alert so we can delete it					
						SQLForDeletes = "DELETE FROM SC_AlertsSent WHERE UniqueRecordIdentifier = " & rsSCAlertsSent("UniqueRecordIdentifier")
						Set rsForDeletes = cnnAlertsSent.Execute(SQLForDeletes)
					End If
					rsSCAlertsSent.MoveNext
				Loop
				
				Set rsForDeletes = Nothing
			End If
			
			Set rsSCAlertsSent = Nothing
			cnnAlertsSent.Close
			Set cnnAlertsSent = Nothing
			'**********************************************
		
			Set cnnAlerts = Server.CreateObject("ADODB.Connection")
			cnnAlerts.open (MUV_Read("ClientCnnString"))
			
			SQLAlerts = "SELECT * FROM SC_Alerts Where AlertType='System' And Enabled = 1" 
				
			Set rsAlert = Server.CreateObject("ADODB.Recordset")
			rsAlert.CursorLocation = 3 
			Set rsAlert = cnnAlerts.Execute(SQLAlerts)
			
			If not rsAlert.EOF Then
						
				Do While Not rsAlert.EOF
				
					Response.Write("<b>Found Alert Named: " & rsAlert("AlertName") &"</b> Alert Record #:" & rsAlert("InternalAlertRecNumber") &"<br>") 
			
					'Now the real work begins
					
					'Broken into include files because there is so
					'much code (over 13 condiitons%>	
					<!--#include file="include_System_import_related.asp"-->
					<!--#include file="include_System_deliveryboard_related.asp"-->					
					<%


				rsAlert.Movenext
			Loop

		End If
			
		Set rsAlert = Nothing
		cnnAlerts.Close
		Set cnnALerts = Nothing
						
			 				
	
		Response.Write("******** DONE Processing " & ClientKey  & "************<br>")
	End If				
	TopRecordset.movenext
	
	Loop
	
	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type=""text/javascript"">closeme();</script>")	

'*************************
'*************************
'Subs and funcs begin here

Sub SendAlert (passedAlertNumber,PassedDateTimeValueIfApplicable,AlertType,passedNotificationType)

	Send_To = ""
	'*************************
	'First do the email alerts
	'*************************
	'Get user based emails
	If Not IsNull(rsAlert("EmailToUserNos")) Then
		If rsAlert("EmailToUserNos") <> "" And rsAlert("EmailToUserNos") <> "0" Then
			UserNoList = Split(rsAlert("EmailToUserNos"),",")
			For x = 0 To UBound(UserNoList)
				Send_To = Send_To & GetUserEmailByUserNo(UserNoList(x)) & ";"
			Next
		End If
	End If
										
	'Get additional emails if there are any
	If rsAlert("AdditionalEmails") <> "" and not IsNull(rsAlert("AdditionalEmails")) Then
		tmpSendAlertToAdditionalEmails = trim(rsAlert("AdditionalEmails"))		
		If Len(tmpSendAlertToAdditionalEmails) > 1 Then
			If Right(tmpSendAlertToAdditionalEmails,1) <> ";" Then tmpSendAlertToAdditionalEmails = tmpSendAlertToAdditionalEmails & ";"
			Send_To = Send_To & tmpSendAlertToAdditionalEmails
		End If
	End If
	
	'Only do this if there are actually emails to send
	If Send_To <> "" Then
	
		If rsAlert("EmailVerbiage") <> "" Then
			If Not IsNull(rsAlert("EmailVerbiage")) Then
				AdditionalEmailVerbiage = rsAlert("EmailVerbiage")
			End If
		End If
		
		Select Case AlertType
			Case "BackendStarted"
				emailSubject = "Backend data import started "
				emailHeadLineText = "Backend data import started at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Backend data import started at " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "BackendFinished"
				emailSubject = "Backend data import finished "
				emailHeadLineText = "Backend data import finished at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Backend data import finished at " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "BackendNoStart"
				emailSubject = "Backend data import did not start "
				emailHeadLineText = "Backend data import did not start by " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Backend data import did not start by " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "BackendRunTooLong"
				emailSubject = "Backend data import running too long "
				emailHeadLineText = "Backend data import running too long: " & MUV_READ("RunningMinutes") & " minutes so far."
				txtMessage = "Backend data import running too long: " & MUV_READ("RunningMinutes") & " minutes so far."
			Case "RebuildStarted"
				emailSubject = "Daily data rebuild started "
				emailHeadLineText = "Daily data rebuild started at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Daily data rebuild started at " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "RebuildFinished"
				emailSubject = "Daily data rebuild finished "
				emailHeadLineText = "Daily data rebuild finished at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Daily data rebuild finished at " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "RebuildNotRun"
				emailSubject = "Daily data rebuild did not start "
				emailHeadLineText = "Daily data rebuild did not start by " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Daily data rebuild did not start by " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "RebuildRunTooLong"
				emailSubject = "Daily data rebuild running too long "
				emailHeadLineText = "Daily data rebuild running too long: " & MUV_READ("RunningMinutes") & " minutes so far."
				txtMessage = "Daily data rebuild running too long: " & MUV_READ("RunningMinutes") & " minutes so far."
			Case "DBoardFinished"
				emailSubject = "Nightly delivery board update finished "
				emailHeadLineText = "Nightly delivery board update finished at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Nightly delivery board update finished at " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "DBoardNotRun"
				emailSubject = "Nightly delivery board update did not start "
				emailHeadLineText = "Nightly delivery board update did not start by " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Nightly delivery board update did not start by " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "DBoardSkipped"
				emailSubject = "Nightly delivery board update SKIPPED "
				emailHeadLineText = "Nightly delivery board update SKIPPED at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Nightly delivery board update SKIPPED at " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "DBoardOnDemandRun"
				emailSubject = "On demand delivery board update was run "
				emailHeadLineText = "On demand delivery board update was run at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "On demand delivery board update was run at " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "AutoCompJSONNotRun"		
				emailSubject = "Autocomplete JSON File Problem"
				emailHeadLineText = "Autocomplete JSON file(s) missing or out-of-date at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Autocomplete JSON file(s) missing or out-of-date at " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "HistOldInvoice"
				emailSubject = "Invoice history out of date"
				emailHeadLineText = "Most recent invoice in history is more than " & ABS(DateDiff("d",MostRecentInvoiceDate,Now())) & " days old at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Invoice history out of date at " & FormatDateTime(PassedDateTimeValueIfApplicable)
			Case "RouteFileEmpty"
				emailSubject = "Route file empty"
				emailHeadLineText = "The route table is empty at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "The route table is empty at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
			Case "ProspectNoNextActivity"
				emailSubject = "Prospect with no next activity"
				emailHeadLineText = "Prospect number " & rsProspecting("InternalRecordIdentifier") & " has no next activity at " & FormatDateTime(PassedDateTimeValueIfApplicable) 
				txtMessage = "Prospect with no next activity at " & FormatDateTime(PassedDateTimeValueIfApplicable) 

		End Select
		
		%>
		<!--#include file="../../emails/system_realtime_alert.asp"-->	
		<%
		
		'Now Send the emails
		'Got all the addresses so now break them up
		Send_To_Array = Split(Send_To,";")
	
		For x = 0 to Ubound(Send_To_Array) -1
			Send_To = Send_To_Array(x)
			
			Response.Write("<font color='green'><b>Sending alerts to Send_To: " & Send_To & "</font><br>") 
			
			If passedNotificationType = "Alert" Then
				SendFrom = "MDS Insight Alerts"
			Elseif passedNotificationType = "Notification" Then
				SendFrom = "MDS Insight Notifications"
			Else
				SendFrom = "MDS Insight"
			End If
			SendFrom = SendFrom & " (" & MUV_READ("ClientID") & ")"
			
			SendMail "mailsender@" & maildomain,Send_To, emailSubject,emailBody,"Alerts","System Alert",SendFrom
			
			CreateAuditLogEntry "System Alert Sent","System Alert Sent","Minor",0,"System Realtime Alert Sent to " & Send_To & " for invoice #: " & PassedInvoiceNumber & " - " & emailSubject 
			
		Next 
	
	End If
	
	Send_To=""
	'**********************
	'Now do the text alerts
	'**********************
	'Get user based Texts
	If Not IsNull(rsAlert("TextToUserNos")) Then
		If rsAlert("TextToUserNos") <> "" And rsAlert("TextToUserNos") <> "0" Then
			UserNoList = Split(rsAlert("TextToUserNos"),",")
			For x = 0 To UBound(UserNoList)
				Send_To = Send_To & getUserCellNumber(UserNoList(x)) & ","
			Next
		End If
	End If
	
	'Get additional texts if there are any
	If rsAlert("AdditionalText") <> "" and not IsNull(rsAlert("AdditionalText")) Then
		tmpSendAlertToAdditionalTexts = trim(rsAlert("AdditionalText"))		
		If Len(tmpSendAlertToAdditionalTexts) > 1 Then
			If Right(tmpSendAlertToAdditionalTexts,1) <> "," Then tmpSendAlertToAdditionalTexts = tmpSendAlertToAdditionalTexts & ","
			 Send_To = Send_To & tmpSendAlertToAdditionalTexts
		End If
	End If
	
	'Only do this if there are actually texts to send
	If Send_To <> "" Then
								
		Send_To = Replace(Send_To,"-","") ' EZ Texting doesn't like dashes
	
		If rsAlert("TextVerbiage") <> "" Then
			If Not IsNull(rsAlert("TextVerbiage")) Then
				AdditionalTextVerbiage = rsAlert("TextVerbiage")
			End If
		End If
	

		txtSubject = passedNotificationType 
		
		txtMessage = txtMessage & AdditionalTextVerbiage & " "
		txtMessage = txtMessage & "(Alert name: " & rsAlert("AlertName")& ")"
	
		
		'*****Text numbers don't get split into an array, the php takes multiple #'s seprated by commas	
			
		If Right(Send_To,1) = "," Then Send_To = Left(Send_To,Len(Send_To)-1)
	
		TEXT_TO = Send_To
		
		'Split Text_To for recording the alerts sent
		TextNumberArray = Split(TEXT_TO&",",",")
		
		For x = 0 to ubound(TextNumberArray) -1
			Response.Write("TextNumberArray (x):" & TextNumberArray (x) & "<br>")
		Next
	
		Response.Write("<font color='green'><b>Sending text to: "& TEXT_TO & "</b></font><br>")
		
		CreateAuditLogEntry "System Alert Sent","System Alert Sent","Minor",0,"System Realtime Alert Sent to " & TEXT_TO & " for invoice #: " & PassedInvoiceNumber & " - " & emailSubject 	
		
		Response.Write("POST TO: " & BaseURL & "inc/sendtext_post.php<br>")
	
		str_data="txtSubject=" & txtSubject  & "&txtMessage=" & txtMessage & "&txtTEXT_TO=" & TEXT_TO
	
		str_data = str_data & "&txtu1=" & EzTextingUserID() & "&txtu2=" & EzTextingPassword()
		
		str_data = str_data & "&txtCountry=" & GetCompanyCountry()
		
		Response.Write("str_data= " & str_data & "<br>")
	
		Set obj_post=Server.CreateObject("Msxml2.SERVERXMLHTTP")
		obj_post.Open "POST", BaseURL & "inc/sendtext_post.php",False
		obj_post.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		obj_post.Send str_data
		
	
		Response.Write("obj_post.responseText: " & obj_post.responseText & "<br>")
	
	End If

	If AlertType <> "ProspectNoNextActivity" Then
		Call WriteAlertSentRecord (rsAlert("InternalAlertRecNumber"),PassedDateTimeValueIfApplicable)
	Else
		Call WriteAlertSentRecordProspecting (rsAlert("InternalAlertRecNumber"),PassedDateTimeValueIfApplicable,rsProspecting("InternalRecordIdentifier"))
	End IF

End Sub


Sub SetClientCnnString

	dummy=MUV_WRITE("cnnStatus","")

	SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open InsightCnnString
	
	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3

	
	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection
	'and exit
	If Recordset.recordcount <= 0 then
		Recordset.close
		Connection.close
		set Recordset=nothing
		set Connection=nothing
	Else
		ClientCnnString = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
		ClientCnnString = ClientCnnString & ";Database=" & Recordset.Fields("dbCatalog")
		ClientCnnString = ClientCnnString & ";Uid=" & Recordset.Fields("dbLogin")
		ClientCnnString = ClientCnnString & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
		dummy = MUV_Write("ClientCnnString",ClientCnnString)
		dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
		Session("SQL_Owner") = Recordset.Fields("dbLogin")
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub


Sub WriteAlertSentRecord (PassedAlertRecIdentifier,PassedDateTimeValueIfApplicable)

		Set cnnAlertsSent = Server.CreateObject("ADODB.Connection")
		cnnAlertsSent.open (MUV_Read("ClientCnnString"))
		Set rsSCAlertsSent = Server.CreateObject("ADODB.Recordset")
		rsSCAlertsSent.CursorLocation = 3 

		rsSCAlertsSent = "Insert Into SC_AlertsSent (Alert_InternalAlertRecNumber,DateTimeValueIfApplicable) Values (" &  PassedAlertRecIdentifier & ",'" & PassedDateTimeValueIfApplicable & "')"

		Set rsSCAlertsSent = cnnAlertsSent.Execute(rsSCAlertsSent)

		Set rsSCAlertsSent = Nothing
		cnnAlertsSent.Close
		Set cnnAlertsSent = Nothing

End Sub

Function AlertSent(passedInternalAlertRecNumber,passedDateTimeString)

		resultAlertSent = False
		
		Set cnnAlertSent  = Server.CreateObject("ADODB.Connection")
		cnnAlertSent.open (MUV_Read("ClientCnnString"))
		Set rsAlertSent  = Server.CreateObject("ADODB.Recordset")
		rsAlertSent.CursorLocation = 3 

		SQLAlertSent  = "Select * from SC_AlertsSent  Where Alert_InternalAlertRecNumber = " & passedInternalAlertRecNumber & " AND DateTimeValueIfApplicable = '" & passedDateTimeString& "'"

		Set rsAlertSent = cnnAlertSent.Execute(SQLAlertSent)
		If not rsAlertSent.Eof Then
			resultAlertSent = True
			Response.Write("ALERT SENT TRUE: SC_AlertsSent UniqueRecordIdentifier: " & rsAlertSent("UniqueRecordIdentifier") & "<br>")
		End IF

		Set rsAlertSent = Nothing
		cnnAlertSent.Close
		Set cnnAlertSent  = Nothing
		
		AlertSent = resultAlertSent

End Function 

Function AlertSentProspecting(passedInternalAlertRecNumber,passedDateTimeString,passedProspectIntRecID)

		resultAlertSentProspecting = False
		
		Set cnnAlertSentProspecting  = Server.CreateObject("ADODB.Connection")
		cnnAlertSentProspecting.open (MUV_Read("ClientCnnString"))
		Set rsAlertSentProspecting  = Server.CreateObject("ADODB.Recordset")
		rsAlertSentProspecting.CursorLocation = 3 

		SQLAlertSentProspecting  = "Select * from SC_AlertsSent WHERE ProspectRecIDIfApplicable = " & passedProspectIntRecID & " AND Alert_InternalAlertRecNumber = " & passedInternalAlertRecNumber & " AND DateTimeValueIfApplicable = '" & passedDateTimeString& "'"

		Set rsAlertSentProspecting = cnnAlertSentProspecting.Execute(SQLAlertSentProspecting)
		If not rsAlertSentProspecting.Eof Then
			resultAlertSentProspecting = True
			Response.Write("ALERT SENT TRUE: SC_AlertsSent UniqueRecordIdentifier: " & rsAlertSentProspecting("UniqueRecordIdentifier") & "<br>")
		End IF

		Set rsAlertSentProspecting = Nothing
		cnnAlertSentProspecting.Close
		Set cnnAlertSentProspecting  = Nothing
		
		AlertSentProspecting = resultAlertSentProspecting

End Function 

Sub WriteAlertSentRecordProspecting (PassedAlertRecIdentifier,PassedDateTimeValueIfApplicable,passedProspectIntRecID)

		Set cnnAlertsSent = Server.CreateObject("ADODB.Connection")
		cnnAlertsSent.open (MUV_Read("ClientCnnString"))
		Set rsSCAlertsSent = Server.CreateObject("ADODB.Recordset")
		rsSCAlertsSent.CursorLocation = 3 

		rsSCAlertsSent = "Insert Into SC_AlertsSent (Alert_InternalAlertRecNumber,DateTimeValueIfApplicable,ProspectRecIDIfApplicable) Values (" &  PassedAlertRecIdentifier & ",'" & PassedDateTimeValueIfApplicable & "'," & passedProspectIntRecID & ")"

		Set rsSCAlertsSent = cnnAlertsSent.Execute(rsSCAlertsSent)

		Set rsSCAlertsSent = Nothing
		cnnAlertsSent.Close
		Set cnnAlertsSent = Nothing

End Sub


%>