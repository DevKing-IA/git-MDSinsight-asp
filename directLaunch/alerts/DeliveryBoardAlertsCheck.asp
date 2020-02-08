<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/mailDirectLaunch.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Routing.asp"-->
<script type="text/javascript">
    function closeme() {
window.open('', '_parent', '');
window.close();  }
</script>
<%
'Delivery Board Alert processing page
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/alerts/DeliveryBoardAlertCheck.asp?runlevel=run_now
Server.ScriptTimeout = 2500

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)

If Request.QueryString("runlevel") <> "run_now" then response.end

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

'This single page loops through and handles alerts for ALL clients
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1"
If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then ' Good 'ol dev failsafe
	'SQL = "SELECT * FROM tblServerInfo where dbLogin like '%dev'"
End If

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
			'Begin Delivery Board Alerts
			'****************************************
			Response.Write("Begin Delivery Board Alerts<br>")
			
'**********************************************
			'Delete alerts sent more than 30 days
			'ONLY delete delivery boards, others will do their own
			'Don't really need to dothis, but will delete stuff older
			'Than 30 days just to keep the file from getting
			'too big
			Set cnnAlertsSent = Server.CreateObject("ADODB.Connection")
			cnnAlertsSent.open (MUV_Read("ClientCnnString"))
			Set rsSCAlertsSent = Server.CreateObject("ADODB.Recordset")
			rsSCAlertsSent.CursorLocation = 3 
			'''''on error resume next ' In case they dont have this table
			
			rsSCAlertsSent = "SELECT * FROM SC_AlertsSent WHERE DateTimeSent < dateadd(d,-30,getdate())"
			Set rsSCAlertsSent = cnnAlertsSent.Execute(rsSCAlertsSent)
			on error goto 0
			
			If Not rsSCAlertsSent.EOF Then
			
				Set rsForDeletes = Server.CreateObject("ADODB.Recordset")
				rsForDeletes.CursorLocation = 3 

				Do While Not rsSCAlertsSent.EOF
					If GetAlertType(rsSCAlertsSent("Alert_InternalAlertRecNumber")) = "DeliveryBoardAlert" Then
						'OK, it is an old delivery board alert so we can delete it					
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
			
			SQLAlerts = "SELECT * FROM SC_Alerts Where AlertType='DeliveryBoardAlert' And Enabled = 1" 
				
			Set rsAlert = Server.CreateObject("ADODB.Recordset")
			rsAlert.CursorLocation = 3 
			Set rsAlert = cnnAlerts.Execute(SQLAlerts)
			
			If not rsAlert.EOF Then
						
				Do While Not rsAlert.EOF
				
					Response.Write("<b>Found Alert Named: " & rsAlert("AlertName") &"</b><br>") 
			
						
						'Now the real work begins
						'Run through RT_DeliveryBoard   &  see if this alert needs to be sent
						Select Case rsAlert("Condition")
							Case "AM_Overdue"
						
								Response.Write("Check AM_Overdue <br>")
								
								Set cnnDeliveryBoard = Server.CreateObject("ADODB.Connection")
								cnnDeliveryBoard.open (MUV_Read("ClientCnnString"))
								Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
								rsDeliveryBoard.CursorLocation = 3 
	
								SQL_DeliveryBoard = "SELECT * FROM RT_DeliveryBoard WHERE DeliveryStatus IS NULL AND DeliveryInProgress=0 AND AMorPM = 'AM'"
								
								Set rsDeliveryBoard = cnnDeliveryBoard.Execute(SQL_DeliveryBoard)
								
								If Not rsDeliveryBoard.EOF Then
									Do While Not rsDeliveryBoard.Eof
										
										CurrentHour = Hour(Now())
										CurrentMinute = Minute(Now())
										
										AlertHour = cint(Left(rsAlert("TimeOfDay"),Len(rsAlert("TimeOfDay"))-2))
										AlertMinute = cint(Right(rsAlert("TimeOfDay"),2))
										
										Response.Write("CurrentHour :" & CurrentHour & "<br>")
										Response.Write("CurrentMinute :" & CurrentMinute & "<br>")
										Response.Write("AlertHour :" & AlertHour & "<br>")
										Response.Write("AlertMinute :" & AlertMinute & "<br>")
										
										'OK found unmarked AM tickets, now check the time
										If CurrentHour > AlertHour Then 'OK, hour is great, so send it
										
											If AlertSent(rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum"),rsAlert("Condition")
											
										Elseif CurrentHour = AlertHour Then ' same hour, check minutes
											
											If CurrentMinute > AlertMinute Then ' yes, send it
												
												If AlertSent(rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum"),rsAlert("Condition")
												
											End If
											
										End If
									
										rsDeliveryBoard.MoveNext
									Loop
								End If
								
								Set rsDeliveryBoard = Nothing
								cnnDeliveryBoard.Close
								Set cnnDeliveryBoard = Nothing
								
							Case "Priority_Overdue"
						
								Response.Write("Check Priority_Overdue <br>")
								
								Set cnnDeliveryBoard = Server.CreateObject("ADODB.Connection")
								cnnDeliveryBoard.open (MUV_Read("ClientCnnString"))
								Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
								rsDeliveryBoard.CursorLocation = 3 
	
								SQL_DeliveryBoard = "SELECT * FROM RT_DeliveryBoard WHERE DeliveryStatus IS NULL AND DeliveryInProgress=0 AND Priority = 1"
								
								Set rsDeliveryBoard = cnnDeliveryBoard.Execute(SQL_DeliveryBoard)
								
								If Not rsDeliveryBoard.EOF Then
									Do While Not rsDeliveryBoard.Eof
										
										CurrentHour = Hour(Now())
										CurrentMinute = Minute(Now())
										
										AlertHour = cint(Left(rsAlert("TimeOfDay"),Len(rsAlert("TimeOfDay"))-2))
										AlertMinute = cint(Right(rsAlert("TimeOfDay"),2))
										
										Response.Write("CurrentHour :" & CurrentHour & "<br>")
										Response.Write("CurrentMinute :" & CurrentMinute & "<br>")
										Response.Write("AlertHour :" & AlertHour & "<br>")
										Response.Write("AlertMinute :" & AlertMinute & "<br>")
										
										'OK found unmarked Priority tickets, now check the time
										If CurrentHour > AlertHour Then 'OK, hour is great, so send it
										
											If AlertSent(rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum"),rsAlert("Condition")
											
										Elseif CurrentHour = AlertHour Then ' same hour, check minutes
											
											If CurrentMinute > AlertMinute Then ' yes, send it
												
												If AlertSent(rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum"),rsAlert("Condition")
												
											End If
											
										End If
									
										rsDeliveryBoard.MoveNext
									Loop
								End If
								
								Set rsDeliveryBoard = Nothing
								cnnDeliveryBoard.Close
								Set cnnDeliveryBoard = Nothing
								
								
							Case "Priority No Delivery"

								Response.Write("Check Priority No Delivery <br>")
								
								Set cnnDeliveryBoard = Server.CreateObject("ADODB.Connection")
								cnnDeliveryBoard.open (MUV_Read("ClientCnnString"))
								Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
								rsDeliveryBoard.CursorLocation = 3 
	
								SQL_DeliveryBoard = "SELECT * FROM RT_DeliveryBoard WHERE DeliveryStatus = 'No Delivery' AND Priority = 1"
								
								Set rsDeliveryBoard = cnnDeliveryBoard.Execute(SQL_DeliveryBoard)
								
								If Not rsDeliveryBoard.EOF Then
									Do While Not rsDeliveryBoard.Eof
									
										If PartialDelivery(rsDeliveryBoard("IvsNum")) <> True Then
										
											If AlertSent(rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum"),rsAlert("Condition")
										
										End If
									
										rsDeliveryBoard.MoveNext
									Loop
								End If
								
								Set rsDeliveryBoard = Nothing
								cnnDeliveryBoard.Close
								Set cnnDeliveryBoard = Nothing
								
							Case "No Delivery"

								Response.Write("Check No Delivery <br>")
								
								Set cnnDeliveryBoard = Server.CreateObject("ADODB.Connection")
								cnnDeliveryBoard.open (MUV_Read("ClientCnnString"))
								Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
								rsDeliveryBoard.CursorLocation = 3 
	
								SQL_DeliveryBoard = "SELECT * FROM RT_DeliveryBoard WHERE DeliveryStatus = 'No Delivery'"
								
								Set rsDeliveryBoard = cnnDeliveryBoard.Execute(SQL_DeliveryBoard)
								
								If Not rsDeliveryBoard.EOF Then
									Do While Not rsDeliveryBoard.Eof
									
										If PartialDelivery(rsDeliveryBoard("IvsNum")) <> True Then
										
											If AlertSent(rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum"),rsAlert("Condition")
										
										End If
									
										rsDeliveryBoard.MoveNext
									Loop
								End If
								
								Set rsDeliveryBoard = Nothing
								cnnDeliveryBoard.Close
								Set cnnDeliveryBoard = Nothing
							
							Case "Delivered"
							
								Response.Write("Check Delivered <br>")
								
								Set cnnDeliveryBoard = Server.CreateObject("ADODB.Connection")
								cnnDeliveryBoard.open (MUV_Read("ClientCnnString"))
								Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
								rsDeliveryBoard.CursorLocation = 3 
	
								SQL_DeliveryBoard = "SELECT * FROM RT_DeliveryBoard WHERE DeliveryStatus = 'Delivered'"
								
								Set rsDeliveryBoard = cnnDeliveryBoard.Execute(SQL_DeliveryBoard)
								
								If Not rsDeliveryBoard.EOF Then
									Do While Not rsDeliveryBoard.Eof
										
										If PartialDelivery(rsDeliveryBoard("IvsNum")) <> True Then
										
											If AlertSent(rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum"),rsAlert("Condition")
											
										End If
									
										rsDeliveryBoard.MoveNext
									Loop
								End If
								
								Set rsDeliveryBoard = Nothing
								cnnDeliveryBoard.Close
								Set cnnDeliveryBoard = Nothing
							
							Case "Partial"
							
								Response.Write("Check Partial Deliveries <br>")
								
								Set cnnDeliveryBoard = Server.CreateObject("ADODB.Connection")
								cnnDeliveryBoard.open (MUV_Read("ClientCnnString"))
								Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
								rsDeliveryBoard.CursorLocation = 3 
	
								SQL_DeliveryBoard = "SELECT * FROM RT_DeliveryBoard WHERE DeliveryStatus IS NOT NULL"
								
								Set rsDeliveryBoard = cnnDeliveryBoard.Execute(SQL_DeliveryBoard)
								
								If Not rsDeliveryBoard.EOF Then
									Do While Not rsDeliveryBoard.Eof
										
										If PartialDelivery(rsDeliveryBoard("IvsNum")) = True Then
										
											If AlertSent(rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsDeliveryBoard("IvsNum"),rsAlert("Condition")
											
										End If
									
										rsDeliveryBoard.MoveNext
									Loop
								End If
								
								Set rsDeliveryBoard = Nothing
								cnnDeliveryBoard.Close
								Set cnnDeliveryBoard = Nothing
							
							End Select

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

Sub SendAlert (passedAlertNumber,PassedInvoiceNumber,AlertType)


	Response.write("**********GOT TO SEND ALERT()**************************<br>")
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
	
		Response.write("**********FOUND EMAIL RECIPIENTS TO SEND TO**************************<br>")

	
		If rsAlert("EmailVerbiage") <> "" Then
			If Not IsNull(rsAlert("EmailVerbiage")) Then
				AdditionalEmailVerbiage = rsAlert("EmailVerbiage")
			End If
		End If
			
		
		Select Case AlertType
			Case "AM_Overdue"
				emailSubject = "Invoice " & PassedInvoiceNumber & " AM delivery overdue"
				emailHeadLineText = "Invoice # " & PassedInvoiceNumber & " is flagged for <u>AM Delivery</u> and has not been marked as <i>delivered</i> or <i>no delivery</i> as of " & FormatDateTime(Now(),3) & " today."
				txtMessage = "Invoice # " & PassedInvoiceNumber & " is flagged for AM Delivery and has not been marked as delivered or no delivery."
			Case "Priority_Overdue"
				emailSubject = "Invoice " & PassedInvoiceNumber & " Priority delivery overdue"
				emailHeadLineText = "Invoice # " & PassedInvoiceNumber & " is flagged as a <u>Priority Delivery</u> and has not been marked as <i>delivered</i> or <i>no delivery</i> as of " & FormatDateTime(Now(),3) & " today."
				txtMessage = "Invoice # " & PassedInvoiceNumber & " is flagged as a Priority Delivery and has not been marked as delivered or no delivery."
			Case "Priority No Delivery"
				emailSubject = "Invoice " & PassedInvoiceNumber & " not delivered"
				emailHeadLineText = "Invoice # " & PassedInvoiceNumber & " is flagged as a <u>Priority Delivery</u> and was updated to <i>not delivered</i> at " & FormatDateTime(GetLastDeliveryStatusChangeBYInvoiceNumDelBoard(PassedInvoiceNumber),3) & " today."
				txtMessage = "Invoice # " & PassedInvoiceNumber & " is flagged as a Priority Delivery and was updated to not delivered at " & FormatDateTime(Now(),3) & " today."
			Case "No Delivery"
				emailSubject = "Invoice " & PassedInvoiceNumber & " not delivered"
				emailHeadLineText = "Invoice # " & PassedInvoiceNumber & " was updated to <i>not delivered</i> at " & FormatDateTime(GetLastDeliveryStatusChangeBYInvoiceNumDelBoard(PassedInvoiceNumber),3) & " today."
				txtMessage = "Invoice # " & PassedInvoiceNumber & " was updated to not delivered at " & FormatDateTime(Now(),3) & " today."
			Case "Delivered"
				emailSubject = "Invoice " & PassedInvoiceNumber & " delivered"
				emailHeadLineText = "Invoice # " & PassedInvoiceNumber & " was updated to <i>delivered</i> at " & FormatDateTime(GetLastDeliveryStatusChangeBYInvoiceNumDelBoard(PassedInvoiceNumber),3) & " today."
				txtMessage = "Invoice # " & PassedInvoiceNumber & " was delivered at " & FormatDateTime(Now(),3) & " today."
			Case "Partial"
				emailSubject = "Partial delivery update for " & GetTerm("customer") & " " & GetCustNumberByInvoiceNumDelBoard(PassedInvoiceNumber) 
				emailHeadLineText = "The delivery for " & GetTerm("customer") & " " & GetCustNumberByInvoiceNumDelBoard(PassedInvoiceNumber) & " was <i>partially</i> updated at " & FormatDateTime(GetLastDeliveryStatusChangeBYInvoiceNumDelBoard(PassedInvoiceNumber),3) & " today."
				txtMessage = GetTerm("Customer") & " " & GetCustNumberByInvoiceNumDelBoard(PassedInvoiceNumber) & " partial delivery " & FormatDateTime(Now(),3) & " today."
		End Select
		
		%>
		<!--#include file="../../emails/deliveryboard_realtime_alert.asp"-->	
		<%
		
		'Now Send the emails
		'Got all the addresses so now break them up
		Send_To_Array = Split(Send_To,";")
	
		For x = 0 to Ubound(Send_To_Array) -1
			Send_To = Send_To_Array(x)
			Response.Write("<font color='green'><b>Sending alerts to Send_To: " & Send_To & " for ticket#: " & PassedMemoNumber & "</b></font><br>")
			
			If passedNotificationType = "Alert" Then
				SendFrom = "MDS Insight Alerts"
			Elseif passedNotificationType = "Notification" Then
				SendFrom = "MDS Insight Notifications"
			Else
				SendFrom = "MDS Insight"
			End If
			SendFrom = SendFrom & " (" & MUV_READ("ClientID") & ")"
			
			SendMail "mailsender@" & maildomain,Send_To, emailSubject,emailBody,"Alerts","System Alert",SendFrom
			
		
			CreateAuditLogEntry "Delivery Board Alert Sent","Delivery Board Alert Sent","Minor",0,"Delivery Board Realtime Alert Sent to " & Send_To & " for invoice #: " & PassedInvoiceNumber & " - " & emailSubject 
			
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
	
	
		txtSubject = "DeliveryAlert"
		
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
		
		CreateAuditLogEntry "Delivery Board Alert Sent","Delivery Board Alert Sent","Minor",0,"Delivery Board Realtime Alert Sent to " & TEXT_TO & " for invoice #: " & PassedInvoiceNumber & " - " & emailSubject 	
		
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

	Call WriteAlertSentRecord (rsAlert("InternalAlertRecNumber"),PassedInvoiceNumber)

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


Sub WriteAlertSentRecord (PassedAlertRecIdentifier,PassedInvoiceNumber)

		Set cnnAlertsSent = Server.CreateObject("ADODB.Connection")
		cnnAlertsSent.open (MUV_Read("ClientCnnString"))
		Set rsSCAlertsSent = Server.CreateObject("ADODB.Recordset")
		rsSCAlertsSent.CursorLocation = 3 

		rsSCAlertsSent = "Insert Into SC_AlertsSent (Alert_InternalAlertRecNumber,InvoiceNumberIfApplicable) Values (" &  PassedAlertRecIdentifier & "," & PassedInvoiceNumber & ")"

		Set rsSCAlertsSent = cnnAlertsSent.Execute(rsSCAlertsSent)

		Set rsSCAlertsSent = Nothing
		cnnAlertsSent.Close
		Set cnnAlertsSent = Nothing

End Sub

Function AlertSent(passedInternalAlertRecNumber,passedInvoiceNumber)

		resultAlertSent = False
		
		Set cnnAlertSent  = Server.CreateObject("ADODB.Connection")
		cnnAlertSent.open (MUV_Read("ClientCnnString"))
		Set rsAlertSent  = Server.CreateObject("ADODB.Recordset")
		rsAlertSent.CursorLocation = 3 

		SQLAlertSent  = "Select * from SC_AlertsSent  Where Alert_InternalAlertRecNumber = " & passedInternalAlertRecNumber & " AND InvoiceNumberIfApplicable = '" & passedInvoiceNumber & "'"

		Set rsAlertSent = cnnAlertSent.Execute(SQLAlertSent)
		If not rsAlertSent.Eof Then resultAlertSent = True

		Set rsAlertSent = Nothing
		cnnAlertSent.Close
		Set cnnAlertSent  = Nothing
		
		AlertSent = resultAlertSent

End Function 

Function PartialDelivery(passedInvoiceNumber)

		resultPartialDelivery = False
		
		Set cnnPartialDelivery  = Server.CreateObject("ADODB.Connection")
		cnnPartialDelivery.open (MUV_Read("ClientCnnString"))
		Set rsPartialDelivery  = Server.CreateObject("ADODB.Recordset")
		rsPartialDelivery.CursorLocation = 3 

		SQLPartialDelivery  = "Select Count(*) AS CustCount from RT_DeliveryBoard WHERE CustNum = '" & GetCustNumberByInvoiceNumDelBoard(passedInvoiceNumber) & "'"

		Set rsPartialDelivery = cnnPartialDelivery.Execute(SQLPartialDelivery)
		
		If not rsPartialDelivery.Eof Then
			If rsPartialDelivery("CustCount") > 1 Then 'more than  1 invoice
			
				CustCount = rsPartialDelivery("CustCount")
				
				'There is more than 1 invoice for this customer, so we have to figure more things out
				SQLPartialDelivery  = "Select * from RT_DeliveryBoard WHERE CustNum = '" & GetCustNumberByInvoiceNumDelBoard(passedInvoiceNumber) & "'"
				Set rsPartialDelivery = cnnPartialDelivery.Execute(SQLPartialDelivery)
				DelCount = 0
				NoDelCount = 0				
				NullCount = 0
				
				If Not rsPartialDelivery.EOF Then
					Do While Not rsPartialDelivery.EOF
						
						If Not IsNull(rsPartialDelivery("DeliveryStatus")) Then 
							If rsPartialDelivery("DeliveryStatus") = "Delivered" Then DelCount = DelCount + 1
							If rsPartialDelivery("DeliveryStatus") = "No Delivery" Then DelCount = DelCount + 1
						Else
							NullCount = NullCount + 1
						End If
						
						rsPartialDelivery.movenext
					Loop
					
					'Now see if there are different delivery statuses
					If DelCount <> 0 AND DelCount <> CustCount  Then resultPartialDelivery = True
					If NoDelCount <> 0 AND NoDelCount <> CustCount  Then resultPartialDelivery = True
					If NullCount <> 0 AND NullCount <> CustCount  Then resultPartialDelivery = True

				End If

			End If
		End If

		Set rsPartialDelivery = Nothing
		cnnPartialDelivery.Close
		Set cnnPartialDelivery  = Nothing
		
		'This is a little different from normal operation. It may appear to be partil becuase
		'the driver has not finished marking the delivery yet. So, if it appears to be partial,
		'wait 25 seconds and check again
		'Does the whole thing all over again
		
		If resultPartialDelivery = True And (DelCount + NoDelCount) <> CustCount Then
		
				Response.Write("Waiting 25 secs to check account number " & GetCustNumberByInvoiceNumDelBoard(passedInvoiceNumber) & "<br>")
				Response.Write("DelCount:" & DelCount   & "<br>")
				Response.Write("NoDelCount:" & NoDelCount  & "<br>")
				Response.Write("NullCount:" & NullCount   & "<br>")
				Response.Write("CustCount:" & CustCount  & "<br>")

				Call Delay(25)

				Response.Write("Finsihed waiting 25 secs")
				
				resultPartialDelivery = False
				
				Set cnnPartialDelivery  = Server.CreateObject("ADODB.Connection")
				cnnPartialDelivery.open (MUV_Read("ClientCnnString"))
				Set rsPartialDelivery  = Server.CreateObject("ADODB.Recordset")
				rsPartialDelivery.CursorLocation = 3 
		
				SQLPartialDelivery  = "Select Count(*) AS CustCount from RT_DeliveryBoard WHERE CustNum = '" & GetCustNumberByInvoiceNumDelBoard(passedInvoiceNumber) & "'"
		
				Set rsPartialDelivery = cnnPartialDelivery.Execute(SQLPartialDelivery)
				
				If not rsPartialDelivery.Eof Then
					If rsPartialDelivery("CustCount") > 1 Then 'more than  1 invoice
					
						CustCount = rsPartialDelivery("CustCount") 
					
						'There is more than 1 invoice for this customer, so we have to figure more things out
						SQLPartialDelivery  = "Select * from RT_DeliveryBoard WHERE CustNum = '" & GetCustNumberByInvoiceNumDelBoard(passedInvoiceNumber) & "'"
						Set rsPartialDelivery = cnnPartialDelivery.Execute(SQLPartialDelivery)
						DelCount = 0
						NoDelCount = 0				
						NullCount = 0
						If Not rsPartialDelivery.EOF Then
							Do While Not rsPartialDelivery.EOF
								
								If Not IsNull(rsPartialDelivery("DeliveryStatus")) Then 
									If rsPartialDelivery("DeliveryStatus") = "Delivered" Then DelCount = DelCount + 1
									If rsPartialDelivery("DeliveryStatus") = "No Delivery" Then DelCount = DelCount + 1
								Else
									NullCount = NullCount + 1
								End If
								
								rsPartialDelivery.movenext
							Loop
							
							'Now see if there are different delivery statuses
							If DelCount <> 0 AND DelCount <> CustCount  Then resultPartialDelivery = True
							If NoDelCount <> 0 AND NoDelCount <> CustCount  Then resultPartialDelivery = True
							If NullCount <> 0 AND NullCount <> CustCount  Then resultPartialDelivery = True
		
						End If
		
					End If
				End If
		
				Set rsPartialDelivery = Nothing
				cnnPartialDelivery.Close
				Set cnnPartialDelivery  = Nothing
		
		End If
		
		PartialDelivery = resultPartialDelivery

End Function  

Sub Delay(DelaySeconds)
	SecCount = 0
	Sec2 = 0
	While SecCount < DelaySeconds + 1
		Sec1 = Second(Time())
		If Sec1 <> Sec2 Then
			Sec2 = Second(Time())
			SecCount = SecCount + 1
		End If
	Wend 
End Sub
%>