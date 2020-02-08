<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/mailDirectLaunch.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Orders.asp"-->
<!--#include file="../../inc/InsightFuncs_API.asp"-->
<script type="text/javascript">
    function closeme() {
window.open('', '_parent', '');
window.close();  }
</script>
<%
'Order API Alert processing page
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/alerts/OrderAPIAlertsCheck.asp?runlevel=run_now
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
			'Begin Order API Alerts
			'****************************************
			Response.Write("Begin Order API Alerts<br>")
			
			'**********************************************
			'Delete alerts sent more than 30 days
			'ONLY delete Order API, others will do their own
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
					If GetAlertType(rsSCAlertsSent("Alert_InternalAlertRecNumber")) = "OrderAPIAlert" Then
						'OK, it is an old Order API alert so we can delete it					
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
			
			SQLAlerts = "SELECT * FROM SC_Alerts Where AlertType='OrderAPIAlert' And Enabled = 1" 
				
			Set rsAlert = Server.CreateObject("ADODB.Recordset")
			rsAlert.CursorLocation = 3 
			Set rsAlert = cnnAlerts.Execute(SQLAlerts)
			
			If not rsAlert.EOF Then
						
				Do While Not rsAlert.EOF
				
					Response.Write("<b>Found Alert Named: " & rsAlert("AlertName") &"</b><br>") 
			
						
						'Now the real work begins
						'Run through API_OR_OrderHeader &  see if this alert needs to be sent
						
						Select Case rsAlert("Condition")
						
							Case "Order Contains Driver Notes"
						
								Response.Write("Check Order Contains Driver Notes <br>")
								
								Set cnnOrderAPI = Server.CreateObject("ADODB.Connection")
								cnnOrderAPI.open (MUV_Read("ClientCnnString"))
								Set rsOrderAPI = Server.CreateObject("ADODB.Recordset")
								rsOrderAPI.CursorLocation = 3 
	
								SQL_OrderAPI = "SELECT OrderID FROM API_OR_OrderHeader WHERE Voided <> 1 AND "
								SQL_OrderAPI = SQL_OrderAPI & " (DriverNotes IS NOT NULL AND DriverNotes <> '') AND "
								SQL_OrderAPI = SQL_OrderAPI & " API_OR_OrderHeader.RecordCreationDateTime >" & rsAlert("RecordCreationDate") & " AND"
								SQL_OrderAPI = SQL_OrderAPI & " OrderID NOT IN (SELECT OrderIDIfApplicable FROM SC_AlertsSent WHERE Alert_InternalAlertRecNumber = " & rsAlert("InternalAlertRecNumber") & ") "								
								SQL_OrderAPI = SQL_OrderAPI & " GROUP BY OrderID "
								
								Set rsOrderAPI = cnnOrderAPI.Execute(SQL_OrderAPI)
								
								If Not rsOrderAPI.EOF Then
								
									Do While Not rsOrderAPI.Eof
																				
										If AlertSent(rsAlert("InternalAlertRecNumber"),rsOrderAPI("OrderID")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsOrderAPI("OrderID"),rsAlert("Condition"),rsAlert("NotificationType")
										rsOrderAPI.MoveNext
										
									Loop
									
								End If
								
								Set rsOrderAPI = Nothing
								cnnOrderAPI.Close
								Set cnnOrderAPI = Nothing
								


							Case "Order Contains Warehouse Notes"
						
								Response.Write("Check Order Contains Warehouse Notes <br>")
								
								Set cnnOrderAPI = Server.CreateObject("ADODB.Connection")
								cnnOrderAPI.open (MUV_Read("ClientCnnString"))
								Set rsOrderAPI = Server.CreateObject("ADODB.Recordset")
								rsOrderAPI.CursorLocation = 3 
	
								SQL_OrderAPI = "SELECT OrderID FROM API_OR_OrderHeader WHERE Voided <> 1 AND "
								SQL_OrderAPI = SQL_OrderAPI & "(WarehouseNotes IS NOT NULL AND WarehouseNotes <> '') AND "
								SQL_OrderAPI = SQL_OrderAPI & " API_OR_OrderHeader.RecordCreationDateTime >" & rsAlert("RecordCreationDate") & " AND"
								SQL_OrderAPI = SQL_OrderAPI & " OrderID NOT IN (SELECT OrderIDIfApplicable FROM SC_AlertsSent WHERE Alert_InternalAlertRecNumber = " & rsAlert("InternalAlertRecNumber") & ") "								
								SQL_OrderAPI = SQL_OrderAPI & " GROUP BY OrderID "
							
								
								Set rsOrderAPI = cnnOrderAPI.Execute(SQL_OrderAPI)
								
								If Not rsOrderAPI.EOF Then
								
									Do While Not rsOrderAPI.Eof
																				
										If AlertSent(rsAlert("InternalAlertRecNumber"),rsOrderAPI("OrderID")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsOrderAPI("OrderID"),rsAlert("Condition"),rsAlert("NotificationType")
										rsOrderAPI.MoveNext
										
									Loop
									
								End If
								
								Set rsOrderAPI = Nothing
								cnnOrderAPI.Close
								Set cnnOrderAPI = Nothing

								


							Case "RA Contains Driver Notes"
						
								Response.Write("Check RA Contains Driver Notes <br>")
								
								Set cnnOrderAPI = Server.CreateObject("ADODB.Connection")
								cnnOrderAPI.open (MUV_Read("ClientCnnString"))
								Set rsOrderAPI = Server.CreateObject("ADODB.Recordset")
								rsOrderAPI.CursorLocation = 3 
	
								SQL_OrderAPI = "SELECT OrderID FROM API_OR_RAHeader WHERE Voided <> 1 AND "
								SQL_OrderAPI = SQL_OrderAPI & "(DriverNotes IS NOT NULL AND DriverNotes <> '') AND "
								SQL_OrderAPI = SQL_OrderAPI & " API_OR_RAHeader.RecordCreationDateTime >" & rsAlert("RecordCreationDate") & " AND"
								SQL_OrderAPI = SQL_OrderAPI & " OrderID NOT IN (SELECT OrderIDIfApplicable FROM SC_AlertsSent WHERE Alert_InternalAlertRecNumber = " & rsAlert("InternalAlertRecNumber") & ") "								
								SQL_OrderAPI = SQL_OrderAPI & " GROUP BY OrderID "
							
								
								Set rsOrderAPI = cnnOrderAPI.Execute(SQL_OrderAPI)
								
								If Not rsOrderAPI.EOF Then
								
									Do While Not rsOrderAPI.Eof
																				
										If AlertSent(rsAlert("InternalAlertRecNumber"),rsOrderAPI("OrderID")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsOrderAPI("OrderID"),rsAlert("Condition"),rsAlert("NotificationType")
										rsOrderAPI.MoveNext
										
									Loop
									
								End If
								
								Set rsOrderAPI = Nothing
								cnnOrderAPI.Close
								Set cnnOrderAPI = Nothing

								


							Case "RA Contains Warehouse Notes"
						
								Response.Write("Check RA Contains Warehouse Notes <br>")
								
								Set cnnOrderAPI = Server.CreateObject("ADODB.Connection")
								cnnOrderAPI.open (MUV_Read("ClientCnnString"))
								Set rsOrderAPI = Server.CreateObject("ADODB.Recordset")
								rsOrderAPI.CursorLocation = 3 
	
								SQL_OrderAPI = "SELECT OrderID FROM API_OR_RAHeader WHERE Voided <> 1 AND "
								SQL_OrderAPI = SQL_OrderAPI & "(WarehouseNotes IS NOT NULL AND WarehouseNotes <> '') AND "
								SQL_OrderAPI = SQL_OrderAPI & " API_OR_RAHeader.RecordCreationDateTime >" & rsAlert("RecordCreationDate") & " AND"
								SQL_OrderAPI = SQL_OrderAPI & " OrderID NOT IN (SELECT OrderIDIfApplicable FROM SC_AlertsSent WHERE Alert_InternalAlertRecNumber = " & rsAlert("InternalAlertRecNumber") & ") "								
								SQL_OrderAPI = SQL_OrderAPI & " GROUP BY OrderID "
							
								
								Set rsOrderAPI = cnnOrderAPI.Execute(SQL_OrderAPI)
								
								If Not rsOrderAPI.EOF Then
								
									Do While Not rsOrderAPI.Eof
																				
										If AlertSent(rsAlert("InternalAlertRecNumber"),rsOrderAPI("OrderID")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsOrderAPI("OrderID"),rsAlert("Condition"),rsAlert("NotificationType")
										rsOrderAPI.MoveNext
										
									Loop
									
								End If
								
								Set rsOrderAPI = Nothing
								cnnOrderAPI.Close
								Set cnnOrderAPI = Nothing
							

								


							Case "RA Contains RA Notes"
						
								Response.Write("Check RA Contains Warehouse Notes <br>")
								
								Set cnnOrderAPI = Server.CreateObject("ADODB.Connection")
								cnnOrderAPI.open (MUV_Read("ClientCnnString"))
								Set rsOrderAPI = Server.CreateObject("ADODB.Recordset")
								rsOrderAPI.CursorLocation = 3 
	
								SQL_OrderAPI = "SELECT OrderID FROM API_OR_RAHeader WHERE Voided <> 1 AND "
								SQL_OrderAPI = SQL_OrderAPI & "(RA_Notes IS NOT NULL AND RA_Notes <> '') AND "
								SQL_OrderAPI = SQL_OrderAPI & " API_OR_RAHeader.RecordCreationDateTime >" & rsAlert("RecordCreationDate") & " AND"
								SQL_OrderAPI = SQL_OrderAPI & " OrderID NOT IN (SELECT OrderIDIfApplicable FROM SC_AlertsSent WHERE Alert_InternalAlertRecNumber = " & rsAlert("InternalAlertRecNumber") & ") "								
								SQL_OrderAPI = SQL_OrderAPI & " GROUP BY OrderID "
							
								
								Set rsOrderAPI = cnnOrderAPI.Execute(SQL_OrderAPI)
								
								If Not rsOrderAPI.EOF Then
								
									Do While Not rsOrderAPI.Eof
																				
										If AlertSent(rsAlert("InternalAlertRecNumber"),rsOrderAPI("OrderID")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsOrderAPI("OrderID"),rsAlert("Condition"),rsAlert("NotificationType")
										rsOrderAPI.MoveNext
										
									Loop
									
								End If
								
								Set rsOrderAPI = Nothing
								cnnOrderAPI.Close
								Set cnnOrderAPI = Nothing
							
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

Sub SendAlert(passedAlertNumber,PassedOrderID,AlertType,passedNotificationType)


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
		
			Case "Order Contains Driver Notes"
			
				emailSubject = "Driver notes - Order " & PassedOrderID & " - (" & ClientKey & ")"
				emailHeadLineText = "Order " & PassedOrderID & " received via the order api includes the following driver note: " & GetOrderDriverNotesByOrderID(PassedOrderID)
				txtMessage = "Driver notes - Order " & PassedOrderID & " - (" & ClientKey & ")"
			
			Case "Order Contains Warehouse Notes"
			
				emailSubject = "Warehouse notes - Order " & PassedOrderID & " - (" & ClientKey & ")"
				emailHeadLineText = "Order " & PassedOrderID & " received via the order api includes the following warehouse note: " & GetOrderWarehouseNotesByOrderID(PassedOrderID)
				txtMessage = "Warehouse notes - Order " & PassedOrderID & " - (" & ClientKey & ")"
			
			Case "RA Contains Driver Notes"
			
				emailSubject = "RA " & GetRAIDByOrderID(PassedOrderId) & " has driver notes - Order " & PassedOrderID & " - (" & ClientKey & ")"
				emailHeadLineText = "RA " & GetRAIDByOrderID(PassedOrderId) & " for Order " & PassedOrderID & " received via the order api includes the following driver note: " & GetRADriverNotesByOrderID(PassedOrderID)
				txtMessage = "RA " & GetRAIDByOrderID(PassedOrderId) & " has driver notes - Order " & PassedOrderID & " - (" & ClientKey & ")"
			
			Case "RA Contains Warehouse Notes"
			
				emailSubject = "RA " & GetRAIDByOrderID(PassedOrderId) & " has warehouse notes - Order " & PassedOrderID & " - (" & ClientKey & ")"
				emailHeadLineText = "RA " & GetRAIDByOrderID(PassedOrderId) & " for Order " & PassedOrderID & " received via the order api includes the following warehouse note: " & GetRAWarehouseNotesByOrderID(PassedOrderID)
				txtMessage = "RA " & GetRAIDByOrderID(PassedOrderId) & " has warehouse notes - Order " & PassedOrderID & " - (" & ClientKey & ")"
			
			Case "RA Contains RA Notes"
			
				emailSubject = "RA " & GetRAIDByOrderID(PassedOrderId) & " has RA notes - Order " & PassedOrderID & " - (" & ClientKey & ")"
				emailHeadLineText = "RA " & GetRAIDByOrderID(PassedOrderId) & " for Order " & PassedOrderID & " received via the order api includes the following RA note: " & GetRARANotesByOrderID(PassedOrderID)
				txtMessage = "RA " & GetRAIDByOrderID(PassedOrderId) & " has RA notes - Order " & PassedOrderID & " - (" & ClientKey & ")"
			
			End Select
		
		%>
		<!--#include file="../../emails/orderAPI_realtime_alert.asp"-->	
		<%
		
		'Now Send the emails
		'Got all the addresses so now break them up
		Send_To_Array = Split(Send_To,";")
	
		For x = 0 to Ubound(Send_To_Array) -1
			Send_To = Send_To_Array(x)
			Response.Write("<font color='green'><b>Sending alerts to Send_To: " & Send_To & " for Order #: " & PassedOrderID & "</b></font><br>")
			
			If passedNotificationType = "Alert" Then
				SendFrom = "MDS Insight Alerts"
			Elseif passedNotificationType = "Notification" Then
				SendFrom = "MDS Insight Notifications"
			Else
				SendFrom = "MDS Insight"
			End If
			SendFrom = SendFrom & " (" & MUV_READ("ClientID") & ")"
			
			SendMail "mailsender@" & maildomain,Send_To, emailSubject,emailBody,"Alerts","System Alert",SendFrom
			
			CreateAuditLogEntry "Order API Alert Sent","Order API Alert Sent","Minor",0,"Order API Realtime Alert Sent to " & Send_To & " for invoice #: " & PassedOrderID & " - " & emailSubject 
			
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
	
	
		txtSubject = "OrderAPIAlert"
		
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
		
		CreateAuditLogEntry "Order API Alert Sent","Order API Alert Sent","Minor",0,"Order API Realtime Alert Sent to " & TEXT_TO & " for order #: " & PassedOrderID & " - " & emailSubject 	
		
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

	Call WriteAlertSentRecord (rsAlert("InternalAlertRecNumber"),PassedOrderID)

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


Sub WriteAlertSentRecord (PassedAlertRecIdentifier,PassedOrderID)

		Set cnnAlertsSent = Server.CreateObject("ADODB.Connection")
		cnnAlertsSent.open (MUV_Read("ClientCnnString"))
		Set rsSCAlertsSent = Server.CreateObject("ADODB.Recordset")
		rsSCAlertsSent.CursorLocation = 3 

		rsSCAlertsSent = "INSERT INTO SC_AlertsSent (Alert_InternalAlertRecNumber,OrderIDIfApplicable) Values (" &  PassedAlertRecIdentifier & ",'" & PassedOrderID & "')"

		Set rsSCAlertsSent = cnnAlertsSent.Execute(rsSCAlertsSent)

		Set rsSCAlertsSent = Nothing
		cnnAlertsSent.Close
		Set cnnAlertsSent = Nothing

End Sub

Function AlertSent(passedInternalAlertRecNumber,PassedOrderID)

		resultAlertSent = False
		
		Set cnnAlertSent  = Server.CreateObject("ADODB.Connection")
		cnnAlertSent.open (MUV_Read("ClientCnnString"))
		Set rsAlertSent  = Server.CreateObject("ADODB.Recordset")
		rsAlertSent.CursorLocation = 3 

		SQLAlertSent  = "SELECT * FROM SC_AlertsSent WHERE Alert_InternalAlertRecNumber = " & passedInternalAlertRecNumber & " AND OrderIDIfApplicable = '" & PassedOrderID & "'"

		Set rsAlertSent = cnnAlertSent.Execute(SQLAlertSent)
		If not rsAlertSent.Eof Then resultAlertSent = True

		Set rsAlertSent = Nothing
		cnnAlertSent.Close
		Set cnnAlertSent  = Nothing
		
		AlertSent = resultAlertSent

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