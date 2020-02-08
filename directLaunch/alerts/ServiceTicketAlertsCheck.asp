<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/mailDirectlaunch.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.asp"-->
<script type="text/javascript">
    function closeme() {
window.open('', '_parent', '');
window.close();  }
</script>
<%
Server.ScriptTimeout = 2500
Dim ServiceNotesForEmail 

'Service Alert processing page
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/alerts/ServiceTicketAlertsCheck.asp?runlevel=run_now

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)

If Request.QueryString("runlevel") <> "run_now" then response.end

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
maildomain = Replace(maildomain,"www.","")

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

			'The IF statement below makes sure that when run from DEV it only deos client keys with a d
			'and when run from LIVE it only does client keys without a d
			'Pretty smart, huh
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
			or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then 

				Response.Write("******** Processing <b>" & ClientKey  & "</b>************<br>")
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
					'Begin Service Ticket Alerts
					'****************************************
					Response.Write("<br><font color='blue'><b>Begin Service Ticket Alerts</b></font><br>")
					'**********************************************
					'Delete alerts sent more than 48 hours ago
					'If we don't the system always thinks the alerts
					'have been sent & the max send thresholds have been met
					'ONLY delete Service Alerts, night batch will do it's own
					Set cnnAlertsSent = Server.CreateObject("ADODB.Connection")
					cnnAlertsSent.open (MUV_Read("ClientCnnString"))
					Set rsSCAlertsSent = Server.CreateObject("ADODB.Recordset")
					rsSCAlertsSent.CursorLocation = 3 
					'''''on error resume next ' In case they dont have this table
					
					rsSCAlertsSent = "SELECT * FROM SC_AlertsSent WHERE DateTimeSent < dateadd(hh,-840,getdate())" ' All cleared after 2 weeks, this is OK, different than nightbatch alerts
					Set rsSCAlertsSent = cnnAlertsSent.Execute(rsSCAlertsSent)
					on error goto 0
					
					If Not rsSCAlertsSent.EOF Then
					
						Set rsForDeletes = Server.CreateObject("ADODB.Recordset")
						rsForDeletes.CursorLocation = 3 
		
						Do While Not rsSCAlertsSent.EOF
							If GetAlertType(rsSCAlertsSent("Alert_InternalAlertRecNumber")) = "ServiceElapsed" OR GetAlertType(rsSCAlertsSent("Alert_InternalAlertRecNumber")) = "ServiceNumTick" OR GetAlertType(rsSCAlertsSent("Alert_InternalAlertRecNumber")) = "ServiceOtherConditions" Then
								'OK, it is an old service alert so we can delete it					
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
					
					'*********************************************************************************************
					'CHECK SERVICE TICKET ELAPSED ALERTS
					'*********************************************************************************************
				
					Set cnnAlerts = Server.CreateObject("ADODB.Connection")
					cnnAlerts.open (MUV_Read("ClientCnnString"))
					
					SQLAlerts = "SELECT * FROM SC_Alerts Where AlertType='ServiceElapsed' And Enabled = 1" 
						
					Set rsAlert = Server.CreateObject("ADODB.Recordset")
					rsAlert.CursorLocation = 3 
					Set rsAlert = cnnAlerts.Execute(SQLAlerts)
					
					If not rsAlert.EOF Then
								
						Do While Not rsAlert.EOF
						
							Response.Write("<b>Found Alert Named: " & rsAlert("AlertName") &"</b><br>") 
					
								
								'Now the real work begins
								'Run through all the service memos &  see if this alert needs to be sent
								Select Case rsAlert("Condition")
									Case "NotDispatched"
								
										Set cnnServiceMemos = Server.CreateObject("ADODB.Connection")
										cnnServiceMemos.open (MUV_Read("ClientCnnString"))
										Set rsServiceMemos = Server.CreateObject("ADODB.Recordset")
										rsServiceMemos.CursorLocation = 3 
			
										SQL_ServiceMemos = "SELECT Distinct MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN'"
										
										Set rsServiceMemos = cnnServiceMemos.Execute(SQL_ServiceMemos)
										
										If Not rsServiceMemos.EOF Then
											Do While Not rsServiceMemos.Eof
												
												'Response.Write(rsServiceMemos("MemoNumber") & " - " & GetServiceTicketCurrentStage(rsServiceMemos("MemoNumber")) & " - " & TimeFromStageStartUntilTargetDateTime(rsServiceMemos("MemoNumber"),GetServiceTicketCurrentStage(rsServiceMemos("MemoNumber")),Now()) & "<br>")
												
												If TimeFromStageStartUntilTargetDateTime(rsServiceMemos("MemoNumber"),GetServiceTicketCurrentStage(rsServiceMemos("MemoNumber")),Now()) >   rsAlert("NBMinutes") Then 
												
													If ServiceTicketIsDispatched(rsServiceMemos("MemoNumber")) = False Then
														If AlertSent(rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber"),rsAlert("Condition")
													End If
												
												End If
											
												rsServiceMemos.MoveNext
											Loop
										End If
										
										Set rsServiceMemos = Nothing
										cnnServiceMemos.Close
										Set cnnServiceMemos = Nothing
										
									Case "NoACK"
		
										If advancedDispatchIsOn() Then
											
											Set cnnServiceMemos = Server.CreateObject("ADODB.Connection")
											cnnServiceMemos.open (MUV_Read("ClientCnnString"))
											Set rsServiceMemos = Server.CreateObject("ADODB.Recordset")
											rsServiceMemos.CursorLocation = 3 
				
											SQL_ServiceMemos = "SELECT Distinct MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN'"
											
											Set rsServiceMemos = cnnServiceMemos.Execute(SQL_ServiceMemos)
											
											If Not rsServiceMemos.EOF Then
												Do While Not rsServiceMemos.Eof
													
													If GetServiceTicketCurrentStage(rsServiceMemos("MemoNumber")) = "Dispatched" Then 
														If TimeFromStageStartUntilTargetDateTime(rsServiceMemos("MemoNumber"),"Dispatched",Now()) > rsAlert("NBMinutes") Then
															If ServiceTicketDispatchACKed(rsServiceMemos("MemoNumber")) = False Then
																If AlertSent(rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber"),rsAlert("Condition")
															End If
														End If
													End If
												
													rsServiceMemos.MoveNext
												Loop
											End If
											
											Set rsServiceMemos = Nothing
											cnnServiceMemos.Close
											Set cnnServiceMemos = Nothing
												
										End If
									
									Case "NoOnSite"
									
										If advancedDispatchIsOn() Then
											
											
												Set cnnServiceMemos = Server.CreateObject("ADODB.Connection")
												cnnServiceMemos.open (MUV_Read("ClientCnnString"))
												Set rsServiceMemos = Server.CreateObject("ADODB.Recordset")
												rsServiceMemos.CursorLocation = 3 
					
												SQL_ServiceMemos = "SELECT Distinct MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN'"
												
												Set rsServiceMemos = cnnServiceMemos.Execute(SQL_ServiceMemos)
												
												If Not rsServiceMemos.EOF Then
													Do While Not rsServiceMemos.Eof
													
														'Originally had this check for DISP and DISP ACK but then realized it doesn't need to check that
														'This is for no one on site, regardless of dispatch status
														'It checks the alert minutes against the elapsed minutes from when the ticket originated, regardless of stage
														If CalcServiceTicketElapsedMinutes(rsServiceMemos("MemoNumber")) > rsAlert("NBMinutes") Then
															If ServiceTicketWasOnSite(rsServiceMemos("MemoNumber")) = False Then
																If AlertSent(rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber"),rsAlert("Condition")
															End If
														End If 												
														rsServiceMemos.MoveNext
													Loop
												End If
												
												Set rsServiceMemos = Nothing
												cnnServiceMemos.Close
												Set cnnServiceMemos = Nothing
										End If
									
									Case "OpenTooLong"
									
										Set cnnServiceMemos = Server.CreateObject("ADODB.Connection")
										cnnServiceMemos.open (MUV_Read("ClientCnnString"))
										Set rsServiceMemos = Server.CreateObject("ADODB.Recordset")
										rsServiceMemos.CursorLocation = 3 
			
										SQL_ServiceMemos = "SELECT Distinct MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN'"
										
										Set rsServiceMemos = cnnServiceMemos.Execute(SQL_ServiceMemos)
										
										If Not rsServiceMemos.EOF Then
											Do While Not rsServiceMemos.Eof
											
												If ServiceCallElapsedMinutes(rsServiceMemos("MemoNumber")) > rsAlert("NBMinutes") Then
											
													If AlertSent(rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber"),rsAlert("Condition")
												
												End If
											
												rsServiceMemos.MoveNext
											Loop
										End If
										
										Set rsServiceMemos = Nothing
										cnnServiceMemos.Close
										Set cnnServiceMemos = Nothing
									
									Case "RedispatchTooLong"
									
										Set cnnServiceMemos = Server.CreateObject("ADODB.Connection")
										cnnServiceMemos.open (MUV_Read("ClientCnnString"))
										Set rsServiceMemos = Server.CreateObject("ADODB.Recordset")
										rsServiceMemos.CursorLocation = 3 
			
										SQL_ServiceMemos = "SELECT Distinct MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN'"
										
										Set rsServiceMemos = cnnServiceMemos.Execute(SQL_ServiceMemos)
										
										If Not rsServiceMemos.EOF Then
											Do While Not rsServiceMemos.Eof
												
												If AwaitingRedispatch(rsServiceMemos("MemoNumber")) Then
												
													If TimeFromStageStartUntilTargetDateTime(rsServiceMemos("MemoNumber"),"Redispatch",Now()) > rsAlert("NBMinutes") Then
												
														If AlertSent(rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber"),rsAlert("Condition")
														
													End If
													
												End If
											
												rsServiceMemos.MoveNext
											Loop
										End If
										
										Set rsServiceMemos = Nothing
										cnnServiceMemos.Close
										Set cnnServiceMemos = Nothing
									
									Case "AnyStage"
									
										Set cnnServiceMemos = Server.CreateObject("ADODB.Connection")
										cnnServiceMemos.open (MUV_Read("ClientCnnString"))
										Set rsServiceMemos = Server.CreateObject("ADODB.Recordset")
										rsServiceMemos.CursorLocation = 3 
			
										SQL_ServiceMemos = "SELECT Distinct MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN'"
										
										Set rsServiceMemos = cnnServiceMemos.Execute(SQL_ServiceMemos)
										
										If Not rsServiceMemos.EOF Then
											Do While Not rsServiceMemos.Eof
											
												If TimeFromStageStartUntilTargetDateTime(rsServiceMemos("MemoNumber"),GetServiceTicketCurrentStage(rsServiceMemos("MemoNumber")),Now()) > rsAlert("NBMinutes") Then
											
													If AlertSent(rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber")) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber"),rsAlert("Condition")
														
												
												End If
											
												rsServiceMemos.MoveNext
											Loop
										End If
										
										Set rsServiceMemos = Nothing
										cnnServiceMemos.Close
										Set cnnServiceMemos = Nothing
									
									Case "Declined"
									
										Set cnnServiceMemos = Server.CreateObject("ADODB.Connection")
										cnnServiceMemos.open (MUV_Read("ClientCnnString"))
										Set rsServiceMemos = Server.CreateObject("ADODB.Recordset")
										rsServiceMemos.CursorLocation = 3 
			
										SQL_ServiceMemos = "SELECT Distinct MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN'"
										
										Set rsServiceMemos = cnnServiceMemos.Execute(SQL_ServiceMemos)
										
										If Not rsServiceMemos.EOF Then
											Do While Not rsServiceMemos.Eof
											
												If GetServiceTicketCurrentStage(rsServiceMemos("MemoNumber")) = "Dispatch Declined"  Then
											
													'Get the detail record number of the last decline
													RecNumHolder = MostRecentDispatchDeclineRecordNumberByTicket(rsServiceMemos("MemoNumber"))
													
													
													If AlertSentDetailLevel(rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber"),RecNumHolder) <> True Then SendAlert rsAlert("InternalAlertRecNumber"),rsServiceMemos("MemoNumber"),rsAlert("Condition")
		
												
												End If
											
												rsServiceMemos.MoveNext
											Loop
										End If
										
										Set rsServiceMemos = Nothing
										cnnServiceMemos.Close
										Set cnnServiceMemos = Nothing
									
								End Select
								
							rsAlert.Movenext
					
					
				Loop
	
			End If
				
			Set rsAlert = Nothing
			cnnAlerts.Close
			Set cnnALerts = Nothing
		
						
			
			'*********************************************************************************************
			'END CHECK SERVICE TICKET ELAPSED ALERTS
			'*********************************************************************************************
						


			'*********************************************************************************************
			'CHECK SERVICE TICKET CONTAINS TECH NOTES ALERTS
			'*********************************************************************************************

			Set cnnAlerts = Server.CreateObject("ADODB.Connection")
			cnnAlerts.open (MUV_Read("ClientCnnString"))
			
			SQLAlerts = "SELECT * FROM SC_Alerts Where AlertType='ServiceOtherConditions' And Enabled = 1" 
				
			Set rsAlert = Server.CreateObject("ADODB.Recordset")
			rsAlert.CursorLocation = 3 
			Set rsAlert = cnnAlerts.Execute(SQLAlerts)
			
			If not rsAlert.EOF Then
						
				Do While Not rsAlert.EOF
				
					Response.Write("<b>Found Alert Named: " & rsAlert("AlertName") &"</b><br>") 
		
					'Now the real work begins
					'Run through FS_ServiceMemosDetail and see if this alert needs to be sent
					
					Select Case rsAlert("Condition")
					
						Case "ContainsServiceNotes"
					
							Response.Write("Service Ticket Contains Service Notes <br>")
							
							Set cnnFsServiceMemos = Server.CreateObject("ADODB.Connection")
							cnnFsServiceMemos.open (MUV_Read("ClientCnnString"))
							Set rsFsServiceMemos = Server.CreateObject("ADODB.Recordset")
							rsFsServiceMemos.CursorLocation = 3 
							
							
							SQL_FsServiceMemos = "SELECT DISTINCT MemoNumber, ServiceNotes FROM ("
							
							SQL_FsServiceMemos = SQL_FsServiceMemos & " SELECT MemoNumber , ServiceNotesFromTech AS ServiceNotes FROM FS_ServiceMemos WHERE (CurrentStatus = 'CLOSE' OR CurrentStatus = 'CANCEL') AND "
							SQL_FsServiceMemos = SQL_FsServiceMemos & " (ServiceNotesFromTech IS NOT NULL AND ServiceNotesFromTech <> '') AND "
							SQL_FsServiceMemos = SQL_FsServiceMemos & " RecordCreatedateTime >= CAST('" & rsAlert("RecordCreationDate") & "' AS datetime) "

							SQL_FsServiceMemos = SQL_FsServiceMemos & " UNION "
							
							SQL_FsServiceMemos = SQL_FsServiceMemos & "SELECT DISTINCT MemoNumber, Remarks AS ServiceNotes FROM FS_ServiceMemosDetail "
							SQL_FsServiceMemos = SQL_FsServiceMemos & " WHERE (ClosedOrCancelled <> 1) AND (Remarks IS NOT NULL) AND (Remarks <> '') AND "
							SQL_FsServiceMemos = SQL_FsServiceMemos & " FS_ServiceMemosDetail.RecordCreatedDateTime >" & rsAlert("RecordCreationDate") & " AND"
							SQL_FsServiceMemos = SQL_FsServiceMemos & " (MemoStage <> 'Dispatched') AND "
							SQL_FsServiceMemos = SQL_FsServiceMemos & " (MemoStage <> 'Received') AND "
							SQL_FsServiceMemos = SQL_FsServiceMemos & " (MemoStage <> 'Dispatch Acknowledged') "
							SQL_FsServiceMemos = SQL_FsServiceMemos & " AND (MemoStage <> 'Under Review') "
							
							SQL_FsServiceMemos = SQL_FsServiceMemos & " )  AS derivedtbl_1"
							

							If ClientKey = "1230d" then
								'Response.Write("<br><br>" & SQL_FsServiceMemos & "<br><br><br><br><br>")
							'Response.End
							End If
							
							ServiceNotesForEmail = ""
							
							Set rsFsServiceMemos = cnnFsServiceMemos.Execute(SQL_FsServiceMemos)
							
							If NOT rsFsServiceMemos.EOF Then
							
								Do While NOT rsFsServiceMemos.EOF
																			
									If AlertSent(rsAlert("InternalAlertRecNumber"),rsFsServiceMemos("MemoNumber")) <> True Then 
										
										ServiceNotesForEmail = rsFsServiceMemos("ServiceNotes") ' Global var, dimmed at top, just to pass service notes to email
										
										SendAlert rsAlert("InternalAlertRecNumber"),rsFsServiceMemos("MemoNumber"),rsAlert("Condition")
									End If

									rsFsServiceMemos.MoveNext
									
								Loop
								
							End If
							
							Set rsFsServiceMemos = Nothing
							cnnFsServiceMemos.Close
							Set cnnFsServiceMemos = Nothing
						
					End Select

				rsAlert.Movenext
			Loop

		End If
			
		Set rsAlert = Nothing
		cnnAlerts.Close
		Set cnnALerts = Nothing
		
		'*********************************************************************************************
		'END CHECK SERVICE TICKET CONTAINS TECH NOTES ALERTS
		'*********************************************************************************************

		Response.Write("******** DONE Processing " & ClientKey  & "************<br>")
	End If		
	
	End If ' this is the endif for live doing live & dev doing dev
		
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

Sub SendAlert (passedAlertNumber,PassedMemoNumber,AlertType)

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
			Case "NotDispatched"
				emailSubject = "Ticket " & PassedMemoNumber & " has not been dispatched"
				emailHeadLineText = "Service ticket #" & PassedMemoNumber & " has remained un-dispatched since " & GetServiceTicketSTAGEDateTime(PassedMemoNumber,GetServiceTicketCurrentStage(PassedMemoNumber))
				txtMessage =  "Service ticket #" & PassedMemoNumber & " has remained un-dispatched since " & GetServiceTicketSTAGEDateTime(PassedMemoNumber,GetServiceTicketCurrentStage(PassedMemoNumber))
			Case "NoACK"
				emailSubject = "Ticket " & PassedMemoNumber & " dispatch not acknowledged"
				emailHeadLineText = "Service ticket #" & PassedMemoNumber & "  has been dispatched but the dispatch has not been acknowledged since " & GetServiceTicketSTAGEDateTime(PassedMemoNumber,GetServiceTicketCurrentStage(PassedMemoNumber))
				txtMessage =  "Service ticket #" & PassedMemoNumber & "  has been dispatched but the dispatch has not been acknowledged since " & GetServiceTicketSTAGEDateTime(PassedMemoNumber,GetServiceTicketCurrentStage(PassedMemoNumber))
			Case "NoOnSite"
				emailSubject = "Ticket " & PassedMemoNumber & "No one has been on site"
				emailHeadLineText = "Service ticket #" & PassedMemoNumber & "  no one has been on site since this ticket was originally submitted at " & GetServiceTicketOpenDateTime(PassedMemoNumber)
				txtMessage =  "Service ticket #" & PassedMemoNumber & "  no one has been on site. Originally submitted " & GetServiceTicketOpenDateTime(PassedMemoNumber)
			Case "OpenTooLong"	
				emailSubject = "Ticket " & PassedMemoNumber & " open too long"
				emailHeadLineText = "Service ticket #" & PassedMemoNumber & "  has been open too long. This ticket was originally submitted at " & GetServiceTicketOpenDateTime(PassedMemoNumber)
				txtMessage =  "Service ticket #" & PassedMemoNumber & "  has been open too long. This ticket was originally submitted at " & GetServiceTicketOpenDateTime(PassedMemoNumber)
			Case "RedispatchTooLong"
				emailSubject = "Service ticket #" & PassedMemoNumber & " Awaiting Redispatch"
				emailHeadLineText = "Service ticket #" & PassedMemoNumber & " has been awaiting redispatch since " & AwaitingRedispatchSince_DateTime(PassedMemoNumber)
				txtMessage =  "Service ticket #" & PassedMemoNumber & " has been awaiting redispatch since " & AwaitingRedispatchSince_DateTime(PassedMemoNumber)
			Case "AnyStage"
				emailSubject = "Service ticket #" & PassedMemoNumber & " idle too long"
				emailHeadLineText = "Service ticket #" & PassedMemoNumber & " has been idle in its current stage since " & GetServiceTicketSTAGEDateTime(PassedMemoNumber,GetServiceTicketCurrentStage(PassedMemoNumber))
				txtMessage =  "Service ticket #" & PassedMemoNumber & " has been idle in its current stage since " & GetServiceTicketSTAGEDateTime(PassedMemoNumber,GetServiceTicketCurrentStage(PassedMemoNumber))
			Case "Declined"
				emailSubject = "Dispatch declined by " & GetUserDisplayNameByUserNo(MostRecentDispatchDeclineByTicket(PassedMemoNumber))
				emailHeadLineText = "Service ticket #" & PassedMemoNumber & ". The dispatch was declined by " & GetUserDisplayNameByUserNo(MostRecentDispatchDeclineByTicket(PassedMemoNumber))
				txtMessage =  "Service ticket #" & PassedMemoNumber & ". The dispatch was declined by " & GetUserDisplayNameByUserNo(MostRecentDispatchDeclineByTicket(PassedMemoNumber))& " "
			Case "ContainsServiceNotes"
				emailSubject = "Service ticket #" & PassedMemoNumber & " field tech notes from " & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(PassedMemoNumber))
				txtMessage = "Service ticket #" & PassedMemoNumber & " contains technician notes."
		End Select
		
		%>
		<!--#include file="../../emails/service_realtime_alert.asp"-->	
		<%
		
		'Now Send the emails
		'Got all the addresses so now break them up
		Send_To_Array = Split(Send_To,";")
	
		For x = 0 to Ubound(Send_To_Array) -1
			Send_To = Send_To_Array(x)
			Response.Write("<font color='green'><b>Sending alerts to Send_To: " & Send_To & " for ticket#: " & PassedMemoNumber & "</b></font><br>")
		

			SendMail "mailsender@" & maildomain,Send_To,emailSubject,emailBody,GetTerm("Service"),GetTerm("Service") & " Alert","MDS Insight"

		'	Response.Write("SendMailmailsender@" & maildomain & "," & Send_To & "," & emailSubject & "," & emailBody & "," & GetTerm("Service")& ","& GetTerm("Service") & " Alert<br>")
		'	Response.Write(AlertType & "<br>")
			CreateAuditLogEntry "Service Alert Sent","Service Alert Sent","Minor",0,"Service Realtime Alert Sent to " & Send_To & " for ticket #: " & PassedMemoNumber & " - " & emailSubject 
			
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
	
	
		txtSubject = "Service Alert"
		
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
		
		CreateAuditLogEntry "Service Alert Sent","Service Alert Sent","Minor",0,"Service Realtime Alert Sent to " & TEXT_TO & " for ticket #: " & PassedMemoNumber & " - " & emailSubject 	
		
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

	If AlertType <> "Declined" Then 
		Call WriteAlertSentRecord (rsAlert("InternalAlertRecNumber"),PassedMemoNumber)
	Else
		If GetServiceTicketCurrentStage(PassedMemoNumber) = "Dispatch Declined"  Then
			'Get the detail record number of the last decline
			RecNumHolder = MostRecentDispatchDeclineRecordNumberByTicket(PassedMemoNumber)
			
			Call WriteAlertSentRecordDetailLevel (rsAlert("InternalAlertRecNumber"),PassedMemoNumber,RecNumHolder )
		End If
	End If
	

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


Sub WriteAlertSentRecord (PassedAlertRecIdentifier,PassedMemoNumber)

		Set cnnAlertsSent = Server.CreateObject("ADODB.Connection")
		cnnAlertsSent.open (MUV_Read("ClientCnnString"))
		Set rsSCAlertsSent = Server.CreateObject("ADODB.Recordset")
		rsSCAlertsSent.CursorLocation = 3 

		rsSCAlertsSent = "Insert Into SC_AlertsSent (Alert_InternalAlertRecNumber,ServiceTicketNumberIfApplicable) Values (" &  PassedAlertRecIdentifier & ",'" & PassedMemoNumber & "')"

		Set rsSCAlertsSent = cnnAlertsSent.Execute(rsSCAlertsSent)


		Set rsSCAlertsSent = Nothing
		cnnAlertsSent.Close
		Set cnnAlertsSent = Nothing

End Sub

Sub WriteAlertSentRecordDetailLevel (PassedAlertRecIdentifier,PassedMemoNumber,passedDetailRecNumber)

		Set cnnAlertsSent = Server.CreateObject("ADODB.Connection")
		cnnAlertsSent.open (MUV_Read("ClientCnnString"))
		Set rsSCAlertsSent = Server.CreateObject("ADODB.Recordset")
		rsSCAlertsSent.CursorLocation = 3 

		rsSCAlertsSent = "Insert Into SC_AlertsSent (Alert_InternalAlertRecNumber,ServiceTicketNumberIfApplicable,ServiceMemoDetailRecIfApplicable) Values (" 
		rsSCAlertsSent = rsSCAlertsSent &  PassedAlertRecIdentifier & "," & PassedMemoNumber & "," & passedDetailRecNumber & ")"

		Set rsSCAlertsSent = cnnAlertsSent.Execute(rsSCAlertsSent)

		Set rsSCAlertsSent = Nothing
		cnnAlertsSent.Close
		Set cnnAlertsSent = Nothing

End Sub


Function AlertSent(passedInternalAlertRecNumber,passedMemoNumberSent)

		resultAlertSent = False
		
		Set cnnAlertSent  = Server.CreateObject("ADODB.Connection")
		cnnAlertSent.open (MUV_Read("ClientCnnString"))
		Set rsAlertSent  = Server.CreateObject("ADODB.Recordset")
		rsAlertSent.CursorLocation = 3 

		SQLAlertSent  = "Select * from SC_AlertsSent  Where Alert_InternalAlertRecNumber = " & passedInternalAlertRecNumber & " AND ServiceTicketNumberIfApplicable = '" & passedMemoNumberSent & "'"

		Set rsAlertSent = cnnAlertSent.Execute(SQLAlertSent)
		If not rsAlertSent.Eof Then resultAlertSent = True

		Set rsAlertSent = Nothing
		cnnAlertSent.Close
		Set cnnAlertSent  = Nothing
		
		AlertSent = resultAlertSent

End Function 

Function AlertSentDetailLevel(passedInternalAlertRecNumber,passedMemoNumberSent,passedDetailRecNumber)

		resultAlertSentDetailLevel = False
		
		Set cnnAlertSentDetailLevel  = Server.CreateObject("ADODB.Connection")
		cnnAlertSentDetailLevel.open (MUV_Read("ClientCnnString"))
		Set rsAlertSentDetailLevel  = Server.CreateObject("ADODB.Recordset")
		rsAlertSentDetailLevel.CursorLocation = 3 

		SQLAlertSentDetailLevel  = "Select * from SC_AlertsSent  Where Alert_InternalAlertRecNumber = " & passedInternalAlertRecNumber & " AND "
		SQLAlertSentDetailLevel  = SQLAlertSentDetailLevel  & " ServiceTicketNumberIfApplicable = '" & passedMemoNumberSent & "' AND"
		SQLAlertSentDetailLevel  = SQLAlertSentDetailLevel  & " ServiceMemoDetailRecIfApplicable = " & passedDetailRecNumber

		Set rsAlertSentDetailLevel = cnnAlertSentDetailLevel.Execute(SQLAlertSentDetailLevel)
		If not rsAlertSentDetailLevel.Eof Then resultAlertSentDetailLevel = True

		Set rsAlertSentDetailLevel = Nothing
		cnnAlertSentDetailLevel.Close
		Set cnnAlertSentDetailLevel  = Nothing
		
		AlertSentDetailLevel = resultAlertSentDetailLevel

End Function 

 
%>