<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/mailDirectLaunch.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Routing.asp"-->
<script type="text/javascript">
    function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>

<meta http-equiv="refresh" content="300">

<%
'Response.End
'Delivery Board Nag Message processing page
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check both global settings and user over-rides to see if nag messages need to be sent out
'Usage = "http://{xxx}.{domain}.com/directLaunch/nags/DeliveryBoardNagCheck.asp?runlevel=run_now
Server.ScriptTimeout = 2500

Dim EntryThread

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)

If Request.QueryString("runlevel") <> "run_now" then
	Response.Write("Improper usage, no run level was specified in the query string")	
	response.end
End IF 

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

'This single page loops through and handles alerts for ALL clients
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
	
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3


If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF
	
		ClientKey = TopRecordset.Fields("clientkey")
		PROCESS_NAGS = True
	
		'To begin with, see if this client uses the routing module 
		'If they don't then don't bother checking for Nags
		If TopRecordset.Fields("routingModule") = "Enabled" Then
	
			'The IF statement below makes sure that when run from DEV it only deos client keys with a d
			'and when run from LIVE it only does client keys without a d
			'Pretty smart, huh
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
			or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then 
	
				Call SetClientCnnString
				
				Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
				
			
				
				'**************************************************************
				'Get next Entry Thread for use in the SC_AuditLogDLaunch table
				On Error Goto 0
				Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
				cnnAuditLog.open MUV_READ("ClientCnnString") 
				Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
				rsAuditLog.CursorLocation = 3 
				Set rsAuditLog = cnnAuditLog.Execute("Select TOP 1 * from SC_AuditLogDLaunch order by EntryThread desc")
				If Not rsAuditLog.EOF Then 
					If IsNull(rsAuditLog("EntryThread")) Then EntryThread =1 Else EntryThread = rsAuditLog("EntryThread") + 1
				Else
					EntryThread = 1
				End If
				set rsAuditLog = nothing
				cnnAuditLog.close
				set cnnAuditLog = nothing

				CreateAuditLogEntry "Routing module Nag check","Routing module Nag check","Minor",0,"Routing module Nag check ran."					

				WriteResponse "<font color='purple' size='24'>Start processing " & ClientKey  & "</font><br>"

				WriteResponse ("Setting Stopmail vars for " & ClientKey  & "<br>")
				
				If Session("ClientCnnString") <> ""Then
					'SEE IF MAIL IS ON OR OFF
					SQLtoggle = "Select STOPALLEMAIL from " & MUV_Read("SQL_Owner") & ".Settings_Global"
					
					WriteResponse (SQLtoggle & "<br>")
					Set cnntoggle = Server.CreateObject("ADODB.Connection")
					cnntoggle.open (Session("ClientCnnString"))
					Set rstoggle = Server.CreateObject("ADODB.Recordset")
					rstoggle.CursorLocation = 3 
					Set rstoggle = cnntoggle.Execute(SQLtoggle)
					If rstoggle.Eof Then 
						Session("MAILOFF") = 1 ' If eof then set email to off
						WriteResponse ("<font color='red'>MAIL OFF</font><br>")
					Else
						Session("MAILOFF") = rstoggle("STOPALLEMAIL")
						If Session("MAILOFF") = 1 Then
							WriteResponse ("<font color='red'>MAIL OFF<br>-</font>")				
						Else
						WriteResponse ("<font color='green'>MAIL ON<br></font>")				
						End IF
					End If
					set rstoggle = Nothing
					cnntoggle.close
					set cnntoggle = Nothing
				Else
					Session("MAILOFF") = 0 ' There was no valid ccn string, so assume it is on
				End If
				
				If MUV_READ("cnnStatus") = "OK" Then ' else it loops
				
					
					'*****************************************************************
					'Start by checking the company calendar & see if we are open today
					'*****************************************************************
					If WeekDay(Now()) = 1 or Weekday(Now()) = 7 Then
						'It is Saturday or Sunday
						WriteResponse ("Today is " & WeekDayName(WeekDay(Now())) & " Nags are not processed today<br>")					
						PROCESS_NAGS = False
						If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then
							WriteResponse ("But this is running on dev, so we let it process on weekends.<br>")					
							PROCESS_NAGS = True
						End IF
					End if
					
					
					Set cnnAlerts = Server.CreateObject("ADODB.Connection")
					cnnAlerts.open (MUV_Read("ClientCnnString"))
					
					'See if the master nag message flag is on or off
					If NAGMasterOn() <> True Then 'Function returns true or false
						PROCESS_NAGS = False
						WriteResponse ("NAG Master flag is turned off for this client, NAGs will not be checked.<br>")	
						CreateAuditLogEntry "Routing module Nag check","Routing module Nag check","Minor",0,"NAG Master flag is turned off for this client, NAGs will not be checked."
					End IF						

					If PROCESS_NAGS = True Then
					
						Set cnnCompanyCalendar = Server.CreateObject("ADODB.Connection")
						cnnCompanyCalendar.open MUV_READ("ClientCnnString") 
						Set rsCompanyCalendar = Server.CreateObject("ADODB.Recordset")
						rsCompanyCalendar.CursorLocation = 3 
						
						SQLCompanyCalendar = "SELECT * FROM Settings_CompanyCalendar WHERE "
						SQLCompanyCalendar = SQLCompanyCalendar & "YearNum = " &  Year(Now()) & " AND "
						SQLCompanyCalendar = SQLCompanyCalendar & "MonthNum = " &  Month(Now()) & " AND "
						SQLCompanyCalendar = SQLCompanyCalendar & "DayNum = " &  Day(Now())
						
						Set rsCompanyCalendar = cnnCompanyCalendar.Execute(SQLCompanyCalendar)
						
						If Not rsCompanyCalendar.EOF Then 

							If rsCompanyCalendar("OpenClosedCloseEarly") = "Closed" Then 
								WriteResponse ("Skipping, closed today.<br>")
								PROCESS_NAGS = False
							End If
							If rsCompanyCalendar("OpenClosedCloseEarly") = "Close Early" Then
							
								CurrentTime = TZTime()
								CloseEarlyTime = rsCalendar("ClosingTime")
								
								CurrentTime_hour = Hour(CurrentTime)
								CurrentTime_min = Minute(CurrentTime)
								CloseEarlyTime_hour = Hour(CloseEarlyTime)
								CloseEarlyTime_min = Minute(CloseEarlyTime)
								
								'***************Time Difference Calculation ***************************
								If (CurrentTime_hour > CloseEarlyTime_hour) Then
									WriteResponse ("Skipping, company closed early today. Closed at " & FormatDateTime(CloseEarlyTime,3) & " Time now is " & CurrentTime & "<br>")
									PROCESS_NAGS = False
								ElseIf (CurrentTime_hour = CloseEarlyTime_hour) AND (CurrentTime_min >= CloseEarlyTime_min) Then
									WriteResponse ("Skipping, company closed early today. Closed at " & FormatDateTime(CloseEarlyTime,3) & " Time now is " & CurrentTime & "<br>")
									PROCESS_NAGS = False
								End If		
												
							End If

						End If
						
						set rsCompanyCalendar = nothing
						cnnCompanyCalendar.close
						set cnnCompanyCalendar = nothing
					
					End If
					
					'Is it just a normal day but they are not open yet
					If PROCESS_NAGS = True Then
						If cdate(TZTime()) < cdate(BusinessDayStart()) Then
							WriteResponse ("Skipping, the normal business day has not started yet. Business day start is " & FormatDateTime(BusinessDayStart,3) & " Time now is " & TZTime & "<br>")
							PROCESS_NAGS = False
						Else
							WriteResponse ("OK, the normal business day is started. Business day start is " & FormatDateTime(BusinessDayStart,3) & " Time now is " & TZTime & "<br>")
						End If	
					End If
					
					
					'Is it just a normal day but past closing time					
					If PROCESS_NAGS = True Then
						If cdate(TZTime()) >= cdate(BusinessDayEnd()) Then
							WriteResponse ("Skipping, the normal business day has ended. Business day end is " & FormatDateTime(BusinessDayEnd,3) & " Time now is " & TZTime & "<br>")
							PROCESS_NAGS = False
						Else
							WriteResponse ("OK, the normal business day is ongoing. Business day end is " & FormatDateTime(BusinessDayEnd,3) & " Time now is " & TZTime & "<br>")
						End If	
					End If
					
					
					If PROCESS_NAGS = True Then
					
						WriteResponse ("Begin processing Nags<br>")
				
						'Get all of the drivers who have deliveries
						Set cnnAlerts = Server.CreateObject("ADODB.Connection")
						cnnAlerts.open (MUV_Read("ClientCnnString"))
						
						SQLDrivers = "SELECT DISTINCT RT_DeliveryBoard.TruckNumber, tblUsers.userNo, tblUsers.userNextStopNagMessageOverride, "
						SQLDrivers = SQLDrivers & "userNextStopNagMinutes, userNextStopNagIntervalMinutes, userNextStopNagMessageMaxToSendPerStop, "
                        SQLDrivers = SQLDrivers & "userNextStopNagMessageMaxToSendThisDriverPerDay, userNextStopNagMessageSendMethod, userNoActivityNagMessageOverride, userNoActivityNagMinutes, userNoActivityNagIntervalMinutes, "
                        SQLDrivers = SQLDrivers & "userNoActivityNagMessageMaxToSendPerStop, userNoActivityNagMessageMaxToSendPerDriverPerDay, userNoActivityNagMessageSendMethod, userNoActivityNagTimeOfDay "
						SQLDrivers = SQLDrivers & "FROM RT_DeliveryBoard INNER JOIN "
                        SQLDrivers = SQLDrivers & "tblUsers ON tblUsers.userTruckNumber = RT_DeliveryBoard.TruckNumber WHERE " 
                        SQLDrivers = SQLDrivers & "tblUsers.userArchived <> 1 and tblUsers.userEnabled = 1" 
							
						Set rsDrivers = Server.CreateObject("ADODB.Recordset")
						rsDrivers.CursorLocation = 3 
						Set rsDrivers = cnnAlerts.Execute(SQLDrivers)
						
						If not rsDrivers.EOF Then
									
							Do While Not rsDrivers.EOF
							
								If GetUserDisplayNameByUserNo(rsDrivers("userNo")) = "RichTheDriver" Then Response.write("<font color='red'>")
								WriteResponse ("<b>Checking driver: " & GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" Route#:" & rsDrivers("TruckNumber") & "</b><br>") 
								If GetUserDisplayNameByUserNo(rsDrivers("userNo")) = "RichTheDriver" Then Response.write("</font>")
								
								CurrentDriverUserNo = rsDrivers("userNo")
								CurrentTruckNumber = rsDrivers("TruckNumber")
								SendNagMessage = True
								

								'****************************
								' N E X T   S T O P   N A G S 
								'****************************
								'Before we do anything, see if this driver has marked everything in their route (done for day)
								If GetRemainingStopsByUserNo(CurrentDriverUserNo) = 0 Then
									WriteResponse ("Skipping: " & GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" has finished for the day. There are no unmarked deliveries.<br>") 
									SendNagMessage = False
								Else
									WriteResponse (GetUserDisplayNameByUserNo(CurrentDriverUserNo) & " has " & GetRemainingStopsByUserNo(CurrentDriverUserNo) &" stops unmarked.<br>") 
								End If

								
								If SendNagMessage = True Then
									'See if they are turned off at the user level
									If rsDrivers("userNextStopNagMessageOverride") = "No" Then
										WriteResponse ("Skipping: " & GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" - Next Stop Nag messages is set to No at the user level<br>") 
										SendNagMessage = False
									Else
										If rsDrivers("userNextStopNagMessageOverride") = "Use Global" Then
											'Get the values from the global settings
											WriteResponse (GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" - override setting for No Next Stop is " &  rsDrivers("userNextStopNagMessageOverride") & " <br>") 
											SQLGlobal = "SELECT * FROM Settings_Global "
											Set rsGlobal = Server.CreateObject("ADODB.Recordset")
											rsGlobal.CursorLocation = 3 
											Set rsGlobal = cnnAlerts.Execute(SQLGlobal)
											' See if it's turned on in global, if it's not, we are done
											If rsGlobal("NextStopNagMessageONOFF") = 0 Then
												WriteResponse ("Skipping " & GetUserDisplayNameByUserNo(rsDrivers("userNo")) & " - Next stop nag messages are turned off at the global settings level.<br>")
												SendNagMessage = False 
											Else
		 										NextStopNagMinutes = rsGlobal("NextStopNagMinutes")
		 										NextStopNagIntervalMinutes = rsGlobal("NextStopNagIntervalMinutes")
		 										NextStopNagMessageMaxToSendPerStop = rsGlobal("NextStopNagMessageMaxToSendPerStop")
		 										NextStopNagMessageMaxToSendThisDriverPerDay = rsGlobal("NextStopNagMessageMaxToSendPerDriverPerDay")
		 										NextStopNagMessageSendMethod = rsGlobal("NextStopNagMessageSendMethod")
											End If
											Set rsGlobal = Nothing
										Else
											'Get the values from the user settings
	 										WriteResponse (GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" - override setting  for No Next Stop is " &  rsDrivers("userNextStopNagMessageOverride") & " <br>") 
	 										NextStopNagMinutes = rsDrivers("userNextStopNagMinutes")
	 										NextStopNagIntervalMinutes = rsDrivers("userNextStopNagIntervalMinutes")
	 										NextStopNagMessageMaxToSendPerStop = rsDrivers("userNextStopNagMessageMaxToSendPerStop")
	 										NextStopNagMessageMaxToSendThisDriverPerDay = rsDrivers("userNextStopNagMessageMaxToSendThisDriverPerDay")
	 										NextStopNagMessageSendMethod = rsDrivers("userNextStopNagMessageSendMethod")
										End If	
									End If
									
									If SendNagMessage = True Then
										' OK, we got all the settings
										
										' If no deliveries have been marked yet, the driver hasn't started the route
										SQLDeliveryBoard = "SELECT * FROM RT_DeliveryBoard WHERE TruckNumber = '" & CurrentTruckNumber & "' AND DeliveryStatus IS NOT NULL or DeliveryInProgress = 1"
								
										Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
										rsDeliveryBoard.CursorLocation = 3 
										Set rsDeliveryBoard = cnnAlerts.Execute(SQLDeliveryBoard)
	
										If rsDeliveryBoard.Eof Then ' If EOF, they have not started their route yet, nothing is marked
											WriteResponse (GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" has not marked any deliveries yet, this indicates the route has not been started yet, so we can't nag for a next stop <br>") 
											SendNagMessage = False
										Else
											WriteResponse ("OK " & GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" has started the route for the day. At least one delivery is marked.<br>") 
										End If
									
										Set rsDeliveryBoard = Nothing
									End If
									
									If SendNagMessage = True Then
									
										'So now lets see if they have a Next Stop marked or not
										SQLDeliveryBoard = "SELECT * FROM RT_DeliveryBoard WHERE TruckNumber = '" & CurrentTruckNumber & "' AND ManualNextStop = 1"
								
										Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
										rsDeliveryBoard.CursorLocation = 3 
										Set rsDeliveryBoard = cnnAlerts.Execute(SQLDeliveryBoard)

										If NOT rsDeliveryBoard.Eof Then
											WriteResponse (GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" - The Next Stop is properly marked. Nothing to do.<br>") 
											SendNagMessage = False
										End If
										
										Set rsDeliveryBoard = Nothing
									End If
									
									
									If SendNagMessage = True Then
										'If we are here, they are on the road and a next stop is not marked
										
										'See if they are currently in the SC_NagSkipUsers file today
										SQLDeliveryBoard = "SELECT * FROM SC_NagSkipUsers WHERE UserNo = '" & CurrentDriverUserNo & "' AND NagType = 'routingNoNextStop'"
								
										Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
										rsDeliveryBoard.CursorLocation = 3 
										Set rsDeliveryBoard = cnnAlerts.Execute(SQLDeliveryBoard)

										If NOT rsDeliveryBoard.Eof Then
											WriteResponse ("Skipping - " & GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" is in today's NagSkipUsers List.<br>") 
											SendNagMessage = False
										End If
										
										Set rsDeliveryBoard = Nothing
									End If
										
									If SendNagMessage = True Then
										'See if it has been X minutes since the last delivery was marked
										If GetLastInvoiceMarkedDATETIMEByTruckNumber(CurrentTruckNumber) <> "" Then		
											If DateDiff("n",GetLastInvoiceMarkedDATETIMEByTruckNumber(CurrentTruckNumber),Now()) < NextStopNagMinutes Then
												WriteResponse ("Not time yet, only " & DateDiff("n",GetLastInvoiceMarkedDATETIMEByTruckNumber(CurrentTruckNumber),Now()) & " minutes have passed<br>")
												WriteResponse ("The nag setting says to wait for " & NextStopNagMinutes  & " minutes<BR>")
												SendNagMessage = False
											End If
										End If
									End If							
									

									If SendNagMessage = True Then	 		
										'Now check all the Limits
										'Number of Nags per Driver per day
										If GetNumberOfNagMessagesSentOnDate(CurrentDriverUserNo , "routingNoNextStop", Now()) > NextStopNagMessageMaxToSendThisDriverPerDay Then
											WriteResponse ("Skipping: The maximum number of nag messages per day is set to " & NextStopNagMessageMaxToSendThisDriverPerDay & "<BR>")
											WriteResponse (GetNumberOfNagMessagesSentOnDate(CurrentDriverUserNo , "routingNoNextStop", Now()) & " nag messages have been sent today<BR>")
											SendNagMessage = False
										End If	
	
										'Wait x minutes between nags
										If DateDiff("n",GetLastNagSentTime(CurrentDriverUserNo , "routingNoNextStop"),Now()) < NextStopNagIntervalMinutes Then
											MinutesToGo = NextStopNagIntervalMinutes - DateDiff("n",GetLastNagSentTime(CurrentDriverUserNo , "routingNoNextStop"),Now()) 
											WriteResponse ("Skipping: The interval between Nags is set to " & NextStopNagIntervalMinutes & ". The last Nag message was sent " & DateDiff("n",GetLastNagSentTime(CurrentDriverUserNo , "routingNoNextStop"),Now()) & " minutes ago. " & MinutesToGo  & " minutes to go.<br>")
											SendNagMessage = False
										End If
									End If
										
									If SendNagMessage = True Then	 	
										'Number of Nags per Stop
										'We get the time of the last stop & see how many Nags have been send since then
										
										'Account for the fact that the last activity might not be from today
										DateAndTimeForActivityCheck = GetLastDeliveryStatusChangeBYTruck(CurrentTruckNumber)
										If DateDiff("d",DateAndTimeForActivityCheck,Now()) > 0 Then DateAndTimeForActivityCheck = Now()

										If GetNumberOfNagMessagesSentSinceDateTime(CurrentDriverUserNo , "routingNoNextStop", DateAndTimeForActivityCheck)  >= NextStopNagMessageMaxToSendPerStop Then
											WriteResponse ("Skipping: The maximum number of nag messages per stop is set to " & NextStopNagMessageMaxToSendPerStop & "<BR>")
											WriteResponse (GetNumberOfNagMessagesSentSinceDateTime(CurrentDriverUserNo , "routingNoNextStop", GetLastDeliveryStatusChangeBYTruck(CurrentTruckNumber)) & " nag messages have been sent for the current stop<BR>")
											SendNagMessage = False
										Else
											WriteResponse("Last stop was marked at " & GetLastDeliveryStatusChangeBYTruck(CurrentTruckNumber) & "<br>")
											WriteResponse("Max nags per stop is " & NextStopNagMessageMaxToSendPerStop & "<br>")
											WriteResponse("nags sent this stop is " & GetNumberOfNagMessagesSentSinceDateTime(CurrentDriverUserNo , "routingNoNextStop", GetLastDeliveryStatusChangeBYTruck(CurrentTruckNumber)) & "<br>")	
										End If		
									End If
									
									
									'*****************************************************************
									' Eureeka! If you made it this far we need to send the Nag message
									'*****************************************************************
									If SendNagMessage = True Then
										WriteResponse("     S E N D I N G       N A G       M E S S A G E<br>")
										SendNag CurrentDriverUserNo,CurrentTruckNumber,"routingNoNextStop",NextStopNagMessageSendMethod,0
									End If

								End If

'***********************************************************************************************************************************************************************************	
								'*******************************
								' N O  A C T I V I T Y   N A G S 
								'*******************************
								
								'Start the process over again
								SendNagMessage = True
								
								'Before we do anything, see if this driver has marked everything in their route (done for day)
								If GetRemainingStopsByUserNo(CurrentDriverUserNo) = 0 Then
									WriteResponse ("Skipping: " & GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" has finished for the day. There are no unmarked deliveries.<br>") 
									SendNagMessage = False
								Else
									WriteResponse (GetUserDisplayNameByUserNo(CurrentDriverUserNo) & " has " & GetRemainingStopsByUserNo(CurrentDriverUserNo) &" stops unmarked.<br>") 
								End If
									
								If SendNagMessage = True Then
									'See if they are turned off at the user level
									If rsDrivers("userNoActivityNagMessageOverride") = "No" Then
										WriteResponse ("Skipping: " & GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" - No Activity Nag messages is set to No at the user level<br>") 
										SendNagMessage = False
									Else
										If rsDrivers("userNoActivityNagMessageOverride") = "Use Global" Then
											'Get the values from the global settings
											WriteResponse (GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" - override setting for No Activity is " &  rsDrivers("userNoActivityNagMessageOverride") & " <br>") 
											SQLGlobal = "SELECT * FROM Settings_Global "
											Set rsGlobal = Server.CreateObject("ADODB.Recordset")
											rsGlobal.CursorLocation = 3 
											Set rsGlobal = cnnAlerts.Execute(SQLGlobal)
											' See if it's turned on in global, if it's not, we are done
											If rsGlobal("NoActivityNagMessageONOFF") = 0 Then
												WriteResponse ("Skipping " & GetUserDisplayNameByUserNo(rsDrivers("userNo")) & " - No Activity nag messages are turned off at the global settings level.<br>")
												SendNagMessage = False 
											Else
		 										NoActivityNagMinutes = rsGlobal("NoActivityNagMinutes")
		 										NoActivityNagIntervalMinutes = rsGlobal("NoActivityNagIntervalMinutes")
		 										NoActivityNagMessageMaxToSendPerStop = rsGlobal("NoActivityNagMessageMaxToSendPerStop")
		 										NoActivityNagMessageMaxToSendPerDriverPerDay = rsGlobal("NoActivityNagMessageMaxToSendPerDriverPerDay")
		 										NoActivityNagMessageSendMethod = rsGlobal("NoActivityNagMessageSendMethod")
		 										NoActivityNagTimeOfDay = FormatDateTime(rsGlobal("NoActivityNagTimeOfDay"),3)
											End If
											Set rsGlobal = Nothing
										Else
											'Get the values from the user settings
	 										WriteResponse (GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" - override setting for No Activity is " &  rsDrivers("userNoActivityNagMessageOverride") & " <br>") 
		 										NoActivityNagMinutes = rsDrivers("userNoActivityNagMinutes")
		 										NoActivityNagIntervalMinutes = rsDrivers("userNoActivityNagIntervalMinutes")
		 										NoActivityNagMessageMaxToSendPerStop = rsDrivers("userNoActivityNagMessageMaxToSendPerStop")
		 										NoActivityNagMessageMaxToSendPerDriverPerDay = rsDrivers("userNoActivityNagMessageMaxToSendPerDriverPerDay")
		 										NoActivityNagMessageSendMethod = rsDrivers("userNoActivityNagMessageSendMethod")
		 										NoActivityNagTimeOfDay = FormatDateTime(rsDrivers("userNoActivityNagTimeOfDay"),3)
										End If	
									End If
									
									
									If SendNagMessage = True Then
										'Just see if they need to be skipped before we do anything else
										
										'See if they are currently in the SC_NagSkipUsers file today
										SQLDeliveryBoard = "SELECT * FROM SC_NagSkipUsers WHERE UserNo = '" & CurrentDriverUserNo & "' AND NagType = 'routingNoActivity'"
								
										Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
										rsDeliveryBoard.CursorLocation = 3 
										Set rsDeliveryBoard = cnnAlerts.Execute(SQLDeliveryBoard)

										If NOT rsDeliveryBoard.Eof Then
											WriteResponse ("Skipping - " & GetUserDisplayNameByUserNo(rsDrivers("userNo")) &" is in today's NagSkipUsers List.<br>") 
											SendNagMessage = False
										End If
										
										Set rsDeliveryBoard = Nothing
									End If

									
									If SendNagMessage = True Then
										
										' If no deliveries have been marked yet, and no manual next stops have been set, there is No activity yet
										' Get the last date & time that 'something' happenned for this driver
										SQLDeliveryBoard = ""
										SQLDeliveryBoard = "SELECT MAX(LastActivityDateTime) AS LastActivityDateTime FROM "
										SQLDeliveryBoard = SQLDeliveryBoard &  "(SELECT LastDeliveryStatusChange As LastActivityDateTime FROM RT_DeliveryBoard WHERE TruckNumber = '" & CurrentTruckNumber & "' "
										SQLDeliveryBoard = SQLDeliveryBoard & " UNION "
										SQLDeliveryBoard = SQLDeliveryBoard & "SELECT ManualNextStopChanged As LastActivityDateTime FROM RT_DeliveryBoard WHERE TruckNumber = '" & CurrentTruckNumber & "' "
										SQLDeliveryBoard = SQLDeliveryBoard & ") AS t1"
										
										Set rsDeliveryBoard = Server.CreateObject("ADODB.Recordset")
										rsDeliveryBoard.CursorLocation = 3 

										Set rsDeliveryBoard = cnnAlerts.Execute(SQLDeliveryBoard)

										LastActivityDateTime = ""
	
										If NOT rsDeliveryBoard.Eof Then 
											LastActivityDateTime = rsDeliveryBoard("LastActivityDateTime")
										End If
										
										If IsNull(LastActivityDateTime) Then LastActivityDateTime = ""
										
										If LastActivityDateTime <> "" Then
											WriteResponse (GetUserDisplayNameByUserNo(CurrentDriverUserNo) &" last performed an action at " & LastActivityDateTime & " <br>") 
										Else
											WriteResponse (GetUserDisplayNameByUserNo(CurrentDriverUserNo) &" has not performed any actions today.<br>") 
										End If
									
										Set rsDeliveryBoard = Nothing
									End If
									
									
									If SendNagMessage = True Then
										'See if it has been X minutes since the last actitivty
										If LastActivityDateTime <> "" Then		
											If DateDiff("n",LastActivityDateTime,Now()) < NoActivityNagMinutes Then
												WriteResponse ("Skipping Not time yet, only " & DateDiff("n",LastActivityDateTime,Now()) & " minutes have passed<br>")
												WriteResponse ("The nag setting says to wait for " & NoActivityNagMinutes & " minutes<BR>")
												SendNagMessage = False
											Else
												WriteResponse ("It has been " & DateDiff("n",LastActivityDateTime,Now()) & " minutes since there has been any activity<br>")
												WriteResponse ("The nag setting says to wait for " & NoActivityNagMinutes & " minutes<BR>")
												IdleTime = DateDiff("n",LastActivityDateTime,Now())
											End If
										Else
											IdleTime = 0
										End If
									End If	

									
									If SendNagMessage = True Then
										'If there has been no activity from the driver yet today - check against the time setting
										If LastActivityDateTime = "" Then
											If cdate(TZTime) > cdate(NoActivityNagTimeOfDay)  Then
												WriteResponse ("There has been no actitivy from this driver yet today<br>")
												WriteResponse ("Their start time setting is set to " & NoActivityNagTimeOfDay & " and it is now " & TZTime & ". Better send a nag<BR>")
											Else
												WriteResponse ("There has been no actitivy from this driver yet today<br>")
												WriteResponse ("Skipping Their start time setting is set to " & NoActivityNagTimeOfDay & " and it is now " & TZTime & " not time to start yet<BR>")
												SendNagMessage = False
											End If
										End If
									End If
										

									If SendNagMessage = True Then	
										'Now check all the Limits
										'Number of Nags per Driver per day
										WriteResponse ("Check for The maximum number of nag messages per day<BR>") 
										If GetNumberOfNagMessagesSentOnDate(CurrentDriverUserNo , "routingNoActivity", Now()) > NoActivityNagMessageMaxToSendPerDriverPerDay Then
											WriteResponse ("Skipping: The maximum number of nag messages per day is set to " & NoActivityNagMessageMaxToSendPerDriverPerDay & "<BR>")
											WriteResponse (GetNumberOfNagMessagesSentOnDate(CurrentDriverUserNo , "routingNoActivity", Now()) & " nag messages have been sent today<BR>")
											SendNagMessage = False
										Else
											WriteResponse ("Check for The maximum number of nag messages per day is OK<BR>")
										End If	
									End If
									
									If SendNagMessage = True Then
										'Wait x minutes between nags
										WriteResponse ("Check for The interval between Nag messages<BR>")
										If DateDiff("n",GetLastNagSentTime(CurrentDriverUserNo , "routingNoActivity"),Now()) < NoActivityNagIntervalMinutes Then
											MinutesToGo = NoActivityNagIntervalMinutes - DateDiff("n",GetLastNagSentTime(CurrentDriverUserNo , "routingNoActivity"),Now())
											WriteResponse ("Skipping: The interval between Nags is set to " & NoActivityNagIntervalMinutes & ". The last Nag message was sent " & DateDiff("n",GetLastNagSentTime(CurrentDriverUserNo , "routingNoActivity"),Now()) & " minutes ago. " & MinutesToGo  & " minutes to go.<br>")
											SendNagMessage = False
										Else
											WriteResponse ("Check for The interval between Nag messages is OK<BR>")
										End If
									End If											 	
									
									If SendNagMessage = True Then
										'Number of Nags per No Activity
										'We get the time of the last stop & see how many Nags have been send since then
										
										'Account for the fact that the last activity might not be from today
										If LastActivityDateTime <> "" Then 
											DateAndTimeForActivityCheck = LastActivityDateTime
											WriteResponse ("DateAndTimeForActivityCheck :" & DateAndTimeForActivityCheck  & "<br>")
											WriteResponse ("LastActivityDateTime:" & LastActivityDateTime & "<br>")										
											If DateDiff("d",DateAndTimeForActivityCheck,Now()) > 0 Then DateAndTimeForActivityCheck = Now()
										Else
											DateAndTimeForActivityCheck = Now()
										End If
										
										WriteResponse("Date Time used to check maximum number of nag messages per No Activity events is: " & DateAndTimeForActivityCheck & "<br>")	
										WriteResponse("LastActivityDateTime is: " & LastActivityDateTime & "<br>")	
										
										If GetNumberOfNagMessagesSentSinceDateTime(CurrentDriverUserNo , "routingNoActivity", DateAndTimeForActivityCheck) >= NoActivityNagMessageMaxToSendPerStop Then
											WriteResponse ("Skipping: The maximum number of nag messages per No Activity events is set to " & NoActivityNagMessageMaxToSendPerStop & "<BR>")
											WriteResponse (GetNumberOfNagMessagesSentSinceDateTime(CurrentDriverUserNo , "routingNoActivity", LastActivityDateTime) & " nag messages have been sent for the current stop<BR>")
											SendNagMessage = False
										Else
											If LastActivityDateTime = "" Then
												WriteResponse("This diver has no activity yet today<br>")											
											Else
												WriteResponse("Last activity was marked at " & LastActivityDateTime & "<br>")
											End If
											WriteResponse("Max nags per no activity event is " & NoActivityNagMessageMaxToSendPerStop & "<br>")
											WriteResponse("nags sent this no activity event is " & GetNumberOfNagMessagesSentSinceDateTime(CurrentDriverUserNo , "routingNoActivity", DateAndTimeForActivityCheck ) & "<br>")	
										End If		
									End If
									
									'*****************************************************************
									' Eureeka! If you made it this far we need to send the Nag message
									'*****************************************************************
									If SendNagMessage = True Then
										WriteResponse("     S E N D I N G       N A G       M E S S A G E<br>")
										SendNag CurrentDriverUserNo,CurrentTruckNumber,"routingNoActivity",NoActivityNagMessageSendMethod ,IdleTime
									End If
									
								End If
									
'***********************************************************************************************************************************************************************************									
			
							rsDrivers.Movenext
						Loop
			
					End If
						
					Set rsDrivers = Nothing
					cnnAlerts.Close
					Set cnnALerts = Nothing
									
					WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				End If ' From PROCESS_NAAGS
			
			End If
			
		End If	
		
	Else ' is the routing module enabled
		WriteResponse ("Skipping the client " & ClientKey & " because the routing module is not enabled.<BR>")
	End If ' is the routing module enabled
	
	TopRecordset.movenext
	
	Loop
	
	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

'Response.write("<script type=""text/javascript"">closeme();</script>")	

'*************************
'*************************
'Subs and funcs begin here

Sub SendNag  (passedDriverUserNo,PassedCurrentTruckNumber,passedAlertType,passedSendMethod,passedIdleTime)


	WriteResponse ("**********GOT TO SEND ALERT()**************************<br>")
	Send_To = ""

	If Instr(UCASE(passedSendMethod),"TEXT") <> 0  Then
	
		'Get the text number of the driver
		Send_To = getUserCellNumber(passedDriverUserNo)
		
		'Only do this if we have their cell #
		If Send_To <> "" Then
	
			WriteResponse ("**********FOUND TEXT RECIPIENTS TO SEND TO**************************<br>")

			Select Case passedAlertType
				Case "routingNoNextStop"
					txtSubject = "Set Next Stop"
					RandomLoginValue = GetRandomURLString()
					txtMessage = "Please mark your next stop. You can use this link which will expire in 30 minutes. "
					txtMessage = txtMessage & Server.URLEncode(baseURL & "ql_text.asp?c=" & ClientKey & "&r=" & RandomLoginValue)
				Case "routingNoActivity"
					RandomLoginValue = ""
					txtSubject = "No Activity"
					If passedIdleTime = 0 Then
						txtMessage = "You have not marked any delivery information today. Please remember to mark deliveries and next stops."
					Else
						txtMessage = "You have not marked any delivery information for " & passedIdleTime & " minutes. Please remember to mark deliveries and next stops."
					End If
			End Select

	
			Send_To = Replace(Send_To,"-","") ' EZ Texting doesn't like dashes
	
		
			'*****Text numbers don't get split into an array, the php takes multiple #'s seprated by commas	
			
			If Right(Send_To,1) = "," Then Send_To = Left(Send_To,Len(Send_To)-1)
	
			TEXT_TO = Send_To
		
			' If this is running on dev, send to Rich's text number
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then
				WriteResponse ("Running on dev - changing text number to 6099294430<br>")					
				TEXT_TO = "6099294430"
			Else
				'Only copy rich twice a week, I am getting too many texts
				If Weekday(Now()) = 2 or WeekDay(Now()) = 6 Then TEXT_TO = TEXT_TO & ",6099294430"
			End IF
	
			txtMessage = txtMessage & "(" & GetUserDisplayNameByUserNo(passedDriverUserNo) & ")"' So rich knows who it is for

			WriteResponse ("<font color='green'><b>Sending text to: "& TEXT_TO & "</b></font><br>")
		
			CreateAuditLogEntry "Nag Message Sent","Nag Message Sent","Minor",0,"Nag message texted to " & GetUserDisplayNameByUserNo(passedDriverUserNo) & " at " & TEXT_TO & " Subject: - " & txtSubject 	
		
			WriteResponse ("POST TO: " & BaseURL & "inc/sendtext_post.php<br>")
	
			str_data="txtSubject=" & txtSubject  & "&txtMessage=" & txtMessage & "&txtTEXT_TO=" & TEXT_TO
	
			str_data = str_data & "&txtu1=" & EzTextingUserID() & "&txtu2=" & EzTextingPassword()
		
			str_data = str_data & "&txtCountry=" & GetCompanyCountry()
		
			WriteResponse ("str_data= " & str_data & "<br>")
	
			Set obj_post=Server.CreateObject("Msxml2.SERVERXMLHTTP")
			obj_post.Open "POST", BaseURL & "inc/sendtext_post.php",False
			obj_post.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			obj_post.Send str_data
		
			WriteResponse ("obj_post.responseText: " & obj_post.responseText & "<br>")
	
		End If
	
	End If

	Call WriteNagSentRecord (passedDriverUserNo,PassedCurrentTruckNumber,passedAlertType,RandomLoginValue)

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


Sub WriteNagSentRecord (passedDriverUserNo,PassedCurrentTruckNumber,passedAlertType,passedRandomLoginValue)

		Set cnnNagSent = Server.CreateObject("ADODB.Connection")
		cnnNagSent.open (MUV_Read("ClientCnnString"))
		Set rsNagSent = Server.CreateObject("ADODB.Recordset")
		rsNagSent.CursorLocation = 3 

		If passedDriverUserNo <> "" Then UserNameSentToIfApplicable = GetUserDisplayNameByUserNo(passedDriverUserNo)

		If passedAlertType = "routingNoNextStop" Then	
			rsNagSent = "Insert Into SC_NagsSent ( NagType, UserNoSentToIfApplicable,RandomLoginValueIfApplicable, UserNameSentToIfApplicable) "
			rsNagSent = rsNagSent & "Values ('" &  passedAlertType & "'," & passedDriverUserNo & ",'" & passedRandomLoginValue & "','" & UserNameSentToIfApplicable  & "')"
		End If

		If passedAlertType = "routingNoActivity" Then	
			rsNagSent = "Insert Into SC_NagsSent ( NagType, UserNoSentToIfApplicable, UserNameSentToIfApplicable ) "
			rsNagSent = rsNagSent & "Values ('" &  passedAlertType & "'," & passedDriverUserNo & ",'" & UserNameSentToIfApplicable & "')"
		End If

		Set rsNagSent = cnnNagSent.Execute(rsNagSent)

		Set rsNagSent = Nothing
		cnnNagSent.Close
		Set cnnNagSent = Nothing

End Sub

Sub WriteResponse(passedLogEntry)

	response.write(Now() & "&nbsp;&nbsp;&nbsp;" & passedLogEntry)
	
	passedLogEntry = Replace(passedLogEntry,"'","''")
	
	SQL = "INSERT INTO SC_AuditLogDLaunch (EntryThread, DirectLaunchName, DirectLaunchFile, LogEntry)"
	SQL = SQL &  " VALUES (" & EntryThread & ""
	SQL = SQL & ",'Routing Module Nag Check'"
	SQL = SQL & ",'/directlaunch/nags/RoutingModuleNagCheck.asp'"
	SQL = SQL & ",'"  & passedLogEntry & "'"
	SQL = SQL & ")"
	
	Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
	cnnAuditLog.open MUV_READ("ClientCnnString") 
	Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
	rsAuditLog.CursorLocation = 3 
	
	Set rsAuditLog = cnnAuditLog.Execute(SQL)

	set rsAuditLog = nothing
	cnnAuditLog.close
	set cnnAuditLog = nothing

End Sub

function RandomString()

    Randomize()

    dim CharacterSetArray
    CharacterSetArray = Array(_
        Array(7, "abcdefghijklmnopqrstuvwxyz"), _
        Array(1, "0123456789") _
    )

    dim i
    dim j
    dim Count
    dim Chars
    dim Index
    dim Temp

    for i = 0 to UBound(CharacterSetArray)

        Count = CharacterSetArray(i)(0)
        Chars = CharacterSetArray(i)(1)

        for j = 1 to Count

            Index = Int(Rnd() * Len(Chars)) + 1
            Temp = Temp & Mid(Chars, Index, 1)

        next

    next

    dim TempCopy

    do until Len(Temp) = 0

        Index = Int(Rnd() * Len(Temp)) + 1
        TempCopy = TempCopy & Mid(Temp, Index, 1)
        Temp = Mid(Temp, 1, Index - 1) & Mid(Temp, Index + 1)

    loop

    RandomString = TempCopy

end function

Function GetRandomURLString()

	StringOK = False
	
	Do 
	
		resultGetRandomURLString = RandomString 
	
		'Make sure it was not used in the last week
		SQLGetRandomURLString = "SELECT * FROM SC_NagsSent WHERE RandomLoginValueIfApplicable= '" & resultGetRandomURLString & "' AND RecordCreationDateTime < '" & DateAdd("d",7,Now()) & "'"
	
		Set rsGetRandomURLString = Server.CreateObject("ADODB.Recordset")
		rsGetRandomURLString.CursorLocation = 3 
		Set rsGetRandomURLString = cnnAlerts.Execute(SQLGetRandomURLString)
	
		If rsGetRandomURLString.Eof Then StringOK = True
		
	Loop Until StringOK = True	
		
	Set rsGetRandomURLString = Nothing
	
	GetRandomURLString = resultGetRandomURLString 

End Function



%>