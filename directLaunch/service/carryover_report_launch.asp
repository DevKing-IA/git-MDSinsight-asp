<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/mailDirectLaunch.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Users.asp"-->
<script type="text/javascript">
    function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
 

<%
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page
'Usage = "http://{xxx}.{domain}.com/directLaunch/service/carryover_report_launch.asp?runlevel=run_now
Server.ScriptTimeout = 25000

Dim EntryThread
Dim SendToPrimary, SendToSecondary, SendToTeams

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)

If Request.QueryString("runlevel") <> "run_now" then
	Response.Write("Improper usage, no run level was specified in the query string")	
	response.end
End IF 


'New parameter added, FORCE, tells it to run
'regadless of the scheduled time
If Request.QueryString("force") = "run_now" Then
	dummy = MUV_Write("Force",1)
End If


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
	
		'To begin with, see if this client uses the Service module 
		'If they don't then don't bother running for then
		
		Response.Write("Seeing if we need to run this for " & ClientKey & "<br>")
		
		If TopRecordset.Fields("serviceModule") = "Enabled" Then
	
			'The IF statement below makes sure that when run from DEV it only deos client keys with a d
			'and when run from LIVE it only does client keys without a d
			'Pretty smart, huh
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
			OR (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then 
												
				Call SetClientCnnString
				
				Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
				
				'**********************************************************
				' Now see if the service carry over report is on or off
				'**********************************************************				
				'This is here so we only open it once for the whole page
				Set cnn_Settings_FieldService = Server.CreateObject("ADODB.Connection")
				cnn_Settings_FieldService.open (MUV_READ("ClientCnnString"))
				Set rs_Settings_FieldService = Server.CreateObject("ADODB.Recordset")
				rs_Settings_FieldService.CursorLocation = 3 
				SQL_Settings_FieldService = "SELECT * FROM Settings_FieldService"
				Set rs_Settings_FieldService = cnn_Settings_FieldService.Execute(SQL_Settings_FieldService)
				If not rs_Settings_FieldService.EOF Then
					ServiceTicketCarryoverReportOnOff = rs_Settings_FieldService("ServiceTicketCarryoverReportOnOff")
					SendToPrimary = rs_Settings_FieldService("ServiceTicketCarryoverReportToPrimarySalesman")
					SendToSecondary = rs_Settings_FieldService("ServiceTicketCarryoverReportToSecondarySalesman")
					ServiceTicketCarryoverReportTeamIntRecIDs = rs_Settings_FieldService("ServiceTicketCarryoverReportTeamIntRecIDs")
				Else
					ServiceTicketCarryoverReportOnOff = 0
				End If
				Set rs_Settings_FieldService = Nothing
				cnn_Settings_FieldService.Close
				Set cnn_Settings_FieldService = Nothing

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

				'*****************************
				' S C H E D U L E R  L O G I C
				'*****************************
				If MUV_READ("FORCE") <> "1" Then ' Force parameter tell it not to read the schedule
					' Element 0 is "0" or "1" indicatiing OK to run or not
					' Element 1 hold 'OK' or the reason it should not run
					CheckSchedulerArray = Split(CheckScheduler("Settings_FieldService","Schedule_ServiceTicketCarryoverReportGeneration"),",")
					If CheckSchedulerArray(0) <> "1" Then
						ServiceTicketCarryoverReportOnOff = 0 ' Just turn it off & let the page flow normally
					End If
					Response.Write("<b>CheckScheduler Results: " &  CheckSchedulerArray(1) & "&nbsp;&nbsp;(" & ClientKey  & ")</b><br>")
					CreateAuditLogEntry "Service Ticket Carryover Report Launch","Service Ticket Carryover Report Launch","Minor",0,"Service Ticket Carryover Report Schedule check results: " & CheckSchedulerArray(1)
				End If
				'************************************
				' E O F  S C H E D U L E R  L O G I C
				'************************************

					
				If ServiceTicketCarryoverReportOnOff = 1 Then
	
					CreateAuditLogEntry "Service Ticket Carry Over Report Launch","Service Ticket Carry Over Report Launch","Minor",0,"Service Ticket Carry Over Report Launch ran."					
	
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
					
					If MUV_READ("cnnStatus") = "OK" AND Session("MAILOFF") = 0 Then ' else it loops
					
						'xxxxxxxxxxxxxxxx	
										
						'*************************************************************************************************
						'Create variable to save attachment comma separated files names to pass to SendMailWMultipleAtt()
						'*************************************************************************************************
	
						Dim fnAttachmentArray
						
	
						'*********************************************************************
						'Now create and save Service Ticket Carry Over REPORT PDF
						'*********************************************************************
																						
						
						Set Pdf = Server.CreateObject("Persits.Pdf")
						Set Doc = Pdf.CreateDocument
						
						
						Response.Write("<br>" & baseURL & "directlaunch/service/carryoverReport.asp?c=" & ClientKey & ",scale=0.8; hyperlinks=true; drawbackground=true<br>")
						Doc.ImportFromUrl baseURL & "directlaunch/service/carryoverReport.asp?c=" & ClientKey , "scale=0.8; hyperlinks=true; drawbackground=true; landscape=true"
						Response.Write(baseURL  & "directlaunch/service/carryoverReport.asp?c=" & ClientKey & "<br>")
						
						fn = "\clientfiles\" & trim(ClientKey) &"\z_pdfs\" & formatDateTime(Now(),2) & "-" & formatdatetime(Now(),4) & "_carryoverReport.pdf"
						fn = Replace(fn,"/","-")
						fn = Replace(fn,":","-")
						response.write(fn & "<br>")
	
						fn2 = Left(baseURL,Len(baseURL)-1) & fn
						fn2 = Replace(fn2,"\","/")
						response.write(fn2 & "<br>")
						response.write(Server.MapPath(fn) & "-Server.MapPath(fn)<br>")
						Main_PDF_Filename = fn
						
						fnAttachmentArray = Server.MapPath(Main_PDF_Filename) 
						
						Filename = Doc.Save(Server.MapPath(fn), False)
						
						
						'Now wait until the file exists on the server before we try to mail it
						TimeoutSecs = 60
						TimeoutCounter=0
						FOundFile = False
						Do While TimeoutCounter < TimeoutSecs 
							If CheckRemoteURL(fn2) = True Then
								FoundFile = True
								Exit Do ' The file is there
							End If
							DelayResponse(1) ' wait 1 sec & try again
							TimeoutCounter = TimeoutCounter + 1
						Loop
						
						If FoundFile <> True Then 
							Response.Write ("NO FILE FOUND")
							Response.End ' Could not fine the pdf, so just bail
						End If
	
						'******************************************************
						'Now start figuring out who it needs to be emailed to
						'******************************************************	
						
						
						SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"
						
						Set Connection = Server.CreateObject("ADODB.Connection")
						Set Recordset = Server.CreateObject("ADODB.Recordset")
						
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
								%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect.
								<%
								Response.End
						Else
							Session("ClientCnnString") = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
							Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & Recordset.Fields("dbCatalog")
							Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & Recordset.Fields("dbLogin")
							Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
							Recordset.close
							Connection.close	
						End If	
						
						
						'This is here so we only open it once for the whole page
						Set cnn_Settings_Global = Server.CreateObject("ADODB.Connection")
						cnn_Settings_Global.open (Session("ClientCnnString"))
						Set rs_Settings_Global = Server.CreateObject("ADODB.Recordset")
						rs_Settings_Global.CursorLocation = 3 
						SQL_Settings_Global = "SELECT * FROM Settings_FieldService"
						Set rs_Settings_Global = cnn_Settings_Global.Execute(SQL_Settings_Global)
						If not rs_Settings_Global.EOF Then
							ServiceTicketCarryoverReportUserNos = rs_Settings_Global("ServiceTicketCarryoverReportUserNos")
							ServiceTicketCarryoverReportAdditionalEmails = rs_Settings_Global("ServiceTicketCarryoverReportAdditionalEmails")
							ServiceTicketCarryoverReportEmailSubject = rs_Settings_Global("ServiceTicketCarryoverReportEmailSubject")
						End If
						Set rs_Settings_Global = Nothing
						cnn_Settings_Global.Close
						Set cnn_Settings_Global = Nothing
						
						Response.Write("OK, got here<br>")
						Response.Write("attachments: " & fnAttachmentArray & "<br>")
						Response.Write("<script type=""text/javascript"">closeme();</script>")
						'OK, now start breaking out the email addresses
						'*******************************************************************************************************************************************************************
						
	
						
							Send_To=""
							
							'Now see if there any additionals
							If ServiceTicketCarryoverReportAdditionalEmails <> "" and not IsNull(ServiceTicketCarryoverReportAdditionalEmails) Then
								tmpServiceTicketCarryoverReportAdditionalEmails   = trim(ServiceTicketCarryoverReportAdditionalEmails)		
								If Len(tmpServiceTicketCarryoverReportAdditionalEmails) > 1 Then
									If Right(tmpServiceTicketCarryoverReportAdditionalEmails,1) <> ";" Then tmpServiceTicketCarryoverReportAdditionalEmails = tmpServiceTicketCarryoverReportAdditionalEmails& ";"
									Send_To = Send_To & tmpServiceTicketCarryoverReportAdditionalEmails
								End If	
							End If
							
							'Get user based emails
							If ServiceTicketCarryoverReportUserNos <> "" Then
								UserNoList = Split(ServiceTicketCarryoverReportUserNos ,",")
								For x = 0 To UBound(UserNoList)
									Send_To = Send_To & GetUserEmailByUserNo(UserNoList(x)) & ";"
								Next
							End If
	
									
							'Got all the addresses so now break them up
							Send_To_Array = Split(Send_To,";")
							
							Response.Write("<br>Send_To: " & Send_To & "<br>")
							
							'HERE WE ACTUALLY SEND THE EMAIL
							For x = 0 to Ubound(Send_To_Array) -1
								Send_To = Send_To_Array(x)
	
								If ServiceTicketCarryoverReportEmailSubject = "" Then
									emailSubject = "Service Ticket Carry Over Report (" & ClientKey & ")"
								Else
							 		emailSubject = ServiceTicketCarryoverReportEmailSubject & " (" & ClientKey & ")"
								End If
	
								emailBody = ""
								'Failsafe for dev
								sURL = Request.ServerVariables("SERVER_NAME")
								If Instr(ucase(sURL),"DEV.") <> 0 Then Send_To = "rich@ocsaccess.com"
								If Instr(ucase(sURL),"DEV2.") <> 0 Then Send_To = "cgrecco@ocsaccess.com"
								emailBody = "Your Service Ticket Carry Over Report is attached. (" & ClientKey & ")"
								'fn3=Server.MapPath(fn)
								'Response.Write(fn3 & "<br>")
								
								If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV")) <> 0 Then
									Send_To="rsmith@ocsaccess.com"
								End If

								SendMailWAtt "mailsender@" & maildomain,Send_To,emailSubject,emailBody,fnAttachmentArray,"Service","Service Carry Over Report","MDS Insight"
								
								CreateAuditLogEntry "Automated Service Ticket Carry Over Report","Automated Service Ticket Carry Over Report","Minor",0,"Automated Service Ticket Carry Over Report Sent to " & Send_To 
								Response.Write("Sent the email to " & Send_To & "<br>")
								Response.Write("Sent the email, all done<br>")
							Next 
							

					' Everything above was for the overall summary report
					' Now we need to check & run the primary & secondary salesman
					' if appropriate
					
					'*************************************************************************************************
					'CHECK TO SEE IF WE NEED TO SEND CARRYOVER REPORT TO PRIMARY SALESMEN
					''*************************************************************************************************					

					If SendToPrimary = 1 Then


						Set cnn_Settings_FieldService = Server.CreateObject("ADODB.Connection")
						cnn_Settings_FieldService.open (Session("ClientCnnString"))
						
						Set rsSalesman = Server.CreateObject("ADODB.Recordset")
						rsSalesman.CursorLocation = 3 
						
						SQLSalesman = "Select Distinct Salesman FROM AR_Customer WHERE Salesman IS NOT NULL AND Salesman <> '' ORDER BY Salesman"
						
						Response.Write(SQLSalesman  & "<BR>")
						
						Set rsSalesman = cnn_Settings_FieldService.Execute(SQLSalesman)
												
						If Not rsSalesman.EOF Then 

							'Count Them
							PrimaryCount = 0
							Do While Not rsSalesman.EOF

								PrimaryCount = PrimaryCount +1

								rsSalesman.Movenext
							Loop
						End If
						
						Redim PrimarySalesmanArray(PrimaryCount)
						
						rsSalesman.MoveFirst
						ArrPointer = 0
						
						Do While Not rsSalesman.EOF

							If IsNumeric(rsSalesman("Salesman")) Then 
								PrimarySalesmanArray(ArrPointer) = rsSalesman("Salesman")
								ArrPointer = ArrPointer + 1
							End If

							rsSalesman.Movenext
						Loop

						set rsSalesman = nothing
						cnn_Settings_FieldService.close
						set cnn_Settings_FieldService= nothing

						'Now process each salesman
						For i = 0 To Ubound(PrimarySalesmanArray)-1
						
								response.write("for loop:" & PrimarySalesmanArray(i) & "<br>")
						
								fnAttachmentArray = ""
						
								Set Pdf = Nothing
								Set Doc = Nothing
	
								Set Pdf = Server.CreateObject("Persits.Pdf")
								Set Doc = Pdf.CreateDocument

								SalesmanToProcess = PrimarySalesmanArray(i)
								
								' See if there are any calls for this salesman, otherwise
								' set the variable to ""
								SQL_CheckSalesman = "SELECT * FROM FS_ServiceMemos "
								SQL_CheckSalesman = SQL_CheckSalesman & " INNER JOIN AR_Customer ON AR_Customer.CustNum = AccountNumber "
								SQL_CheckSalesman = SQL_CheckSalesman & " WHERE AR_Customer.Salesman = '" & SalesmanToProcess & "'"
								SQL_CheckSalesman = SQL_CheckSalesman & " AND CurrentStatus = 'OPEN' "
								SQL_CheckSalesman = SQL_CheckSalesman & " AND RecordSubType = 'OPEN' "
								
								Set cnnCheckSalesman = Server.CreateObject("ADODB.Connection")
								cnnCheckSalesman.open (MUV_READ("ClientCnnString"))
								Set rsCheckSalesman = Server.CreateObject("ADODB.Recordset")
								rsCheckSalesman.CursorLocation = 3 
								Set rsCheckSalesman = cnnCheckSalesman.Execute(SQL_CheckSalesman)
								
								If GetUserNoBySalesPersonNo(SalesmanToProcess) = "" Then SalesmanToProcess = ""
								
								If rsCheckSalesman.EOF Then SalesmanToProcess = "" ' Nothing for this salesman
								
								Set rsCheckSalesman = Nothing
								cnnCheckSalesman.Close
								Set cnnCheckSalesman = Nothing
								
								If SalesmanToProcess <> "" Then
								
									Response.Write("<br>" & baseURL & "directlaunch/service/carryoverReport.asp?c=" & ClientKey & "&sls=" & SalesmanToProcess & ",scale=0.8; hyperlinks=true; drawbackground=true<br>")
									Doc.ImportFromUrl baseURL & "directlaunch/service/carryoverReport.asp?c=" & ClientKey & "&sls=" & SalesmanToProcess , "scale=0.8; hyperlinks=true; drawbackground=true; landscape=true"
								
									fn = "\clientfiles\" & trim(ClientKey) &"\z_pdfs\" & formatDateTime(Now(),2) & "-" & formatdatetime(Now(),4) & "_" & "carryoverReport_" & SalesmanToProcess & ".pdf"
									fn = Replace(fn,"/","-")
									fn = Replace(fn,":","-")
									response.write("fn:" & fn & "<br>")
					
									fn2 = Left(baseURL,Len(baseURL)-1) & fn
									fn2 = Replace(fn2,"\","/")
									response.write(fn2 & "<br>")
									response.write(Server.MapPath(fn) & "-Server.MapPath(fn)<br>")
									Main_PDF_Filename = fn
										
									fnAttachmentArray = Server.MapPath(Main_PDF_Filename) 
									
									response.write("---Server.MapPath(fn):" & Server.MapPath(fn) & "<br>")
									
									Filename = Doc.Save(Server.MapPath(fn), False)
									
									response.write("SAVED<br>")
											
									'Now wait until the file exists on the server before we try to mail it
									TimeoutSecs = 60
									TimeoutCounter=0
									FOundFile = False
									Do While TimeoutCounter < TimeoutSecs 
										If CheckRemoteURL(fn2) = True Then
											FoundFile = True
											Exit Do ' The file is there
										End If
										DelayResponse(1) ' wait 1 sec & try again
										TimeoutCounter = TimeoutCounter + 1
									Loop
										
									If FoundFile <> True Then 
										Response.Write ("NO FILE FOUND")
										Response.End ' Could not fine the pdf, so just bail
									End If
									
									UserNoToProcess = GetUserNoBySalesPersonNo(SalesmanToProcess)
									WriteResponse("Lookup User no for salesman : " & SalesmanToProcess & "<br>")
									WriteResponse("Found the user no: " & UserNoToProcess  & "<br>")
									WriteResponse("Sending email to salesperson: " & GetUserDisplayNameByUserNo(UserNoToProcess) & "<br>")
									WriteResponse("Email for this salesperson is : " &  GetUserEmailByUserNo(UserNoToProcess) & "<br>")
										
									Send_To = GetUserEmailByUserNo(UserNoToProcess)

									'HERE WE ACTUALLY SEND THE EMAIL
									If ServiceTicketCarryoverReportEmailSubject = "" Then
										emailSubject = "Service Ticket Carry Over Report - " & GetUserDisplayNameByUserNo(UserNoToProcess) & " - (" & ClientKey & ")"
									Else
								 		emailSubject = ServiceTicketCarryoverReportEmailSubject & " - " & GetUserDisplayNameByUserNo(UserNoToProcess) & " - (" & ClientKey & ")"
									End If
				
									emailBody = ""
									'Failsafe for dev
									sURL = Request.ServerVariables("SERVER_NAME")
									If Instr(ucase(sURL),"DEV.") <> 0 Then Send_To = "rich@ocsaccess.com"
									emailBody = "Your Service Ticket Carry Over Report is attached. (" & ClientKey & ")"
									'fn3=Server.MapPath(fn)
									'Response.Write(fn3 & "<br>")
										
									If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV")) <> 0 Then
										Send_To="rsmith@ocsaccess.com"
									End If
	
									SendMailWAtt "mailsender@" & maildomain,Send_To,emailSubject,emailBody,fnAttachmentArray,"Service","Service Carry Over Report","MDS Insight"
									
									CreateAuditLogEntry "Automated Service Ticket Carry Over Report","Automated Service Ticket Carry Over Report","Minor",0,"Automated Service Ticket Carry Over Report Sent to " & Send_To 
									Response.Write("Sent the email to " & Send_To & "<br>")
									Response.Write("Sent the email, all done<br>")
										
								End If
						Next
						
					
					End If

					'*************************************************************************************************
					'CHECK TO SEE IF WE NEED TO SEND CARRYOVER REPORT TO SECONDARY SALESMEN
					''*************************************************************************************************

					If SendToSecondary = 1 Then


						Set cnn_Settings_FieldService = Server.CreateObject("ADODB.Connection")
						cnn_Settings_FieldService.open (Session("ClientCnnString"))
						
						Set rsSalesman = Server.CreateObject("ADODB.Recordset")
						rsSalesman.CursorLocation = 3 
						
						SQLSalesman = "Select Distinct SecondarySalesman FROM AR_Customer WHERE SecondarySalesman IS NOT NULL AND SecondarySalesman <> '' ORDER BY SecondarySalesman "
						
						Response.Write(SQLSalesman  & "<BR>")
						
						Set rsSalesman = cnn_Settings_FieldService.Execute(SQLSalesman)
												
						If Not rsSalesman.EOF Then 

							'Count Them
							SecondaryCount = 0
							Do While Not rsSalesman.EOF

								SecondaryCount = SecondaryCount +1

								rsSalesman.Movenext
							Loop
						End If
						
						Redim SecondarySalesmanArray(SecondaryCount)
						
						rsSalesman.MoveFirst
						ArrPointer = 0
						
						Do While Not rsSalesman.EOF

							If IsNumeric(rsSalesman("SecondarySalesman")) Then 
								SecondarySalesmanArray(ArrPointer) = rsSalesman("SecondarySalesman")
								ArrPointer = ArrPointer + 1
							End If

							rsSalesman.Movenext
						Loop

						set rsSalesman = nothing
						cnn_Settings_FieldService.close
						set cnn_Settings_FieldService= nothing

						response.write("----------------------------------------------------<br>")
						response.write("----------------------------------------------------<br>")

						'Now process each salesman
						For i = 0 To Ubound(SecondarySalesmanArray) - 1
						
								response.write("for loop:" & SecondarySalesmanArray(i) & "<br>")
						
								fnAttachmentArray = ""
						
								Set Pdf = Nothing
								Set Doc = Nothing
	
								Set Pdf = Server.CreateObject("Persits.Pdf")
								Set Doc = Pdf.CreateDocument

								SalesmanToProcess = SecondarySalesmanArray(i)
								
								' See if there are any calls for this salesman, otherwise
								' set the variable to ""
								SQL_CheckSalesman = "SELECT * FROM FS_ServiceMemos "
								SQL_CheckSalesman = SQL_CheckSalesman & " INNER JOIN AR_Customer ON AR_Customer.CustNum = AccountNumber "
								SQL_CheckSalesman = SQL_CheckSalesman & " WHERE AR_Customer.SecondarySalesman = '" & SalesmanToProcess & "'"
								SQL_CheckSalesman = SQL_CheckSalesman & " AND CurrentStatus = 'OPEN' "
								SQL_CheckSalesman = SQL_CheckSalesman & " AND RecordSubType = 'OPEN' "
								
								Set cnnCheckSalesman = Server.CreateObject("ADODB.Connection")
								cnnCheckSalesman.open (MUV_READ("ClientCnnString"))
								Set rsCheckSalesman = Server.CreateObject("ADODB.Recordset")
								rsCheckSalesman.CursorLocation = 3 
								Set rsCheckSalesman = cnnCheckSalesman.Execute(SQL_CheckSalesman)
								
								If GetUserNoBySalesPersonNo(SalesmanToProcess) = "" Then SalesmanToProcess = ""
								
								If rsCheckSalesman.EOF Then SalesmanToProcess = "" ' Nothing for this salesman
								
								Set rsCheckSalesman = Nothing
								cnnCheckSalesman.Close
								Set cnnCheckSalesman = Nothing
								
								If SalesmanToProcess <> "" Then
								
									Response.Write("<br>" & baseURL & "directlaunch/service/carryoverReport.asp?c=" & ClientKey & "&sls=" & SalesmanToProcess & ",scale=0.8; hyperlinks=true; drawbackground=true<br>")
									Doc.ImportFromUrl baseURL & "directlaunch/service/carryoverReport.asp?c=" & ClientKey & "&sls=" & SalesmanToProcess , "scale=0.8; hyperlinks=true; drawbackground=true; landscape=true"
								
									fn = "\clientfiles\" & trim(ClientKey) &"\z_pdfs\" & formatDateTime(Now(),2) & "-" & formatdatetime(Now(),4) & "_" & "carryoverReport_" & SalesmanToProcess & ".pdf"
									fn = Replace(fn,"/","-")
									fn = Replace(fn,":","-")
									response.write("fn:" & fn & "<br>")
					
									fn2 = Left(baseURL,Len(baseURL)-1) & fn
									fn2 = Replace(fn2,"\","/")
									response.write(fn2 & "<br>")
									response.write(Server.MapPath(fn) & "-Server.MapPath(fn)<br>")
									Main_PDF_Filename = fn
										
									fnAttachmentArray = Server.MapPath(Main_PDF_Filename) 
									
									response.write("---Server.MapPath(fn):" & Server.MapPath(fn) & "<br>")
									
									Filename = Doc.Save(Server.MapPath(fn), False)
									
									response.write("SAVED<br>")
											
									'Now wait until the file exists on the server before we try to mail it
									TimeoutSecs = 60
									TimeoutCounter=0
									FOundFile = False
									Do While TimeoutCounter < TimeoutSecs 
										If CheckRemoteURL(fn2) = True Then
											FoundFile = True
											Exit Do ' The file is there
										End If
										DelayResponse(1) ' wait 1 sec & try again
										TimeoutCounter = TimeoutCounter + 1
									Loop
										
									If FoundFile <> True Then 
										Response.Write ("NO FILE FOUND")
										Response.End ' Could not fine the pdf, so just bail
									End If
									
									UserNoToProcess = GetUserNoBySalesPersonNo(SalesmanToProcess)
									WriteResponse("Lookup User no for salesman : " & SalesmanToProcess & "<br>")
									WriteResponse("Found the user no: " & UserNoToProcess  & "<br>")
									WriteResponse("Sending email to salesperson: " & GetUserDisplayNameByUserNo(UserNoToProcess) & "<br>")
									WriteResponse("Email for this salesperson is : " &  GetUserEmailByUserNo(UserNoToProcess) & "<br>")
										
									Send_To = GetUserEmailByUserNo(UserNoToProcess)

									'HERE WE ACTUALLY SEND THE EMAIL
									If ServiceTicketCarryoverReportEmailSubject = "" Then
										emailSubject = "Service Ticket Carry Over Report - " & GetUserDisplayNameByUserNo(UserNoToProcess) & " - (" & ClientKey & ")"
									Else
								 		emailSubject = ServiceTicketCarryoverReportEmailSubject & " - " & GetUserDisplayNameByUserNo(UserNoToProcess) & " - (" & ClientKey & ")"
									End If
				
									emailBody = ""
									'Failsafe for dev
									sURL = Request.ServerVariables("SERVER_NAME")
									If Instr(ucase(sURL),"DEV.") <> 0 Then Send_To = "rich@ocsaccess.com"
									emailBody = "Your Service Ticket Carry Over Report is attached. (" & ClientKey & ")"
									'fn3=Server.MapPath(fn)
									'Response.Write(fn3 & "<br>")
										
									If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV")) <> 0 Then
										Send_To="rsmith@ocsaccess.com"
									End If
	
									SendMailWAtt "mailsender@" & maildomain,Send_To,emailSubject,emailBody,fnAttachmentArray,"Service","Service Carry Over Report","MDS Insight"
									
									CreateAuditLogEntry "Automated Service Ticket Carry Over Report","Automated Service Ticket Carry Over Report","Minor",0,"Automated Service Ticket Carry Over Report Sent to " & Send_To 
									Response.Write("Sent the email to " & Send_To & "<br>")
									Response.Write("Sent the email, all done<br>")
										
								End If
						Next
						
					
					End If
					
					'*************************************************************************************************
					'CHECK TO SEE IF WE NEED TO SEND CARRYOVER REPORT TO ANY TEAMS
					''*************************************************************************************************

					If ServiceTicketCarryoverReportTeamIntRecIDs <> "" Then
					
						TeamsArray = Split(ServiceTicketCarryoverReportTeamIntRecIDs,",")
						
						For i = 0 to Ubound(TeamsArray)
	
							Set cnn_Settings_FieldService = Server.CreateObject("ADODB.Connection")
							cnn_Settings_FieldService.open (Session("ClientCnnString"))
							
							Set rsTeam = Server.CreateObject("ADODB.Recordset")
							rsTeam.CursorLocation = 3 
							
							SQLTeams = "SELECT * FROM USER_Teams WHERE InternalRecordIdentifier = " & TeamsArray(i)
							
							Response.Write(SQLTeams & "<BR>")
							
							Set rsTeam = cnn_Settings_FieldService.Execute(SQLTeams)
							
							TeamToProcess = ""
							TeamsUserNoArray = ""
													
							If Not rsTeam.EOF Then 
													
								response.write("for loop current team rec id: " & TeamsArray(i) & "<br>")
						
								fnAttachmentArray = ""
						
								Set Pdf = Nothing
								Set Doc = Nothing
	
								Set Pdf = Server.CreateObject("Persits.Pdf")
								Set Doc = Pdf.CreateDocument

								TeamToProcess = rsTeam("TeamUserNos")
								TeamsUserNoArray = Split(TeamToProcess,",")
								
								' See if there are any calls for anyone of this team, otherwise
								' set the variable to ""
								SQL_CheckTeamMember = "SELECT * FROM FS_ServiceMemos "
								SQL_CheckTeamMember = SQL_CheckTeamMember & " INNER JOIN AR_Customer ON AR_Customer.CustNum = AccountNumber "
								SQL_CheckTeamMember = SQL_CheckTeamMember & " WHERE "

								SalespersonNumber = ""
								For x = 0 to Ubound(TeamsUserNoArray)
									SalespersonNumber = GetSalesPersonNoByUserNo(TeamsUserNoArray(x))
									If SalespersonNumber <> "" Then
										SlsmnToProcessList =  SalespersonNumber & ","	
									End If						
								Next 
								
								SlsmnToProcessList = Left(SlsmnToProcessList, Len(SlsmnToProcessList) - 1)
								
								If SlsmnToProcessList <> "" Then
									SQL_CheckTeamMember = SQL_CheckTeamMember & " (AR_Customer.Salesman IN ('" & SlsmnToProcessList & "')"
									SQL_CheckTeamMember = SQL_CheckTeamMember & " OR AR_Customer.SecondarySalesman IN ('" & SlsmnToProcessList & "')) "
								End If
								
								SQL_CheckTeamMember = SQL_CheckTeamMember & " AND CurrentStatus = 'OPEN' "
								SQL_CheckTeamMember = SQL_CheckTeamMember & " AND RecordSubType = 'OPEN' "
								
								response.write("SQL_CheckTeamMember: " & SQL_CheckTeamMember & "<br>")
								
								Set cnnCheckTeamMember = Server.CreateObject("ADODB.Connection")
								cnnCheckTeamMember.open (MUV_READ("ClientCnnString"))
								Set rsCheckTeamMember = Server.CreateObject("ADODB.Recordset")
								rsCheckTeamMember.CursorLocation = 3 
								Set rsCheckTeamMember = cnnCheckTeamMember.Execute(SQL_CheckTeamMember)
								
								
								If rsCheckTeamMember.EOF Then TeamToProcess = "" ' Nothing for this team
								
								Set rsCheckTeamMember = Nothing
								cnnCheckTeamMember.Close
								Set cnnCheckTeamMember = Nothing
								
								If TeamToProcess <> "" Then
								
								
									Response.Write("<br>" & baseURL & "directlaunch/service/carryoverReport.asp?c=" & ClientKey & "&tm=" & TeamsArray(i) & ",scale=0.8; hyperlinks=true; drawbackground=true<br>")
									Doc.ImportFromUrl baseURL & "directlaunch/service/carryoverReport.asp?c=" & ClientKey & "&tm=" & TeamsArray(i) ,"scale=0.8; hyperlinks=true; drawbackground=true; landscape=true"

						
									fn = "\clientfiles\" & trim(ClientKey) &"\z_pdfs\" & formatDateTime(Now(),2) & "-" & formatdatetime(Now(),4) & "_" & "carryoverReport_Team" & TeamsArray(i) & ".pdf"
									fn = Replace(fn,"/","-")
									fn = Replace(fn,":","-")
									response.write("fn:" & fn & "<br>")
					
									fn2 = Left(baseURL,Len(baseURL)-1) & fn
									fn2 = Replace(fn2,"\","/")
									response.write(fn2 & "<br>")
									response.write(Server.MapPath(fn) & "-Server.MapPath(fn)<br>")
									Main_PDF_Filename = fn
										
									fnAttachmentArray = Server.MapPath(Main_PDF_Filename) 
									
									response.write("---Server.MapPath(fn):" & Server.MapPath(fn) & "<br>")
									
									Filename = Doc.Save(Server.MapPath(fn), False)
									
									response.write("SAVED<br>")
											
									'Now wait until the file exists on the server before we try to mail it
									TimeoutSecs = 60
									TimeoutCounter=0
									FOundFile = False
									Do While TimeoutCounter < TimeoutSecs 
										If CheckRemoteURL(fn2) = True Then
											FoundFile = True
											Exit Do ' The file is there
										End If
										DelayResponse(1) ' wait 1 sec & try again
										TimeoutCounter = TimeoutCounter + 1
									Loop
										
									If FoundFile <> True Then 
										Response.Write ("NO FILE FOUND")
										Response.End ' Could not fine the pdf, so just bail
									End If
									
									'********************************************************************
									'LOOP THROUGH ALL THE MEMBERS OF A SINGLE TEAM AND SEND EACH MEMBER
									'A COPY OF THEIR TEAM'S REPORT
									'********************************************************************
									For z = 0 to Ubound(TeamsUserNoArray)
									
										UserNoToProcess = TeamsUserNoArray(z)
										WriteResponse("Sending email to team member: " & GetUserDisplayNameByUserNo(UserNoToProcess) & "<br>")
										WriteResponse("Email for this team member is : " &  GetUserEmailByUserNo(UserNoToProcess) & "<br>")
											
										Send_To = GetUserEmailByUserNo(UserNoToProcess)
										
										TeamName = GetTeamNameByTeamIntRecID(TeamsArray(i))
		
										'HERE WE ACTUALLY SEND THE EMAIL
										If ServiceTicketCarryoverReportEmailSubject = "" Then
											emailSubject = "Service Ticket Carry Over Team Report - " & TeamName & " - (" & ClientKey & ")"
										Else
									 		emailSubject = ServiceTicketCarryoverReportEmailSubject & " - " & TeamName & " - (" & ClientKey & ")"
										End If
					
										emailBody = ""
										'Failsafe for dev
										sURL = Request.ServerVariables("SERVER_NAME")
										If Instr(ucase(sURL),"DEV.") <> 0 Then Send_To = "cgrecco@ocsaccess.com"
										emailBody = "Your Service Ticket Team (" & TeamName & ") Carry Over Report is attached. (" & ClientKey & ")"
											
										If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV")) <> 0 Then
											Send_To="cgrecco@ocsaccess.com"
										End If
		
										SendMailWAtt "mailsender@" & maildomain,Send_To,emailSubject,emailBody,fnAttachmentArray,"Service","Service Carry Over Team Report","MDS Insight"
										
										CreateAuditLogEntry "Automated Service Ticket Carry Over Report","Automated Service Ticket Carry Over Report","Minor",0,"Automated Service Ticket Carry Over Report Sent to " & Send_To 
										Response.Write("Sent the email to " & Send_To & "<br>")
										Response.Write("Sent the email, all done<br>")
										
									Next
										
								End If
								
							End If
	
						Next
						
					End If
					'*************************************************************************************************
					'END TEAMS
					'*************************************************************************************************
					

					'****************************************************************************************************
					' Now see if the service carry over report user summary text message feature is on or off
					'****************************************************************************************************				

					Set cnn_Settings_FieldService = Server.CreateObject("ADODB.Connection")
					cnn_Settings_FieldService.open (MUV_READ("ClientCnnString"))
					Set rs_Settings_FieldService = Server.CreateObject("ADODB.Recordset")
					rs_Settings_FieldService.CursorLocation = 3 
					SQL_Settings_FieldService = "SELECT * FROM Settings_FieldService"
					Set rs_Settings_FieldService = cnn_Settings_FieldService.Execute(SQL_Settings_FieldService)
					If not rs_Settings_FieldService.EOF Then
						ServiceTicketCarryoverReportTextSummaryOnOff = rs_Settings_FieldService("ServiceTicketCarryoverReportTextSummaryOnOff")
						ServiceTicketCarryoverReportTextSummaryUserNos = rs_Settings_FieldService("ServiceTicketCarryoverReportTextSummaryUserNos")
					Else
						ServiceTicketCarryoverReportTextSummaryOnOff = 0
					End If
					Set rs_Settings_FieldService = Nothing
					cnn_Settings_FieldService.Close
					Set cnn_Settings_FieldService = Nothing
					
					If ServiceTicketCarryoverReportTextSummaryOnOff = 1 Then
					
						'Get user numbers to send text message to
						If ServiceTicketCarryoverReportTextSummaryUserNos <> "" Then
						
							UserNoTextList = Split(ServiceTicketCarryoverReportTextSummaryUserNos,",")
							
							For x = 0 To UBound(UserNoTextList)
							
								CurrentUserNo = UserNoTextList(x)
	
								'Get the text number of the current user in the loop
								Send_To = getUserCellNumber(CurrentUserNo)
							
								'Only do this if we have their cell #
								If Send_To <> "" Then
							
									WriteResponse("**********FOUND TEXT RECIPIENTS TO SEND TO**************************<br>")
						
									txtSubject = "Service Tickets - Client (" & ClientKey & ")"
									
									'***************************************************************************************
									'Generate message to send in text
									'***************************************************************************************
											
									SQL = "SELECT COUNT(*) AS TotalTickets FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange <> 1"	
						
									Set cnnCarryOver = Server.CreateObject("ADODB.Connection")
									cnnCarryOver.open (MUV_READ("ClientCnnString"))
									Set rsCarryOver  = Server.CreateObject("ADODB.Recordset")
									rsCarryOver.CursorLocation = 3 
									rsCarryOver.Open SQL, cnnCarryOver 
									
									TotalNumberOfTickets = 0
												
									If Not rsCarryOver.EOF Then
										TotalNumberOfTickets = rsCarryOver("TotalTickets")		
									End If
									
									txtMessage = TotalNumberOfTickets & " OPEN SERVICE TICKETS"
												
									rsCarryOver.Close
									'***************************************************************************************
											
									Send_To = Replace(Send_To,"-","") ' EZ Texting doesn't like dashes
							
								
									'*****Text numbers don't get split into an array, the php takes multiple #'s separated by commas	
									
									If Right(Send_To,1) = "," Then Send_To = Left(Send_To,Len(Send_To)-1)
							
									TEXT_TO = Send_To
								
									' If this is running on dev, send to Rich's text number
									If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then
										WriteResponse ("Running on dev - changing text number to 6099294430<br>")					
										TEXT_TO = "6099294430"
									End IF
							
TEXT_TO = "6099294430"
									'TEXT_TO = TEXT_TO & ",6099294430"
									'If Instr(TEXT_TO,"6099294430") <> 0 Then
										txtMessage = txtMessage & "(" & GetUserDisplayNameByUserNo(CurrentUserNo) & ")"' So rich knows who it is for
										If Len(txtMessage) > 160 Then txtMessage = Left(txtMessage,160) ' Max text length of 160
									'End If
									
									WriteResponse ("<font color='green'><b>Sending text to: "& TEXT_TO & "</b></font><br>")
								
									CreateAuditLogEntry "Service Ticket Carryover Report Summary Text Sent","Service Ticket Carryover Report Summary Text Sent","Minor",0,"Summary message texted to " & GetUserDisplayNameByUserNo(CurrentUserNo) & " at " & TEXT_TO & " Subject: - " & txtSubject 	
								
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
							Next
							
						End If
				
					End If
				
										
					WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				
				End If
			
		Else
		
			WriteResponse ("Skipping the client " & ClientKey & " because the service ticket carry over report is turned off.<BR>")
		
		End If ' for the report being turned off
			
		End If	
		
	Else ' is the Service  module enabled
	
		Call SetClientCnnString
				
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
			
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


		WriteResponse ("Skipping the client " & ClientKey & " because the Service Module module is not enabled.<BR>")
		
	End If ' is the Service  module enabled
	
	TopRecordset.movenext
	
	Loop
	
	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type=""text/javascript"">closeme();</script>")	


'************************************************************************************
'************************************************************************************
'Subs and funcs begin here
'************************************************************************************

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


Sub WriteResponse(passedLogEntry)

	response.write(Now() & "&nbsp;&nbsp;&nbsp;" & passedLogEntry)
	
	passedLogEntry = Replace(passedLogEntry,"'","''")
	
	SQL = "INSERT INTO SC_AuditLogDLaunch (EntryThread, DirectLaunchName, DirectLaunchFile, LogEntry)"
	SQL = SQL &  " VALUES (" & EntryThread & ""
	SQL = SQL & ",'Service ticket carry over Report'"
	SQL = SQL & ",'/directlaunch/service/carryover_report_launch.asp'"
	SQL = SQL & ",'"  & passedLogEntry & "'"
	SQL = SQL & ")"
	
	'Response.write("<BR>" & SQL & "<BR>")
	
	Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
	cnnAuditLog.open Session("ClientCnnString") 
	Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
	rsAuditLog.CursorLocation = 3 
	
	Set rsAuditLog = cnnAuditLog.Execute(SQL)

	set rsAuditLog = nothing
	cnnAuditLog.close
	set cnnAuditLog = nothing

End Sub


Sub DelayResponse(numberOfseconds)
 Dim WshShell
 Set WshShell=Server.CreateObject("WScript.Shell")
 WshShell.Run "waitfor /T " & numberOfSecond & "SignalThatWontHappen", , True
End Sub

Function CheckRemoteURL(fileURL)
    ON ERROR RESUME NEXT
    Dim xmlhttp

    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

    xmlhttp.open "GET", fileURL, False
    xmlhttp.send
    If(Err.Number<>0) then
        Response.Write "Could not connect to remote server"
    else
        Select Case Cint(xmlhttp.status)
            Case 200, 202, 302
                Set xmlhttp = Nothing
                CheckRemoteURL = True
            Case Else
                Set xmlhttp = Nothing
                CheckRemoteURL = False
        End Select
    end if
    ON ERROR GOTO 0
End Function


'************************************************************************************
'************************************************************************************
'Subs and funcs end here
'************************************************************************************


%>