<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/mailDirectLaunch.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.asp"-->

<script type="text/javascript">
    function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
 

<%
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page
'Usage = "http://{xxx}.{domain}.com/directLaunch/service/serviceTicketthresholdReport_launch.asp?runlevel=run_now
Server.ScriptTimeout = 25000

Dim EntryThread

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
	
		'To begin with, see if this client uses the service 
		'If they don't then don't bother running for then
		
		Response.Write("Seeing if we need to run this for " & ClientKey & "<br>")
		
		If TopRecordset.Fields("ServiceModule") = "Enabled" Then
	
			'The IF statement below makes sure that when run from DEV it only deos client keys with a d
			'and when run from LIVE it only does client keys without a d
			'Pretty smart, huh
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
			or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then 
												
				Call SetClientCnnString
				
				Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
				
				'**************************************************************
				' Now see if the service ticket threshold report is on or off
				'**************************************************************				
				'This is here so we only open it once for the whole page
				Set cnn_Settings_FieldService = Server.CreateObject("ADODB.Connection")
				cnn_Settings_FieldService.open (MUV_READ("ClientCnnString"))
				Set rs_Settings_FieldService = Server.CreateObject("ADODB.Recordset")
				rs_Settings_FieldService.CursorLocation = 3 
				SQL_Settings_FieldService = "SELECT * FROM Settings_FieldService"
				Set rs_Settings_FieldService = cnn_Settings_FieldService.Execute(SQL_Settings_FieldService)
				If not rs_Settings_FieldService.EOF Then
					ServiceTicketthresholdReportONOFF = rs_Settings_FieldService("ServiceTicketthresholdReportONOFF")
				Else
					ServiceTicketthresholdReportONOFF = 0
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
					CheckSchedulerArray = Split(CheckScheduler("Settings_FieldService","Schedule_ServiceTicketthresholdReportGeneration"),",")
					If CheckSchedulerArray(0) <> "1" Then
						ServiceTicketthresholdReportONOFF = 0 ' Just turn it off & let the page flow normally
					End If
					Response.Write("<b>CheckScheduler Results: " &  CheckSchedulerArray(1) & "&nbsp;&nbsp;(" & ClientKey  & ")</b><br>")
					CreateAuditLogEntry "Service Ticket threshold Report Launch","Service Ticket threshold Report Launch","Minor",0,"Service Ticket threshold Report Schedule check results: " & CheckSchedulerArray(1)
				End If
				'************************************
				' E O F  S C H E D U L E R  L O G I C
				'************************************

					
				If ServiceTicketthresholdReportONOFF = 1 Then
	
					CreateAuditLogEntry "Service Ticket threshold Report Launch","Service Ticket threshold Report Launch","Minor",0,"Service Ticket threshold Report ran."					
	
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
						'Now create and save Field Service Notes Report PDF
						'*********************************************************************
																						
						
						Set Pdf = Server.CreateObject("Persits.Pdf")
						Set Doc = Pdf.CreateDocument
						
						
						Response.Write("<br>" & baseURL & "directlaunch/service/serviceTicketThresholdReport.asp?c=" & ClientKey & ",scale=0.8; hyperlinks=true; drawbackground=true<br>")
						Doc.ImportFromUrl baseURL & "directlaunch/service/serviceTicketThresholdReport.asp?c=" & ClientKey , "scale=0.8; hyperlinks=true; drawbackground=true; landscape=true"
						Response.Write(baseURL  & "directlaunch/service/serviceTicketThresholdReport.asp?c=" & ClientKey & "<br>")
						
						fn = "\clientfiles\" & trim(ClientKey) &"\z_pdfs\" & formatDateTime(Now(),2) & "-" & formatdatetime(Now(),4) & "_serviceTicketthresholdReport.pdf"
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
						Set cnn_Settings_FieldService = Server.CreateObject("ADODB.Connection")
						cnn_Settings_FieldService.open (Session("ClientCnnString"))
						Set rs_Settings_FieldService = Server.CreateObject("ADODB.Recordset")
						rs_Settings_FieldService.CursorLocation = 3 
						SQL_Settings_FieldService = "SELECT * FROM Settings_FieldService"
						Set rs_Settings_FieldService = cnn_Settings_FieldService.Execute(SQL_Settings_FieldService)
						If not rs_Settings_FieldService.EOF Then
							FieldServiceNotesReportUserNos = rs_Settings_FieldService("ServiceTicketthresholdReportUserNos")
							FieldServiceNotesReportAdditionalEmails = rs_Settings_FieldService("ServiceTicketthresholdReportAdditionalEmails")
						End If
						Set rs_Settings_FieldService = Nothing
						cnn_Settings_FieldService.Close
						Set cnn_Settings_FieldService = Nothing
						
						Response.Write("OK, got here<br>")
						Response.Write("attachments: " & fnAttachmentArray & "<br>")
						'Response.Write("<script type=""text/javascript"">closeme();</script>")
						'OK, now start breaking out the email addresses
						'*******************************************************************************************************************************************************************
						

					
						Send_To=""
						
						'Now see if there any additionals
						If FieldServiceNotesReportAdditionalEmails <> "" and not IsNull(FieldServiceNotesReportAdditionalEmails) Then
							tmpFieldServiceNotesReportAdditionalEmails   = trim(FieldServiceNotesReportAdditionalEmails)		
							If Len(tmpFieldServiceNotesReportAdditionalEmails) > 1 Then
								If Right(tmpFieldServiceNotesReportAdditionalEmails,1) <> ";" Then tmpFieldServiceNotesReportAdditionalEmails = tmpFieldServiceNotesReportAdditionalEmails& ";"
								Send_To = Send_To & tmpFieldServiceNotesReportAdditionalEmails
							End If	
						End If
						
						'Get user based emails
						If FieldServiceNotesReportUserNos <> "" Then
							UserNoList = Split(FieldServiceNotesReportUserNos ,",")
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

							emailSubject = "Insight Service Ticket threshold Report (" & ClientKey & ")"

							emailBody = ""
							'Failsafe for dev
							sURL = Request.ServerVariables("SERVER_NAME")
							If Instr(ucase(sURL),"DEV.") <> 0 Then Send_To = "rich@ocsaccess.com"
							emailBody = "Your Insight Service Ticket threshold Report is attached. (" & ClientKey & ")"
							'fn3=Server.MapPath(fn)
							'Response.Write(fn3 & "<br>")
							
							If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV")) <> 0 Then
								Send_To="rsmith@ocsaccess.com"
							End If

							SendMailWAtt "mailsender@" & maildomain,Send_To,emailSubject,emailBody,fnAttachmentArray,"Service","Service Ticket threshold Report","MDS Insight"
							
							CreateAuditLogEntry "Automated Service Ticket threshold Report","Automated Service Ticket threshold Report","Minor",0,"Automated Service Ticket threshold Report Sent to " & Send_To 
							Response.Write("Sent the email to " & Send_To & "<br>")
							Response.Write("Sent the email, all done<br>")
						Next 
					
			
						WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				
				End If
			
			Else
			
				WriteResponse ("Skipping the client " & ClientKey & " because the Field Service Notes Report is turned off.<BR>")
			
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
	SQL = SQL & ",'Field Service Notes Report'"
	SQL = SQL & ",'/directlaunch/service/serviceTicketthresholdReport_launch.asp'"
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