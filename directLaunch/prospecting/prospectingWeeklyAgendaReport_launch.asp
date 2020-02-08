<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/mail.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->

<script type="text/javascript">
    function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
 

<%
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page
'Usage = "http://{xxx}.{domain}.com/directLaunch/prospecting/prospectingWeeklyAgendaReport_launch.asp?runlevel=run_now&force=run_now
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
	
		'To begin with, see if this client uses the prospecting module 
		'If they don't then don't bother running for then
		
		Response.Write("Seeing if we need to run this for " & ClientKey & "<br>")
		
		If TopRecordset.Fields("prospectingModule") = "Enabled" Then
	
			'The IF statement below makes sure that when run from DEV it only deos client keys with a d
			'and when run from LIVE it only does client keys without a d
			'Pretty smart, huh
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
			or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then 
												
				Call SetClientCnnString
				
				Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
				
				'**************************************************************
				' Now see if the prospecting weekly agenda report is on or off
				'**************************************************************				
				Set cnn_Settings_Prospecting = Server.CreateObject("ADODB.Connection")
				cnn_Settings_Prospecting.open (Session("ClientCnnString"))
				Set rs_Settings_Prospecting = Server.CreateObject("ADODB.Recordset")
				rs_Settings_Prospecting.CursorLocation = 3 
				SQL_Settings_Prospecting = "SELECT * FROM Settings_Prospecting"
				Set rs_Settings_Prospecting = cnn_Settings_Prospecting.Execute(SQL_Settings_Prospecting)
				If not rs_Settings_Prospecting.EOF Then
					ProspectingWeeklyAgendaReportOnOff = rs_Settings_Prospecting("ProspectingWeeklyAgendaReportOnOff")
					ProspectingWeeklyAgendaReportUserNos = rs_Settings_Prospecting("ProspectingWeeklyAgendaReportUserNos")
					ProspectingWeeklyAgendaReportAdditionalEmails = rs_Settings_Prospecting("ProspectingWeeklyAgendaReportAdditionalEmails")
					ProspectingWeeklyAgendaReportEmailSubject = rs_Settings_Prospecting("ProspectingWeeklyAgendaReportEmailSubject")
				Else
					ProspectingWeeklyAgendaReportOnOff = vbFalse
				End If
				Set rs_Settings_Prospecting = Nothing
				cnn_Settings_Prospecting.Close
				Set cnn_Settings_Prospecting = Nothing

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
					CheckSchedulerArray = Split(CheckScheduler("Settings_Prospecting","Schedule_ProspectingWeeklyAgendaReportGeneration"),",")
					If CheckSchedulerArray(0) <> "1" Then
						ProspectingWeeklyAgendaReportOnOff = 0 ' Just turn it off & let the page flow normally
					End If
					Response.Write("<b>CheckScheduler Results: " &  CheckSchedulerArray(1) & "&nbsp;&nbsp;(" & ClientKey  & ")</b><br>")
					CreateAuditLogEntry "Prospecting Weekly Agenda Report Launch","Prospecting Weekly Agenda Report Launch","Minor",0,"Prospecting Weekly Agenda Report Schedule check results: " & CheckSchedulerArray(1)
				End If
				'************************************
				' E O F  S C H E D U L E R  L O G I C
				'************************************

					
				If ProspectingWeeklyAgendaReportOnOff = 1 Then
	
					CreateAuditLogEntry "Prospecting Weekly Agenda Report Launch","Prospecting Weekly Agenda Report Launch","Minor",0,"Prospecting Weekly Agenda Report ran."					
	
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
					
						'*************************************************************************************************
						'CHECK TO SEE IF WE NEED TO SEND WEEKLY AGENDA REPORT TO ANY USER NOS
						''*************************************************************************************************
	
						If ProspectingWeeklyAgendaReportUserNos <> "" Then
	
							UserNoList = Split(ProspectingWeeklyAgendaReportUserNos,",")
							
							For x = 0 To UBound(UserNoList)-1
							
								response.write("----------------------------------------------------<br>")
								response.write("----------------------------------------------------<br>")
	
								response.write("for loop user email: " & GetUserEmailByUserNo(UserNoList(x)) & "<br>")
								response.write("for loop user no: " & UserNoList(x) & "<br>")
							
								fnAttachmentArray = ""
						
								Set Pdf = Nothing
								Set Doc = Nothing
	
								Set Pdf = Server.CreateObject("Persits.Pdf")
								Set Doc = Pdf.CreateDocument
	
								'xxxxxxxxxxxxxxxx	
								' Start individual guts of the report										
								
								Set Pdf = Server.CreateObject("Persits.Pdf")
								Set Doc = Pdf.CreateDocument
								Response.Write("<br>" & baseURL & "directlaunch/prospecting/prospectingWeeklyAgendaReport.asp?c=" & ClientKey & "&u=" & UserNoList(x) & ",scale=0.6; hyperlinks=true; drawbackground=true; landscape=false<br>")
								
								Doc.ImportFromUrl baseURL & "directlaunch/prospecting/prospectingWeeklyAgendaReport.asp?c=" & ClientKey & "&u=" & UserNoList(x) , "scale=0.6; hyperlinks=true; drawbackground=true; landscape=false"
																
								Response.Write(baseURL  & "directlaunch/prospecting/prospectingWeeklyAgendaReport.asp?c=" & ClientKey & "&u=" & UserNoList(x) & "<br>")
		
								fn = "\clientfilesV\" & trim(ClientKey) &"\z_pdfs\" & formatDateTime(Now(),2) & "-" & formatdatetime(Now(),4) & "_prospectingWeeklyAgendaReport_" & UserNoList(x) & ".pdf"
								fn = Replace(fn,"/","-")
								fn = Replace(fn,":","-")
								fn2 = baseURL & fn
								fn2 = Replace(fn2,"\","/")
								response.write(fn & "-fn<br>")
								response.write(fn2 & "-fn2<br>")
								response.write(Server.MapPath(fn) & "-Server.MapPath(fn)<br>")
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
								
								Send_To=""
								
								'Current User Email Address
								Send_To = Send_To & GetUserEmailByUserNo(UserNoList(x)) & ";"
	
								'Now see if there any additionals
								If ProspectingWeeklyAgendaReportAdditionalEmails <> "" and not IsNull(ProspectingWeeklyAgendaReportAdditionalEmails) Then
									tmpProspectingWeeklyAgendaReportAdditionalEmails  = trim(ProspectingWeeklyAgendaReportAdditionalEmails)		
									If Len(tmpProspectingWeeklyAgendaReportAdditionalEmails) > 1 Then
										If Right(tmpProspectingWeeklyAgendaReportAdditionalEmails,1) <> ";" Then tmpProspectingWeeklyAgendaReportAdditionalEmails = tmpProspectingWeeklyAgendaReportAdditionalEmails & ";"
										Send_To = Send_To & tmpProspectingWeeklyAgendaReportAdditionalEmails
									End If	
								End If
								
								'Got all the addresses so now break them up
								Send_To_Array = Split(Send_To,";")
								
								Response.Write("<br>Send_To: " & Send_To & "<br>")
		
								'HERE WE ACTUALLY SEND THE EMAIL
								For i = 0 to Ubound(Send_To_Array) -1
									'Send_To = "cgrecco@ocsaccess.com"
									Send_To = Send_To_Array(i)
									If ProspectingWeeklyAgendaReportEmailSubject <> "" Then 
										emailSubject = ProspectingWeeklyAgendaReportEmailSubject 
									Else 
										emailSubject = GetTerm("Prospecting") & " Weekly Agenda Report"
									End If
									emailBody = ""
									'Failsafe for dev
									sURL = Request.ServerVariables("SERVER_NAME")
									If Instr(ucase(sURL),"DEV.") <> 0 Then Send_To = "cgrecco@ocsaccess.com"
									emailBody = "Your " & GetTerm("Prospecting") & " Weekly Agenda is attached."
									fn3=Server.MapPath(fn)
									Response.Write(fn3 & "<br>")
'									SendMailWatt "mailsender@" & maildomain,Send_To,emailSubject,emailBody,fn3,GetTerm("Prospecting"),"Weekly Agenda Report"
									SendMailWatt "mailsender@" & maildomain,"rsmith@ocsaccess.com",emailSubject,emailBody,fn3,GetTerm("Prospecting"),"Weekly Agenda Report"
									CreateAuditLogEntry "Prospecting Weekly Agenda Report","Prospecting Weekly Agenda Report","Minor",0,"Prospecting Weekly Agenda Report Sent to " & Send_To 
									Response.Write("Sent the email to " & Send_To & "<br>")
									Response.Write("Sent the email, all done<br>")
								Next 
							Next
					End If

					WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				
				End If
			
			Else
			
				WriteResponse ("Skipping the client " & ClientKey & " because the Prospecting Weekly Agenda Report is turned off.<BR>")
			
			End If ' for the report being turned off
			
		End If	
		
	Else ' is the prospecting  module enabled
	
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


		WriteResponse ("Skipping the client " & ClientKey & " because the prospecting Module is not enabled.<BR>")
		
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
	SQL = SQL & ",'/directlaunch/prospecting/prospectingWeeklyAgendaReport_launch.asp'"
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