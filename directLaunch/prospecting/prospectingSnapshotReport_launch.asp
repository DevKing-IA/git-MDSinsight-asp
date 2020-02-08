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
'Usage = "http://{xxx}.{domain}.com/directLaunch/prospecting/prospectingSnapshotReport_launch.asp?runlevel=run_now&force=run_now
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
				' Now see if the prospecting snapshot report is on or off
				'**************************************************************				
				Set cnn_Settings_Global = Server.CreateObject("ADODB.Connection")
				cnn_Settings_Global.open (Session("ClientCnnString"))
				Set rs_Settings_Global = Server.CreateObject("ADODB.Recordset")
				rs_Settings_Global.CursorLocation = 3 
				SQL_Settings_Global = "SELECT * FROM Settings_Global"
				Set rs_Settings_Global = cnn_Settings_Global.Execute(SQL_Settings_Global)
				If not rs_Settings_Global.EOF Then
					ProspSnapshotOnOff = rs_Settings_Global("ProspSnapshotOnOff")
					ProspSnapshotInsideSales = rs_Settings_Global("ProspSnapshotInsideSales")
					ProspSnapshotOutsideSales = rs_Settings_Global("ProspSnapshotOutsideSales")
					ProspSnapshotUserNos = rs_Settings_Global("ProspSnapshotUserNos")
					ProspSnapshotAdditionalEmails = rs_Settings_Global("ProspSnapshotAdditionalEmails")
					ProspSnapshotEmailSubject = rs_Settings_Global("ProspSnapshotEmailSubject")
				Else
					ProspSnapshotOnOff = vbFalse
				End If
				Set rs_Settings_Global = Nothing
				cnn_Settings_Global.Close
				Set cnn_Settings_Global = Nothing

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
					CheckSchedulerArray = Split(CheckScheduler("Settings_Prospecting","Schedule_ProspectingSnapshotReportGeneration"),",")
					If CheckSchedulerArray(0) <> "1" Then
						ProspSnapshotOnOff = 0 ' Just turn it off & let the page flow normally
					End If
					Response.Write("<b>CheckScheduler Results: " &  CheckSchedulerArray(1) & "&nbsp;&nbsp;(" & ClientKey  & ")</b><br>")
					CreateAuditLogEntry "Prospecting Snapshot Report Launch","Prospecting Snapshot Report Launch","Minor",0,"Prospecting Snapshot Report Schedule check results: " & CheckSchedulerArray(1)
				End If
				'************************************
				' E O F  S C H E D U L E R  L O G I C
				'************************************

				If MUV_READ("FORCE") = "1" Then
					CreateAuditLogEntry "Prospecting Snapshot Report Launch","Prospecting Snapshot Report Forced","Minor",0,"Prospecting Snapshot report FORCE is set to 1."									
				End If
					
				If ProspSnapshotOnOff = 1 Then
	
					CreateAuditLogEntry "Prospecting Snapshot Report Launch","Prospecting Snapshot Report Launch","Minor",0,"Prospecting Snapshot Report ran."					
	
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
						' Start individual guts of the report										
						
						Set Pdf = Server.CreateObject("Persits.Pdf")
						Set Doc = Pdf.CreateDocument
						Response.Write("<br>" & baseURL & "directlaunch/prospecting/prospectingSnapshotReport.asp?c=" & ClientKey & ",scale=0.6; hyperlinks=true; drawbackground=true; landscape=true<br>")
						Doc.ImportFromUrl baseURL & "directlaunch/prospecting/prospectingSnapshotReport.asp?c=" & ClientKey, "scale=0.6; hyperlinks=true; drawbackground=true; landscape=true"
						Response.Write(baseURL  & "directlaunch/prospecting/prospectingSnapshotReport.asp?c=" & ClientKey & "<br>")


						fn = "\clientfilesV\" & trim(ClientKey) &"\z_pdfs\" & formatDateTime(Now(),2) & "-" & formatdatetime(Now(),4) & "_prospectingSnapshot.pdf"
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



						Response.Write("OK, got here<br>")
						'OK, now start breaking out the email addresses


						Send_To=""
						
						If ProspSnapshotInsideSales = 1 Then
							'Get all the inside sales manager email addresses
							Set cnn_CheckAlerts = Server.CreateObject("ADODB.Connection")
							cnn_CheckAlerts.open Session("ClientCnnString")
							Set rs_CheckAlerts = Server.CreateObject("ADODB.Recordset")
							rs_CheckAlerts.CursorLocation = 3 
							SQL_CheckAlerts = "SELECT userEmail FROM tblUsers WHERE userType = 'Inside Sales Manager' and userArchived <> 1" 
							Set rs_CheckAlerts = cnn_CheckAlerts.Execute(SQL_CheckAlerts)
							If not rs_CheckAlerts.EOF Then
								Do
									If rs_CheckAlerts("userEmail") <> "" AND Not IsNull(rs_CheckAlerts("userEmail")) Then Send_To = Send_To & rs_CheckAlerts("userEmail") & ";"
									rs_CheckAlerts.MoveNext
								Loop Until rs_CheckAlerts.Eof
							End If
							Set rs_CheckAlerts = Nothing
							cnn_CheckAlerts.Close
							Set cnn_CheckAlerts = Nothing
						End If		


						If ProspSnapshotOutsideSales = 1 Then
						
							'Get all the outside sales manager email addresses
							Set cnn_CheckAlerts = Server.CreateObject("ADODB.Connection")
							cnn_CheckAlerts.open Session("ClientCnnString")
							Set rs_CheckAlerts = Server.CreateObject("ADODB.Recordset")
							rs_CheckAlerts.CursorLocation = 3 
							SQL_CheckAlerts = "SELECT userEmail FROM tblUsers WHERE userType = 'Outside Sales Manager' and userArchived <> 1" 
							Set rs_CheckAlerts = cnn_CheckAlerts.Execute(SQL_CheckAlerts)
							If not rs_CheckAlerts.EOF Then
								Do
									If rs_CheckAlerts("userEmail") <> "" AND Not IsNull(rs_CheckAlerts("userEmail")) Then Send_To = Send_To & rs_CheckAlerts("userEmail") & ";"
									rs_CheckAlerts.MoveNext
								Loop Until rs_CheckAlerts.Eof
							End If
							Set rs_CheckAlerts = Nothing
							cnn_CheckAlerts.Close
							Set cnn_CheckAlerts = Nothing
						End If	
			
						'Now see if there any additionals
						If ProspSnapshotAdditionalEmails <> "" and not IsNull(ProspSnapshotAdditionalEmails) Then
							tmpProspSnapshotAdditionalEmails  = trim(ProspSnapshotAdditionalEmails)		
							If Len(tmpProspSnapshotAdditionalEmails) > 1 Then
								If Right(tmpProspSnapshotAdditionalEmails,1) <> ";" Then tmpProspSnapshotAdditionalEmails = tmpProspSnapshotAdditionalEmails & ";"
								Send_To = Send_To & tmpProspSnapshotAdditionalEmails
							End If	
						End If
						
						'Get user based emails
						If ProspSnapshotUserNos <> "" Then
							UserNoList = Split(ProspSnapshotUserNos,",")
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
							If ProspSnapshotEmailSubject <> "" Then emailSubject = ProspSnapshotEmailSubject Else ProspSnapshotEmailSubject = GetTerm("Prospecting") & " Weekly Snapshot"
							emailBody = ""
							'Failsafe for dev
							sURL = Request.ServerVariables("SERVER_NAME")
							If Instr(ucase(sURL),"DEV.") <> 0 Then Send_To = "rich@ocsaccess.com"
							If Instr(ucase(sURL),"DEV2.") <> 0 Then Send_To = "cgrecco@ocsaccess.com"
							emailBody = "Your automatic " & GetTerm("Prospecting") & " Weekly Snapshot is attached."
							fn3=Server.MapPath(fn)
							Response.Write(fn3 & "<br>")
							SendMailWatt "mailsender@" & maildomain,Send_To,emailSubject,emailBody,fn3,GetTerm("Service"),"Threshold Report"
							CreateAuditLogEntry "Automatic Prospecting Weekly Snapshot Report","Automatic Prospecting Weekly Snapshot Report","Minor",0,"Automatic Prospecting Weekly Snapshot Report Sent to " & Send_To 
							Response.Write("Sent the email to " & Send_To & "<br>")
							Response.Write("Sent the email, all done<br>")
						Next 
						
						

						
						' End individual guts of the report										
			
						WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				
				End If
			
			Else
			
				WriteResponse ("Skipping the client " & ClientKey & " because the Prospecting Snapshot Report is turned off.<BR>")
			
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


		WriteResponse ("Skipping the client " & ClientKey & " because the prospecting Module module is not enabled.<BR>")
		
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