<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/mail.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.asp"-->

<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
 
<%

dummy = MUV_WRITE("NewFilterLogic",1)

'Response.Buffer = True  <-----
'Response.Expires = 0  <-----	These lines commented purposely. They keep the page from close when launched automatically. Can't use them.
'Response.Clear  <-----


'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page
'Usage = "http://{xxx}.{domain}.com/directLaunch/service/auto_generate_filter_changes_launch.asp?runlevel=run_now
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
	
		'To begin with, see if this client uses the the service module & the filter module
		'If they don't then don't bother running for then
		
		Response.Write("Seeing if we need to run this for " & ClientKey & "<br>")
		
		If TopRecordset.Fields("serviceModule") = "Enabled" AND TopRecordset.Fields("filterchangeModule") = 1 Then
	
			'The IF statement below makes sure that when run from DEV it only deos client keys with a d
			'and when run from LIVE it only does client keys without a d
			'Pretty smart, huh
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
			or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0)_
			or  ucase(ClientKey) = "DEMO" Then

												
				Call SetClientCnnString
				
				Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
				

				'****************************************************
				' Now see if automatic filter generation is on or off
				'****************************************************				
				'This is here so we only open it once for the whole page
				Set cnn_Settings_FieldService = Server.CreateObject("ADODB.Connection")
				cnn_Settings_FieldService.open (MUV_READ("ClientCnnString"))
				Set rs_Settings_FieldService = Server.CreateObject("ADODB.Recordset")
				rs_Settings_FieldService.CursorLocation = 3 
				SQL_Settings_FieldService = "SELECT * FROM Settings_FieldService"
				Set rs_Settings_FieldService = cnn_Settings_FieldService.Execute(SQL_Settings_FieldService)
				If not rs_Settings_FieldService.EOF Then
					AutoFilterChangeGenerationONOFF = rs_Settings_FieldService("AutoFilterChangeGenerationONOFF")
					AutoFilterChangeUseRegions = rs_Settings_FieldService("AutoFilterChangeUseRegions")
					AutoFilterChangeMaxNumTicketsPerDay = rs_Settings_FieldService("AutoFilterChangeMaxNumTicketsPerDay")
				Else
					AutoFilterChangeGenerationONOFF = 0
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
					CheckSchedulerArray = Split(CheckScheduler("Settings_FieldService","Schedule_FilterGeneration"),",")
					If CheckSchedulerArray(0) <> "1" Then
						AutoFilterChangeGenerationONOFF = 0 ' Just turn it off & let the page flow normally
					End If
					Response.Write("<b>CheckScheduler Results: " &  CheckSchedulerArray(1) & "&nbsp;&nbsp;(" & ClientKey  & ")</b><br>")
					CreateAuditLogEntry "Auto generate filter changes launch","Auto generate filter changes launch","Minor",0,"Auto generate filter changes launch Schedule check results: " & CheckSchedulerArray(1)
				End If
				'************************************
				' E O F  S C H E D U L E R  L O G I C
				'************************************


				WriteResponse "<font color='purple' size='24'>Start processing " & ClientKey  & "</font><br>"

			
				If MUV_READ("cnnStatus") = "OK" Then ' else it loops
					
					Response.Write("blah,blah,blah blah,blah,blah blah,blah,blah blah,blah,blah blah,blah,blah <br>")
						
						' First it will handle filters for existing service tickets %>
						<!--#include file="auto_generate_filter_changes_inc_existingTickets.asp"-->
						
						<%
						'OK, now lets start getting them based on their schedules
						
						'First thing:
						' If there is a MAX in place, see if all the posible tickets to generate is less than or equal to the max
						' is it is, just generate all of them
						
						







If 1 = 2 Then						
						Response.Write("<br><br> OK, we finished everything related to open service tickets, now lets start getting them based on their schedules <br><br>")
						
						
						'Before we do anything see if there is a maximum number set per day & if we have exceeded it
						
						If AutoFilterChangeMaxNumTicketsPerDay = 0 OR AutoFilterChangeMaxNumTicketsPerDay => NumberOfTicketsGeneratedSoFar Then
										
							' All filters sure or overdue today	
							SQLPendingFilterChange = " SELECT DISTINCT * CustID "
							SQLPendingFilterChange = SQLPendingFilterChange & " FROM FS_CustomerFilters WHERE "
							SQLPendingFilterChange = SQLPendingFilterChange & " CASE WHEN FS_CustomerFilters.FrequencyType='D' THEN DATEADD(day, FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
							SQLPendingFilterChange = SQLPendingFilterChange & " WHEN FS_CustomerFilters.FrequencyType='M' THEN DATEADD(day, 28*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
							SQLPendingFilterChange = SQLPendingFilterChange & " WHEN FS_CustomerFilters.FrequencyType='W' THEN DATEADD(day, 7*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
							SQLPendingFilterChange = SQLPendingFilterChange & " ELSE FS_CustomerFilters.LastChangeDateTime END"
							SQLPendingFilterChange = SQLPendingFilterChange & " <= DateAdd(day,0,getdate()) "
							SQLPendingFilterChange = SQLPendingFilterChange & " AND "
							SQLPendingFilterChange = SQLPendingFilterChange & "("
							SQLPendingFilterChange = SQLPendingFilterChange & " FS_CustomerFilters.CustID NOT IN "
							SQLPendingFilterChange = SQLPendingFilterChange & "(SELECT CustID FROM FS_ServiceMemosFilterInfo WHERE ServiceTicketID IN (Select MemoNumber "
							SQLPendingFilterChange = SQLPendingFilterChange & " FROM FS_ServiceMemos WHERE CurrentStatus='OPEN')) "
							SQLPendingFilterChange = SQLPendingFilterChange & ") "
							SQLPendingFilterChange = SQLPendingFilterChange & " ORDER BY CustID "
							
							Response.Write(SQLPendingFilterChange)
			
							Set rsPendingFilterChange = Server.CreateObject("ADODB.Recordset")
							rsPendingFilterChange.CursorLocation = 3 
							Set rsPendingFilterChange = cnnFilterChanges.Execute(SQLPendingFilterChange)
							Set rsCreateTicket = Server.CreateObject("ADODB.Recordset")
	
	
							
							If NOT rsPendingFilterChange.EOF Then
							
								Do While Not rsPendingFilterChange.EOF
								
									CustID = rsPendingFilterChange("CustID")
									FiltersToDo  = ""
									
									
									ReDim FilterList(1)
									
									SQLCreateTicket = "SELECT * FROM FS_CustomerFilters WHERE CustID = '" & CustID & "'"
									
									Set rsCreateTicket = cnnFilterChanges.Execute(SQLCreateTicket)
	
									If NOT rsCreateTicket.EOF Then
									
										Do While Not rsCreateTicket.EOF
										
											FilterList(Ubound(FilterList)) = rsCreateTicket("FilterIntRecID")
									
											FiltersToDo  = FiltersToDo  & rsCreateTicket("InternalRecordIdentifier") & ","
									
											ReDim Preserve FilterList(Ubound(FilterList) + 1 )
											
											'Response.Write("<br>Adding " & rsCreateTicket("InternalRecordIdentifier"))
											
											rsCreateTicket.Movenext
										Loop
										
									End If
									
									If Right(FiltersToDo,1) = "," Then FiltersToDo = Left(FiltersToDo,Len(FiltersToDo)-1)
									
									filters = ""
									
									For x = 0 to Ubound(FilterList)
										If FilterList(x) <> "" Then
											filters = filters & vbcrlf & "Filter: " & GetFilterIDByIntRecID(FilterList(x)) & " - " & GetFilterDescByIntRecID(FilterList(x))
											Response.Write("<br>Adding Filter: " & GetFilterIDByIntRecID(FilterList(x)) & " - " & GetFilterDescByIntRecID(FilterList(x)))
										End If
									Next
	
									NumberOfTicketsGeneratedSoFar = NumberOfTicketsGeneratedSoFar + 1
									
									If AutoFilterChangeMaxNumTicketsPerDay = 0 Then
										msgtext = "0 (Unlimited)"
									Else
										msgtext = AutoFilterChangeMaxNumTicketsPerDay 
									End If
							
									Response.Write("<br>NumberOfTicketsGeneratedSoFar  " & NumberOfTicketsGeneratedSoFar  & "<br> of the maximum of " & msgtext )

									Call SubmitTicket(CustID,FiltersToDo,filters)
							
									rsPendingFilterChange.Movenext
								Loop
							End If
						Else
							' The maximum allowed to be generated in 1 day has been met
							WriteResponse "<font color='purple' size='24'>Maximum tickets to generated has been met. Stopping further ticket creation. Generated " & NumberOfTicketsGeneratedSoFar  & "<br> of the maximum of " & msgtext & "  " & ClientKey  & "</font><br>"
						End If ' for max tickets check
End If						
						'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''								
						'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
						'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''								
						'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
						'EOF OK, we finished everything related to open service tickets, now lets start getting them based on their schedules
						'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''								
						'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
						'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''								
						'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	

					
''''''''''''''''''
''''''''''''''''''
				''''''''''''''''''
				''''''''''''''''''
				''''''''''''''''''
				''''''''''''''''''
				''''''''''''''''''
				''''''''''''''''''
							
				''''''''''''''''''
				''''''''''''''''''
				''''''''''''''''''
				''''''''''''''''''
				''''''''''''''''''
				''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
					
					Else
						If AutoFilterChangeGenerationONOFF <> 1 Then 
							WriteResponse "<font color='purple' size='24'>Auto generate filter changes turned off for " & ClientKey  & "</font><br>"
						Else
							WriteResponse "<font color='purple' size='24'>settings_filedservice has no records for client " & ClientKey  & "</font><br>"
						End If
					End If

					
					cnnFilterChanges.Close
					Set rsServiceTickets = Nothing
					Set rsFilterChanges = Nothing
					Set cnnFilterChanges = Nothing

''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
						
	
										
						WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				
				End If
			
		End If	
		
	Else ' is the module enabled
	
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


		WriteResponse ("Skipping the client " & ClientKey & " because either the service module or the filter change module is not enabled.<BR>")
		
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
	SQL = SQL & ",'MCS Rebuild Helper'"
	SQL = SQL & ",'/directlaunch/bizintel/mcs_rebuild_helper_launch.asp'"
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


Sub SubmitTicket(passedCustID,passedFiltersToDo,passedfilters)

	On Error Goto 0 
	
	'Lookup the cust to get all the contact fields, etc


	If GetPOSTParams("SERVICEMEMOURL1MPLEXFORMAT") = 1 then 'see if using metroplex format
		
		
		'Post to MDS goes here
		data = "create_service_request&account="
		data = data & passedCustID
		data = data & "&prob=FILTER CHANGE - " & passedfilters & "%0A%0D"
		data = data & "Opened by: System Auto Filter Change Generation" 
		data = data & "&probloc=" & ProblemLocation 
		data = data & "&sbn=" & ContactName 
		data = data & "&sbe=" & ContactEmail
		data = data & "&sbp=" & ContactPhone 
		data = data & "&contnm=" & ContactName
		data = data & "&st=" & "OPEN"
		data = data & "&md=" & GetPOSTParams("Mode")
		data = data & "&serno=" & GetPOSTParams("Serno")
		data = data & "&src=" & "MDS Insight"
		data = data & "&usr=0" 
		data = data & "&pm=" & Server.URLencode("Filter Change")
		data = data & "&rids=" & passedFiltersToDo 

writeresponse data
	
		If GetPOSTParams("NeverPutOnHold") = 0  Then data = data & "&hld=1" Else data = data & "&hld=0" 'I know it's wierd. It's the opposite of how it is stored in the table
		
		
	Else
	
		'Post to APIs goes here
				
		data = ""
				
		data = data & "<POST_DATA>"
		data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
		data = data & "<ACCOUNT_NUM>" & AccountNumber & "</ACCOUNT_NUM>"
		data = data & "<PROBLEM_LOCATION>" & ProblemLocation & "</PROBLEM_LOCATION>"
		data = data & "<PROBLEM_DESCRIPTION>FILTER CHANGE - " & filters 
		data = data & "<RECORD_TYPE>SENDSERVICEMSG</RECORD_TYPE>"
		data = data & "<RECORD_SUBTYPE>OPEN</RECORD_SUBTYPE>"
		data = data & "<SERVICE_TICKET_NUMBER>AUTO</SERVICE_TICKET_NUMBER>"
		data = data & "<COMPANY_NAME>" & Replace(GetCustNameByCustNum(AccountNumber),"&","&amp;") & "</COMPANY_NAME>"
		data = data & "<SUBMISSION_SOURCE>MDS Insight</SUBMISSION_SOURCE>"
		data = data & "<SERNO>" & GetPOSTParams("Serno") & "</SERNO>"
		data = data & "<CLIENT_ID>" & MUV_READ("ClientID") & "</CLIENT_ID>"
		data = data & "<USER_NO>" & Session("UserNo") & "</USER_NO>"
		data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
		data = data & "<SUBMITTED_BY_PHONE>" & ContactPhone & "</SUBMITTED_BY_PHONE>"
		data = data & "<SUBMITTED_BY_EMAIL>" & ContactEmail & "</SUBMITTED_BY_EMAIL>"
		data = data & "<SUBMITTED_BY_NAME>" & ContactEmail & "</SUBMITTED_BY_NAME>"
		data = data & "<FILTER_REC_IDS>" & passedFiltersToDo & "</FILTER_REC_IDS>"
		data = data & "</POST_DATA>"
		data = Replace(data ,"&","&amp;")
		data = Replace(data ,chr(34),"")
	
	
	End If
	

	DO_Post = 0
	
	Do_Post = GetPOSTParams("ServiceMemoURL1ONOFF") 
			
	If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0

	If cint(Do_Post) = 1 Then
			
		Description = "Post to " & GetPOSTParams("ServiceMemoURL1")


		CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
		Description = "data:" & data 
		CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")

	
		CreateINSIGHTAuditLogEntry sURL,"Post Loop "& x,GetPOSTParams("Mode")

		Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
			
		httpRequest.Open "POST", GetPOSTParams("ServiceMemoURL1"), False
			
		httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		httpRequest.Send data
		
		If httpRequest.status = 200 THEN 
						
			If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
						
				Description ="success! httpRequest.status returned " & httpRequest.status & " when posting a filter change ticket"& "<br>"
				Description = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
				Description = Description & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				Description = Description & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
				Description = Description & "POSTED DATA:" & data & "<br>"
				Description = Description & "SERNO:" & GetPOSTParams("Serno") & "<br>"
				Description = Description & "MODE:" & GetPOSTParams("Mode") & "<br>"

				CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
	
				Description = "OPEN,    "
				Description = Description & "Account: "  & passedCustID & " - " & GetCustNameByCustNum(passedCustID)
				Description = Description & ",    Description: "  & "FILTER CHANGE"
				CreateAuditLogEntry "Service Memo Added","Service Memo Added","Major",0,Description

	
			Else
				'FAILURE
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting a filter change ticket"& "<br>"
				emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
				emailBody = emailBody & "POSTED DATA:" & data & "<br>"
				emailBody = emailBody & "SERNO:" & GetPOSTParams("Serno") & "<br>"
				emailBody = emailBody & "MODE:" & GetPOSTParams("Mode") & "<br>"
							
				SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"

				Description = emailBody 
				CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
				
			End If
						
		Else
							
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting a filter change ticket"& "<br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
			emailBody = emailBody & "POSTED DATA:" & data & "<br>"
			emailBody = emailBody & "SERNO:" & GetPOSTParams("Serno") & "<br>"
			emailBody = emailBody & "MODE:" & GetPOSTParams("Mode") & "<br>"
	
			SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
					
			Description = emailBody 
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
										
		End If
		
	End IF


WriteResponse httpRequest.status & "<br>"
WriteResponse httpRequest.responseText & "<br>"
WriteResponse Description & "<br>"
	
	' Write Audit trail first, then post
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	
	Set rs8 = Server.CreateObject("ADODB.Recordset")
	rs8.CursorLocation = 3 
	'Set rs8 = cnn8.Execute(SQL)
	set rs8 = Nothing
		

End Sub

Sub ArrayCull(ByRef arr)
  Dim i, dict
  If IsArray(arr) Then
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(arr)
      If Not dict.Exists(arr(i)) Then
        Call dict.Add(arr(i), arr(i))
      End If
    Next
    arr = dict.Items
  End If
End Sub
'************************************************************************************
'************************************************************************************
'Subs and funcs end here
'************************************************************************************

%>