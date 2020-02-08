<% @ Language = VBScript %>
<%
Response.Buffer = True
Response.Expires = 0
Response.Clear
%>
<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.asp"-->
<%dummy=MUV_Write("ClientID","") 'Need this here

ClientKey = Request.QueryString("c")

SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open InsightCnnString
'Response.Write("InsightCnnString:" & InsightCnnString & "<br>")

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3
'Response.Write("SQL:" & SQL& "<br>")

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Automatic service ticket threashold report<%
	Response.End
Else
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & Recordset.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & Recordset.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	Recordset.close
	Connection.close	
End If	



'This is here so we only open it once for the whole page
Set cnnthresholdReport = Server.CreateObject("ADODB.Connection")
cnnthresholdReport.open (Session("ClientCnnString"))
Set rsthresholdReport = Server.CreateObject("ADODB.Recordset")

SQLthresholdReport = "SELECT * FROM Settings_FieldService"
Set rsthresholdReport = cnnthresholdReport.Execute(SQLthresholdReport)
If not rsthresholdReport.EOF Then
   ServiceTicketthresholdReportONOFF = rsthresholdReport("ServiceTicketthresholdReportONOFF")
   ServiceTicketthresholdReportOnlyUndispatched = rsthresholdReport("ServiceTicketthresholdReportOnlyUndispatched")
   ServiceTicketthresholdReportOnlySkipFilterChanges = rsthresholdReport("ServiceTicketthresholdReportOnlySkipFilterChanges")
   ServiceTicketthresholdReportthresholdHours = rsthresholdReport("ServiceTicketthresholdReportthresholdHours")
Else
	ServiceTicketthresholdReportONOFF = 0
End If


Set rsthresholdReport = Nothing
cnnthresholdReport.Close
Set cnnthresholdReport = Nothing

If ServiceTicketthresholdReportONOFF <> 1 Then
	%>MDS Insight: The automatic service ticket threashold report is not turned on.
	<%
	Response.End
End IF

%>



<body bgcolor="#FFFFFF" text="#000000" link="#000080" vlink="#000080" alink="#000080" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">

<table border="0" width="650" align="center">
	<tr>
		<td width="100%">
		<table border="0" width="100%">
			<tr><td><img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png"></td></tr>
			<tr>
				<td width="20%" height="16">
				<p align="right"><font face="Arial" size="1">
				&nbsp;</font> </p>
				</td>
			</tr>
		</table>
		<table border="1" width="855" style="border-collapse: collapse" bordercolor="#111111" cellpadding="2">
			<tr>
				<td width="50%" height="16">
					<p align="left"><b><font face="Arial" size="2">MDS Insight: Service Ticket Threshold Report</font></b></p>
				</td>
				<td width="50%" height="16" valign="middle" align="right">
					<font face="Arial" size="1">Report Generated: <%= WeekDayName(WeekDay(DateValue(Now()))) %>&nbsp;<%= Now() %><br></font>
				</td>
			</tr>
			<tr><td width="100%" colspan="2"><br><br></td></tr>
		</table>
		<%
					
			'**************************************************************************'
			'***Now perform outer SQL STMT to get the appropriate service calls file***'
			'**************************************************************************'

			SQL = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' "
			SQL = SQL & "order by submissionDateTime"

'Response.Write("<br><br>" & SQL & "<br><br>")

			Set cnnThreshold = Server.CreateObject("ADODB.Connection")
			cnnThreshold.open (Session("ClientCnnString"))
			Set rsThreshold  = Server.CreateObject("ADODB.Recordset")
			rsThreshold.CursorLocation = 3 
			rsThreshold.Open SQL, cnnThreshold
			
			
			If Not rsThreshold.EOF Then
			
				nRecCount = rsThreshold.RecordCount
				RowCount = 0
				CurrentPage = 1
				NumberPages = Int(nRecCount/27) + 1 '***27 Table Cells Fit On One Printable Page***'

		%>
				<table border="1" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32">

				<tr>
				<td width="95" align="center" height="18">
				<font face="Arial" style="font-size: 8pt"><b>
				Ticket #</b></font></td>
				<td width="320" align="center" height="25">
				<font face="Arial" style="font-size: 8pt"><b>
				Account</b></font></td>
				<td width="135" align="center" height="18"><b>
				<font face="Arial" style="font-size: 8pt">Submitted</font></b></td>
				<td width="100" align="center" height="18"><b>
				<font face="Arial" style="font-size: 8pt">Elapsed Time</font></b></td>
				<td width="45" align="center" height="18">
				<font face="Arial" style="font-size: 8pt"><b>
				Stage</b></font></td>
				</tr>
			<%
			
			Do While Not (rsThreshold.EOF)
			
				elapsedMinutes = ServiceCallElapsedMinutesOpenTicket(rsThreshold("MemoNumber"))
				If elapsedMinutes < 1 Then elapsedMinutes = 1 ' If it has been less than 1 minute, just show 1 anyway
				elapsedString = ""
				minutesInServiceDay = GetNumberOfMinutesInServiceDay()
				elapsedDays = 	elapsedMinutes \ minutesInServiceDay
				If int(elapsedDays) > 0 Then
					elapsedMinutes = elapsedMinutes - (int(elapsedDays) * minutesInServiceDay)
					elapsedString = elapsedDays & "d "
				End If
				elapsedHours = elapsedMinutes \ 60
				If int(elapsedHours) > 0 Then 
					elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
					elapsedString = elapsedString  & elapsedHours & "h "
				End If
				If int(elapsedMinutes) > 0 Then
					elapsedString = elapsedString  & elapsedMinutes & "m"
				End If

				ShowRecord = True
				
				If ServiceTicketthresholdReportOnlyUndispatched = 1 Then
					If GetServiceTicketCurrentStage(rsThreshold("MemoNumber")) <> "Received" Then ShowRecord = False ' Only show awaiting dispatch
				End If

				If ServiceTicketthresholdReportOnlySkipFilterChanges = 1 Then
					If rsThreshold("FilterChange") = 1 Then ShowRecord = False ' Skip filter chanegs
				End If

				
				If ShowRecord = True Then

					If Round(ServiceCallElapsedMinutesOpenTicket(rsThreshold("MemoNumber"))/60,0) > ServiceTicketthresholdReportthresholdHours Then 
		
	
						RowCount = RowCount + 1
						submissionDateTime=rsThreshold("SubmissionDateTime")
						TicketStageVar = GetServiceTicketCurrentStage(rsThreshold("MemoNumber"))
						TicketStageVar = Replace(TicketStageVar ,"Dispatch Acknowledged","Dispatch Ack")
						TicketStageVar = Replace(TicketStageVar ,"Received","Awaiting Dispatch")
						TicketStageVar = Replace(TicketStageVar ,"Released","Awaiting Dispatch")
						TicketStageVar = Replace(TicketStageVar ,"Dispatch","Disp")
						TicketStageVar = Replace(TicketStageVar ,"Disped","Dispatched")
	
					%>
						<tr>
							<td width="15%" height="22" align="center">
							<font face="Arial" style="font-size: 8pt"><%= rsThreshold("MemoNumber") %></font></td>
							
							<td width="45%" height="25" align="center">
							<font face="Arial" style="font-size: 8pt"><%= rsThreshold("AccountNumber") & " " & rsThreshold("Company") %></font></td>
							
							<td width="10%" height="17" align="center">
							<font face="Arial" style="font-size: 8pt"><%= padDate(MONTH(submissionDateTime),2) & "/" & padDate(DAY(submissionDateTime),2) & "/" & padDate(RIGHT(YEAR(submissionDateTime),2),2) %>&nbsp;</font></td>
							
							<td width="15%" height="17" align="center"><font face="Arial" style="font-size: 8pt"><%= elapsedString %></font></td>
							
							<td width="15%" height="17" align="center"><font style="font-size: 8pt"><%= TicketStageVar %></font></td>
						</tr>
						
	<%				End If
	
				End If ' for show record
		
			rsThreshold.MoveNext
			Loop
			
			End If %>
				<br/><br/>
			</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</body>

</html>
