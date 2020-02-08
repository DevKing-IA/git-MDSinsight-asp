<% @ Language = VBScript %>

<head>
<style type="text/css">
.auto-style1 {
	text-align: right;
}
</style>
</head>

<%
Response.Buffer = True
Response.Expires = 0
Response.Clear
Dim PageNum, RowCount, FontSizeVar
FontSizeVar = 9
PageNum = 0
NoBreak = False
'Adjust = -3
'MAdjust = -1
PageWidth = 1200


%>
<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.asp"-->
<%

dummy=MUV_Write("ClientID","") 'Need this here

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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Service Ticket Daily Notes Report<%
	Response.End
Else
	ClientCnnStringvar = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	ClientCnnStringvar = ClientCnnStringvar  & ";Database=" & Recordset.Fields("dbCatalog")
	ClientCnnStringvar = ClientCnnStringvar & ";Uid=" & Recordset.Fields("dbLogin")
	ClientCnnStringvar = ClientCnnStringvar & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	dummy = MUV_Write("ClientCnnString",ClientCnnStringvar)
	Recordset.close
	Connection.close	
End If	

Session("ClientCnnString") = MUV_READ("ClientCnnString")


'This is here so we only open it once for the whole page
Set cnn_Settings_FieldService = Server.CreateObject("ADODB.Connection")
cnn_Settings_FieldService.open (MUV_READ("ClientCnnString"))
Set rs_Settings_FieldService = Server.CreateObject("ADODB.Recordset")
rs_Settings_FieldService.CursorLocation = 3 
SQL_Settings_FieldService = "SELECT * FROM Settings_FieldService"
Set rs_Settings_FieldService = cnn_Settings_FieldService.Execute(SQL_Settings_FieldService)
If not rs_Settings_FieldService.EOF Then
	FieldServiceNotesReportOnOff = rs_Settings_FieldService("FieldServiceNotesReportOnOff")
Else
	FieldServiceNotesReportOnOff = 0
End If
Set rs_Settings_FieldService = Nothing
cnn_Settings_FieldService.Close
Set cnn_Settings_FieldService = Nothing


Yesterday = DateAdd("d",-1, Now())
'Yesterday = cDate("12-27-2017")
YesterdayDayName = WeekdayName(Weekday(Yesterday))

If YesterdayDayName = "Saturday" Then
	LastBusinessDay = DateAdd("d",-1, Yesterday)
	LastBusinessDayName = WeekdayName(Weekday(LastBusinessDay))
ElseIf YesterdayDayName = "Sunday" Then
	LastBusinessDay = DateAdd("d",-2, Yesterday)
	LastBusinessDayName = WeekdayName(Weekday(LastBusinessDay))
Else
	LastBusinessDay = Yesterday
	LastBusinessDayName = YesterdayDayName
End If


FieldServiceNotesDay = Day(LastBusinessDay)
FieldServiceNotesMonth = Month(LastBusinessDay)
FieldServiceNotesYear = Year(LastBusinessDay)



If FieldServiceNotesReportOnOff <> 1 Then
	%>MDS Insight: The Service Ticket Daily Notes Report is not turned on.
	<%
	Response.End
End IF

%>



<body bgcolor="#FFFFFF" text="#000000" link="#000080" vlink="#000080" alink="#000080" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
<table border="0" width="<%=PageWidth%>" align="center">
	<tr>
		<td width="100%" align="center">

		<%
		
			'*******************************************************
			'*** First get all of the summary information
			'*** The summary is a listing of technicians who
			'*** have tickets on a given business day, and the 
			'*** number of notes they have entered in total
			'*******************************************************
		
			Call PageHeader

			LinesPerPage = 2
			

			SQLDailyFieldServiceNotesTechsLoop = "SELECT DISTINCT EnteredByUserNo FROM FS_ServiceMemosNotes WHERE "
			SQLDailyFieldServiceNotesTechsLoop = SQLDailyFieldServiceNotesTechsLoop & " (DAY(RecordCreationDateTime) = " & FieldServiceNotesDay & " AND "
			SQLDailyFieldServiceNotesTechsLoop = SQLDailyFieldServiceNotesTechsLoop & "MONTH(RecordCreationDateTime) =  " & FieldServiceNotesMonth & " AND "
			SQLDailyFieldServiceNotesTechsLoop = SQLDailyFieldServiceNotesTechsLoop & "YEAR(RecordCreationDateTime) =  " & FieldServiceNotesYear & ")"
			
			'Response.Write("SQLDailyFieldServiceNotesTechsLoop :" & SQLDailyFieldServiceNotesTechsLoop & "<br>")
			
			Set cnnDailyFieldServiceNotesTechsLoop = Server.CreateObject("ADODB.Connection")
			cnnDailyFieldServiceNotesTechsLoop.open(Session("ClientCnnString"))
			Set rsDailyFieldServiceNotesTechsLoop = Server.CreateObject("ADODB.Recordset")
			rsDailyFieldServiceNotesTechsLoop.CursorLocation = 3 
			Set rsDailyFieldServiceNotesTechsLoop = cnnDailyFieldServiceNotesTechsLoop.Execute(SQLDailyFieldServiceNotesTechsLoop)
	
			'************************************************************
			'If there are technicians with service calls on the
			'requested business day, loop through and check to see
			'the number of notes each one has entered
			'************************************************************
			
			If NOT rsDailyFieldServiceNotesTechsLoop.EOF Then
			
				%>
				<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
					<tr>
						<td>
						<font face="Consolas">
						<hr>
						<center><h2><%= GetTerm("Field Service") %> Notes Summary for <%= LastBusinessDayName %>&nbsp;<%= FormatDateTime(LastBusinessDay,2) %></h2></center>
						<hr>
						</font>
						</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
			
				<%
			
				Do While Not rsDailyFieldServiceNotesTechsLoop.EOF
								
					notesEnteredByUserNo = rsDailyFieldServiceNotesTechsLoop("EnteredByUserNo")
					notesEnteredByUserName = GetUserDisplayNameByUserNo(notesEnteredByUserNo)
					userType = getUserType(notesEnteredByUserNo)
									
					Set cnnCurrentUserNotesCount = Server.CreateObject("ADODB.Connection")
					cnnCurrentUserNotesCount.open Session("ClientCnnString")
					Set rsCurrentUserNotesCount = Server.CreateObject("ADODB.Recordset")
					rsCurrentUserNotesCount.CursorLocation = 3 					
				
					'***********************************************************************************************
					'Get all the notes for this technician from the new FS_ServiceMemosNotes table
					'on the requested busines day
					'***********************************************************************************************
					SQLcurrentUserNotesCount = "SELECT COUNT(*) AS TechNotesCount FROM FS_ServiceMemosNotes WHERE EnteredByUserNo = " & notesEnteredByUserNo
					SQLcurrentUserNotesCount = SQLcurrentUserNotesCount & " AND DAY(RecordCreationDateTime) = " & FieldServiceNotesDay & " AND "
					SQLcurrentUserNotesCount = SQLcurrentUserNotesCount & "MONTH(RecordCreationDateTime) =  " & FieldServiceNotesMonth & " AND "
					SQLcurrentUserNotesCount = SQLcurrentUserNotesCount & "YEAR(RecordCreationDateTime) =  " & FieldServiceNotesYear
					
					'Response.Write("SQLcurrentUserNotesCount :" & SQLcurrentUserNotesCount  & "<br>")
					
					Set rsCurrentUserNotesCount = cnnCurrentUserNotesCount.Execute(SQLcurrentUserNotesCount)
											
					If NOT rsCurrentUserNotesCount.EOF Then
					
						'*************************************************************************
						'Determine if the tickets has notes or service images
						'*************************************************************************
						currentUserNotesCount = rsCurrentUserNotesCount("TechNotesCount")
					
						%>
						<tr>
							<td>
								<font face="Consolas" style="font-size: 14pt">
								<%= userType %>: <strong><%= notesEnteredByUserName %></strong> has <strong><%= currentUserNotesCount %></strong> Tickets With Notes</font>
							</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
						</tr>
						
						<%
												
					End If
					
					'*************************************************************************
					'Set the next technician user number
					'*************************************************************************
					
					rsDailyFieldServiceNotesTechsLoop.MoveNext
			
				Loop
		
			End If
		 %>
			
					<br/><br/>
			</table>
			</td>
		</tr>
		<tr>
		<td>
	
<% 

	set rsDailyFieldServiceNotesTechsLoop = Nothing
	cnnDailyFieldServiceNotesTechsLoop.Close	
	set cnnDailyFieldServiceNotesTechsLoop = Nothing
		
	set rsCurrentUserNotesCount = Nothing
	cnnCurrentUserNotesCount.Close	
	set cnnCurrentUserNotesCount = Nothing


	Call Footer


		
	'*******************************************************
	'*** END END END END END END END END END END END END END 
	'*** This section is the first page which prints all the
	'*** service ticket notes summary information
	'*******************************************************


	'*******************************************************
	'*** This section is the detail section of daily field
	'*** service notes - it shows only the user that
	'*** have entered notes
	'*******************************************************
	
	SQLDailyFieldServiceNotesTechsLoop = "SELECT DISTINCT EnteredByUserNo FROM FS_ServiceMemosNotes WHERE "
	SQLDailyFieldServiceNotesTechsLoop = SQLDailyFieldServiceNotesTechsLoop & " (DAY(RecordCreationDateTime) = " & FieldServiceNotesDay & " AND "
	SQLDailyFieldServiceNotesTechsLoop = SQLDailyFieldServiceNotesTechsLoop & "MONTH(RecordCreationDateTime) =  " & FieldServiceNotesMonth & " AND "
	SQLDailyFieldServiceNotesTechsLoop = SQLDailyFieldServiceNotesTechsLoop & "YEAR(RecordCreationDateTime) =  " & FieldServiceNotesYear & ")"
	
	'Response.Write("SQLDailyFieldServiceNotesTechsLoop :" & SQLDailyFieldServiceNotesTechsLoop & "<br>")
	
	Set cnnDailyFieldServiceNotesTechsLoop = Server.CreateObject("ADODB.Connection")
	cnnDailyFieldServiceNotesTechsLoop.open(Session("ClientCnnString"))
	Set rsDailyFieldServiceNotesTechsLoop = Server.CreateObject("ADODB.Recordset")
	rsDailyFieldServiceNotesTechsLoop.CursorLocation = 3 
	Set rsDailyFieldServiceNotesTechsLoop = cnnDailyFieldServiceNotesTechsLoop.Execute(SQLDailyFieldServiceNotesTechsLoop)

	'************************************************************
	'If there are technicians with service calls on the
	'requested business day, loop through and check to see
	'the number of notes each one has entered
	'************************************************************
	
	If NOT rsDailyFieldServiceNotesTechsLoop.EOF Then
		
		Call PageHeader
		
		PreviousNotesEnteredByUserNo = 0
		
		Set cnnCurrentUserNotesCount = Server.CreateObject("ADODB.Connection")
		cnnCurrentUserNotesCount.open Session("ClientCnnString")
		Set rsCurrentUserNotesCount = Server.CreateObject("ADODB.Recordset")
		rsCurrentUserNotesCount.CursorLocation = 3 	
		
		Set cnnServiceTicketDetails = Server.CreateObject("ADODB.Connection")
		cnnServiceTicketDetails.open Session("ClientCnnString")
		Set rsServiceTicketDetails = Server.CreateObject("ADODB.Recordset")
		rsServiceTicketDetails.CursorLocation = 3 		
			
		Do While Not rsDailyFieldServiceNotesTechsLoop.EOF

			FontSizeVar = 9
			LinesPerPage = 32
	
			TechnicianUserNo = rsDailyFieldServiceNotesTechsLoop("EnteredByUserNo")
			TechnicianName = GetUserDisplayNameByUserNo(TechnicianUserNo)
	
			'***********************************************************************************************
			'Get all the notes for this technician from the new FS_ServiceMemosNotes table
			'on the requested busines day
			'***********************************************************************************************
			SQLcurrentUserNotesCount = "SELECT * FROM FS_ServiceMemosNotes WHERE EnteredByUserNo = " & TechnicianUserNo 
			SQLcurrentUserNotesCount = SQLcurrentUserNotesCount & " AND DAY(RecordCreationDateTime) = " & FieldServiceNotesDay & " AND "
			SQLcurrentUserNotesCount = SQLcurrentUserNotesCount & "MONTH(RecordCreationDateTime) =  " & FieldServiceNotesMonth & " AND "
			SQLcurrentUserNotesCount = SQLcurrentUserNotesCount & "YEAR(RecordCreationDateTime) =  " & FieldServiceNotesYear & " "
			SQLcurrentUserNotesCount = SQLcurrentUserNotesCount & "ORDER BY ServiceTicketID, RecordCreationDateTime DESC "
			
			'Response.Write("SQLcurrentUserNotesCount :" & SQLcurrentUserNotesCount  & "<br>")
			
			Set rsCurrentUserNotesCount = cnnCurrentUserNotesCount.Execute(SQLcurrentUserNotesCount)
									
			If NOT rsCurrentUserNotesCount.EOF Then
	
				If (cInt(PreviousNotesEnteredByUserNo) <> cInt(TechnicianUserNo)) Then
					%>
					<table border="0" width="<%=PageWidth%>">
						<tr>
							<td colspan="5">
								<font face="Consolas">
									<hr>
										<center><h2><%= GetTerm("Field Service") %> Notes for <%= TechnicianName %></h2></center>
									<hr>
								</font>
							</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td width="20%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=GetTerm("Customer")%></font></u></strong>
							</td>
							<td width="5%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Ticket #</font></u></strong>
							</td>
							<td width="5%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Status</font></u></strong>
							</td>		
							<td width="35%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Original Problem</font></u></strong>
							</td>
							<td width="35%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Service Notes</font></u></strong>
							</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
					</table>
					<%
				End If
				
				
				PreviousNotesServiceTicketID = 0
	
				Do While Not rsCurrentUserNotesCount.EOF
				
					notesEntered = rsCurrentUserNotesCount("Note")
					notesServiceTicketID = rsCurrentUserNotesCount("ServiceTicketID")
					notesEnteredByUserNo = rsCurrentUserNotesCount("EnteredByUserNo")
					notesEnteredDateTime = rsCurrentUserNotesCount("RecordCreationDateTime")
					
					submissionHour = Hour(notesEnteredDateTime)
					submissionMinute = Minute(notesEnteredDateTime)
					submissionZeroFactor = "0" & submissionMinute
					submissionAMPM = "AM"
					If submissionHour >= 12 then submissionAMPM = "PM"
					If submissionHour > 12 then submissionHour = submissionHour - 12
					If submissionMinute <= 9 then submissionMinute = submissionZeroFactor	

					notesOriginalProblem = GetServiceTicketProblemByTicketNumber(notesServiceTicketID)
					notesEnteredByUserName = GetUserDisplayNameByUserNo(notesEnteredByUserNo)
					userType = getUserType(notesEnteredByUserNo)
					notesCustID = GetServiceTicketCust(notesServiceTicketID)
					notesCustName = GetCustNameByCustNum(notesCustID)
					notesServiceTicketStatus = GetServiceTicketStatus(notesServiceTicketID)							
		
					%>
					<table border="0" width="<%=PageWidth%>">
					<tr valign="top">
		
						<!-- If we are writing additional notes for the same ticket, do not write the customer info all over again -->
						<%
						
						If cInt(PreviousNotesServiceTicketID) = cInt(notesServiceTicketID) Then		
												
							notesOriginalProblem = ""
							%>
							<td width="20%">
								<strong><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></strong>
							</td>
							<td width="5%" class="auto-style1">
								<strong><font face="Consolas" style="font-size: <%= FontSizeVar %>pt; color:#0076D3;">---</font></strong>
							</td>
							<td width="5%">
								<strong><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">---</font></strong>
							</td>
							<%
						Else %>
						
							<td width="20%">
								<strong><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= notesCustName %>, <%= notesCustID %></font></strong>
							</td>
							<td width="5%" class="auto-style1">
								<strong><font face="Consolas" style="font-size: <%= FontSizeVar %>pt; color:#0076D3;"><%= notesServiceTicketID %></font></strong>
							</td>
							<td width="5%">
								<strong><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= notesServiceTicketStatus %></font></strong>
							</td>
						
						<%
						End If
						
						'***********************************************************
						'First write the original problem notes
						'***********************************************************
						
						ServiceTicketNotes = notesOriginalProblem
						
						'' Now handle the notes spanning multiple lines - the actual printing
						NumNoteLines = 0 
						NumNoteLines = int(Len(ServiceTicketNotes) / 50)
						If Len(ServiceTicketNotes) MOD 50 <> 0 Then NumNoteLines = NumNoteLines + 1
		
						ReDim DetailLinesArray(NumNoteLines)
						For x = 0 to NumNoteLines -1
							If x = 0 Then
								DetailLinesArray(x) = Mid(ServiceTicketNotes,1,50)
							ElseIf x = 1 Then
								DetailLinesArray(x) = Mid(ServiceTicketNotes,51,50)
							ElseIf x = NumNoteLines -1 Then 
								DetailLinesArray(x) = Mid(ServiceTicketNotes,(x*50)+1,Len(ServiceTicketNotes)- ((x*50)))
							Else
								DetailLinesArray(x) = Mid(ServiceTicketNotes,(x*50)+1,50)
							End If
						Next
							
						
						If notesOriginalProblem = "" Then
						%>
							<td width="35%">
								<strong><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">---</font></strong>
							</td>
						<%
						Else
							%><td width="35%" style="white-space: nowrap;"><%
							For z = 0 to Ubound(DetailLinesArray) -1
								%>
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt;color:#666B85;"><%= DetailLinesArray(z) %></font>
								<%
								Response.Write("<br>")
								RowCount = RowCount + 1
							Next
							%></td><%
						End If
						

						'***********************************************************
						'Then write the notes entered by the user
						'***********************************************************
						
						ServiceTicketNotes = notesEntered & " (entered at " & submissionHour & ":" & submissionMinute & " " & submissionAMPM & ")"					
						
						'' Now handle the notes spanning multiple lines - the actual printing
						NumNoteLines = 0 
						NumNoteLines = int(Len(ServiceTicketNotes) / 53)
						If Len(ServiceTicketNotes) MOD 53 <> 0 Then NumNoteLines = NumNoteLines + 1
		
						ReDim DetailLinesArrayNotes(NumNoteLines)
						For x = 0 to NumNoteLines -1
							If x = 0 Then
								DetailLinesArrayNotes(x) = Mid(ServiceTicketNotes,1,53)
							ElseIf x = 1 Then
								DetailLinesArrayNotes(x) = Mid(ServiceTicketNotes,54,53)
							ElseIf x = NumNoteLines -1 Then 
								DetailLinesArrayNotes(x) = Mid(ServiceTicketNotes,(x*53)+1,Len(ServiceTicketNotes)- ((x*53)))
							Else
								DetailLinesArrayNotes(x) = Mid(ServiceTicketNotes,(x*53)+1,53)
							End If
						Next						
						
						
							
						%><td width="35%" style="white-space: nowrap;"><%
						For z = 0 to Ubound(DetailLinesArrayNotes) -1
							%>
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt;color:#28a745!important;"><%= DetailLinesArrayNotes(z) %></font>
							<%
							Response.Write("<br>")
							RowCount = RowCount + 1
						Next
						%></td>
							 
					</tr>
					<tr>
						<td colspan="5">&nbsp;</td>
					</tr>
					<%					
					
					RowCount = RowCount + 2
					
					If (cInt(RowCount) > cInt(LinesPerPage)) Then
						%></table><%
						Call Footer
						Call PageHeader
						%>
						<table border="0" width="<%=PageWidth%>">
						<tr>
							<td colspan="5">
								<font face="Consolas">
									<hr>
										<% If (cInt(PreviousNotesEnteredByUserNo) <> cInt(notesEnteredByUserNo)) Then %>
											<center><h2><%= GetTerm("Field Service") %> Notes for <%= TechnicianName %>&nbsp;(<%= TechnicianUserNo %>) CONTINUED</h2></center>
										<% Else %>
											<center><h2><%= GetTerm("Field Service") %> Notes for <%= TechnicianName %>&nbsp;(<%= TechnicianUserNo %>)</h2></center>
										<% End If %>
									<hr>
								</font>
							</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td width="20%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=GetTerm("Customer")%></font></u></strong>
							</td>
							<td width="5%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Ticket #</font></u></strong>
							</td>
							<td width="5%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Status</font></u></strong>
							</td>		
							<td width="35%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Original Problem</font></u></strong>
							</td>
							<td width="35%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Service Notes</font></u></strong>
							</td>
						</tr>
					<tr>
						<td colspan="5">&nbsp;</td>
					</tr>
						<%				
					End If
					
					PreviousNotesServiceTicketID = notesServiceTicketID
					rsCurrentUserNotesCount.MoveNext
				Loop


				If (cInt(RowCount) > cInt(LinesPerPage)) Then
				
					If (cInt(PreviousNotesEnteredByUserNo) <> cInt(notesEnteredByUserNo)) Then
						%></table><%
						Call Footer
						Call PageHeader
					Else
						%></table><%
						Call Footer
						Call PageHeader					
						%>
						<table border="0" width="<%=PageWidth%>">
						<tr>
							<td colspan="5">
								<font face="Consolas">
									<hr>
										<center><h2><%= GetTerm("Field Service") %> Notes for <%= TechnicianName %>&nbsp;(<%= TechnicianUserNo %>) CONTINUED</h2></center>
									<hr>
								</font>
							</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td width="20%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=GetTerm("Customer")%></font></u></strong>
							</td>
							<td width="5%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Ticket #</font></u></strong>
							</td>
							<td width="5%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Status</font></u></strong>
							</td>		
							<td width="35%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Original Problem</font></u></strong>
							</td>
							<td width="35%">
								<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Service Notes</font></u></strong>
							</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
						</tr>
						<%
					End If
				
				Else
	
					If (cInt(PreviousNotesEnteredByUserNo) <> cInt(notesEnteredByUserNo)) Then
						%></table><%
						Call Footer
						Call PageHeader
					End If
					
				End If
		
			End If
			
			rsDailyFieldServiceNotesTechsLoop.Movenext	
					
			PreviousNotesEnteredByUserNo = notesEnteredByUserNo
	
		Loop
		
		NoBreak = True
		Call Footer	
		
		
		set rsDailyFieldServiceNotesTechsLoop = Nothing
		cnnDailyFieldServiceNotesTechsLoop.Close	
		set cnnDailyFieldServiceNotesTechsLoop = Nothing
			
		set rsCurrentUserNotesCount = Nothing
		cnnCurrentUserNotesCount.Close	
		set cnnCurrentUserNotesCount = Nothing
			
		set rsServiceTicketDetails = Nothing
		cnnServiceTicketDetails.Close	
		set cnnServiceTicketDetails = Nothing
		
	End If


	'*******************************************************
	'*** END END END END END END END END END END END END END 
	'*** This section is the detail section of daily field
	'*** service notes - it shows only the user that
	'*** have entered notes
	'*******************************************************
	%>

</td>
</tr>
</table>

</body>
</html>


<%
Sub PageHeader

	RowCount = 0
	%>

	<table border="0" width="<%=PageWidth%>" cellspacing="0" align="center">
		<tr>
			<td width="50%"><img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png" style="height:55px;"></td>
			<td width="50%">
				<p align="center"><b><font face="Consolas" size="4">MDS Insight Service Ticket Daily Notes Report</font></b></p>
				<p align="center"><font face="Consolas" size="2">Report Generated: <%= WeekDayName(WeekDay(DateValue(Now()))) %>&nbsp;<%= Now() %><br></font></p>
			</td>
		</tr>
		<tr>
			<td width="20%" height="16">
				<p align="right"><font face="Consolas" size="1">&nbsp;</font></p>
			</td>
		</tr>
	</table>
	<%
	PageNum = PageNum + 1
End Sub

Sub SubHeader
	%> 
	<br><br><br>
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
		<tr>
			<td>
			<font face="Consolas">
			<hr>
			</font>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
	
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">
	<tr>
		<td width="20%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=GetTerm("Customer")%></font></u></strong>
		</td>
		<td width="5%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Ticket #</font></u></strong>
		</td>
		<td width="5%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Status</font></u></strong>
		</td>		
		<td width="35%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Original Problem</font></u></strong>
		</td>
		<td width="35%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Service Notes</font></u></strong>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<%
End Sub


Sub Footer

	'Now get us to the next page
	Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'><table>")
	For x = 1 to LinesPerPage - RowCount
		Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'><tr><td border='1'>&nbsp;</td></tr>")
	Next
	Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'></table>")
	%>
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
		<tr>
			<td colspan="3">
				<hr>
			</td>
		</tr>
		<tr>
			<td width="33%">
				<font face="Consolas" style="font-size: 9pt">directlaunch/service/serviceTicketDailyNotesReport.asp</font>
			</td>
			<td width="33%" align="center">
				<font face="Consolas" style="font-size: 12pt">Page:&nbsp;<%=PageNum%></font>
			</td>
			<td width="33%">
				<font face="Consolas" style="font-size: 12pt">&nbsp;</font>
			</td>
		</tr>
	</table>
	<% If NoBreak <> True Then %>
		<BR style="page-break-after: always">
	<% End If

End Sub

%>