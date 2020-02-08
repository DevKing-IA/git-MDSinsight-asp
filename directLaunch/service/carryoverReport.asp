<% @ Language = VBScript %>
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
<!--#include file="../../inc/InsightFuncs_Users.asp"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Service Ticket Carry Over report<%
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



'****************************************************
'CHECK FOR INDIVIDUAL SALESMEN OR TEAMS
'****************************************************

Slsmn = Request.QueryString("sls") ' Gets passed if only being run for one primary or secondary salesman 
TeamIntRecID = Request.QueryString("tm") ' Gets passed if only being run for a team 

'****************************************************
'BUILD AN ARRAY OF SALESMEN OR TEAM USERS NOS FOR
'LOOPING THROUGH REPORT BUILD
'****************************************************
If Slsmn <> "" Then

	SlsmArray = Split(Slsmn,",")
	
ElseIf TeamIntRecID <> "" Then 

	TeamUserNos = GetTeamUserNosByTeamIntRecID(TeamIntRecID)
	TeamUserNosArray = Split(TeamUserNos,",")
	
	For x = 0 to Ubound(TeamUserNosArray)
		SalespersonNumber = GetSalesPersonNoByUserNo(TeamUserNosArray(x))
		If SalespersonNumber <> "" Then
			SlsmString =  SalespersonNumber & ","	
		End If						
	Next 
	
	SlsmString = Left(SlsmString, Len(SlsmString) - 1)
	SlsmArray = Split(SlsmString,",")
	
Else
	'Create an empty array so checks don't throw errors
	Dim SlsmArray()
	ReDim SlsmArray(-1)
End If

'****************************************************



'This is here so we only open it once for the whole page
Set cnn_Settings_FieldService = Server.CreateObject("ADODB.Connection")
cnn_Settings_FieldService.open (MUV_READ("ClientCnnString"))
Set rs_Settings_FieldService = Server.CreateObject("ADODB.Recordset")
rs_Settings_FieldService.CursorLocation = 3 
SQL_Settings_FieldService = "SELECT * FROM Settings_FieldService"
Set rs_Settings_FieldService = cnn_Settings_FieldService.Execute(SQL_Settings_FieldService)
If not rs_Settings_FieldService.EOF Then
	ServiceTicketCarryoverReportOnOff = rs_Settings_FieldService("ServiceTicketCarryoverReportOnOff")
	CarryoverReportInclCustType = rs_Settings_FieldService("CarryoverReportInclCustType")
	CarryoverReportInclTicketNum = rs_Settings_FieldService("CarryoverReportInclTicketNum")
	CarryoverReportShowRedoBreakdown = rs_Settings_FieldService("CarryoverReportShowRedoBreakdown")	
	ServiceTicketCarryoverReportIncludeRegions = rs_Settings_FieldService("ServiceTicketCarryoverReportIncludeRegions")
	ShowSeparateFilterChangesTabOnServiceScreen = rs_Settings_FieldService("ShowSeparateFilterChangesTabOnServiceScreen")	
Else
	ServiceTicketCarryoverReportOnOff = 0
End If
Set rs_Settings_FieldService = Nothing
cnn_Settings_FieldService.Close
Set cnn_Settings_FieldService = Nothing

If ServiceTicketCarryoverReportOnOff <> 1 Then
	%>MDS Insight: The service ticket carry over report is not turned on.
	<%
	Response.End
End IF

%>



<body bgcolor="#FFFFFF" text="#000000" link="#000080" vlink="#000080" alink="#000080" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">


<table border="0" width="<%=PageWidth%>" align="center">

	<tr>
		<td width="100%" align="center">

		<%
		
		If Ubound(SlsmArray) < 0 Then


			'*******************************************************
			'*** This section is the first page which prints all the
			'*** service ticket summary info
			'*** Does NOT get included if being run for one salesman
			'*** or for teams
			'*******************************************************
		
			Call PageHeader

			LinesPerPage = 42
			
			SQL = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange <> 1 ORDER BY submissionDateTime DESC"	


			Set cnnCarryOver = Server.CreateObject("ADODB.Connection")
			cnnCarryOver.open (MUV_READ("ClientCnnString"))
			Set rsCarryOver  = Server.CreateObject("ADODB.Recordset")
			rsCarryOver.CursorLocation = 3 
			rsCarryOver.Open SQL, cnnCarryOver 
						
			'Response.Write(SQL & "<br>")
			
			If Not rsCarryOver.EOF Then
			
				
				%>
				<br><br><br>
				<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
					<tr>
						<td colspan="5">
						<font face="Consolas">
						<hr>
						<center><h2>Service Ticket Summary for <%= Date() %></h2></center>
						<hr>
						</font>
						</td>
					</tr>
					<tr>
						<td colspan="5">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="5">&nbsp;</td>
					</tr>

						
				<%
					If Not rsCarryOver.EOF Then
					
						NumberOfMinutesInServiceDayVar = GetNumberOfMinutesInServiceDay()
						
					
						TotalNumberOfTickets = 0
						TotalOpenedToday = 0
						AwaitingDispatch = 0
						AwaitingAcknowledgement = 0
						Acknowledged = 0
						EnRoute = 0
						OnSite = 0
						Swap = 0
						WaitForParts = 0
						UnableToWork = 0
						Followup = 0
						Closed = 0
						Open1Day = 0
						Open2Days = 0
						Open3To5Days = 0
						OpenOver5Days = 0
						
						Do While Not (rsCarryOver.EOF)
						
							TotalNumberOfTickets = TotalNumberOfTickets + 1
							If DateDiff("d",rsCarryOver("submissionDateTime"),DateAdd("d",Adjust,Now())) = 0 Then TotalOpenedToday = TotalOpenedToday + 1
							
							ServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsCarryOver("MemoNumber"))
							
							Select Case ServiceTicketCurrentStageVar 
								Case "Received"
									AwaitingDispatch = AwaitingDispatch + 1
								Case "Released"
									AwaitingDispatch = AwaitingDispatch + 1
								Case "Dispatched"
									AwaitingAcknowledgement = AwaitingAcknowledgement + 1						
								Case "Dispatch Acknowledged"
									Acknowledged = Acknowledged  + 1
								Case "En Route"
									EnRoute = EnRoute +1
								Case "On Site"
									OnSite = OnSite +1
								Case "Unable To Work"
									UnableToWork = UnableToWork + 1
								Case "Swap"
									Swap = Swap +1
								Case "Wait for parts"
									WaitForParts = WaitForParts + 1
								Case "Follow Up"
									Followup = Followup + 1
							End Select
							
													
							RowCount = RowCount + 2.5
							
							rsCarryOver.MoveNext
						Loop
						
						End IF
						rsCarryOver.Close
						
						%>
						
						<tr>
							<td colspan="5">
								<font face="Consolas" style="font-size: 24pt">
								<%= TotalNumberOfTickets %>&nbsp;OPEN SERVICE TICKETS</font>
							</td>
						</tr>

						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						
						<%
						' Now get the closed tickets
						
						' I know this is lazy
						
						CloseDate = DateAdd("d",Adjust,Date())
						
						SQL = "SELECT Count(*) As CloseCount FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE'"
						SQL = SQL & " AND Month(RecordCreatedateTime) = " & Month(CloseDate) + MAdjust
						SQL = SQL & " AND Year(RecordCreatedateTime) = " &  Year(CloseDate)
						SQL = SQL & " AND Day(RecordCreatedateTime) = " & Day(CloseDate)
						SQL = SQL & " AND FilterChange <> 1 "
						rsCarryOver.Open SQL, cnnCarryOver 

						'Response.Write(SQL & "<br>")
			
						If Not rsCarryOver.Eof Then
						
							Closed = rsCarryOver("CloseCount")
						
						End If
						
						Open1Day = GetNumberOfServiceTicketsInTimeRange(0, NumberOfMinutesInServiceDayVar)
						Open2Days = GetNumberOfServiceTicketsInTimeRange(NumberOfMinutesInServiceDayVar+1, NumberOfMinutesInServiceDayVar*2)
						Open3To5Days = GetNumberOfServiceTicketsInTimeRange((NumberOfMinutesInServiceDayVar*2)+1, NumberOfMinutesInServiceDayVar*5)
						OpenOver5Days = GetNumberOfServiceTicketsInTimeRange((NumberOfMinutesInServiceDayVar*5)+1, 99999)

						%>
						<tr>
							<td colspan="2" style="border-bottom: 1px solid black;" width="40%">
								<font face="Consolas" style="font-size: 14pt"><strong>OVERVIEW</strong></font>
							</td>
							<td style="border-bottom: 1px solid white;" width="10%">&nbsp;</td>
							<td colspan="2" style="border-bottom: 1px solid black;" width="40%">
								<font face="Consolas" style="font-size: 14pt"><strong>AGING</strong></font>
							</td>							
						</tr>
						
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>						
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%=(TotalNumberOfTickets - TotalOpenedToday + Closed)%></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Carried Over From Previous Days</font></td>
							<td width="5%">&nbsp;</td>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= Open1Day %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Open 1 Day (0-8 hrs)</font></td>
						</tr>
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalOpenedToday %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Created <%= WeekdayName(Weekday(Date())) %> (today)</font></td>
							<td width="5%">&nbsp;</td>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= Open2Days %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Open 2 Days (8-16 hrs)</font></td>
						</tr>
						
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= Closed %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Closed <%= WeekdayName(Weekday(Date())) %> (today)</font></td>
							<td width="5%">&nbsp;</td>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= Open3To5Days %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Open 3-5 Days (16-40 hrs)</font></td>
						</tr>
						
						<tr>
							<td colspan="2">
								<font face="Consolas" style="font-size: 14pt"><hr width="50%" align="left"></font>
							</td>
							<td width="5%">&nbsp;</td>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= OpenOver5Days %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Open Over 5 Days (Over 40 hrs)</font></td>
						</tr>
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfTickets %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">TOTAL</font></td>
							<td width="5%">&nbsp;</td>
							<td colspan="2" align="right">
								<font face="Consolas" style="font-size: 14pt"><hr width="50%" align="left"></font>
							</td>							
						</tr>
						
						
						<tr>
							<td colspan="2">
								<font face="Consolas" style="font-size: 14pt">&nbsp;</font>
							</td>
							<td width="5%">&nbsp;</td>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfTickets %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">TOTAL</font></td>						
						</tr>
						
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						
						
						<tr>
							<td colspan="2" style="border-bottom: 1px solid black;" width="40%">
								<font face="Consolas" style="font-size: 14pt"><strong>STATUS</strong></font>
							</td>
							<td style="border-bottom: 1px solid white;" width="10%">&nbsp;</td>
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>
								<td colspan="2" style="border-bottom: 1px solid black;" width="40%">
									<font face="Consolas" style="font-size: 14pt"><strong>REDO</strong></font>
								</td>							
							<% Else %>
								<td colspan="2" style="border-bottom: 1px solid white;" width="40%">
									&nbsp;
								</td>							
							<%End If%>
						</tr>
						
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						

						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= AwaitingDispatch %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Awaiting Dispatch</font></td>
							<td width="5%">&nbsp;</td>
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= Swap %></font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">Swap</font></td>
							<% Else %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt">&nbsp;</font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>
							<% End If %>								
						</tr>
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= AwaitingAcknowledgement %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Dispatched, Awaiting Acknowledgement</font></td>
							<td width="5%">&nbsp;</td>
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= UnableToWork %></font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">Unable To Work</font></td>
							<% Else %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt">&nbsp;</font>&nbsp;&nbsp;</td>							
								<td width="42%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>
							<% End If %>								
						</tr>

						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= Acknowledged %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Dispatch Acknowledged</font></td>
							<td width="5%">&nbsp;</td>
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= FollowUp %></font>&nbsp;&nbsp;</td>							
								<td width="42%"><font face="Consolas" style="font-size: 14pt">Follow Up</font></td>
							<% Else %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt">&nbsp;</font>&nbsp;&nbsp;</td>							
								<td width="42%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>
							<% End If %>									
						</tr>
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= EnRoute %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">En Route</font></td>
							<td width="5%">&nbsp;</td>
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= WaitForParts %></font>&nbsp;&nbsp;</td>							
								<td width="42%"><font face="Consolas" style="font-size: 14pt">Wait For Parts</font></td>
							<% Else %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt">&nbsp;</font>&nbsp;&nbsp;</td>							
								<td width="42%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>
							<% End If %>									
						</tr>
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= OnSite %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">On Site</font></td>
							<td width="5%">&nbsp;</td>
							<td colspan="2">
								<% If CarryoverReportShowRedoBreakdown = 1 Then %>
									<font face="Consolas" style="font-size: 14pt"><hr width="50%" align="left"></font>
								<% Else %>
									<font face="Consolas" style="font-size: 14pt">&nbsp;</font>
								<% End If %>										
							</td>							
						</tr>
						

						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= UnableToWork + Swap + WaitForParts + Followup %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Re-Dispatched</font></td>
							<td width="5%">&nbsp;</td>
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= UnableToWork + Swap + WaitForParts + Followup %></font>&nbsp;&nbsp;</td>							
								<td width="42%"><font face="Consolas" style="font-size: 14pt">TOTAL</font></td>
							<% Else %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt">&nbsp;</font>&nbsp;&nbsp;</td>							
								<td width="42%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>							
							<% End If %>
						</tr>

						<tr>
							<td colspan="2">
								<font face="Consolas" style="font-size: 14pt"><hr width="50%" align="left"></font>
							</td>
							<td width="5%">&nbsp;</td>
							<td colspan="2">
								<font face="Consolas" style="font-size: 14pt">&nbsp;</font>
							</td>							
						</tr>
						
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfTickets %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">TOTAL</font></td>
							<td width="5%">&nbsp;</td>
							<td colspan="2">
								<font face="Consolas" style="font-size: 14pt">&nbsp;</font>
							</td>							
						</tr>
						
						<% RowCount = LinesPerPage - 1

					Else%>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">
								<font face="Consolas" style="font-size: 14pt">
								<hr>
								<center>Great news! There are no issues to report.</center></font>
								<hr>
								<% NoBreak = True %>
							</td>
						</tr>
						<%RowCount = 0

					End If %>
		
				<br/><br/>
		</table>
		</td>
	</tr>
	<tr>
	<td>
	
	<% 

			Call Footer

				
			'*******************************************************
			'*** END END END END END END END END END END END END END 
			'*** This section is the first page which prints all the
			'*** service ticket summary info
			'*** Does NOT get included if being run for one salesman
			'*******************************************************
		End If	
	%>
	</table>
	
	<table border="0" width="<%=PageWidth%>" align="center">
	<tr>
		<td width="100%" align="center">
	
	<%
	'*******************************************************
	'*** This section is the detail section of all the open
	'*** service tickets
	'*** DOES get included in all versions of the report
	'*** no matter who it is being sent to
	'*******************************************************
	

	'Now we start doing all the individual detail sections
	HeldGroupDescription = ""
	
	SQL = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange <> 1 ORDER BY submissionDateTime"
	
	Set cnnCarryOver = Server.CreateObject("ADODB.Connection")
	cnnCarryOver.open (MUV_READ("ClientCnnString"))
	Set rsCarryOver = Server.CreateObject("ADODB.Recordset")
	rsCarryOver.CursorLocation = 3 
	rsCarryOver.Open SQL, cnnCarryOver
	

	If Not rsCarryOver.EOF Then
		
		Call PageHeader
		Call SubHeader
	
		Do While Not rsCarryOver.EOF

			FontSizeVar = 9
			LinesPerPage = 41
			
			%>
			<tr>
			<%

			' Check to see if the Salesman file exists
			SalesmanFileExists = True
								
			'Only if the backend has a salesman table
			On Error Goto 0
			Set rsSLsmn = Server.CreateObject("ADODB.Recordset")
		
			Err.Clear
			on error resume next
			Set rsSlsmn = cnnCarryOver.Execute("SELECT TOP 1 * FROM Salesman")
			If Err.Description <> "" Then
				If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
					SalesmanFileExists = False
				End If
			End IF
			On error goto 0
						
			SQL_ARCUST = "SELECT * FROM AR_Customer  WHERE CustNum = '" & rsCarryOver("AccountNumber") & "'" 
			Set rs_ARCUST = Server.CreateObject("ADODB.Recordset")
			rs_ARCUST.CursorLocation = 3
			Set rs_ARCUST = cnnCarryOver.Execute(SQL_ARCUST)

			If Not rs_ARCUST.Eof Then
				Primary = rs_ARCUST("Salesman")
				Secondary = rs_ARCUST("SecondarySalesman")
				CustType = rs_ARCUST("CustType")
				'NOTE: FOR TEAM FILTERTING, THE PRIMARY OR SECONDARY, MUST BE EQUAL TO ONE OF THE TEAM MEMBER'S USER NUMBERS
			End If

			If Not IsNumeric(Primary) Then Primary = -1 ' We have to set it to something numeric
			If Not IsNumeric(Secondary) Then Secondary = -1 ' We have to set it to something numeric
			
			If SalesmanFileExists = False Then
				Primary = -1 ' We have to set it to something numeric
				Secondary = -1 ' We have to set it to something numeric
			End If
					
			'Response.Write("Primary Sales No: " & Primary & ".....Secondary Sales No: " & Secondary & "<br>")	
			
				
			If (Ubound(SlsmArray) < 0) OR (UBound(SlsmArray) >= 0 AND (InArray(SlsmArray,Primary) OR InArray(SlsmArray,Secondary))) Then ' for salesman filtering
						
			
				'SEE IF YOU HAVE ALREADY PRINTED THE TICKET
				'''''''''''''CODE HERE
					
				' DO THE ELAPSED TIME
				
				elapsedMinutes = ServiceCallElapsedMinutesOpenTicket(rsCarryOver("MemoNumber"))
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
					elapsedString = elapsedString  & elapsedHours & "h"
				End If
				If int(elapsedMinutes) > 0 Then
					elapsedString = elapsedString  & elapsedMinutes & "m"
				End If

				' Nicer looking submission date & time
				
				submissionHour = Hour(rsCarryOver("submissionDateTime"))
				submissionMinute = Minute(rsCarryOver("submissionDateTime"))
				submissionZeroFactor = "0" & submissionMinute
				submissionAMPM = "AM"
				If submissionHour >= 12 then submissionAMPM = "PM"
				If submissionHour > 12 then submissionHour = submissionHour - 12
				If submissionMinute <= 9 then submissionMinute = submissionZeroFactor	
				
				submissionDateTime = rsCarryOver("submissionDateTime")
				
				ticketDateDisplay = padDate(MONTH(submissionDateTime),2) & "/" & padDate(DAY(submissionDateTime),2) & "/" & padDate(RIGHT(YEAR(submissionDateTime),2),2)
				
						%>
						<td width="5%">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsCarryOver("AccountNumber") %></font>
						</td>							
						<% If CarryoverReportInclCustType = 1 OR CarryoverReportInclTicketNum = 1 Then %>
							<td width="20%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsCarryOver("Company")) > 35 Then Response.Write(Left(rsCarryOver("Company"),35)) Else Response.Write(rsCarryOver("Company"))  %></font>
							</td>
						<% Else %>
							<td width="42%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsCarryOver("Company")) > 35 Then Response.Write(Left(rsCarryOver("Company"),35)) Else Response.Write(rsCarryOver("Company"))  %></font>
							</td>								
						<% End If %>
						<% If CarryoverReportInclCustType = 1 Then %>
							<td width="5%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= CustType %></font>
							</td>
						<% End If %>
						<% If CarryoverReportInclTicketNum = 1 Then %>
							<td width="5%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsCarryOver("MemoNumber") %></font>
							</td>
						<% End If %>
						<%
						If SalesmanFileExists = True Then
							PrimarySalesPerson = GetSalesmanNameBySlsmnSequence(Primary)
		    				SecondarySalesPerson = GetSalesmanNameBySlsmnSequence(Secondary)
						    If Instr(PrimarySalesPerson ," ") <> 0 Then
								PrimarySalesPerson =  Left(PrimarySalesPerson,Instr(PrimarySalesPerson ," ")+1)
							End If
							If Instr(SecondarySalesPerson," ") <> 0 Then
								SecondarySalesPerson = Left(SecondarySalesPerson,Instr(SecondarySalesPerson," ")+1) 
							End If
							%>
							<td width="8%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=PrimarySalesPerson%></font>
							</td>
							<% If CarryoverReportInclTicketNum = 0 Then %>
								<td width="7%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=SecondarySalesPerson%></font>
								</td>
							<% End If %>
						<% End If %>
						<td width="7%">
							<!--<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=ticketDateDisplay & "&nbsp;" & submissionHour & ":" & submissionMinute & " " & submissionAMPM%></font>-->
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=ticketDateDisplay%></font>
						</td>
						<td width="7%">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=elapsedString%></font>
						</td>
						<% If SalesmanFileExists = True Then %>
							<td width="5%" align="right">
							<%
								TicketStageVar = GetServiceTicketCurrentStage(rsCarryOver("MemoNumber"))
								TicketStageVar = Replace(TicketStageVar ,"Dispatch Acknowledged","Dispatch Ack")
								TicketStageVar = Replace(TicketStageVar ,"Received","Awaiting Dispatch")
								TicketStageVar = Replace(TicketStageVar ,"Released","Awaiting Dispatch")
								TicketStageVar = Replace(TicketStageVar ,"Dispatch","Disp")
								TicketStageVar = Replace(TicketStageVar ,"Disped","Dispatched")
							Else ' we can give it more room %>
							<td width="20%" align="right">
							<%
								TicketStageVar = GetServiceTicketCurrentStage(rsCarryOver("MemoNumber"))
							End If%>	
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=TicketStageVar %></font>
						</td> <%
					ServiceTicketNotes = GetLastServiceTicketNotesByTicket(rsCarryOver("MemoNumber"))
					If Len(ServiceTicketNotes) > 50 Then ServiceTicketNotes = Left(ServiceTicketNotes,50)
						If Len(ServiceTicketNotes) > 1 Then %>
							<td width="31%" style="white-space: nowrap;">
								<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= ServiceTicketNotes %></font>
							</td>
						<% 	Else %>
						<td width="31%" style="white-space: nowrap;">
							<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt">&nbsp;</font>
						</td>
					<% 	End If  %>
			 			 
			</tr>
	<%
			RowCount = RowCount + 1

		End If ' For Slsmn filtering


			rsCarryOver.Movenext	

			If RowCount > LinesPerPage Then
				%></table><%
				Call Footer
				Call PageHeader
				Call SubHeader
			End If
			
		
		Loop

		
		If Ubound(SlsmArray) >= 0 OR Closed < 1 Then NoBreak = True
		
		Call Footer	
		
	End If


	'*******************************************************
	'*** END END END END END END END END END END END END END 
	'*** This section is the detail section of all the open
	'*** service tickets
	'*** DOES get included in all versions of the report
	'*** no matter who it is being sent to
	'*******************************************************
	
	
	%>
	
	
	


<table border="0" width="<%=PageWidth%>" align="center">

	<tr>
		<td width="100%" align="center">

		<%
		
		If Ubound(SlsmArray) < 0 Then

			'*******************************************************
			'*** This section is the first page which prints all the
			'*** filter ticket summary info
			'*** Does NOT get included if being run for one salesman
			'*** or for a team
			'*******************************************************
		
			Call PageHeader

			LinesPerPage = 42
			
			SQL = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange = 1 ORDER BY submissionDateTime DESC"	


			Set cnnCarryOver = Server.CreateObject("ADODB.Connection")
			cnnCarryOver.open (MUV_READ("ClientCnnString"))
			Set rsCarryOver  = Server.CreateObject("ADODB.Recordset")
			rsCarryOver.CursorLocation = 3 
			rsCarryOver.Open SQL, cnnCarryOver 
						
			'Response.Write(SQL & "<br>")
			
			If Not rsCarryOver.EOF Then
			
				
				%>
				<br><br><br>
				<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
					<tr>
						<td colspan="5">
						<font face="Consolas">
						<hr>
						<center><h2>Filter Ticket Summary for <%= Date() %></h2></center>
						<hr>
						</font>
						</td>
					</tr>
					<tr>
						<td colspan="5">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="5">&nbsp;</td>
					</tr>

						
				<%
					If Not rsCarryOver.EOF Then
					
					
						TotalNumberOfTicketsFilters = 0
						TotalOpenedTodayFilters = 0
						AwaitingDispatchFilters = 0
						AwaitingAcknowledgementFilters = 0
						AcknowledgedFilters = 0
						EnRouteFilters = 0
						OnSiteFilters = 0
						SwapFilters = 0
						WaitForPartsFilters = 0
						UnableToWorkFilters = 0
						FollowupFilters = 0
						ClosedFilters = 0
						
						TotalNumberOfTicketFilterChanges = 0
						TotalOpenedTodayFilterChanges = 0
						AwaitingDispatchFilterChanges = 0
						AwaitingAcknowledgementFilterChanges = 0
						AcknowledgedFilterChanges = 0
						EnRouteFilterChanges = 0
						OnSiteFilterChanges = 0
						SwapFilterChanges = 0
						WaitForPartsFilterChanges = 0
						UnableToWorkFilterChanges = 0
						FollowupFilterChanges = 0
						ClosedFilterChanges = 0
						
						
						Do While Not (rsCarryOver.EOF)
						
							TotalNumberOfTicketsFilters = TotalNumberOfTicketsFilters + 1
							TotalNumberOfTicketFilterChanges = TotalNumberOfTicketFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
							
							If DateDiff("d",rsCarryOver("submissionDateTime"),DateAdd("d",Adjust,Now())) = 0 Then TotalOpenedTodayFilters = TotalOpenedTodayFilters + 1
							If DateDiff("d",rsCarryOver("submissionDateTime"),DateAdd("d",Adjust,Now())) = 0 Then TotalOpenedTodayFilterChanges = TotalOpenedTodayFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
							
							ServiceTicketCurrentStageVarFilters = GetServiceTicketCurrentStage(rsCarryOver("MemoNumber"))
							
							Select Case ServiceTicketCurrentStageVarFilters 
								Case "Received"
									AwaitingDispatchFilters = AwaitingDispatchFilters + 1
									AwaitingDispatchFilterChanges = AwaitingDispatchFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
								Case "Released"
									AwaitingDispatchFilters = AwaitingDispatchFilters + 1
									AwaitingDispatchFilterChanges = AwaitingDispatchFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
								Case "Dispatched"
									AwaitingAcknowledgementFilters = AwaitingAcknowledgementFilters + 1		
									AwaitingAcknowledgementFilterChanges = AwaitingAcknowledgementFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))				
								Case "Dispatch Acknowledged"
									AcknowledgedFilters = AcknowledgedFilters + 1
									AcknowledgedFilterChanges = AcknowledgedFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
								Case "En Route"
									EnRouteFilters = EnRouteFilters + 1
									EnRouteFilterChanges = EnRouteFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
								Case "On Site"
									OnSiteFilters = OnSiteFilters + 1
									OnSiteFilterChanges = OnSiteFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
								Case "Unable To Work"
									UnableToWorkFilters = UnableToWorkFilters + 1
									UnableToWorkFilterChanges = UnableToWorkFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
								Case "Swap"
									Swap = Swap + 1
									SwapFilterChanges = SwapFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
								Case "Wait for parts"
									WaitForPartsFilters = WaitForPartsFilters + 1
									WaitForPartsFilterChanges = WaitForPartsFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
								Case "Follow Up"
									FollowupFilters = FollowupFilters + 1
									FollowupFilterChanges = FollowupFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
							End Select
							
													
							RowCount = RowCount + 2.5
							
							rsCarryOver.MoveNext
						Loop
						
						End IF
						rsCarryOver.Close
						
						' Now get the closed tickets
						
						' I know this is lazy
						
						CloseDate = DateAdd("d",Adjust,Date())
						
						SQL = "SELECT Count(*) As CloseCount FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE'"
						SQL = SQL & " AND Month(RecordCreatedateTime) = " & Month(CloseDate) + MAdjust
						SQL = SQL & " AND Year(RecordCreatedateTime) = " &  Year(CloseDate)
						SQL = SQL & " AND Day(RecordCreatedateTime) = " & Day(CloseDate)
						SQL = SQL & " AND FilterChange = 1 "
						rsCarryOver.Open SQL, cnnCarryOver 

						'Response.Write(SQL & "<br>")
			
						If Not rsCarryOver.Eof Then
							ClosedFilters = rsCarryOver("CloseCount")
						End If
						
						rsCarryOver.Close
						
						SQL = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE'"
						SQL = SQL & " AND Month(RecordCreatedateTime) = " & Month(CloseDate) + MAdjust
						SQL = SQL & " AND Year(RecordCreatedateTime) = " &  Year(CloseDate)
						SQL = SQL & " AND Day(RecordCreatedateTime) = " & Day(CloseDate)
						SQL = SQL & " AND FilterChange = 1 "
						rsCarryOver.Open SQL, cnnCarryOver 

						'Response.Write(SQL & "<br>")
			
						If Not rsCarryOver.Eof Then
							Do While NOT rsCarryOver.Eof 
								ClosedFilterChanges = ClosedFilterChanges + GetNumberOfFilterChangesForServiceTicket(rsCarryOver("MemoNumber"))
								rsCarryOver.MoveNext
							Loop
						End If

						Open1DayFilters = GetNumberOfFilterTicketsInTimeRange(0, NumberOfMinutesInServiceDayVar)
						Open1DayFilterChanges = GetNumberOfFilterChangesInTimeRange(0, NumberOfMinutesInServiceDayVar)
						Open2DaysFilters = GetNumberOfFilterChangesInTimeRange(NumberOfMinutesInServiceDayVar+1, NumberOfMinutesInServiceDayVar*2)
						Open2DaysFilterChanges = GetNumberOfFilterChangesInTimeRange(NumberOfMinutesInServiceDayVar+1, NumberOfMinutesInServiceDayVar*2)
						Open3To5DaysFilters = GetNumberOfFilterChangesInTimeRange((NumberOfMinutesInServiceDayVar*2)+1, NumberOfMinutesInServiceDayVar*5)
						Open3To5DaysFilterChanges = GetNumberOfFilterChangesInTimeRange((NumberOfMinutesInServiceDayVar*2)+1, NumberOfMinutesInServiceDayVar*5)
						OpenOver5DaysFilters = GetNumberOfFilterChangesInTimeRange((NumberOfMinutesInServiceDayVar*5)+1, 99999)
						OpenOver5DaysFilterChanges = GetNumberOfFilterChangesInTimeRange((NumberOfMinutesInServiceDayVar*5)+1, 99999)
						
						%>
						
						<tr>
							<td colspan="5">
								<font face="Consolas" style="font-size: 24pt">
								&nbsp;<%= TotalNumberOfTicketsFilters %>&nbsp;OPEN FILTER TICKETS&nbsp;(<%= TotalNumberOfTicketFilterChanges %> Filters)</font>
							</td>
						</tr>

						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						
						<tr>
							<td colspan="2" style="border-bottom: 1px solid black;" width="40%">
								<font face="Consolas" style="font-size: 14pt"><strong>OVERVIEW</strong></font>
							</td>
							<td style="border-bottom: 1px solid white;" width="10%">&nbsp;</td>
							<td colspan="2" style="border-bottom: 1px solid black;" width="40%">
								<font face="Consolas" style="font-size: 14pt"><strong>AGING</strong></font>
							</td>							
						</tr>
						
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>						
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfTicketsFilters - TotalOpenedTodayFilters + ClosedFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Carried Over From Previous Days&nbsp;(<%=TotalNumberOfTicketFilterChanges%>)</font></td>
							<td width="5%">&nbsp;</td>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= Open1DayFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Open 1 Day (0-8 hrs)&nbsp;(<%= Open1DayFilterChanges %>)</font></td>
						</tr>
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalOpenedTodayFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Created <%= WeekdayName(Weekday(Date())) %>&nbsp;(<%= TotalOpenedTodayFilterChanges %>)</font></td>
							<td width="5%">&nbsp;</td>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= Open2DaysFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Open 2 Days (8-16 hrs)&nbsp;(<%= Open2DaysFilterChanges %>)</font></td>							
						</tr>
						
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= ClosedFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Closed <%= WeekdayName(Weekday(Date())) %>&nbsp;(<%= ClosedFilterChanges %>)</font></td>
							<td width="5%">&nbsp;</td>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= Open3To5DaysFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Open 3-5 Days (16-40 hrs)&nbsp;(<%= Open3To5DaysFilterChanges %>)</font></td>
						</tr>
						
						<tr>
							<td colspan="2">
								<font face="Consolas" style="font-size: 14pt"><hr width="50%" align="left"></font>
							</td>
							<td width="5%">&nbsp;</td>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= OpenOver5DaysFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Open Over 5 Days (Over 40 hrs)&nbsp;(<%= OpenOver5DaysFilterChanges %>)</font></td>
						</tr>
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfTicketsFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">TOTAL&nbsp;(<%= GetNumberOfFilterChangesInTimeRange(0, 99999) %> filters)</font></td>
							<td width="5%">&nbsp;</td>
							<td colspan="2">
								<font face="Consolas" style="font-size: 14pt"><hr width="50%" align="left"></font>
							</td>							
						</tr>
						
						<tr>
							<td colspan="2">
								<font face="Consolas" style="font-size: 14pt">&nbsp;</font>
							</td>
							<td width="5%">&nbsp;</td>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfTicketsFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">TOTAL&nbsp;(<%= TotalNumberOfTicketFilterChanges %> filters)</font></td>
						</tr>
						
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						
						
						<tr>
							<td colspan="2" style="border-bottom: 1px solid black;" width="40%">
								<font face="Consolas" style="font-size: 14pt"><strong>STATUS</strong></font>
							</td>
							<td style="border-bottom: 1px solid white;" width="10%">&nbsp;</td>
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>
								<td colspan="2" style="border-bottom: 1px solid black;" width="40%">
								<font face="Consolas" style="font-size: 14pt"><strong>REDO</strong></font>
							<% Else %>
								<td colspan="2" style="border-bottom: 1px solid white;" width="40%">
								<font face="Consolas" style="font-size: 14pt"><strong>&nbsp;</strong></font>							
							<% End If %>
							</td>							
						</tr>
						
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						

						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= AwaitingDispatchFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Awaiting Dispatch&nbsp;(<%= AwaitingDispatchFilterChanges %>)</font></td>
							<td width="5%">&nbsp;</td>
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= SwapFilters %></font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">Swap&nbsp;(<%= SwapFilterChanges %> fltrs)</font></td>	
							<% Else %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt">&nbsp;</font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>	
							<% End If %>
							
						</tr>
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= AwaitingAcknowledgementFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Dispatched, Awaiting Acknowledgement&nbsp;(<%= AwaitingAcknowledgementFilterChanges %>)</font></td>
							<td width="5%">&nbsp;</td>
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= UnableToWorkFilters %></font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">Unable To Work&nbsp;(<%= UnableToWorkFilterChanges %>)</font></td>
							<% Else %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt">&nbsp;</font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>	
							<% End If %>
						
						</tr>

						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= AcknowledgedFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Dispatch Acknowledged&nbsp;(<%= AcknowledgedFilterChanges %>)</font></td>
							<td width="5%">&nbsp;</td>
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= FollowUpFilters %></font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">Follow Up&nbsp;(<%= FollowUpFilterChanges %>)</font></td>
							<% Else %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt">&nbsp;</font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>	
							<% End If %>
						</tr>
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= EnRouteFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">En Route&nbsp;(<%= EnRouteFilterChanges %>)</font></td>
							<td width="5%">&nbsp;</td>	
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>						
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= WaitForPartsFilters %></font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">Wait For Parts&nbsp;(<%= WaitForPartsFilterChanges %>)</font></td>
							<% Else %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt">&nbsp;</font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>	
							<% End If %>
						</tr>
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= OnSiteFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">On Site&nbsp;(<%= OnSiteFilterChanges %>)</font></td>
							<td width="5%">&nbsp;</td>
							<td colspan="2">
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>	
								<font face="Consolas" style="font-size: 14pt"><hr width="50%" align="left"></font>
							<% Else %>
								<font face="Consolas" style="font-size: 14pt">&nbsp;</font>
							<% End If %>
								
							</td>							
						</tr>
						

						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= UnableToWorkFilters + SwapFilters + WaitForPartsFilters + FollowupFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">Re-Dispatched&nbsp;(<%= UnableToWorkFilterChanges + SwapFilterChanges + WaitForPartsFilterChanges + FollowupFilterChanges %>)</font></td>
							<td width="5%">&nbsp;</td>
							<% If CarryoverReportShowRedoBreakdown = 1 Then %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= UnableToWorkFilters + SwapFilters + WaitForPartsFilters + FollowupFilters %></font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">TOTAL&nbsp;(<%= UnableToWorkFilterChanges + SwapFilterChanges + WaitForPartsFilterChanges + FollowupFilterChanges %> filters)</font></td>
							<% Else %>
								<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt">&nbsp;</font>&nbsp;&nbsp;</td>
								<td width="42%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>
							<% End If %>
								
						</tr>

						<tr>
							<td colspan="2">
								<font face="Consolas" style="font-size: 14pt"><hr width="50%" align="left"></font>
							</td>
							<td width="5%">&nbsp;</td>
							<td colspan="2">
								<font face="Consolas" style="font-size: 14pt">&nbsp;</font>
							</td>							
						</tr>
						
						
						<tr>
							<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfTicketsFilters %></font>&nbsp;&nbsp;</td>
							<td width="42%"><font face="Consolas" style="font-size: 14pt">TOTAL&nbsp;(<%= TotalNumberOfTicketFilterChanges %> filters)</font></td>
							<td width="5%">&nbsp;</td>
							<td colspan="2">
								<font face="Consolas" style="font-size: 14pt">&nbsp;</font>
							</td>							
						</tr>
						
						<%
							 RowCount = LinesPerPage - 1
					Else%>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">
								<font face="Consolas" style="font-size: 14pt">
								<hr>
								<center>There are no filter tickets.</center></font>
								<hr>
								<% NoBreak = True %>
							</td>
						</tr>
						<%RowCount = LinesPerPage -27

					End If %>
		
				<br/><br/>
		</table>
		</td>
	</tr>
	<tr>
	<td>
	
	<% 

			Call Footer

				
			'*******************************************************
			'*** END END END END END END END END END END END END END 
			'*** This section is the first page which prints all the
			'*** filter ticket summary info
			'*** Does NOT get included if being run for one salesman
			'*******************************************************
		End If	

	%>
	</table>

<table border="0" width="<%=PageWidth%>" align="center">

	<tr>
		<td width="100%" align="center">

<%

	'*******************************************************
	'*** This section is the detail section of all the open
	'*** service tickets THAT ARE FILTER TICKETS
	'*** DOES get included in all versions of the report
	'*** no matter who it is being sent to
	'*******************************************************
	

	'Now we start doing all the individual detail sections
	HeldGroupDescription = ""
	
	SQL = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange = 1 ORDER BY submissionDateTime"
	
	Set cnnCarryOver = Server.CreateObject("ADODB.Connection")
	cnnCarryOver.open (MUV_READ("ClientCnnString"))
	Set rsCarryOver = Server.CreateObject("ADODB.Recordset")
	rsCarryOver.CursorLocation = 3 
	rsCarryOver.Open SQL, cnnCarryOver
	

	If Not rsCarryOver.EOF AND Ubound(SlsmArray) < 0 Then
		
		Call PageHeader
		Call SubHeader
	
		Do While Not rsCarryOver.EOF

			FontSizeVar = 9
			LinesPerPage = 27
			
			%>
			<tr>
			<%
					
			SQL_ARCUST = "SELECT * FROM AR_Customer  WHERE CustNum = '" & rsCarryOver("AccountNumber") & "'" 
			Set rs_ARCUST = Server.CreateObject("ADODB.Recordset")
			rs_ARCUST.CursorLocation = 3
			Set rs_ARCUST = cnnCarryOver.Execute(SQL_ARCUST)

			If Not rs_ARCUST.Eof Then
				Primary = rs_ARCUST("Salesman")
				Secondary = rs_ARCUST("SecondarySalesman")
				CustType = rs_ARCUST("CustType")
			End If

			If Not IsNumeric(Primary) Then Primary = -1 ' We have to set it to something numeric
			If Not IsNumeric(Secondary) Then Secondary = -1 ' We have to set it to something numeric
			
			'Response.Write("Primary Sales No: " & Primary & ".....Secondary Sales No: " & Secondary & "<br>")						
				
			'''''''''''''If Slsmn = "" OR (Slsmn <> "" And (cint(Slsmn) = cint(Primary) OR cint(Slsmn) = cint(Secondary))) Then ' OLD WAY for salesman filtering				
				
			If (Ubound(SlsmArray) < 0) OR (UBound(SlsmArray) >= 0 AND (InArray(SlsmArray,Primary) OR InArray(SlsmArray,Secondary))) Then ' for salesman filtering
				
				' DO THE ELAPSED TIME
				
				elapsedMinutes = ServiceCallElapsedMinutesOpenTicket(rsCarryOver("MemoNumber"))
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
					elapsedString = elapsedString  & elapsedHours & "h"
				End If
				If int(elapsedMinutes) > 0 Then
					elapsedString = elapsedString  & elapsedMinutes & "m"
				End If

				' Nicer looking submission date & time
				
				submissionHour = Hour(rsCarryOver("submissionDateTime"))
				submissionMinute = Minute(rsCarryOver("submissionDateTime"))
				submissionZeroFactor = "0" & submissionMinute
				submissionAMPM = "AM"
				If submissionHour >= 12 then submissionAMPM = "PM"
				If submissionHour > 12 then submissionHour = submissionHour - 12
				If submissionMinute <= 9 then submissionMinute = submissionZeroFactor	
				
				submissionDateTime = rsCarryOver("submissionDateTime")
				
				ticketDateDisplay = padDate(MONTH(submissionDateTime),2) & "/" & padDate(DAY(submissionDateTime),2) & "/" & padDate(RIGHT(YEAR(submissionDateTime),2),2)
				
						%>
						<td width="5%">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsCarryOver("AccountNumber") %></font>
						</td>							
						<% If CarryoverReportInclCustType = 1 OR CarryoverReportInclTicketNum = 1 Then %>
							<td width="20%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsCarryOver("Company")) > 35 Then Response.Write(Left(rsCarryOver("Company"),35)) Else Response.Write(rsCarryOver("Company"))  %></font>
							</td>
						<% Else %>
							<td width="42%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsCarryOver("Company")) > 35 Then Response.Write(Left(rsCarryOver("Company"),35)) Else Response.Write(rsCarryOver("Company"))  %></font>
							</td>								
						<% End If %>
						<% If CarryoverReportInclCustType = 1 Then %>
							<td width="5%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= CustType %></font>
							</td>
						<% End If %>
						<% If CarryoverReportInclTicketNum = 1 Then %>
							<td width="5%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsCarryOver("MemoNumber") %></font>
							</td>
						<% End If %>
						<%
						PrimarySalesPerson = GetSalesmanNameBySlsmnSequence(Primary)
	    				SecondarySalesPerson = GetSalesmanNameBySlsmnSequence(Secondary)
					    If Instr(PrimarySalesPerson ," ") <> 0 Then
							PrimarySalesPerson =  Left(PrimarySalesPerson,Instr(PrimarySalesPerson ," ")+1)
						End If
						If Instr(SecondarySalesPerson," ") <> 0 Then
							SecondarySalesPerson = Left(SecondarySalesPerson,Instr(SecondarySalesPerson," ")+1) 
						End If
						%>
						<td width="8%">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=PrimarySalesPerson%></font>
						</td>
						<% If CarryoverReportInclTicketNum = 0 Then %>
							<td width="7%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=SecondarySalesPerson%></font>
							</td>
						<% End If %>
						<td width="7%">
							<!--<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=ticketDateDisplay & "&nbsp;" & submissionHour & ":" & submissionMinute & " " & submissionAMPM%></font>-->
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=ticketDateDisplay%></font>
						</td>
						<td width="7%">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=elapsedString%></font>
						</td>
						<td width="5%" align="right">
							<%
								TicketStageVar = GetServiceTicketCurrentStage(rsCarryOver("MemoNumber"))
								TicketStageVar = Replace(TicketStageVar ,"Dispatch Acknowledged","Dispatch Ack")
								TicketStageVar = Replace(TicketStageVar ,"Received","Awaiting Dispatch")
								TicketStageVar = Replace(TicketStageVar ,"Released","Awaiting Dispatch")
								TicketStageVar = Replace(TicketStageVar ,"Dispatch","Disp")
								TicketStageVar = Replace(TicketStageVar ,"Disped","Dispatched")
							%>	
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=TicketStageVar %></font>
						</td>
						<%
						
					ServiceTicketNotes = GetLastServiceTicketNotesByTicket(rsCarryOver("MemoNumber"))
					If Len(ServiceTicketNotes) > 50 Then ServiceTicketNotes = Left(ServiceTicketNotes,50)
						If Len(ServiceTicketNotes) > 1 Then %>
							<td width="31%" style="white-space: nowrap;">
								<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= ServiceTicketNotes %></font>
							</td>
						<% 	Else %>
						<td width="31%" style="white-space: nowrap;">
							<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt">&nbsp;</font>
						</td>
					<% 	End If  %>
			 			 
			</tr>
			<%
			
			'***************************************************************************
			'WRITE OUT THE FILTER DETAILS FOR EACH FILTER ON THIS TICKET
			'***************************************************************************

			NumFiltersThisTicket = 0
			
			Set cnnCustFilters = Server.CreateObject("ADODB.Connection")
			cnnCustFilters.open (MUV_READ("ClientCnnString"))

			SQL_CustFilters = "SELECT * FROM FS_ServiceMemosFilterInfo WHERE ServiceTicketID = '" & rsCarryOver("MemoNumber") & "'" 
			
			'Response.Write("<br>" & SQL_CustFilters & "<br>")
			
			Set rs_CustFilters = Server.CreateObject("ADODB.Recordset")
			rs_CustFilters.CursorLocation = 3
			Set rs_CustFilters = cnnCustFilters.Execute(SQL_CustFilters)

			If Not rs_CustFilters.EOF Then
			
				Do While NOT rs_CustFilters.EOF
				
					CustFilterIntRecID = rs_CustFilters("CustFilterIntRecID")
					ICFilterIntRecID = rs_CustFilters("ICFilterIntRecID")
					Completed = rs_CustFilters("Completed")
					If Completed = 0 Then Completed = "N" Else Completed = "Y"
					CompletedDate = rs_CustFilters("CompletedDate")
					CompletedByUserNo = rs_CustFilters("CompletedByUserNo")
					CompletedByUserName = GetUserFirstAndLastNameByUserNo(CompletedByUserNo)
					FilterDescription = GetFilterDescByFilterIntRecID(ICFilterIntRecID)
					FilterLocation = GetFilterLocationByFilterIntRecID(ICFilterIntRecID)
					
					If FilterLocation = "" Then
						'FilterLocation = "NO LOCATION"
					End If
					
					If CompletedByUserName = "" Then
						CompletedByUserName = "NO TECH NAME"
					End If
				
					If CompletedDate = "" OR IsNull(CompletedDate) Then
						CompletedDate = "NO COMPLETED DATE"
					End If

					%>
					<tr>
						<td>
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font>
						</td>
						<% If Completed = "Y" Then CompletedText = "(Completed)" Else CompletedText = "" %>									
						<td colspan="3">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= GetFilterIDByIntRecID(ICFilterIntRecID) %>&nbsp;-&nbsp;<%= GetFilterDescByIntRecID(ICFilterIntRecID)%>&nbsp;<%=CompletedText%></strong></font>
						</td>								
						<!--<td colspan="2">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FilterLocation %></strong></font>
						</td>-->
						<td colspan="4">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong>&nbsp;</strong></font>
						</td>
					</tr>
					<!--<tr>
						<td colspan="10" align="center">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font>
						</td>																					 
					</tr>-->
					
					<%
					rs_CustFilters.MoveNext
					NumFiltersThisTicket = NumFiltersThisTicket + 2
					
					If rs_CustFilters.EOF Then
						%>
						<tr>
							<td colspan="10" align="center">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font>
							</td>																					 
						</tr>
						<%
						NumFiltersThisTicket = NumFiltersThisTicket + 1
					End If
					
				Loop
				
			Else
			
					%>
					<tr>
						<td>
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font>
						</td>							
						<td colspan="9">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">No filter data available (Ticket:<%= rsCarryOver("MemoNumber")%>)</font>
						</td>								
					</tr>
					<tr>
						<td colspan="10" align="center">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font>
						</td>																					 
					</tr>
					<%
					
					NumFiltersThisTicket = NumFiltersThisTicket + 1

			End If 'EOF Check
		
		RowCount = RowCount + NumFiltersThisTicket
		
	End If ' For Slsmn filtering

		rsCarryOver.Movenext	

		If RowCount > LinesPerPage Then
			%></table><%
			Call Footer
			Call PageHeader
			Call SubHeader
		End If
		
	
	Loop

	
	If Ubound(SlsmArray) >= 0 OR Closed < 1 Then NoBreak = True
	
	Call Footer	
	End If


	'*******************************************************
	'*** END END END END END END END END END END END END END 
	'*** This section is the detail section of all the open
	'*** service tickets
	'*** DOES get included in all versions of the report
	'*** no matter who it is being sent to
	'*******************************************************
	


	If Ubound(SlsmArray) < 0 AND Closed > 0 Then
					
		'*******************************************************
		'*** This section is the breakdown of all closed tickets
		'*** summarized by service tech
		'*** Does NOT get included if being run for one salesman
		'*******************************************************
			
%>
<table border="0" width="<%=PageWidth%>" align="center">
	<tr>
		<td width="100%" align="center">

		<% Call PageHeader
					
			'*************************************************
			'*** Outer loop to compile the Summary Information
			'*************************************************
			LinesPerPage = 30
			
			CloseDate = DateAdd("d",Adjust,Date())
			
			SQL = " SELECT UserNoOfServiceTech, COUNT(*) AS NumCallsForTech FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE'"
			SQL = SQL & " AND Month(RecordCreatedateTime) = " & Month(CloseDate) + MAdjust
			SQL = SQL & " AND Year(RecordCreatedateTime) = " & Year(CloseDate)
			SQL = SQL & " AND Day(RecordCreatedateTime) = " & Day(CloseDate) 
			'SQL = SQL & " AND FilterChange <> 1 "
			SQL = SQL & " GROUP BY UserNoOfServiceTech"
			SQL = SQL & " ORDER BY UserNoOfServiceTech"

			'Response.Write(SQL & "<br>")

			Set cnnCarryOver = Server.CreateObject("ADODB.Connection")
			cnnCarryOver.open (MUV_READ("ClientCnnString"))
			Set rsCarryOver  = Server.CreateObject("ADODB.Recordset")
			rsCarryOver.CursorLocation = 3 
			rsCarryOver.Open SQL, cnnCarryOver 
			
			
			If Not rsCarryOver.EOF Then
			
				%>
				<br><br><br>
				<table border="0" width="<%=PageWidth-200%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
					<tr>
						<td colspan="5">
						<font face="Consolas">
						<hr>
						<center><h2>Closed Calls Summary By Service Tech</h2></center>
						<hr>
						</font>
						</td>
					</tr>
					<tr>
						<td colspan="5">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="5"><font face="Consolas" style="font-size: 12pt">&nbsp;</font></td>
					</tr>
					<tr>
						<td colspan="5">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="5">&nbsp;</td>
					</tr>
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">&nbsp;</font>
						</td>
						<td align="center">
							<font face="Consolas" style="font-size: 14pt"><strong>Total Tickets</strong></font>
						</td>						
						<td align="center">
							<font face="Consolas" style="font-size: 14pt"><strong>Service Tickets</strong></font>
						</td>
						<td align="center">
							<font face="Consolas" style="font-size: 14pt"><strong>Filter Tickets</strong></font>
						</td>
						<td align="center">
							<font face="Consolas" style="font-size: 14pt"><strong>Filters</strong></font>
						</td>
					</tr>

					<tr>
						<td colspan="5">&nbsp;</td>
					</tr>
						
				<%
					If Not rsCarryOver.EOF Then
					
						
						Do While Not (rsCarryOver.EOF)
						
							TechUserNo = rsCarryOver("UserNoOfServiceTech")
							TechName = GetUserDisplayNameByUserNo(TechUserNo)
							TotalCalls =  rsCarryOver("NumCallsForTech")
							ClosedServiceCalls = GetNumberOfClosedServiceTicketsForTech(CloseDate,TechUserNo)
							ClosedFilterCalls = GetNumberOfClosedFilterTicketsForTech(CloseDate,TechUserNo)
							ClosedFilterChanges = GetNumberOfClosedFilterChangesForTech(CloseDate,TechUserNo)
						%>
						<tr>
							<td>
								<font face="Consolas" style="font-size: 14pt"><%= TechName %></font>
							</td>
							<td align="center">
								<font face="Consolas" style="font-size: 14pt"><%= TotalCalls %></font>
							</td>
							<td align="center">
								<font face="Consolas" style="font-size: 14pt"><%= ClosedServiceCalls %></font>
							</td>							
							<td align="center">
								<font face="Consolas" style="font-size: 14pt"><%= ClosedFilterCalls %></font>
							</td>
							<td align="center">
								<font face="Consolas" style="font-size: 14pt"><%= ClosedFilterChanges %></font>
							</td>
						</tr>
						
						<%	
						RowCount = RowCount + 2.5
						rsCarryOver.MoveNext
						Loop

					Else%>
							<tr><td colspan="5">&nbsp;</td></tr>
							<tr><td colspan="5">&nbsp;</td></tr>
							<tr><td colspan="5">&nbsp;</td></tr>
							<tr><td colspan="5">&nbsp;</td></tr>
							<tr>
								<td colspan="5">
									<font face="Consolas" style="font-size: 14pt">
									<hr>
									<center>No service or filter tickets closed today.</center></font>
									<hr>
									<% NoBreak = True %>
								</td>
							</tr>
							<%RowCount = 0

					End If %>
		
				<br/><br/>
		</table>
		</td>
	</tr>				
		<%		
			NoBreak = True	
			Call Footer	

			End If
			 %>
			</table>

<%
	
			'*******************************************************
			'*** END END END END END END END END END END END END END 
			'*** This section is the breakdown of all closed tickets
			'*** summarized by service tech
			'*** Does NOT get included if being run for one salesman
			'*******************************************************
		End IF
		
 %>


</td>
</tr>
</table>


<% If ServiceTicketCarryoverReportIncludeRegions = 1 Then %>
		
	<%

	If Ubound(SlsmArray) < 0 AND Closed > 0 Then
					
		'*******************************************************
		'*** This section is the breakdown of all closed tickets
		'*** summarized by REGION
		'*** Does NOT get included if being run for one salesman
		'*******************************************************
				
	%>
	<table border="0" width="<%=PageWidth%>" align="center">
		<tr>
			<td width="100%" align="center">
	
			<% Call PageHeader
						
				'*************************************************
				'*** Outer loop to compile the Summary Information
				'*************************************************
				LinesPerPage = 30
				RowCount = 0
				
				CloseDate = DateAdd("d",Adjust,Date())

				'******************************************************************************************************
				'Loop Through All The Regions Defined in AR_Reions
				'******************************************************************************************************
				
				SQLRegions = "SELECT * FROM AR_Regions ORDER BY Region ASC"
	
				'Response.Write(SQL & "<br>")
	
				Set cnnCarryOverRegions = Server.CreateObject("ADODB.Connection")
				cnnCarryOverRegions.open (MUV_READ("ClientCnnString"))
				Set rsCarryOverRegions  = Server.CreateObject("ADODB.Recordset")
				rsCarryOverRegions.CursorLocation = 3 
				rsCarryOverRegions.Open SQLRegions, cnnCarryOverRegions 
				
				
				If Not rsCarryOverRegions.EOF Then
				
					%>
					<br><br><br>
					<table border="0" width="<%=PageWidth-200%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
						<tr>
							<td colspan="5">
							<font face="Consolas">
							<hr>
							<center><h2>Closed Calls Summary By Region</h2></center>
							<hr>
							</font>
							</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5"><font face="Consolas" style="font-size: 12pt">&nbsp;</font></td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td>
								<font face="Consolas" style="font-size: 14pt">&nbsp;</font>
							</td>
							<td align="center">
								<font face="Consolas" style="font-size: 14pt"><strong>Total Tickets</strong></font>
							</td>						
							<td align="center">
								<font face="Consolas" style="font-size: 14pt"><strong>Service Tickets</strong></font>
							</td>
							<td align="center">
								<font face="Consolas" style="font-size: 14pt"><strong>Filter Tickets</strong></font>
							</td>
							<td align="center">
								<font face="Consolas" style="font-size: 14pt"><strong>Filters</strong></font>
							</td>
						</tr>
	
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
							
						<%
						
						Do While Not (rsCarryOverRegions.EOF)
						
							CurrentRegionName = rsCarryOverRegions("Region")
							
							SQLCustomer = " SELECT AccountNumber, COUNT(*) AS NumCallsForAccount FROM FS_ServiceMemos WHERE CurrentStatus = 'CLOSE' AND RecordSubType = 'CLOSE'"
							SQLCustomer = SQLCustomer & " AND Month(RecordCreatedateTime) = " & Month(CloseDate) + MAdjust
							SQLCustomer = SQLCustomer & " AND Year(RecordCreatedateTime) = " & Year(CloseDate)
							SQLCustomer = SQLCustomer & " AND Day(RecordCreatedateTime) = " & Day(CloseDate) 
							SQLCustomer = SQLCustomer & " GROUP BY AccountNumber"
							SQLCustomer = SQLCustomer & " ORDER BY AccountNumber"

							Set cnnCustomer = Server.CreateObject("ADODB.Connection")
							cnnCustomer.open (MUV_READ("ClientCnnString"))
							Set rsCustomer  = Server.CreateObject("ADODB.Recordset")
							rsCustomer.CursorLocation = 3 
							
							rsCustomer.Open SQLCustomer, cnnCustomer
							
							'******************************************************************************************************
							'Loop Through All The Customer Accounts With Closed Tickets 
							'If the region of the account, matches the current region, then add to the regions running ticket tally
							'******************************************************************************************************
							
							If NOT rsCustomer.EOF Then
					
								TotalCallsThisRegion =  0			
								ClosedServiceCallsThisRegion = 0
								ClosedFilterCallsThisRegion = 0
								ClosedFilterChangesThisRegion = 0
					
								Do While Not rsCustomer.EOF
							
									CustID = rsCustomer("AccountNumber")
									CustRegion = GetCustRegionByCustID(CustID)
								
									If CurrentRegionName = CustRegion Then
									
										TotalCallsThisRegion =  TotalCallsThisRegion + rsCustomer("NumCallsForAccount")				
										ClosedServiceCallsThisRegion = ClosedServiceCallsThisRegion + GetNumberOfClosedServiceTicketsForCustomerAcct(CloseDate,CustID)
										ClosedFilterCallsThisRegion = ClosedFilterCallsThisRegion + GetNumberOfClosedFilterTicketsForCustomerAcct(CloseDate,CustID)
										ClosedFilterChangesThisRegion = ClosedFilterChangesThisRegion + GetNumberOfClosedFilterChangesForCustomerAcct(CloseDate,CustID)
									
									End If
									
								rsCustomer.MoveNext
								Loop
							
							End If
				
							If TotalCallsThisRegion > 0 Then
							%>
							<tr>
								<td>
									<font face="Consolas" style="font-size: 14pt"><%= CurrentRegionName %></font>
								</td>
								<td align="center">
									<font face="Consolas" style="font-size: 14pt"><%= TotalCallsThisRegion %></font>
								</td>
								<td align="center">
									<font face="Consolas" style="font-size: 14pt"><%= ClosedServiceCallsThisRegion %></font>
								</td>							
								<td align="center">
									<font face="Consolas" style="font-size: 14pt"><%= ClosedFilterCallsThisRegion %></font>
								</td>
								<td align="center">
									<font face="Consolas" style="font-size: 14pt"><%= ClosedFilterChangesThisRegion %></font>
								</td>
							</tr>
							
							<%	
							End If
							
							RowCount = RowCount + 1
							rsCarryOverRegions.MoveNext
							Loop
	
					Else%>
							<tr><td colspan="5">&nbsp;</td></tr>
							<tr><td colspan="5">&nbsp;</td></tr>
							<tr><td colspan="5">&nbsp;</td></tr>
							<tr><td colspan="5">&nbsp;</td></tr>
							<tr>
								<td colspan="5">
									<font face="Consolas" style="font-size: 14pt">
									<hr>
									<center>No service or filter tickets closed today.</center></font>
									<hr>
									<% NoBreak = True %>
								</td>
							</tr>
							<%RowCount = 0

					End If %>
		
				<br/><br/>
			</table>
			</td>
		</tr>				
		<%		
			NoBreak = True	
			Call Footer	
			 %>
			</table>
	
	<%
		
				'*******************************************************
				'*** END END END END END END END END END END END END END 
				'*** This section is the breakdown of all closed tickets
				'*** summarized by service tech
				'*** Does NOT get included if being run for one salesman
				'*******************************************************
		End IF
			
	 %>
	
	
	</td>
	</tr>
	</table>
<% End If %>


</body>
</html>


<%
Sub PageHeader

	RowCount = 0
	%>

	<table border="0" width="100%">
		<tr>
			<td width="50%"><img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png" style="height:55px;"></td>
			<td width="50%">
				<p align="center"><b><font face="Consolas" size="4">MDS Insight Service Ticket Carry Over Report</font></b></p>
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
			<td width="5%">&nbsp;</td>
		</tr>
	</table>
	
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">
	<tr>
		<td width="5%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=GetTerm("Customer")%>#</font></u></strong>
		</td>
		<% If CarryoverReportInclCustType = 1 OR CarryoverReportInclTicketNum = 1 Then %>
			<td width="20%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Name</font></u></strong>
			</td>
		<% Else ' If we are not showing the cust type, then we give that 5%  to the name column%>
			<td width="42%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Name</font></u></strong>
			</td>
		<% End If %>	
		<% If CarryoverReportInclCustType = 1 Then %>
			<td width="5%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Type</font></u></strong>
			</td>
		<% End If %>
		<% If CarryoverReportInclTicketNum = 1 Then %>
			<td width="5%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Ticket</font></u></strong>
			</td>
		<% End If %>	
		<% If SalesmanFileExists = True Then %>
			<td width="8%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=GetTerm("Primary Salesman")%></font></u></strong>
			</td>
		<% End If
		If CarryoverReportInclTicketNum = 0 Then
			If SalesmanFileExists = True Then %>
				<td width="7%">
					<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=GetTerm("Secondary Salesman")%></font></u></strong>
				</td>
			<% End If
		End If %>
		<td width="7%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Orig Date</font></u></strong>
		</td>
		<td width="7%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Age</font></u></strong>
		</td>
		<% If SalesmanFileExists = True Then %>
			<td width="5%" align="right">
		<% Else %>
			<td width="20%" align="right">
		<%End IF%>
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Status</font></u></strong>
		</td>
		<td width="31%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Last Note</font></u></strong>
		</td>
	</tr>
	<tr>
		<td width="5%">&nbsp;</td>
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
				<font face="Consolas" style="font-size: 9pt">directlaunch/service/carryoverReport.asp</font>
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

Function InArray(theArray,theValue)
    dim i, fnd
    fnd = False
    For i = 0 to UBound(theArray)
    	If theValue <> "" Then
	        If cInt(theArray(i)) = cInt(theValue) Then
	            fnd = True
	            Exit For
	        End If
	    End If
    Next
    InArray = fnd
End Function

%>