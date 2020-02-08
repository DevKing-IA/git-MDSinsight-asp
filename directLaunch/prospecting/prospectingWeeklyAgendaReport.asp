<% @ Language = VBScript %>
<%
Response.Buffer = True
Response.Expires = 0
Response.Clear

FontSizeVar = 9
PageNum = 0
NoBreak = False
PageWidth = 1450

Server.ScriptTimeout = 25000
%>
<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Prospecting.asp"-->
<%
dummy=MUV_Write("ClientID","") 'Need this here

ClientKey = Request.QueryString("c")
ownerUserNo = Request.QueryString("u")

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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Prospecting weekly agenda report.<%
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
Set cnn_Settings_Prospecting = Server.CreateObject("ADODB.Connection")
cnn_Settings_Prospecting.open (Session("ClientCnnString"))
Set rs_Settings_Prospecting = Server.CreateObject("ADODB.Recordset")
rs_Settings_Prospecting.CursorLocation = 3 
SQL_Settings_Prospecting = "SELECT * FROM Settings_Prospecting"
Set rs_Settings_Prospecting = cnn_Settings_Prospecting.Execute(SQL_Settings_Prospecting)
If not rs_Settings_Prospecting.EOF Then
	ProspectingWeeklyAgendaReportOnOff = rs_Settings_Prospecting("ProspectingWeeklyAgendaReportOnOff")
	ProspectingWeeklyAgendaReportUserNos = rs_Settings_Prospecting("ProspectingWeeklyAgendaReportUserNos")
	ProspectingWeeklyAgendaReportEmailSubject = rs_Settings_Prospecting("ProspectingWeeklyAgendaReportEmailSubject")
	ProspectingWeeklyAgendaReportAdditionalEmails = rs_Settings_Prospecting("ProspectingWeeklyAgendaReportAdditionalEmails")
Else
	ProspectingWeeklyAgendaReportOnOff = vbFalse
End If
Set rs_Settings_Prospecting = Nothing
cnn_Settings_Prospecting.Close
Set cnn_Settings_Prospecting = Nothing

mondayOfThisWeek = DateAdd("d", -((Weekday(Date()) + 7 - 2) Mod 7), Date())
tuesdayOfThisWeek = DateAdd("d","1",mondayOfThisWeek)
wednesdayOfThisWeek = DateAdd("d","2",mondayOfThisWeek)
thursdayOfThisWeek = DateAdd("d","3",mondayOfThisWeek)
fridayOfThisWeek = DateAdd("d","4",mondayOfThisWeek)

mondayDayNumber = Day(mondayOfThisWeek)
tuesdayDayNumber = Day(tuesdayOfThisWeek)
wednesdayDayNumber = Day(wednesdayOfThisWeek)
thursdayDayNumber = Day(thursdayOfThisWeek)
fridayDayNumber = Day(fridayOfThisWeek)

mondayMonthName = MonthName(Month(mondayOfThisWeek))
tuesdayMonthName = MonthName(Month(tuesdayOfThisWeek))
wednesdayMonthName = MonthName(Month(wednesdayOfThisWeek))
thursdayMonthName = MonthName(Month(thursdayOfThisWeek))
fridayMonthName = MonthName(Month(fridayOfThisWeek))

mondayYearNumber = Year(mondayOfThisWeek)
fridayYearNumber = Year(mondayOfThisWeek)

mondayOfThisWeekTextName = "Monday, " & mondayMonthName & " " & mondayDayNumber
tuesdayOfThisWeekTextName = "Tuesday, " & tuesdayMonthName & " " & tuesdayDayNumber
wednesdayOfThisWeekTextName = "Wednesday, " & wednesdayMonthName & " " & wednesdayDayNumber
thursdayOfThisWeekTextName = "Thursday, " & thursdayMonthName & " " & thursdayDayNumber
fridayOfThisWeekTextName = "Friday, " & fridayMonthName & " " & fridayDayNumber

mondayOfThisWeekHeaderTextName = mondayMonthName & " " & mondayDayNumber & ", " & mondayYearNumber
fridayOfThisWeekHeaderTextName = fridayMonthName & " " & fridayDayNumber & ", " & fridayYearNumber


firstDateCurrentMonth = Month(Date()) & "/01/" & Year(Date())
firstDateNextMonth = DateAdd("m",1,firstDateCurrentMonth)
lastDateCurrentMonth = DateAdd("d",-1,firstDateNextMonth)

If Month(Date()) = 12 Then
	firstDateNextMonth = "01/01/" & (Year(Date()) + 1)
Else
	firstDateNextMonth = (Month(Date()) + 1) & "/01/" & Year(Date())
End If

firstDateFollowingMonth = DateAdd("m",1,firstDateNextMonth)
lastDateNextMonth = DateAdd("d",-1,firstDateFollowingMonth)

nextMonthName = MonthName(Month(firstDateNextMonth))
nextMonthYearNumber = Year(firstDateNextMonth)
%>

<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
	<title>Test Weekly Agenda</title>
	<style>
	
		h1{
			font-family:Arial, Helvetica, sans-serif;
			font-size:45px;
			font-weight:bold;
		}
		
		h2 {
			font-family:Arial, Helvetica, sans-serif;
			text-align:left;
			font-size:40px;
			font-weight:bold;
			margin-bottom:-40px;
		}
		
		.header {
			width:100%;
			border:1px solid #C0C0C0;
			border-collapse:collapse;
			border-spacing:10px;
			padding:5px;
			font-family:Arial, Helvetica, sans-serif;
		}
		.header td {
			border:1px solid #C0C0C0;
			padding:5px;
			background:#fff;
			height:30px;
			width:50%;
		}
	
		.agenda {
			width:100%;
			border:1px solid #C0C0C0;
			border-collapse:collapse;
			border-spacing:10px;
			padding:5px;
			font-family:Arial, Helvetica, sans-serif;
		}
		.agenda caption {
			caption-side:top;
			text-align:left;
			font-size:40px;
			font-weight:bold;
			margin-bottom:20px;
		}
		.agenda th {
			border:1px solid #C0C0C0;
			padding:5px;
			background:#D8D8D8;
			height:30px;
			width:50%;
			text-align:left;
			font-size:25px;
		}
		
		.agenda th.blue {
			border:1px solid #C0C0C0;
			padding:5px;
			background:#C2D8E4;
			height:30px;
			width:50%;
			text-align:left;
			font-size:25px;
		}
		
		.agenda td {
			border:1px solid #C0C0C0;
			padding:5px;
			height:315px;
			width:50%;
			vertical-align: top;
		}
	

		.agenda-expired {
			width:100%;
			border:1px solid #C0C0C0;
			border-collapse:collapse;
			border-spacing:10px;
			padding:5px;
			font-family:Arial, Helvetica, sans-serif;
		}
		.agenda-expired caption {
			caption-side:top;
			text-align:left;
			font-size:40px;
			font-weight:bold;
			margin-bottom:20px;
		}
		.agenda-expired th {
			border:1px solid #C0C0C0;
			padding:5px;
			background:#D8D8D8;
			height:30px;
			width:50%;
			text-align:left;
			font-size:25px;
		}
		
		.agenda-expired th.blue {
			border:1px solid #C0C0C0;
			padding:5px;
			background:#C2D8E4;
			height:30px;
			width:50%;
			text-align:left;
			font-size:25px;
		}
		
		.agenda-expired td {
			border:1px solid #C0C0C0;
			padding:5px;
			width:50%;
			vertical-align: top;
		}	
		ul {
		    margin: 10px 0px;
		    padding: 0 25px;
		    list-style: none;
	   	}
		li {
		    /*border-bottom-style: dotted;
		    border-bottom-width: 1px;
		    border-bottom-color: #C2D8E4;*/
		    padding:3px 4px;
		    line-height: 1.5
		}
		li:before {
		    /*content:"";
		    border-radius: 50%;
		    border-style: solid;
		    border-width: 1px;*/
		    width: 10px;
		    height: 10px;
		    margin-right:0.3em;
		    position: relative;
		    display: inline-block;
		}
		.company{
			font-weight:bold;
			color:blue;
		}
		.company a{
			font-weight:bold;
			color:#4A86E8;
			text-decoration:none;
		}
		
		.calendar {
			 table-display: fixed;
			 border: 2px solid #4e4f4a;
			 width: 25%;
			 font-family:Arial, Helvetica, sans-serif;
		}
		 .calendar_day_header {
			 text-align: center;
			 width: 14.2857142857%;
			 vertical-align: middle;
		}
		 .calendar_day_cell {
			 border: 1px solid #D8D8D8;
			 text-align: center;
			 width: 14.2857142857%;
			 vertical-align: middle;
		}
		
		 .calendar_day_header:first-child, .calendar_day_cell:first-child {
			 border-left: none;
		}
		 .calendar_day_header:last-child, .calendar_day_cell:last-child {
			 border-right: none;
		}
		 .calendar_day_header, .calendar_day_cell {
			 padding: 0.75rem 0 1.5rem;
		}
		 .calendar_banner_month {
			 border: 2px solid #4e4f4a;
			 border-bottom: none;
			 text-align: center;
			 padding: 0.75rem;
			 background-color:#4e4f4a;
		}
		 .calendar_banner_month h1 {
			 color: #fff;
			 display: inline-block;
			 font-size: 30px;
			 font-weight: 700;
			 text-transform: uppercase;
		}
		 .calendar_day_header {
			 font-size: 20px;
			 letter-spacing: 0.1em;
			 text-transform: uppercase;
		}
		 .calendar_day_cell {
			 font-size: 20px;
			 position: relative;
			 color: #6A6C6E;
		}
		 


	</style>
  </head>
  <body bgcolor="#FFFFFF" text="#000000" link="#000080" vlink="#000080" alink="#000080" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
  
  		<table class="header" width="<%=PageWidth%>">
  			<tr>
  				<td width="80%" align="left" style="vertical-align:middle; border:0px;width:80%">
  					<img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png" style="height:55px;">
  					<h2><%= GetUserDisplayNameByUserNo(ownerUserNo) %>'s Prospecting Agenda</h2>
  					<h1><br><%= mondayOfThisWeekHeaderTextName %> - <br><%= fridayOfThisWeekHeaderTextName %></h1>
    			</td>
  				<td width="10%" align="right" style="vertical-align:middle; border:0px; width:10%;">
 					<table class="calendar">
				        <caption class="calendar_banner_month">
				            <h1><%= mondayMonthName %>&nbsp;<%= mondayYearNumber %></h1>
				        </caption>
				        <thead>
				            <tr>
				                <th class="calendar_day_header">Su</th>
				                <th class="calendar_day_header">Mo</th>
				                <th class="calendar_day_header">Tu</th>
				                <th class="calendar_day_header">We</th>
				                <th class="calendar_day_header">Th</th>
				                <th class="calendar_day_header">Fr</th>
				                <th class="calendar_day_header">Sa</th>
				            </tr>
				        </thead>
				        <tbody>
				        
  					
		  					<%
								loopDate = firstDateCurrentMonth
								dayCounter = 0
								
								do
								    If dayCounter = 0 Then 
								    	%><tr><%
								    End If
								    
								    If loopDate = firstDateCurrentMonth Then
								    	for i = 0 to weekday(firstDateCurrentMonth)-2
								    		dayCounter = dayCounter + 1
								    		%><td class="calendar_day_cell"></td><%
								    	next
								    End If
								
								    CurrentDayNumber = Day((loopDate))
								    CurrentDayName = WeekdayName(weekday(loopDate))
								    
								    If loopDate = mondayOfThisWeek OR loopDate = tuesdayOfThisWeek OR loopDate = wednesdayOfThisWeek OR loopDate = thursdayOfThisWeek OR loopDate = fridayOfThisWeek Then
								    	%><td class="calendar_day_cell"><strong><%= CurrentDayNumber %></strong></td><%
								    Else
								    	%><td class="calendar_day_cell"><%= CurrentDayNumber %></td><%
								    End If
								    
								    If loopDate = lastDateCurrentMonth Then
								    	for x = weekday(lastDateCurrentMonth) to 6
								    		dayCounter = dayCounter + 1
								    		%><td class="calendar_day_cell"></td><%
								    	next
								    End If
								    
								    loopDate = DateAdd("d",1,loopDate)
								    dayCounter = dayCounter + 1
								    
								    If dayCounter >= 7 Then 
								    	dayCounter = 0
								    	%></tr><%
								    End If
								    
								loop until loopDate > lastDateCurrentMonth					
		  					%>
		  					
				        </tbody>
				    </table>
				 </td>
				 <td width="10%" align="right" style="vertical-align:middle; border:0px;width:10%">
 					<table class="calendar">
				        <caption class="calendar_banner_month">
				            <h1><%= nextMonthName %>&nbsp;<%= nextMonthYearNumber %></h1>
				        </caption>
				        <thead>
				            <tr>
				                <th class="calendar_day_header">Su</th>
				                <th class="calendar_day_header">Mo</th>
				                <th class="calendar_day_header">Tu</th>
				                <th class="calendar_day_header">We</th>
				                <th class="calendar_day_header">Th</th>
				                <th class="calendar_day_header">Fr</th>
				                <th class="calendar_day_header">Sa</th>
				            </tr>
				        </thead>
				        <tbody>
				        
  					
		  					<%
								loopDate = firstDateNextMonth 
								dayCounter = 0
								
								do
								    If dayCounter = 0 Then 
								    	%><tr><%
								    End If
								    
								    If loopDate = firstDateNextMonth Then
								    	for i = 0 to weekday(firstDateNextMonth)-2
								    		dayCounter = dayCounter + 1
								    		%><td class="calendar_day_cell"></td><%
								    	next
								    End If
								
								    CurrentDayNumber = Day((loopDate))
								    CurrentDayName = WeekdayName(weekday(loopDate))
								    
								    If loopDate = mondayOfThisWeek OR loopDate = tuesdayOfThisWeek OR loopDate = wednesdayOfThisWeek OR loopDate = thursdayOfThisWeek OR loopDate = fridayOfThisWeek Then
								    	%><td class="calendar_day_cell"><strong><%= CurrentDayNumber %></strong></td><%
								    Else
								    	%><td class="calendar_day_cell"><%= CurrentDayNumber %></td><%
								    End If
								    
								    If loopDate = lastDateNextMonth Then
								    	for x = weekday(lastDateNextMonth) to 6
								    		dayCounter = dayCounter + 1
								    		%><td class="calendar_day_cell"></td><%
								    	next
								    End If
								    
								    loopDate = DateAdd("d",1,loopDate)
								    dayCounter = dayCounter + 1
								    
								    If dayCounter >= 7 Then 
								    	dayCounter = 0
								    	%></tr><%
								    End If
								    
								loop until loopDate > lastDateNextMonth					
		  					%>
		  					
				        </tbody>
				    </table>
  				</td>
  				
  			</tr>
  		</table>
  		
  		<%
  		
		SQLWeeklyAgendaReport = "SELECT * FROM PR_ProspectActivities INNER JOIN PR_Prospects ON PR_Prospects.InternalRecordIdentifier = PR_ProspectActivities.ProspectRecID "
		SQLWeeklyAgendaReport = SQLWeeklyAgendaReport & " WHERE "
		SQLWeeklyAgendaReport = SQLWeeklyAgendaReport & " (PR_Prospects.OwnerUserNo = " & ownerUserNo & ") AND "
		SQLWeeklyAgendaReport = SQLWeeklyAgendaReport & " (PR_Prospects.Pool = 'Live') AND "
		SQLWeeklyAgendaReport = SQLWeeklyAgendaReport & " (PR_ProspectActivities.Status <> 'Completed' OR PR_ProspectActivities.Status <> 'Cancelled' OR PR_ProspectActivities.Status IS NULL) AND "
		SQLWeeklyAgendaReport = SQLWeeklyAgendaReport & " (PR_ProspectActivities.ActivityDueDate >= '" & mondayOfThisWeek & "') AND "
		SQLWeeklyAgendaReport = SQLWeeklyAgendaReport & " (PR_ProspectActivities.ActivityDueDate <= '" & fridayOfThisWeek & "') "
		SQLWeeklyAgendaReport = SQLWeeklyAgendaReport & " ORDER BY PR_ProspectActivities.ActivityDueDate DESC "
	
		Set cnnWeeklyAgendaReport = Server.CreateObject("ADODB.Connection")
		cnnWeeklyAgendaReport.open(Session("ClientCnnString"))
		Set rsWeeklyAgendaReport  = Server.CreateObject("ADODB.Recordset")
		rsWeeklyAgendaReport.CursorLocation = 3 
		rsWeeklyAgendaReport.Open SQLWeeklyAgendaReport, cnnWeeklyAgendaReport 
					
		'Response.Write(SQLWeeklyAgendaReport & "<br>")
		
		If Not rsWeeklyAgendaReport.EOF Then
		
			Set cnnWeeklyAgendaReportSingleDay = Server.CreateObject("ADODB.Connection")
			cnnWeeklyAgendaReportSingleDay.open(Session("ClientCnnString"))
			Set rsWeeklyAgendaReportSingleDay  = Server.CreateObject("ADODB.Recordset")
			rsWeeklyAgendaReportSingleDay.CursorLocation = 3 
		
			%>
			
				<table class="agenda">
					<thead>
						<tr>
							<th><%= mondayOfThisWeekTextName %></th>
							<th><%= tuesdayOfThisWeekTextName %></th>
						</tr>
					</thead>
					<tbody>
						<tr>
							<!-------------------------------------------------------------------------------------------------------------------------------------->
							<!-- MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY MONDAY 
							<!-------------------------------------------------------------------------------------------------------------------------------------->
							<td>
							<%
							SQLWeeklyAgendaReportSingleDay = "SELECT * FROM PR_ProspectActivities "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " INNER JOIN PR_Prospects ON "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_Prospects.InternalRecordIdentifier = PR_ProspectActivities.ProspectRecID "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " WHERE "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_Prospects.OwnerUserNo = " & ownerUserNo & ") AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_Prospects.Pool = 'Live') AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_ProspectActivities.Status <> 'Completed' OR "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_ProspectActivities.Status <> 'Cancelled' OR "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_ProspectActivities.Status IS NULL) AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (Cast(PR_ProspectActivities.ActivityDueDate as Date) = '" & mondayOfThisWeek & "') "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " ORDER BY PR_ProspectActivities.ActivityDueDate ASC "
	
							rsWeeklyAgendaReportSingleDay.Open SQLWeeklyAgendaReportSingleDay, cnnWeeklyAgendaReportSingleDay
		
							If Not rsWeeklyAgendaReportSingleDay.EOF Then 
							
								%><ul><%
								
								Do While NOT rsWeeklyAgendaReportSingleDay.EOF
								
									ActivityDueDate = rsWeeklyAgendaReportSingleDay("ActivityDueDate")
																		
									ActivityProspectRecID = rsWeeklyAgendaReportSingleDay("ProspectRecID")
									ProspectName = GetProspectNameByNumber(ActivityProspectRecID)
									ProspectLocationStreet = GetProspectStreetByNumber(ActivityProspectRecID)
									ProspectLocationCity = GetProspectCityByNumber(ActivityProspectRecID)
									ProspectLocationState = GetProspectStateByNumber(ActivityProspectRecID)
									ProspectLocationPostalCode = GetProspectPostalCodeByNumber(ActivityProspectRecID)
									
									ActivityRecID = rsWeeklyAgendaReportSingleDay("ActivityRecID")
									ActivityDesc = GetActivityByNum(ActivityRecID)
									ActivityNotes = rsWeeklyAgendaReportSingleDay("Notes")
									
									ActivityIsAppointment = rsWeeklyAgendaReportSingleDay("ActivityIsAppointment")
									ActivityAppointmentDuration = rsWeeklyAgendaReportSingleDay("ActivityAppointmentDuration")
									
									ActivityIsMeeting = rsWeeklyAgendaReportSingleDay("ActivityIsMeeting")
									ActivityMeetingDuration = rsWeeklyAgendaReportSingleDay("ActivityMeetingDuration")
									ActivityMeetingLocation = rsWeeklyAgendaReportSingleDay("ActivityMeetingLocation")
									
									If ActivityIsAppointment = "" OR IsNull(ActivityIsAppointment) Then
										ActivityIsAppointment = 0
									End If
									
									If ActivityIsMeeting = "" OR IsNull(ActivityIsMeeting) Then
										ActivityIsMeeting = 0
									End If
									
									'**********************************************************************************************
									'Obtain the start time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									UnformattedActivityStartTime = timevalue(hour(ActivityDueDate) & ":" & minute(ActivityDueDate))
									
									If minute(ActivityDueDate) < 10 AND minute(ActivityDueDate) > 0 Then
										minuteActivityDueDate = "0" & minute(ActivityDueDate)
									ElseIf minute(ActivityDueDate) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDate = minute(ActivityDueDate)
									End If
									
									If hour(ActivityDueDate) > 12 Then
										ActivityStartTime = hour(ActivityDueDate) - 12  & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									Else
										ActivityStartTime = hour(ActivityDueDate) & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									End If
									
									'**********************************************************************************************
									'Determine if the activity is an appointment or meeting, to see which duration to add
									' to the start time, to arrive at the end time
									''**********************************************************************************************
									
									If cInt(ActivityIsAppointment) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityAppointmentDuration,ActivityDueDate)
									ElseIf cInt(ActivityIsMeeting) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityMeetingDuration,ActivityDueDate)
									Else
										ActivityDueDateEndTime = DateAdd("n",0,ActivityDueDate)
									End If
									
									'**********************************************************************************************
									'Obtain the end time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									
									UnformattedActivityEndTime = timevalue(hour(ActivityDueDateEndTime) & ":" & minute(ActivityDueDateEndTime))
									
									If minute(ActivityDueDateEndTime) < 10 AND minute(ActivityDueDateEndTime) > 0 Then
										minuteActivityDueDateEndTime = "0" & minute(ActivityDueDateEndTime)
									ElseIf minute(ActivityDueDateEndTime) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDateEndTime = minute(ActivityDueDateEndTime)
									End If
									
									If hour(ActivityDueDateEndTime) > 12 Then
										ActivityEndTime = hour(ActivityDueDateEndTime) - 12  & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									Else
										ActivityEndTime = hour(ActivityDueDateEndTime) & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									End If
									
									'''http://dev2.mdsinsight.com/prospecting/viewProspectDetail.asp?i=3462
									
									QuickLinkURLDestination = "viewProspect-" & ActivityProspectRecID
									QuickLoginURL = baseURL & "ql.asp?c=" & ClientKey & "&u=" & ownerUserNo & "&d=" & QuickLinkURLDestination

									
									%>
									<% If cInt(ActivityIsAppointment) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %>&nbsp;-&nbsp;<%= ActivityEndTime %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span>.
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>
								    	</li>
								    <% ElseIf cInt(ActivityIsMeeting) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %>&nbsp;-&nbsp;<%= ActivityEndTime %></strong>&nbsp;
								    		<%= ActivityDesc %><br>
								    		<strong>Location</strong>:&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
									<% Else %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;
								    		<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
								    <% End If %>
													    
								<%
								rsWeeklyAgendaReportSingleDay.MoveNext
								Loop
								
								%></ul><%
								
							End If

							%>					
							</td>
							
							<!-------------------------------------------------------------------------------------------------------------------------------------->
							<!-- TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY TUESDAY 
							<!-------------------------------------------------------------------------------------------------------------------------------------->
							
							<td>
							<%
							
							rsWeeklyAgendaReportSingleDay.close
							cnnWeeklyAgendaReportSingleDay.close
							set rsWeeklyAgendaReportSingleDay=nothing
							set cnnWeeklyAgendaReportSingleDay=nothing
							
							Set cnnWeeklyAgendaReportSingleDay = Server.CreateObject("ADODB.Connection")
							cnnWeeklyAgendaReportSingleDay.open(Session("ClientCnnString"))
							Set rsWeeklyAgendaReportSingleDay  = Server.CreateObject("ADODB.Recordset")
							rsWeeklyAgendaReportSingleDay.CursorLocation = 3 
			
							SQLWeeklyAgendaReportSingleDay = "SELECT * FROM PR_ProspectActivities "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " INNER JOIN PR_Prospects ON "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_Prospects.InternalRecordIdentifier = PR_ProspectActivities.ProspectRecID "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " WHERE "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_Prospects.OwnerUserNo = " & ownerUserNo & ") AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_Prospects.Pool = 'Live') AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_ProspectActivities.Status <> 'Completed' OR "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_ProspectActivities.Status <> 'Cancelled' OR "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_ProspectActivities.Status IS NULL) AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (Cast(PR_ProspectActivities.ActivityDueDate as Date) = '" & tuesdayOfThisWeek & "') "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " ORDER BY PR_ProspectActivities.ActivityDueDate ASC "
	
							rsWeeklyAgendaReportSingleDay.Open SQLWeeklyAgendaReportSingleDay, cnnWeeklyAgendaReportSingleDay
		
							If Not rsWeeklyAgendaReportSingleDay.EOF Then 
							
								%><ul><%
								
								Do While NOT rsWeeklyAgendaReportSingleDay.EOF
								
									ActivityDueDate = rsWeeklyAgendaReportSingleDay("ActivityDueDate")
																		
									ActivityProspectRecID = rsWeeklyAgendaReportSingleDay("ProspectRecID")
									ProspectName = GetProspectNameByNumber(ActivityProspectRecID)
									ProspectLocationStreet = GetProspectStreetByNumber(ActivityProspectRecID)
									ProspectLocationCity = GetProspectCityByNumber(ActivityProspectRecID)
									ProspectLocationState = GetProspectStateByNumber(ActivityProspectRecID)
									ProspectLocationPostalCode = GetProspectPostalCodeByNumber(ActivityProspectRecID)
									
									ActivityRecID = rsWeeklyAgendaReportSingleDay("ActivityRecID")
									ActivityDesc = GetActivityByNum(ActivityRecID)
									ActivityNotes = rsWeeklyAgendaReportSingleDay("Notes")
									
									ActivityIsAppointment = rsWeeklyAgendaReportSingleDay("ActivityIsAppointment")
									ActivityAppointmentDuration = rsWeeklyAgendaReportSingleDay("ActivityAppointmentDuration")
									
									ActivityIsMeeting = rsWeeklyAgendaReportSingleDay("ActivityIsMeeting")
									ActivityMeetingDuration = rsWeeklyAgendaReportSingleDay("ActivityMeetingDuration")
									ActivityMeetingLocation = rsWeeklyAgendaReportSingleDay("ActivityMeetingLocation")
									
									If ActivityIsAppointment = "" OR IsNull(ActivityIsAppointment) Then
										ActivityIsAppointment = 0
									End If
									
									If ActivityIsMeeting = "" OR IsNull(ActivityIsMeeting) Then
										ActivityIsMeeting = 0
									End If
									
									'**********************************************************************************************
									'Obtain the start time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									UnformattedActivityStartTime = timevalue(hour(ActivityDueDate) & ":" & minute(ActivityDueDate))
									
									If minute(ActivityDueDate) < 10 AND minute(ActivityDueDate) > 0 Then
										minuteActivityDueDate = "0" & minute(ActivityDueDate)
									ElseIf minute(ActivityDueDate) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDate = minute(ActivityDueDate)
									End If
									
									If hour(ActivityDueDate) > 12 Then
										ActivityStartTime = hour(ActivityDueDate) - 12  & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									Else
										ActivityStartTime = hour(ActivityDueDate) & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									End If
									
									'**********************************************************************************************
									'Determine if the activity is an appointment or meeting, to see which duration to add
									' to the start time, to arrive at the end time
									''**********************************************************************************************
									
									If cInt(ActivityIsAppointment) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityAppointmentDuration,ActivityDueDate)
									ElseIf cInt(ActivityIsMeeting) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityMeetingDuration,ActivityDueDate)
									Else
										ActivityDueDateEndTime = DateAdd("n",0,ActivityDueDate)
									End If
									
									'**********************************************************************************************
									'Obtain the end time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									
									UnformattedActivityEndTime = timevalue(hour(ActivityDueDateEndTime) & ":" & minute(ActivityDueDateEndTime))
									
									If minute(ActivityDueDateEndTime) < 10 AND minute(ActivityDueDateEndTime) > 0 Then
										minuteActivityDueDateEndTime = "0" & minute(ActivityDueDateEndTime)
									ElseIf minute(ActivityDueDateEndTime) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDateEndTime = minute(ActivityDueDateEndTime)
									End If
									
									If hour(ActivityDueDateEndTime) > 12 Then
										ActivityEndTime = hour(ActivityDueDateEndTime) - 12  & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									Else
										ActivityEndTime = hour(ActivityDueDateEndTime) & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									End If
									
									
									QuickLinkURLDestination = "viewProspect-" & ActivityProspectRecID
									QuickLoginURL = baseURL & "ql.asp?c=" & ClientKey & "&u=" & ownerUserNo & "&d=" & QuickLinkURLDestination

									
									%>
									<% If cInt(ActivityIsAppointment) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %>&nbsp;-&nbsp;<%= ActivityEndTime %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span>.
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>
								    	</li>
								    <% ElseIf cInt(ActivityIsMeeting) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %>&nbsp;-&nbsp;<%= ActivityEndTime %></strong>&nbsp;
								    		<%= ActivityDesc %><br>
								    		<strong>Location</strong>:&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
									<% Else %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;
								    		<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
								    <% End If %>
													    
								<%
								rsWeeklyAgendaReportSingleDay.MoveNext
								Loop
								
								%></ul><%
								
							End If

							%>					
							</td>
						</tr>
					</tbody>
				</table>
				
				<table class="agenda">
					<thead>
						<tr>
							<th><%= wednesdayOfThisWeekTextName %></th>
							<th><%= thursdayOfThisWeekTextName %></th>
						</tr>
					</thead>
					<tbody>
						<tr>
							<!-------------------------------------------------------------------------------------------------------------------------------------->
							<!-- WEDNESDAY  WEDNESDAY  WEDNESDAY  WEDNESDAY  WEDNESDAY  WEDNESDAY  WEDNESDAY  WEDNESDAY  WEDNESDAY  WEDNESDAY  WEDNESDAY  WEDNESDAY  
							<!-------------------------------------------------------------------------------------------------------------------------------------->
							
							
							<td>
							<%
							
							rsWeeklyAgendaReportSingleDay.close
							cnnWeeklyAgendaReportSingleDay.close
							set rsWeeklyAgendaReportSingleDay=nothing
							set cnnWeeklyAgendaReportSingleDay=nothing
							
							Set cnnWeeklyAgendaReportSingleDay = Server.CreateObject("ADODB.Connection")
							cnnWeeklyAgendaReportSingleDay.open(Session("ClientCnnString"))
							Set rsWeeklyAgendaReportSingleDay  = Server.CreateObject("ADODB.Recordset")
							rsWeeklyAgendaReportSingleDay.CursorLocation = 3 
							
							SQLWeeklyAgendaReportSingleDay = "SELECT * FROM PR_ProspectActivities "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " INNER JOIN PR_Prospects ON "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_Prospects.InternalRecordIdentifier = PR_ProspectActivities.ProspectRecID "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " WHERE "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_Prospects.OwnerUserNo = " & ownerUserNo & ") AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_Prospects.Pool = 'Live') AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_ProspectActivities.Status <> 'Completed' OR "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_ProspectActivities.Status <> 'Cancelled' OR "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_ProspectActivities.Status IS NULL) AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (Cast(PR_ProspectActivities.ActivityDueDate as Date) = '" & wednesdayOfThisWeek & "') "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " ORDER BY PR_ProspectActivities.ActivityDueDate ASC "
	
							rsWeeklyAgendaReportSingleDay.Open SQLWeeklyAgendaReportSingleDay, cnnWeeklyAgendaReportSingleDay
		
							If Not rsWeeklyAgendaReportSingleDay.EOF Then 
							
								%><ul><%
								
								Do While NOT rsWeeklyAgendaReportSingleDay.EOF
								
									ActivityDueDate = rsWeeklyAgendaReportSingleDay("ActivityDueDate")
																		
									ActivityProspectRecID = rsWeeklyAgendaReportSingleDay("ProspectRecID")
									ProspectName = GetProspectNameByNumber(ActivityProspectRecID)
									ProspectLocationStreet = GetProspectStreetByNumber(ActivityProspectRecID)
									ProspectLocationCity = GetProspectCityByNumber(ActivityProspectRecID)
									ProspectLocationState = GetProspectStateByNumber(ActivityProspectRecID)
									ProspectLocationPostalCode = GetProspectPostalCodeByNumber(ActivityProspectRecID)
									
									ActivityRecID = rsWeeklyAgendaReportSingleDay("ActivityRecID")
									ActivityDesc = GetActivityByNum(ActivityRecID)
									ActivityNotes = rsWeeklyAgendaReportSingleDay("Notes")
									
									ActivityIsAppointment = rsWeeklyAgendaReportSingleDay("ActivityIsAppointment")
									ActivityAppointmentDuration = rsWeeklyAgendaReportSingleDay("ActivityAppointmentDuration")
									
									ActivityIsMeeting = rsWeeklyAgendaReportSingleDay("ActivityIsMeeting")
									ActivityMeetingDuration = rsWeeklyAgendaReportSingleDay("ActivityMeetingDuration")
									ActivityMeetingLocation = rsWeeklyAgendaReportSingleDay("ActivityMeetingLocation")
									
									If ActivityIsAppointment = "" OR IsNull(ActivityIsAppointment) Then
										ActivityIsAppointment = 0
									End If
									
									If ActivityIsMeeting = "" OR IsNull(ActivityIsMeeting) Then
										ActivityIsMeeting = 0
									End If
									
									'**********************************************************************************************
									'Obtain the start time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									UnformattedActivityStartTime = timevalue(hour(ActivityDueDate) & ":" & minute(ActivityDueDate))
									
									If minute(ActivityDueDate) < 10 AND minute(ActivityDueDate) > 0 Then
										minuteActivityDueDate = "0" & minute(ActivityDueDate)
									ElseIf minute(ActivityDueDate) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDate = minute(ActivityDueDate)
									End If
									
									If hour(ActivityDueDate) > 12 Then
										ActivityStartTime = hour(ActivityDueDate) - 12  & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									Else
										ActivityStartTime = hour(ActivityDueDate) & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									End If
									
									'**********************************************************************************************
									'Determine if the activity is an appointment or meeting, to see which duration to add
									' to the start time, to arrive at the end time
									''**********************************************************************************************
									
									If cInt(ActivityIsAppointment) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityAppointmentDuration,ActivityDueDate)
									ElseIf cInt(ActivityIsMeeting) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityMeetingDuration,ActivityDueDate)
									Else
										ActivityDueDateEndTime = DateAdd("n",0,ActivityDueDate)
									End If
									
									'**********************************************************************************************
									'Obtain the end time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									
									UnformattedActivityEndTime = timevalue(hour(ActivityDueDateEndTime) & ":" & minute(ActivityDueDateEndTime))
									
									If minute(ActivityDueDateEndTime) < 10 AND minute(ActivityDueDateEndTime) > 0 Then
										minuteActivityDueDateEndTime = "0" & minute(ActivityDueDateEndTime)
									ElseIf minute(ActivityDueDateEndTime) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDateEndTime = minute(ActivityDueDateEndTime)
									End If
									
									If hour(ActivityDueDateEndTime) > 12 Then
										ActivityEndTime = hour(ActivityDueDateEndTime) - 12  & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									Else
										ActivityEndTime = hour(ActivityDueDateEndTime) & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									End If
									
									QuickLinkURLDestination = "viewProspect-" & ActivityProspectRecID
									QuickLoginURL = baseURL & "ql.asp?c=" & ClientKey & "&u=" & ownerUserNo & "&d=" & QuickLinkURLDestination

									
									%>
									<% If cInt(ActivityIsAppointment) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %>&nbsp;-&nbsp;<%= ActivityEndTime %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span>.
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>
								    	</li>
								    <% ElseIf cInt(ActivityIsMeeting) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %>&nbsp;-&nbsp;<%= ActivityEndTime %></strong>&nbsp;
								    		<%= ActivityDesc %><br>
								    		<strong>Location</strong>:&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
									<% Else %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;
								    		<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
								    <% End If %>
													    
								<%
								rsWeeklyAgendaReportSingleDay.MoveNext
								Loop
								
								%></ul><%
								
							End If

							%>					
							</td>
							
							<!-------------------------------------------------------------------------------------------------------------------------------------->
							<!-- THURSDAY  THURSDAY  THURSDAY  THURSDAY  THURSDAY  THURSDAY  THURSDAY  THURSDAY  THURSDAY  THURSDAY  THURSDAY  THURSDAY  THURSDAY  
							<!-------------------------------------------------------------------------------------------------------------------------------------->
							
							
							<td>
							<%
							
							rsWeeklyAgendaReportSingleDay.close
							cnnWeeklyAgendaReportSingleDay.close
							set rsWeeklyAgendaReportSingleDay=nothing
							set cnnWeeklyAgendaReportSingleDay=nothing
							
							Set cnnWeeklyAgendaReportSingleDay = Server.CreateObject("ADODB.Connection")
							cnnWeeklyAgendaReportSingleDay.open(Session("ClientCnnString"))
							Set rsWeeklyAgendaReportSingleDay  = Server.CreateObject("ADODB.Recordset")
							rsWeeklyAgendaReportSingleDay.CursorLocation = 3 
							
							SQLWeeklyAgendaReportSingleDay = "SELECT * FROM PR_ProspectActivities "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " INNER JOIN PR_Prospects ON "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_Prospects.InternalRecordIdentifier = PR_ProspectActivities.ProspectRecID "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " WHERE "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_Prospects.OwnerUserNo = " & ownerUserNo & ") AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_Prospects.Pool = 'Live') AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_ProspectActivities.Status <> 'Completed' OR "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_ProspectActivities.Status <> 'Cancelled' OR "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_ProspectActivities.Status IS NULL) AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (Cast(PR_ProspectActivities.ActivityDueDate as Date) = '" & thursdayOfThisWeek & "') "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " ORDER BY PR_ProspectActivities.ActivityDueDate ASC "
							
							rsWeeklyAgendaReportSingleDay.Open SQLWeeklyAgendaReportSingleDay, cnnWeeklyAgendaReportSingleDay
		
							If Not rsWeeklyAgendaReportSingleDay.EOF Then 
							
								%><ul><%
								
								Do While NOT rsWeeklyAgendaReportSingleDay.EOF
								
									ActivityDueDate = rsWeeklyAgendaReportSingleDay("ActivityDueDate")
																		
									ActivityProspectRecID = rsWeeklyAgendaReportSingleDay("ProspectRecID")
									ProspectName = GetProspectNameByNumber(ActivityProspectRecID)
									ProspectLocationStreet = GetProspectStreetByNumber(ActivityProspectRecID)
									ProspectLocationCity = GetProspectCityByNumber(ActivityProspectRecID)
									ProspectLocationState = GetProspectStateByNumber(ActivityProspectRecID)
									ProspectLocationPostalCode = GetProspectPostalCodeByNumber(ActivityProspectRecID)
									
									ActivityRecID = rsWeeklyAgendaReportSingleDay("ActivityRecID")
									ActivityDesc = GetActivityByNum(ActivityRecID)
									ActivityNotes = rsWeeklyAgendaReportSingleDay("Notes")
									
									ActivityIsAppointment = rsWeeklyAgendaReportSingleDay("ActivityIsAppointment")
									ActivityAppointmentDuration = rsWeeklyAgendaReportSingleDay("ActivityAppointmentDuration")
									
									ActivityIsMeeting = rsWeeklyAgendaReportSingleDay("ActivityIsMeeting")
									ActivityMeetingDuration = rsWeeklyAgendaReportSingleDay("ActivityMeetingDuration")
									ActivityMeetingLocation = rsWeeklyAgendaReportSingleDay("ActivityMeetingLocation")
									
									If ActivityIsAppointment = "" OR IsNull(ActivityIsAppointment) Then
										ActivityIsAppointment = 0
									End If
									
									If ActivityIsMeeting = "" OR IsNull(ActivityIsMeeting) Then
										ActivityIsMeeting = 0
									End If
									
									'**********************************************************************************************
									'Obtain the start time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									UnformattedActivityStartTime = timevalue(hour(ActivityDueDate) & ":" & minute(ActivityDueDate))
									
									If minute(ActivityDueDate) < 10 AND minute(ActivityDueDate) > 0 Then
										minuteActivityDueDate = "0" & minute(ActivityDueDate)
									ElseIf minute(ActivityDueDate) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDate = minute(ActivityDueDate)
									End If
									
									If hour(ActivityDueDate) > 12 Then
										ActivityStartTime = hour(ActivityDueDate) - 12  & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									Else
										ActivityStartTime = hour(ActivityDueDate) & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									End If
									
									'**********************************************************************************************
									'Determine if the activity is an appointment or meeting, to see which duration to add
									' to the start time, to arrive at the end time
									''**********************************************************************************************
									
									If cInt(ActivityIsAppointment) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityAppointmentDuration,ActivityDueDate)
									ElseIf cInt(ActivityIsMeeting) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityMeetingDuration,ActivityDueDate)
									Else
										ActivityDueDateEndTime = DateAdd("n",0,ActivityDueDate)
									End If
									
									'**********************************************************************************************
									'Obtain the end time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									
									UnformattedActivityEndTime = timevalue(hour(ActivityDueDateEndTime) & ":" & minute(ActivityDueDateEndTime))
									
									If minute(ActivityDueDateEndTime) < 10 AND minute(ActivityDueDateEndTime) > 0 Then
										minuteActivityDueDateEndTime = "0" & minute(ActivityDueDateEndTime)
									ElseIf minute(ActivityDueDateEndTime) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDateEndTime = minute(ActivityDueDateEndTime)
									End If
									
									If hour(ActivityDueDateEndTime) > 12 Then
										ActivityEndTime = hour(ActivityDueDateEndTime) - 12  & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									Else
										ActivityEndTime = hour(ActivityDueDateEndTime) & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									End If
									
									QuickLinkURLDestination = "viewProspect-" & ActivityProspectRecID
									QuickLoginURL = baseURL & "ql.asp?c=" & ClientKey & "&u=" & ownerUserNo & "&d=" & QuickLinkURLDestination

									
									%>
									<% If cInt(ActivityIsAppointment) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %>&nbsp;-&nbsp;<%= ActivityEndTime %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span>.
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>
								    	</li>
								    <% ElseIf cInt(ActivityIsMeeting) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %>&nbsp;-&nbsp;<%= ActivityEndTime %></strong>&nbsp;
								    		<%= ActivityDesc %><br>
								    		<strong>Location</strong>:&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
									<% Else %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;
								    		<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
								    <% End If %>
													    
								<%
								rsWeeklyAgendaReportSingleDay.MoveNext
								Loop
								
								%></ul><%
								
							End If

							%>					
							</td>

						</tr>
					</tbody>
				</table>
		
				<table class="agenda">	
					<thead>
						<tr>
							<th><%=fridayOfThisWeekTextName %></th>
							<th class="blue">Notes:</th>
						</tr>
					</thead>
					<tbody>
						<tr>
							<!-------------------------------------------------------------------------------------------------------------------------------------->
							<!-- FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY FRIDAY 
							<!-------------------------------------------------------------------------------------------------------------------------------------->
							
							<td>
							<%
							
							rsWeeklyAgendaReportSingleDay.close
							cnnWeeklyAgendaReportSingleDay.close
							set rsWeeklyAgendaReportSingleDay=nothing
							set cnnWeeklyAgendaReportSingleDay=nothing
							
							Set cnnWeeklyAgendaReportSingleDay = Server.CreateObject("ADODB.Connection")
							cnnWeeklyAgendaReportSingleDay.open(Session("ClientCnnString"))
							Set rsWeeklyAgendaReportSingleDay  = Server.CreateObject("ADODB.Recordset")
							rsWeeklyAgendaReportSingleDay.CursorLocation = 3 
							
							SQLWeeklyAgendaReportSingleDay = "SELECT * FROM PR_ProspectActivities "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " INNER JOIN PR_Prospects ON "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_Prospects.InternalRecordIdentifier = PR_ProspectActivities.ProspectRecID "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " WHERE "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_Prospects.OwnerUserNo = " & ownerUserNo & ") AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_Prospects.Pool = 'Live') AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (PR_ProspectActivities.Status <> 'Completed' OR "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_ProspectActivities.Status <> 'Cancelled' OR "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " PR_ProspectActivities.Status IS NULL) AND "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " (Cast(PR_ProspectActivities.ActivityDueDate as Date) = '" & fridayOfThisWeek & "') "
							SQLWeeklyAgendaReportSingleDay = SQLWeeklyAgendaReportSingleDay & " ORDER BY PR_ProspectActivities.ActivityDueDate ASC "
	
							rsWeeklyAgendaReportSingleDay.Open SQLWeeklyAgendaReportSingleDay, cnnWeeklyAgendaReportSingleDay
		
							If Not rsWeeklyAgendaReportSingleDay.EOF Then 
							
								%><ul><%
								
								Do While NOT rsWeeklyAgendaReportSingleDay.EOF
								
									ActivityDueDate = rsWeeklyAgendaReportSingleDay("ActivityDueDate")
																		
									ActivityProspectRecID = rsWeeklyAgendaReportSingleDay("ProspectRecID")
									ProspectName = GetProspectNameByNumber(ActivityProspectRecID)
									ProspectLocationStreet = GetProspectStreetByNumber(ActivityProspectRecID)
									ProspectLocationCity = GetProspectCityByNumber(ActivityProspectRecID)
									ProspectLocationState = GetProspectStateByNumber(ActivityProspectRecID)
									ProspectLocationPostalCode = GetProspectPostalCodeByNumber(ActivityProspectRecID)
									
									ActivityRecID = rsWeeklyAgendaReportSingleDay("ActivityRecID")
									ActivityDesc = GetActivityByNum(ActivityRecID)
									ActivityNotes = rsWeeklyAgendaReportSingleDay("Notes")
									
									ActivityIsAppointment = rsWeeklyAgendaReportSingleDay("ActivityIsAppointment")
									ActivityAppointmentDuration = rsWeeklyAgendaReportSingleDay("ActivityAppointmentDuration")
									
									ActivityIsMeeting = rsWeeklyAgendaReportSingleDay("ActivityIsMeeting")
									ActivityMeetingDuration = rsWeeklyAgendaReportSingleDay("ActivityMeetingDuration")
									ActivityMeetingLocation = rsWeeklyAgendaReportSingleDay("ActivityMeetingLocation")
									
									If ActivityIsAppointment = "" OR IsNull(ActivityIsAppointment) Then
										ActivityIsAppointment = 0
									End If
									
									If ActivityIsMeeting = "" OR IsNull(ActivityIsMeeting) Then
										ActivityIsMeeting = 0
									End If

									'**********************************************************************************************
									'Obtain the start time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									UnformattedActivityStartTime = timevalue(hour(ActivityDueDate) & ":" & minute(ActivityDueDate))
									
									If minute(ActivityDueDate) < 10 AND minute(ActivityDueDate) > 0 Then
										minuteActivityDueDate = "0" & minute(ActivityDueDate)
									ElseIf minute(ActivityDueDate) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDate = minute(ActivityDueDate)
									End If
									
									If hour(ActivityDueDate) > 12 Then
										ActivityStartTime = hour(ActivityDueDate) - 12  & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									Else
										ActivityStartTime = hour(ActivityDueDate) & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									End If
									
									'**********************************************************************************************
									'Determine if the activity is an appointment or meeting, to see which duration to add
									' to the start time, to arrive at the end time
									''**********************************************************************************************
									
									If cInt(ActivityIsAppointment) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityAppointmentDuration,ActivityDueDate)
									ElseIf cInt(ActivityIsMeeting) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityMeetingDuration,ActivityDueDate)
									Else
										ActivityDueDateEndTime = DateAdd("n",0,ActivityDueDate)
									End If
									
									'**********************************************************************************************
									'Obtain the end time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									
									UnformattedActivityEndTime = timevalue(hour(ActivityDueDateEndTime) & ":" & minute(ActivityDueDateEndTime))
									
									If minute(ActivityDueDateEndTime) < 10 AND minute(ActivityDueDateEndTime) > 0 Then
										minuteActivityDueDateEndTime = "0" & minute(ActivityDueDateEndTime)
									ElseIf minute(ActivityDueDateEndTime) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDateEndTime = minute(ActivityDueDateEndTime)
									End If
									
									If hour(ActivityDueDateEndTime) > 12 Then
										ActivityEndTime = hour(ActivityDueDateEndTime) - 12  & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									Else
										ActivityEndTime = hour(ActivityDueDateEndTime) & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									End If
									
									QuickLinkURLDestination = "viewProspect-" & ActivityProspectRecID
									QuickLoginURL = baseURL & "ql.asp?c=" & ClientKey & "&u=" & ownerUserNo & "&d=" & QuickLinkURLDestination

									
									%>
									<% If cInt(ActivityIsAppointment) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %>&nbsp;-&nbsp;<%= ActivityEndTime %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span>.
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>
								    	</li>
								    <% ElseIf cInt(ActivityIsMeeting) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %>&nbsp;-&nbsp;<%= ActivityEndTime %></strong>&nbsp;
								    		<%= ActivityDesc %><br>
								    		<strong>Location</strong>:&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
									<% Else %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= ActivityStartTime %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;
								    		<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
								    <% End If %>
													    
								<%
								rsWeeklyAgendaReportSingleDay.MoveNext
								Loop
								
								%></ul><%
								
							End If

							rsWeeklyAgendaReportSingleDay.close
							cnnWeeklyAgendaReportSingleDay.close
							set rsWeeklyAgendaReportSingleDay=nothing
							set cnnWeeklyAgendaReportSingleDay=nothing

							%>					
							</td>
							<td>&nbsp;</td>
						</tr>
					</tbody>
				</table> 
		
		<% Else %>
		
				<table class="agenda">
					<thead>
						<tr>
							<th><%= mondayOfThisWeekTextName %></th>
							<th><%= tuesdayOfThisWeekTextName %></th>
						</tr>
					</thead>
					<tbody>
						<tr>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
						</tr>
					</tbody>
				</table>
				
				<table class="agenda">
					<thead>
						<tr>
							<th><%= wednesdayOfThisWeekTextName %></th>
							<th><%= thursdayOfThisWeekTextName %></th>
						</tr>
					</thead>
					<tbody>
						<tr>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
						</tr>
					</tbody>
				</table>
		
				<table class="agenda">	
					<thead>
						<tr>
							<th><%=fridayOfThisWeekTextName %></th>
							<th class="blue">Notes:</th>
						</tr>
					</thead>
					<tbody>
						<tr>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
						</tr>
					</tbody>
				</table> 
			
		
		<% End If %>
		
		
  		<%
  		
		SQLWeeklyAgendaReportExpired = "SELECT * FROM PR_ProspectActivities INNER JOIN PR_Prospects ON PR_Prospects.InternalRecordIdentifier = PR_ProspectActivities.ProspectRecID "
		SQLWeeklyAgendaReportExpired = SQLWeeklyAgendaReportExpired & " WHERE "
		SQLWeeklyAgendaReportExpired = SQLWeeklyAgendaReportExpired & " (PR_Prospects.OwnerUserNo = " & ownerUserNo & ") AND "
		SQLWeeklyAgendaReportExpired = SQLWeeklyAgendaReportExpired & " (PR_Prospects.Pool = 'Live') AND "
		SQLWeeklyAgendaReportExpired = SQLWeeklyAgendaReportExpired & " (PR_ProspectActivities.Status IS NULL) AND "
		SQLWeeklyAgendaReportExpired = SQLWeeklyAgendaReportExpired & " (Cast(PR_ProspectActivities.ActivityDueDate as Date) < '" & mondayOfThisWeek & "') "
		SQLWeeklyAgendaReportExpired = SQLWeeklyAgendaReportExpired & " ORDER BY PR_ProspectActivities.ActivityDueDate DESC "
	
		Set cnnWeeklyAgendaReportExpired = Server.CreateObject("ADODB.Connection")
		cnnWeeklyAgendaReportExpired.open(Session("ClientCnnString"))
		Set rsWeeklyAgendaReportExpired  = Server.CreateObject("ADODB.Recordset")
		rsWeeklyAgendaReportExpired.CursorLocation = 3 
		rsWeeklyAgendaReportExpired.Open SQLWeeklyAgendaReportExpired, cnnWeeklyAgendaReportExpired 
					
		'Response.Write(SQLWeeklyAgendaReportExpired & "<br>")
		
		If Not rsWeeklyAgendaReportExpired.EOF Then
		%>
		
			<BR style="page-break-before: always">
		  		
	  		<h2 style="margin-bottom:30px"><img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png" style="height:55px;"></h2>
	  		
			<table class="agenda-expired">
				<thead>
					<tr>
						<th>Expired Activity Dates</th>
					</tr>
				</thead>
				<tbody>
					<tr>
						<td>
							<% If Not rsWeeklyAgendaReportExpired.EOF Then 
							
								%><ul><%
								
								Do While NOT rsWeeklyAgendaReportExpired.EOF
								
									ActivityDueDate = rsWeeklyAgendaReportExpired("ActivityDueDate")
									
									FormattedActivityDueDate = Month(ActivityDueDate) & "/" &  Day(ActivityDueDate) & "/" &  Right(Year(ActivityDueDate),2)
																		
									ActivityProspectRecID = rsWeeklyAgendaReportExpired("ProspectRecID")
									ProspectName = GetProspectNameByNumber(ActivityProspectRecID)
									ProspectLocationStreet = GetProspectStreetByNumber(ActivityProspectRecID)
									ProspectLocationCity = GetProspectCityByNumber(ActivityProspectRecID)
									ProspectLocationState = GetProspectStateByNumber(ActivityProspectRecID)
									ProspectLocationPostalCode = GetProspectPostalCodeByNumber(ActivityProspectRecID)
									
									ActivityRecID = rsWeeklyAgendaReportExpired("ActivityRecID")
									ActivityDesc = GetActivityByNum(ActivityRecID)
									ActivityNotes = rsWeeklyAgendaReportExpired("Notes")
									
									ActivityIsAppointment = rsWeeklyAgendaReportExpired("ActivityIsAppointment")
									ActivityAppointmentDuration = rsWeeklyAgendaReportExpired("ActivityAppointmentDuration")
									
									ActivityIsMeeting = rsWeeklyAgendaReportExpired("ActivityIsMeeting")
									ActivityMeetingDuration = rsWeeklyAgendaReportExpired("ActivityMeetingDuration")
									ActivityMeetingLocation = rsWeeklyAgendaReportExpired("ActivityMeetingLocation")
									
									If ActivityIsAppointment = "" OR IsNull(ActivityIsAppointment) Then
										ActivityIsAppointment = 0
									End If
									
									If ActivityIsMeeting = "" OR IsNull(ActivityIsMeeting) Then
										ActivityIsMeeting = 0
									End If
						
									
									'**********************************************************************************************
									'Obtain the start time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									UnformattedActivityStartTime = timevalue(hour(ActivityDueDate) & ":" & minute(ActivityDueDate))
									
									If minute(ActivityDueDate) < 10 AND minute(ActivityDueDate) > 0 Then
										minuteActivityDueDate = "0" & minute(ActivityDueDate)
									ElseIf minute(ActivityDueDate) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDate = minute(ActivityDueDate)
									End If
									
									If hour(ActivityDueDate) > 12 Then
										ActivityStartTime = hour(ActivityDueDate) - 12  & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									Else
										ActivityStartTime = hour(ActivityDueDate) & ":" & minuteActivityDueDate & " " & right(UnformattedActivityStartTime, 2)
									End If
									
									'**********************************************************************************************
									'Determine if the activity is an appointment or meeting, to see which duration to add
									' to the start time, to arrive at the end time
									''**********************************************************************************************
									
									If cInt(ActivityIsAppointment) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityAppointmentDuration,ActivityDueDate)
									ElseIf cInt(ActivityIsMeeting) = 1 Then
										ActivityDueDateEndTime = DateAdd("n",ActivityMeetingDuration,ActivityDueDate)
									Else
										ActivityDueDateEndTime = DateAdd("n",0,ActivityDueDate)
									End If
									
									'**********************************************************************************************
									'Obtain the end time of the activity and format to 00:00 AM/PM
									''**********************************************************************************************
									
									UnformattedActivityEndTime = timevalue(hour(ActivityDueDateEndTime) & ":" & minute(ActivityDueDateEndTime))
									
									If minute(ActivityDueDateEndTime) < 10 AND minute(ActivityDueDateEndTime) > 0 Then
										minuteActivityDueDateEndTime = "0" & minute(ActivityDueDateEndTime)
									ElseIf minute(ActivityDueDateEndTime) = 0 Then
										minuteActivityDueDate = "00"
									Else
										minuteActivityDueDateEndTime = minute(ActivityDueDateEndTime)
									End If
									
									If hour(ActivityDueDateEndTime) > 12 Then
										ActivityEndTime = hour(ActivityDueDateEndTime) - 12  & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									Else
										ActivityEndTime = hour(ActivityDueDateEndTime) & ":" & minuteActivityDueDateEndTime & " " & right(UnformattedActivityEndTime, 2)
									End If
									
									
									
									QuickLinkURLDestination = "viewProspect-" & ActivityProspectRecID
									QuickLoginURL = baseURL & "ql.asp?c=" & ClientKey & "&u=" & ownerUserNo & "&d=" & QuickLinkURLDestination

									
									%>
									<% If cInt(ActivityIsAppointment) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= FormattedActivityDueDate %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span>.
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>
								    	</li>
								    <% ElseIf cInt(ActivityIsMeeting) = 1 Then %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= FormattedActivityDueDate %></strong>&nbsp;
								    		<%= ActivityDesc %><br>
								    		<strong>Location</strong>:&nbsp;<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
									<% Else %>
								    	<li>
								    		<input type="checkbox" disabled="true">&nbsp;
								    		<strong><%= FormattedActivityDueDate %></strong>&nbsp;
								    		<%= ActivityDesc %> with&nbsp;
								    		<span class="company"><a href="<%= QuickLoginURL %>"><%= ProspectName %></a></span> (<%= ProspectLocationStreet %> <%= ProspectLocationCity %>, <%= ProspectLocationState %> <%= ProspectLocationPostalCode %>)
								    		<% If ActivityNotes <> "" Then %>
								    			<br><strong>Notes</strong>: <%= ActivityNotes %>
								    		<% End If %>								    		
								    	</li>
								    <% End If %>
													    
								<%
								rsWeeklyAgendaReportExpired.MoveNext
								Loop
								
								%></ul><%
								
							End If %>

					
						</td>
					</tr>
				</tbody>
			</table>
			
		
		
		<% End If %>
		


	
</body>
</html>