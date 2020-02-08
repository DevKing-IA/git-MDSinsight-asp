<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->
<!--#include file="../css/fa_animation_styles.css"-->

<script type="text/javascript">
    $(document).ready(function() {
    	$(".segment-select").Segment();
    });
</script>

<style type="text/css">
	.fieldservice-heading{
		line-height: 1;
		padding-top: 5px;
		padding-bottom: 5px;
	}
	
	.btn-home{
		color: #fff;
		margin-top: -2px;
		margin-left: 5px;
		float: left;
 	}
	
 ul{
	 color: #666;
	 font-size: 13px;
	 text-transform: uppercase;
	 list-style-type: none;
	     -webkit-margin-before: 0px;
    -webkit-margin-after: 0px;
    -webkit-margin-start: 0px;
    -webkit-margin-end: 0px;
    -webkit-padding-start: 0px;
 }
 
 .enroute{
	 color: green;
 }
 
 .btn-spacing{
	 margin-bottom: 40px;
 }
 
 
 
.btn-block {
    width: auto;
    display: inline-block;
    margin-right:2em;
}
 
 .row{
 	/* flex-wrap: nowrap !important; */
 }

 
 @media (max-width: 767px) {
 	.mob-col{
 		/* width: auto !important;  */
 	}
 }
 
 .driver-menu{
 	text-align:center;
 	margin-bottom:10px;
 	margin-top:10px;
 }
 
.badge-pill-icon-letter {
    padding-right: .3em;
    padding-left: .3em;
    border-radius: 8rem;
}	

 .fa-stack  { font-size: 0.7em; }
  i { vertical-align: middle; }
	
</style>       


 
<h1 class="fieldservice-heading"><a class="btn-home" href="main_menu.asp" role="button"><i class="fa fa-bars"></i></a> Your Stops</h1>

<div class="driver-menu">		
	  <select class="segment-select" id="driverTicketMenu">
	      <option value="main_UnacknowledgedTickets.asp">UNACK (<%=NumberOfServiceTicketsAwaitingACKFromTech(Session("UserNo"))%>)</option>
	      <option value="main_OpenTickets.asp" selected="selected">OPEN (<%=NumberOfServiceTicketsAcknowledgedByTech(Session("UserNo"))%>)</option>
	      <option value="main_ClosedRedoTickets.asp">CLOSED/REDO (<%=NumberOfServiceTicketsClosedOrRedoByTech(Session("UserNo"))%>)</option>
	  </select>
</div><!-- driver-menu-->


<div class="container-fluid">


<%
'Now lookup the other info

SQL = "SELECT * FROM "
SQL = SQL & "FS_ServiceMemosDetail INNER JOIN FS_ServiceMemos ON FS_ServiceMemos.MemoNumber = FS_ServiceMemosDetail.MemoNumber AND FS_ServiceMemos.RecordSubType = 'OPEN' WHERE serviceDetailRecNumber IN "
SQL = SQL & "(SELECT MAX(serviceDetailRecNumber) FROM FS_ServiceMemosDetail WHERE UserNoOfServiceTech = " & Session("UserNo") & " GROUP BY memonumber ) "
SQL = SQL & " AND ClosedorCancelled <> 1 AND MemoStage <> 'Dispatched'"
SQL = SQL & "ORDER BY FS_ServiceMemos.Urgent Desc, OriginalDispatchDateTime"
 
'Response.Write(SQL)

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rsCust = Server.CreateObject("ADODB.Recordset")
rsCust.CursorLocation = 3 


Set rs = cnn8.Execute(SQL)
			
	If not rs.EOF Then
	
		NumberOfMinutesInServiceDayVar = GetNumberOfMinutesInServiceDay()
	
		%><div class="list-group"><%
		
		Do While Not rs.EOF
		
			If LastTechUserNo(rs("MemoNumber")) = Session("UserNo") Then ' If we are not the latest tech, it was reassigned & isnt ours anymore
			
				If AwaitingRedispatch(rs("MemoNumber")) <> True Then
				
					MemoNumber = rs("MemoNumber")
					custNum = rs("Custnum")

					SQL = "SELECT Name,Addr1,Addr2,City,CityStateZip,Phone,Contact FROM AR_Customer WHERE Custnum = '" & rs("Custnum") & "'"
					Set rsCust = cnn8.Execute(SQL)
					If NOT rsCust.EOF Then
						custName = rsCust("Name")
						custAddr1 = rsCust("Addr1")
						custAddr2 = rsCust("Addr2")
						custCity = rsCust("City")
						custCityStateZip = rsCust("CityStateZip")
						custPhone = rsCust("Phone")
						custContact = rsCust("Contact")
					End If 
					%>
								
					<span class="list-group-item list-group-item-action flex-column align-items-start">
						<div class="d-flex w-100 justify-content-between">
							<h6 class="mb-1 font-weight-bold" style="font-size:1.1em;"><%= custName %></h6>
							
							<%
								elapsedMinutes = ServiceCallElapsedMinutesOpenTicket(rs.Fields("MemoNumber"))
								minutesInServiceDay = NumberOfMinutesInServiceDayVar
								
								If elapsedMinutes < 1 Then elapsedMinutes = 1 ' If it has been less than 1 minute, just show 1 anyway
								elapsedMinutesForSorting = elapsedMinutes
								elapsedString = ""
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

							%>							
							<small><%= elapsedString %></small>

						</div>
													
						<small><%= custAddr1 %>&nbsp;<%= custAddr2 %>&nbsp;<%= custCity %></small>

						<h6 class="mb-1 mt-1">Ticket #<%= MemoNumber %>

								<% If TicketIsUrgent(MemoNumber) Then %>
									<span class="badge badge-danger badge-pill"><i class="fas fa-exclamation"></i></span>
								<% End If %>
								
								<% If filterChangeModuleOn() = True Then %>
									<% If TicketIsFilterChange(MemoNumber) Then %>
										<span class="badge badge-warning badge-pill">F</span>
									<% Else %>
										<span class="badge badge-info badge-pill badge-pill-icon-letter"><i class="fas fa-cog"></i></span>
									<% End If %>
								<% Else %>
									<span class="badge badge-info badge-pill badge-pill-icon-letter"><i class="fas fa-cog"></i></span>
								<% End If %>
								
								
								<% If CustHasServiceTicketNotes(MemoNumber) = True Then %>
									<% If NoteNewServiceTicketForUser(MemoNumber) = True Then %>
										<!-- Unread Note Icon-->
										<a href="viewServiceMemoNotes.asp?t=<%= MemoNumber %>&tab=open">
											<span class="fa-stack faa-pulse animated" style="vertical-align: top;">
												<i class="fas fa-sticky-note fa-stack-2x" style="color:#dc3545;"></i>
												<i class="fas fa-envelope fa-stack-1x fa-inverse"></i>
											</span>	
										</a>										
									<% Else %>
										<!-- Edit Existing Note Icon-->
										<a href="viewServiceMemoNotes.asp?t=<%= MemoNumber %>&tab=open">
											<span class="fa-stack" style="vertical-align: top;">
												<i class="fas fa-sticky-note fa-stack-2x" style="color:#28a745;"></i>
												<i class="fas fa-envelope fa-stack-1x fa-inverse"></i>
											</span>	
										</a>											
									<% End If %>								
								<% Else %>
									<!-- Add New Note Icon -->
									<a href="addServiceMemoNote.asp?t=<%= MemoNumber %>&tab=open">
										<span class="fa-stack" style="vertical-align: top;">
											<i class="fas fa-sticky-note fa-stack-2x" style="color:#28a745;"></i>
											<i class="fas fa-plus fa-stack-1x fa-inverse"></i>
										</span>	
									</a>								
								<% End If %>								
						
						</h6>
						
						<!--<p class="mb-1"><%= GetTerm("Account") %>&nbsp;<%= custNum %></p>-->
						<small class="mb-2 d-block"><%= custContact %>&nbsp;<%= custPhone %></small>
						
						<a href="viewTicket.asp?t=<%= MemoNumber %>&c=<%= custNum %>&u=<%= Session("Userno") %>">		 
							<button type="button" class="btn btn-primary btn-sm">Details</button>
						</a>

						<% Select Case GetServiceTicketCurrentStage(MemoNumber)
						
							Case "Dispatched"
							
								If FS_TechCanDecline() Then %>
									<a href="tap_decline.asp?t=<%= MemoNumber %>&c=<%= custNum %>&u=<%= Session("Userno") %>">
										<button type="button" class="btn btn-danger btn-sm">Decline</button>
									</a>
								<% End If %> 
								
								<a href="tap_ack.asp?t=<%= MemoNumber %>&c=<%= custNum %>&u=<%= Session("Userno") %>">
									<button type="button" class="btn btn-warning btn-sm">Tap To Acknowledge</button>
								</a>
								
							<% Case "Dispatch Acknowledged" %>
							
								<a href="tap_enroute.asp?t=<%= MemoNumber %>&c=<%= custNum %>&u=<%= Session("Userno") %>">
									<button type="button" class="btn btn-primary btn-sm">Tap When En Route</button>
								</a>
								
							<% Case "En Route" %>
							
								<a href="tap_onsite.asp?t=<%= MemoNumber %>&c=<%= custNum %>&u=<%= Session("Userno") %>">
									<button type="button" class="btn btn-primary btn-sm">Tap When On Site</button>
								</a>
								
							<% Case Else %>
							
								<a href="onSite.asp?t=<%= MemoNumber %>">
									<button type="button" class="btn btn-primary btn-sm">Tap To Complete This Stop</button>
								</a>
								
						<% End Select %>
						
					</span>

		
					<%
					End If
				End If
			rs.movenext
		loop
		%></div><%
	Else
		%>You Have No Open Service Calls.<%
	End IF

cnn8.close
Set rsCust = Nothing
Set rs = Nothing
Set cnn8 = Nothing				
%></div><!--#include file="../inc/footer-field-service-noTimeout.asp"-->