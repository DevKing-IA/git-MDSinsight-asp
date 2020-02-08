<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->
<%

SelectedMemoNumber = Request.Form("txtTicketNumber")

If SelectedMemoNumber = "" Then
	SelectedMemoNumber = Request.QueryString("t")
End If

%>

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

	.input-lg::-webkit-input-placeholder, textarea::-webkit-input-placeholder {
	  color: #666;
	}
	.input-lg:-moz-placeholder, textarea:-moz-placeholder {
	  color: #666;
	}
	
	.checkboxes label{
		font-weight: normal;
		margin-right: 20px;
	}

	.badge-pill-icon-letter {
	    padding-right: .3em;
	    padding-left: .3em;
	    border-radius: 8rem;
	}
	
	.list-group{
		margin:5px;
	}

</style>

<h1 class="fieldservice-heading" ><a class="btn-home" href="main_OpenTickets.asp" role="button"><i class="fa fa-arrow-left"></i></a>Actions</h1>


<% NumberOfMinutesInServiceDayVar = GetNumberOfMinutesInServiceDay() %>

<div class="list-group">

	<%

	MemoNumber = SelectedMemoNumber
	custNum = GetServiceTicketCust(SelectedMemoNumber)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rsCust = Server.CreateObject("ADODB.Recordset")
	rsCust.CursorLocation = 3 

	SQL = "SELECT Name,Addr1,Addr2,City,CityStateZip,Phone,Contact FROM AR_Customer WHERE Custnum = '" & custNum & "'"
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
				elapsedMinutes = ServiceCallElapsedMinutesOpenTicket(MemoNumber)
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
		
		</h6>
		
		<!--<p class="mb-1"><%= GetTerm("Account") %>&nbsp;<%= custNum %></p>-->
		<small class="mb-2 d-block"><%= custContact %>&nbsp;<%= custPhone %></small>
		
		
	</span>

</div>


<!-- buttons start here !-->
<div class="container-fluid">

	<a href="CloseService_PassThru.asp?t=<%= MemoNumber %>">
		<button type="button" class="btn btn-primary btn-block mt-2">Close</button>
	</a>
	
	<a href="waitForParts_PassThru.asp?t=<%= MemoNumber %>">
		<button type="button" class="btn btn-primary btn-block mt-2">Wait For Parts</button>
	</a>

	<a href="followup_PassThru.asp?t=<%= MemoNumber %>">
		<button type="button" class="btn btn-primary btn-block mt-2">Follow Up</button>
	</a>

	<a href="swap_PassThru.asp?t=<%= MemoNumber %>">
		<button type="button" class="btn btn-primary btn-block mt-2">Swap</button>
	</a>

	<a href="unableToWork_PassThru.asp?t=<%= MemoNumber %>">
		<button type="button" class="btn btn-primary btn-block mt-2">Unable To Work</button>
	</a>
		
</div>
<!-- buttons end here !-->



<%
If FS_TechCanDecline() <> True Then 
	'A sneaky thing to do, if you tap to view the ticket, but you haven't ACKed yet
	'We ACK it for you. After all, you did just look at the details
	
	'See if there is a dispatch ACKed record & if not, ACK it
	SQL = "Select * from FS_ServiceMemosDetail Where MemoNumber = '" & SelectedMemoNumber  & "' AND MemoStage='Dispatch Acknowledged'"
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs = cnn8.Execute(SQL)
	If rs.Eof then ' not found, so ACK it
		SQL = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
		SQL = SQL & " SubmissionDateTime, USerNoSubmittingRecord,UserNoOfServiceTech,OriginalDispatchDateTime)"
		SQL = SQL &  " VALUES (" 
		SQL = SQL & "'"  & SelectedMemoNumber & "'"
		SQL = SQL & ",'"  & SelectedCustomer & "'"
		SQL = SQL & ",'Dispatch Acknowledged'"
		SQL = SQL & ",getdate() "
		SQL = SQL & ","  & Session("UserNo") 
		SQL = SQL & ","  & Session("UserNo")
		SQL = SQL & ", '" & TicketOriginalDispatchDateTime(SelectedMemoNumber) & "')"
		response.write(SQL)
		Set rs = cnn8.Execute(SQL)
	
		
		'Write audit trail for dispatch
		'*******************************
		Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " acknowldged dispatch notification for service ticket number " & ServiceTicketNumber & " at " & NOW()
		CreateAuditLogEntry "Service Ticket System","Dispatch Acknowledged","Minor",0,Description 
		
		Set rs = Nothing
		cnn8.Close
		Set cnn8 = Nothing
	End If
End If
%><!--#include file="../inc/footer-field-service-noTimeout.asp"-->