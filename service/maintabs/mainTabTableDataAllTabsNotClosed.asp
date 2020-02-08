<%	
	SQLCustInfo = "SELECT * FROM AR_Customer WHERE CustNum = '" & rs.Fields("AccountNumber") & "'"

	Set rsCustInfo = cnnCustInfo.Execute(SQLCustInfo)
		
	If Not rsCustInfo.EOF Then CustTypeVar = rsCustInfo("CustType") Else CustType=""
	
	If ServiceColorsOn AND (CustTypeVar = "1" or CustTypeVar = "2" or CustTypeVar = "3") Then 
			If Isnull(rs.Fields("AlertEmailSent")) Then
				If rs.Fields("Urgent") = 1 Then Response.Write("<tr class='urgent-priority'>") Else Response.Write("<tr class='high-priority'>")
			Else
				If rs.Fields("Urgent") = 1 Then Response.Write("<tr class='urgent-priority'>") Else Response.Write("<tr class='alert-high-priority'>")
			End If
	Else 
		'Not high priority but see if an alert was ever sent
		If Isnull(rs.Fields("AlertEmailSent")) or ServiceColorsOn <> 1 Then
			If LineX Mod 2 = 0 then
				'THESE ARE EVEN LINES
				If rs.Fields("Urgent") = 1 Then Response.Write("<tr class='urgent-priority'>") Else Response.Write("<tr class='tr-even'>")
			Else
				'THESE ARE ODD LINE
				If rs.Fields("Urgent") = 1 Then Response.Write("<tr class='urgent-priority'>") Else Response.Write("<tr class='tr-odd'>")
			End If
		Else
			If rs.Fields("Urgent") = 1 Then Response.Write("<tr class='urgent-priority'>") Else Response.Write("<tr class='alert-priority'>")
		End If
	 End If	
	 
	submissionHour = Hour(rs("submissionDateTime"))
	submissionMinute = Minute(rs("submissionDateTime"))
	submissionZeroFactor = "0" & submissionMinute
	submissionAMPM = "AM"
	If submissionHour >= 12 then submissionAMPM = "PM"
	If submissionHour > 12 then submissionHour = submissionHour - 12
	If submissionMinute <= 9 then submissionMinute = submissionZeroFactor	
	
	submissionDateTime = rs("submissionDateTime")
	
	ticketDateDisplay = padDate(MONTH(submissionDateTime),2) & "/" & padDate(DAY(submissionDateTime),2) & "/" & padDate(RIGHT(YEAR(submissionDateTime),2),2)

	Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(rs("submissionDateTime")) & ">" & ticketDateDisplay & "<br>" & submissionHour & ":" & submissionMinute & " " & submissionAMPM & "</td>")%>
	
	<%If rs.Fields("CurrentStatus")="OPEN" Then %>
		<!--<td><a href='./editServiceMemo.asp?memo=<%= rs.Fields("MemoNumber")%>' target='_blank'><%= rs.Fields("MemoNumber")%></a></td>-->
		<td><%= rs.Fields("MemoNumber")%></td>	
	<% Else %>
		<td><%= rs.Fields("MemoNumber")%></td>
	<% End If %>
	<td>
	<%
		NumOpenCallsByAcctVar = NumOpenCallsByAcct(rs.Fields("AccountNumber"))
		If NumOpenCallsByAcctVar  > 1 Then
			Response.Write("<a href='oneCust.asp?cust=" & rs.Fields("AccountNumber") & "' target='_blank'><span class='bluecircle'>" & NumOpenCallsByAcctVar & "</span></a>")
		End If
	%>
	</td>
	<td>
	
	<%
	
	If MUV_READ("EQUIPMENTMODULEON")="Enabled" Then
		TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(rs.Fields("AccountNumber"))
	Else
		TotalEquipmentValue = 0 
	End If
		
	'LCPGP = TotalSalesByPeriodSeq(PeriodSeqBeingEvaluated) - TotalCostByPeriodSeq(PeriodSeqBeingEvaluated)
	LCPGP = 0 
	
	%>
	<%= rs.Fields("AccountNumber") %>

	
	</td>
	<td>
	<%= rs.Fields("Company") %><br>
	<% If NOT rsCustInfo.EOF Then %>
		<%= rsCustInfo("Addr1") %><br>
		<% If Trim(rsCustInfo("Addr2")) <> "" Then Response.Write(rsCustInfo("Addr2")& "<br>") %>
		<%= rsCustInfo("City") %>, <%= rsCustInfo("State") %>&nbsp;<%= rsCustInfo("Zip") %>
		<% If Trim(rsCustInfo("Addr2")) = "" Then Response.Write("<br><br>")
	End If%>
	</td>
	<td><span id="td-padding">
	<%
		If NOT IsNull(rs.Fields("ProblemCode")) Then
			Response.Write("SYMPTOM: " & rs.Fields("ProblemCode") & " - " & GetServiceTicketProblemCodeDescByIntRecID(rs.Fields("ProblemCode")))
		End If
		
		CompressLen = 27
		'See if there are linefeeds in there that need to come out
		If Instr(rs.Fields("ProblemDescription"),"<br>") <> 0 Then CompressLen = Instr(rs.Fields("ProblemDescription"),"<br>")
		If CompressLen > 27 Then CompressLen = 27
		'If len(rs.Fields("ProblemDescription")) > CompressLen Then Response.Write(Left(rs.Fields("ProblemDescription"),CompressLen)) Else 
		Response.Write("<br>")
		Response.Write(rs.Fields("ProblemDescription"))
		
	%>
	</span></td>
	<td sorttable_customkey="<%= rs.Fields("Dispatched") %>">
	<%
	'****************************
	' New Dispatch Code Goes Here
	'****************************
	'Only service managers or admins can do the dispatching
	GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rs.Fields("MemoNumber"))
	
	ticketStageDateTime = GetServiceTicketSTAGEDateTime(rs.Fields("MemoNumber"),GetServiceTicketCurrentStageVar)
	
	ticketStageHour = Hour(ticketStageDateTime)
	ticketStageMinute = Minute(ticketStageDateTime)
	ticketStageZeroFactor = "0" & ticketStageMinute
	ticketStageAMPM = "AM"
	If ticketStageHour >= 12 then ticketStageAMPM = "PM"
	If ticketStageHour > 12 then ticketStageHour = ticketStageHour - 12
	If ticketStageMinute <= 9 then ticketStageMinute = ticketStageZeroFactor	 

	ticketStageDateTimeDisplay = padDate(MONTH(ticketStageDateTime),2) & "/" & padDate(DAY(ticketStageDateTime),2) & "/" & padDate(RIGHT(YEAR(ticketStageDateTime),2),2)

	
	'*************************************************************
	'We need to collect more information about this ticket
	'to see what actions are available under the actions modal
	' and the dispatch modal
	'*************************************************************

	TicketNumber = rs.Fields("MemoNumber")
	UserNoOfServiceTech = GetServiceTicketDispatchedTech(TicketNumber)
	
	CustID = rs.Fields("AccountNumber")

	If len(rs.Fields("Company")) > 19 then 
		Cnam = left(rs.Fields("Company"),19) 
	Else 
		Cnam = rs.Fields("Company")
	End If
	
			
	If userIsServiceManager(Session("userNo")) or userIsAdmin(Session("userNo")) Then%>  
		<% If rs.Fields("RecordSubType") <> "HOLD" AND (GetServiceTicketCurrentStageVar = "Received" or GetServiceTicketCurrentStageVar = "Released") Then
			DynamicFormCounter = DynamicFormCounter  + 1
			
			Response.Write("<span class='labelAwaitingDispatch'>Awaiting Dispatch</span><br>")
			%>
			<% If userCanAccessServiceDispatchButton(Session("UserNo")) = true Then %>
				<span data-toggle="modal" data-target="#serviceBoardDispatchModal" data-service-ticket-number="<%= TicketNumber %>">
				    <button type="button" class="btn btn-primary btn-sm" data-tooltip data-placement="bottom" title="Dispatch Service Ticket" data-toggle="tooltip">Dispatch</button>
				</span>	
			<% End If %>
			
			<% If userCanAccessServiceCloseCancelButton(Session("UserNo")) = true Then %>
				<span data-toggle="modal" data-target="#modalEditExistingServiceTicketForClient" data-memo-number="<%= TicketNumber %>" data-cust-id="<%= CustID %>">
				    <button type="button" class="btn btn-danger btn-sm" data-tooltip data-placement="bottom" title="Close/Cancel Service Ticket" data-toggle="tooltip" style="margin-top:-4px">Close/Cancel</button>
				</span>	
			<% End If %>			
	
		<% Else 
			If AwaitingRedispatch(rs.Fields("MemoNumber")) = True Then
				
				If GetServiceTicketCurrentStageVar = "Awaiting Acknowledgement" Then
					Response.Write("<span class='labelAwaitingAcknowledgement'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "Dispatch Acknowledged" Then
					Response.Write("<span class='labelDispatchAcknowledged'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "En Route" Then
					Response.Write("<span class='labelEnRoute'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "On Site" Then
					Response.Write("<span class='labelOnSite'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "Swap" Then
					Response.Write("<span class='labelSwap'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "Wait for parts" Then
					Response.Write("<span class='labelWaitForParts'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "Follow Up" Then
					Response.Write("<span class='labelFollowUp'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "Unable To Work" Then
					Response.Write("<span class='labelUnableToWork'>" & GetServiceTicketCurrentStageVar  & "</span>")
				Else
					Response.Write("<span class='label-default'>" & GetServiceTicketCurrentStageVar  & "</span>")
				End If
				
				Response.Write("<br>" & GetServiceTicketSTAGEUser(rs.Fields("MemoNumber"),GetServiceTicketCurrentStageVar) & "<br>")
				
				
				'Response.Write(GetServiceTicketSTAGEDateTime(rs.Fields("MemoNumber"),GetServiceTicketCurrentStageVar))
				Response.Write(ticketStageDateTimeDisplay & " " & ticketStageHour & ":" & ticketStageMinute & " " & ticketStageAMPM)
				DynamicFormCounter = DynamicFormCounter  + 1%> 
				<!-- new dispatch button !-->
				<br>
				
				<% If userCanAccessServiceDispatchButton(Session("UserNo")) = true Then %>
					<span data-toggle="modal" data-target=".bs-example-modal-lg-customize<%=DynamicFormCounter%>" data-memo-number="<%= TicketNumber %>" data-cust-id="<%= CustID %>">
					    <button type="button" class="btn btn-primary btn-sm" data-tooltip data-placement="bottom" title="Redispatch Service Ticket" data-toggle="tooltip" style="margin-top:-4px">Redispatch</button>
					</span>
				<% End If %>		
					
				<% If userCanAccessServiceCloseCancelButton(Session("UserNo")) = true Then %>			
					<span data-toggle="modal" data-target="#modalEditExistingServiceTicketForClient" data-memo-number="<%= TicketNumber %>" data-cust-id="<%= CustID %>">
					    <button type="button" class="btn btn-danger btn-sm" data-tooltip data-placement="bottom" title="Close/Cancel Service Ticket" data-toggle="tooltip" style="margin-top:-4px">Close/Cancel</button>
					</span>		
				<% End If %>				
				
				<div class="modal fade bs-example-modal-lg-customize<%=DynamicFormCounter%>" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel" aria-hidden="true">
				<!--#include file="../dispatch_modal.asp"-->	
				<!-- eof new dispatch button !-->
			<% Else
			
				If GetServiceTicketCurrentStageVar = "Awaiting Acknowledgement" Then
					Response.Write("<span class='labelAwaitingAcknowledgement'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "Dispatch Acknowledged" Then
					Response.Write("<span class='labelDispatchAcknowledged'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "En Route" Then
					Response.Write("<span class='labelEnRoute'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "On Site" Then
					Response.Write("<span class='labelOnSite'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "Swap" Then
					Response.Write("<span class='labelSwap'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "Wait for parts" Then
					Response.Write("<span class='labelWaitForParts'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "Follow Up" Then
					Response.Write("<span class='labelFollowUp'>" & GetServiceTicketCurrentStageVar  & "</span>")
				ElseIf GetServiceTicketCurrentStageVar = "Unable To Work" Then
					Response.Write("<span class='labelUnableToWork'>" & GetServiceTicketCurrentStageVar  & "</span>")
				Else
					Response.Write("<span class='label-default'>" & GetServiceTicketCurrentStageVar  & "</span>")
				End If
				
				If GetServiceTicketCurrentStageVar = "Received" OR GetServiceTicketCurrentStageVar = "Released"_
					OR GetServiceTicketCurrentStageVar = "Awaiting Acknowledgement"_
					OR GetServiceTicketCurrentStageVar = "Dispatch Acknowledged"_
					OR GetServiceTicketCurrentStageVar = "En Route"_
					OR GetServiceTicketCurrentStageVar = "On Site"	Then %>	
					
					<% If userCanAccessServiceCloseCancelButton(Session("UserNo")) = true Then %>
						<br>
						<span data-toggle="modal" data-target="#modalEditExistingServiceTicketForClient" data-memo-number="<%= TicketNumber %>" data-cust-id="<%= CustID %>">
						    <button type="button" class="btn btn-danger btn-sm" data-tooltip data-placement="bottom" title="Close/Cancel Service Ticket" data-toggle="tooltip" >Close/Cancel</button>
						</span>		
					<% End If %>		
				<%
				End If
				
				'Response.Write(GetServiceTicketSTAGEUser(rs.Fields("MemoNumber"),GetServiceTicketCurrentStageVar ) & "<br>")
				Response.Write("<br>" & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(rs.Fields("MemoNumber"))) & "<br>")
				'Response.Write(GetServiceTicketSTAGEDateTime(rs.Fields("MemoNumber"),GetServiceTicketCurrentStageVar))
				Response.Write(ticketStageDateTimeDisplay & " " & ticketStageHour & ":" & ticketStageMinute & " " & ticketStageAMPM)
			End If
		 End If
		 
	 ELSE 'This is what CSRs see
	 
			If GetServiceTicketCurrentStageVar = "Awaiting Acknowledgement" Then
				Response.Write("<span class='labelAwaitingAcknowledgement'>" & GetServiceTicketCurrentStageVar  & "</span>")
			ElseIf GetServiceTicketCurrentStageVar = "Dispatch Acknowledged" Then
				Response.Write("<span class='labelDispatchAcknowledged'>" & GetServiceTicketCurrentStageVar  & "</span>")
			ElseIf GetServiceTicketCurrentStageVar = "En Route" Then
				Response.Write("<span class='labelEnRoute'>" & GetServiceTicketCurrentStageVar  & "</span>")
			ElseIf GetServiceTicketCurrentStageVar = "On Site" Then
				Response.Write("<span class='labelOnSite'>" & GetServiceTicketCurrentStageVar  & "</span>")
			ElseIf GetServiceTicketCurrentStageVar = "Swap" Then
				Response.Write("<span class='labelSwap'>" & GetServiceTicketCurrentStageVar  & "</span>")
			ElseIf GetServiceTicketCurrentStageVar = "Wait for parts" Then
				Response.Write("<span class='labelWaitForParts'>" & GetServiceTicketCurrentStageVar  & "</span>")
			ElseIf GetServiceTicketCurrentStageVar = "Follow Up" Then
				Response.Write("<span class='labelFollowUp'>" & GetServiceTicketCurrentStageVar  & "</span>")
			ElseIf GetServiceTicketCurrentStageVar = "Unable To Work" Then
				Response.Write("<span class='labelUnableToWork'>" & GetServiceTicketCurrentStageVar  & "</span>")
			Else
				Response.Write("<span class='label-default'>" & GetServiceTicketCurrentStageVar  & "</span>")
			End If

			If GetServiceTicketDispatchedTech(rs.Fields("MemoNumber")) <> "N/A" Then
				Response.Write("<br>" & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(rs.Fields("MemoNumber"))))
				Response.Write("<br>")
				'Response.Write(GetServiceTicketDispatchedDateTime(rs.Fields("MemoNumber")))
				Response.Write(ticketStageDateTimeDisplay & " " & ticketStageHour & ":" & ticketStageMinute & " " & ticketStageAMPM)
			End If
			
	End If
	'****************************
 	' New Dispatch Code Ends Here
	'****************************
	
	OpenedDate1 = FormatDateTime(GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),2)
	OpenedTime1 = FormatDateTime(GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),4)
	OpenedDateTime1 = GetServiceTicketOpenDateTime(rs.Fields("MemoNumber"))
	
	
	 %>

	<style>
	
		.notesContainer {
		  position: absolute; 
		  width: 600px;
		}	
		
		.spacer-top {
		    margin-top: 2px;
		}	
		
		.spacer-bottom {
		    margin-bottom: 30px;
		}	
		
		.btn-primary {
		    color: #fff;
		    background-color: #337ab7;
		    border-color: #2e6da4;
		    margin-bottom: 5px;
		}		
	</style>


	<%
	
	ServiceTicketNotes = GetLastServiceTicketNotesByTicket(rs.Fields("MemoNumber"))
	
	If Len(ServiceTicketNotes) > 103 Then ServiceTicketNotes = Left(ServiceTicketNotes,103)
	
	%>
	
	<% If Len(ServiceTicketNotes) > 0 Then %>
	
	
		<div class="spacer-top"></div>
		<span class="notesContainer">
		 	<%= ServiceTicketNotes %>
		</span>
		<div class="spacer-bottom"></div>
	
	<% End If %>

	</td>

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
	Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
	'''''''''''Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedMinutes & ":" & ServiceCallElapsedMinutesOpenTicket(rs.Fields("MemoNumber")) & "X<br>") 

	%>
	</td>
	
	
	<td>
		<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
			<button type="button" class="btn btn-success btn-sm" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= TicketNumber %>" data-customer-id="<%= CustID %>" data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>" data-target="#serviceBoardTicketOptionsModal" style="cursor:pointer;">Actions</button>
		<% End If %>
		
		<% If CustHasServiceTicketNotes(TicketNumber) = True Then %>
		
			<% If NoteNewServiceTicketForUser(TicketNumber) = True Then %>
				<!-- Pulsing note-->
				<span data-toggle="modal" data-target="#modalEditServiceTicketNotes" data-invoice-number="<%= TicketNumber %>" data-customer-id="<%= CustID %>" data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>">
				    <button type="button" class="btn btn-success btn-sm" data-tooltip data-placement="bottom" title="Service Ticket Notes" data-toggle="tooltip"><i class="fa fa-file-text-o faa-pulse animated fa-2x" aria-hidden="true"></i></button>
				</span>	
			<% Else %>
				<!-- Note Icon-->
				<span data-toggle="modal" data-target="#modalEditServiceTicketNotes" data-invoice-number="<%= TicketNumber %>" data-customer-id="<%= CustID %>" data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>">
				    <button type="button" class="btn btn-success btn-sm" data-tooltip data-placement="bottom" title="Service Ticket Notes" data-toggle="tooltip"><i class="fa fa-file-text-o" aria-hidden="true"></i></button>
				</span>					
			<% End If %>								
		<% Else %>
			<!-- Pencil Icon -->
			<span data-toggle="modal" data-target="#modalEditServiceTicketNotes" data-invoice-number="<%= TicketNumber %>" data-customer-id="<%= CustID %>" data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>">
			    <button type="button" class="btn btn-success btn-sm" data-tooltip data-placement="bottom" title="Service Ticket Notes" data-toggle="tooltip"><i class="fa fa-pencil" aria-hidden="true"></i></button>
			</span>				
		<% End If %>
		
		<span data-toggle="modal" data-target="#modalViewOpenClosedServiceTicketDetailsForClient" data-memo-number="<%= TicketNumber %>" data-cust-id="<%= CustID %>">
		    <button type="button" class="btn btn-success btn-sm" data-tooltip data-placement="bottom" title="View Service Ticket Details" data-toggle="tooltip"><i class="fas fa-eye"></i></button>
		</span>	
			
	    <% If TotalEquipmentValue <> 0 Then %>
			<span data-toggle="modal" data-target="#modalEquipmentVPC" data-lcp-gp="<%= LCPGP %>" data-invoice-number="<%= TicketNumber %>" data-cust-id="<%= CustID %>" data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>">
			    <button type="button" class="btn btn-success btn-sm" data-tooltip data-placement="bottom" title="View Customer Equipment" data-toggle="tooltip"><i class="fas fa-plug"></i></button>
			</span>			    	
	    <% End If %>

		<%
		If filterChangeModuleOn() Then
			If rs.Fields("RecordSubType") <> "HOLD" AND (GetServiceTicketCurrentStageVar = "Received" or GetServiceTicketCurrentStageVar = "Released") Then
				If rs.Fields("FilterChange") <> 1 Then
					If CustHasPendingFilterChange(CustID) = True Then
						DaysTilChange = datediff("d",Date(),FChange_NextDate(CustID))
						If DaysTilChange < 0 Then ' overdue	%>
							<br><span class="labelFilterChangeIndicatorAndButtonColorOverDue">Filter change<br>overdue <%=DaysTilChange%> days</span>
						<% Else %>
							<br><span class="labelFilterChangeIndicatorAndButtonColor">Filter change due<br>in <%=DaysTilChange%> days</span>
						<% End If		
					Else
						OpenTicks = ""
						OpenTicks = GetOpenFilterTicketsByCustID(CustID)
						'Need to split these up beacuse there may be one or more &
						'now we want to have links to the ticket details
						OpenTickArray = Split(OpenTicks,",")
						If Ubound(OpenTickArray) > 0 Then ' They have OPEN Filter tickets already %>
								<br><span class="labelFilterChangeIndicatorAndButtonColor">Filter on ticket<br>
								<% For x = 0 to Ubound(OpenTickArray)-1
									'Response.Write("<a href='./editServiceMemo.asp?memo=" & OpenTickArray(x) & "' target='_blank'>" & OpenTickArray(x) & "</a><br>")
									%><%= OpenTickArray(x) %><br><%
								Next 
								%></span>
						<%End If
				 	End If
				End If
			ElseIf AwaitingRedispatch(rs.Fields("MemoNumber")) = True Then
				If rs.Fields("FilterChange") <> 1 Then
					If CustHasPendingFilterChange(CustID) = True Then
						DaysTilChange = datediff("d",Date(),FChange_NextDate(CustID))
						If DaysTilChange < 0 Then ' overdue	%>
							<br><span class="labelFilterChangeIndicatorAndButtonColorOverDue">Filter change<br>overdue <%=DaysTilChange%> days</span>
						<% Else %>
							<br><span class="labelFilterChangeIndicatorAndButtonColor">Filter change due<br>in <%=DaysTilChange%> days</span>
						<% End If	
					Else
						OpenTicks = ""
						OpenTicks = GetOpenFilterTicketsByCustID(CustID)
						'Need to split these up beacuse there may be one or more &
						'now we want to have links to the ticket details
						OpenTickArray = Split(OpenTicks,",")
						If Ubound(OpenTickArray) > 0 Then ' They have OPEN Filter tickets already %>
								<br><span class="labelFilterChangeIndicatorAndButtonColor">Filter on ticket<br>
								<% For x = 0 to Ubound(OpenTickArray)-1
									'Response.Write("<a href='./editServiceMemo.asp?memo=" & OpenTickArray(x) & "' target='_blank'>" & OpenTickArray(x) & "</a><br>")
									%><%= OpenTickArray(x) %><br><%
								Next 
								%></span>
						<%End If
				 	End If
				End If
			End If
		End If%></td>
		<td><%= rs("SubmissionSource") %></td>
</tr><!-- eof table line !-->



