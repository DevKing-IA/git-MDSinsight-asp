<%	

	SQLCustInfo = "SELECT * FROM AR_Customer WHERE CustNum = '" & rs.Fields("AccountNumber") & "'"
		
	Set rsCustInfo = cnnCustInfo.Execute(SQLCustInfo)
		
	If Not rsCustInfo.EOF Then CustTypeVar = rsCustInfo("CustType") Else CustType=""


		'No priority alert colors on closed tickets

		If LineX Mod 2 = 0 then
			'THESE ARE EVEN LINES
			Response.Write("<tr class='tr-even'>")
		Else
			'THESE ARE ODD LINE
			Response.Write("<tr class='tr-odd'>")
		End If

	 

	submissionOpenedDateTime = GetServiceTicketOpenDateTime(rs.Fields("MemoNumber"))
	 
	submissionHour = Hour(submissionOpenedDateTime)
	submissionMinute = Minute(submissionOpenedDateTime)
	submissionZeroFactor = "0" & submissionMinute
	submissionAMPM = "AM"
	If submissionHour >= 12 then submissionAMPM = "PM"
	If submissionHour > 12 then submissionHour = submissionHour - 12
	If submissionMinute <= 9 then submissionMinute = submissionZeroFactor
	
	submissionDateTime = GetServiceTicketOpenDateTime(rs.Fields("MemoNumber"))
	
	ticketDateDisplay = padDate(MONTH(submissionDateTime),2) & "/" & padDate(DAY(submissionDateTime),2) & "/" & padDate(RIGHT(YEAR(submissionDateTime),2),2)

	Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(rs("submissionDateTime")) & ">" & ticketDateDisplay & "<br>" & submissionHour & ":" & submissionMinute & " " & submissionAMPM & "</td>")%>
	
	<td><%= rs.Fields("MemoNumber")%></td>

	<td>
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
	<%= rsCustInfo("Addr1") %><br>
	<% If Trim(rsCustInfo("Addr2")) <> "" Then Response.Write(rsCustInfo("Addr2")& "<br>") %>
	<%= rsCustInfo("City") %>, <%= rsCustInfo("State") %>&nbsp;<%= rsCustInfo("Zip") %>
	<% If Trim(rsCustInfo("Addr2")) = "" Then Response.Write("<br><br>")%>
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
	
		closedHour = Hour(rs("RecordCreateDateTime"))
		closedMinute = Minute(rs("RecordCreateDateTime"))
		closedZeroFactor = "0" & closedMinute
		closedAMPM = "AM"
		If closedHour >= 12 then closedAMPM = "PM"
		If closedHour > 12 then closedHour = closedHour - 12
		If closedMinute <= 9 then closedMinute = closedZeroFactor	 
	
		Response.Write(GetUserDisplayNameByUserNo(rs.Fields("UserNoOfServiceTech")))
		Response.Write("<br>")
			
		closedDateTime = rs("RecordCreateDateTime")
		
		closedDateTimeDisplay = padDate(MONTH(closedDateTime),2) & "/" & padDate(DAY(closedDateTime),2) & "/" & padDate(RIGHT(YEAR(closedDateTime),2),2)
			
		Response.Write(closedDateTimeDisplay & " " & closedHour & ":" & closedMinute & " " & closedAMPM)

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

		
	elapsedMinutes = ServiceCallElapsedMinutesClosedTicket(rs.Fields("MemoNumber"))
	minutesInServiceDay = NumberOfMinutesInServiceDayVar
	
	If elapsedMinutes < 1 Then elapsedMinutes = 1 ' If it has been less than 1 minute, just show 1 anyway
	elapsedMinutesForSorting = elapsedMinutes
	elapsedString = ""
	elapsedDays = 	elapsedMinutes \ minutesInServiceDay
	If int(elapsedDays) > 0 Then
		elapsedMinutes = elapsedMinutes - (int(elapsedDays) * minutesInServiceDay)
		elapsedString = elapsedDays & "d "
	End If
	elapsedHours = 	elapsedMinutes \ 60
	If int(elapsedHours) > 0 Then 
		elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
		elapsedString = elapsedString  & elapsedHours & "h "
	End IF	
	If int(elapsedMinutes) > 0 Then
		elapsedString = elapsedString  & elapsedMinutes & "m"
	End If
	
	Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 


	%>
	</td>
	
	<%
	
	'*************************************************************
	'We need to collect more information about this ticket
	'to see what actions are available under the actions modal
	'*************************************************************

	TicketNumber = rs.Fields("MemoNumber")
	ServiceTicketCurrentStage = GetServiceTicketCurrentStageVar 
	
	CustID = rs.Fields("AccountNumber")

	If len(rs.Fields("Company")) > 19 then 
		Cnam = left(rs.Fields("Company"),19) 
	Else 
		Cnam = rs.Fields("Company")
	End If
	
	%>
	
	<td>
	
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

		</td>
		
		<td><%=rs("SubmissionSource")%></td></tr><!-- eof table line !-->