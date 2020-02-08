<style>

	.signature-width{
		max-width: 100px;
	}

</style>

<div role="tabpanel" class="tab-pane fade in" id="service">

	<% ' S E R V I C E  T I C K E T S
	SQL = "SELECT * FROM FS_ServiceMemos "
	SQL = SQL & " WHERE AccountNumber ='" & customerID & "' "
	SQL = SQL & " ORDER BY submissionDateTime desc"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	%>
	<!-- row !-->
	<div class="row-line">
		<div class="table-responsive">
 			<table class="table sortable table-striped">
              <thead>
                <tr>
	              <th class="sorttable_numeric">Date</th>
	              <th>Ticket #</th>	              
                  <th>Status</th>
                  <th class="sorttable_nosort">&nbsp;</th>
                  <th class="sorttable_nosort">Description</th>
                  <% If advancedDispatchIsOn() Then %>
		              <th>Stage</th>
	              <% Else %>
	                  <th>Dispatched</th>
	              <% End If %>
                  <th class="sorttable_numeric">Elapsed<br>Time</th>
                  <th class="sorttable_nosort">PIC</th>
                  <th class="sorttable_nosort">SIG</th>
                  <th>Submitted Via</th>
                </tr>
              </thead>
              
              <tbody class='searchable'>
				<%
				Do While Not rs.EOF
						If rs.Fields("CurrentStatus") = rs.Fields("RecordSubType") Then ' Show only 1 line per memo, the most current status
				        %>
							<!-- table line !-->
							<tr class="low-priority">
							<%Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(rs("submissionDateTime")) & ">" & FormatDateTime(rs("submissionDateTime")) & "</td>")%>
							<%If rs.Fields("CurrentStatus")="OPEN" Then %>
								<td><a href='<%= BaseURL %>service/editServiceMemo.asp?memo=<%= rs.Fields("MemoNumber")%>'><%= rs.Fields("MemoNumber")%></a></td>
							<% Else %>
								<td><a href='<%= BaseURL %>service/viewServiceMemo.asp?memo=<%= rs.Fields("MemoNumber")%>'><%= rs.Fields("MemoNumber")%></a></td>
							<% End If %>
							<td><%= rs.Fields("RecordSubType") %></td>
							<td>&nbsp;</td>
							<td>
							<%
								CompressLen = 27
								'See if there are linefeeds in there that need to come out
								If Instr(rs.Fields("ProblemDescription"),"<br>") <> 0 Then CompressLen = Instr(rs.Fields("ProblemDescription"),"<br>")
								If CompressLen > 27 Then CompressLen = 27
								If len(rs.Fields("ProblemDescription")) > CompressLen Then Response.Write(Left(rs.Fields("ProblemDescription"),CompressLen)) Else Response.Write(rs.Fields("ProblemDescription"))%>
							</td>
							<%
								If rs.Fields("CurrentStatus") <> "CLOSE" and rs.Fields("CurrentStatus") <> "CANCEL" Then ' dont show a stage if they are closed or cancelled
									Response.Write("<td><b>"& GetServiceTicketCurrentStage(rs.Fields("MemoNumber")) & "</b><br>")
									Response.Write(GetServiceTicketSTAGEUser(rs.Fields("MemoNumber"),GetServiceTicketCurrentStage(rs.Fields("MemoNumber"))) & "<br>")
									Response.Write(GetServiceTicketSTAGEDateTime(rs.Fields("MemoNumber"),GetServiceTicketCurrentStage(rs.Fields("MemoNumber")))& "</td>")
								Else
									Response.Write("<td>&nbsp;</td>")
								End If
							If ElapsedTimeCalcMethod() = "Actual" Then
								If rs.Fields("CurrentStatus") = "CLOSE" or rs.Fields("CurrentStatus") = "CANCEL" Then
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),GetServiceTicketCloseDateTime(rs.Fields("MemoNumber")))
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									minutesInServiceDay = GetNumberOfMinutesInServiceDay()
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
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),GetServiceTicketCloseDateTime(rs.Fields("MemoNumber")))
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								ElseIf rs.Fields("CurrentStatus") = "OPEN" Then 
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									minutesInServiceDay = GetNumberOfMinutesInServiceDay()
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
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								Elseif rs.Fields("CurrentStatus") = "HOLD" Then
									'Response.Write("<td sorttable_customkey='" & 0 & "'>" & "Hold<br>") 
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									minutesInServiceDay = GetNumberOfMinutesInServiceDay()
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
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
								End If
							Else
								If rs.Fields("CurrentStatus") = "CLOSE" or rs.Fields("CurrentStatus") = "CANCEL" Then
									elapsedMinutes = ServiceCallElapsedMinutes(rs.Fields("MemoNumber"))
									elapsedMinutesForSorting = elapsedMinutes 
									elapsedString = ""
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									If elapsedMinutes = 0 Then elapsedString = "0"
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								ElseIf rs.Fields("CurrentStatus") = "OPEN" Then 
									elapsedMinutes = ServiceCallElapsedMinutes(rs.Fields("MemoNumber"))
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								Elseif rs.Fields("CurrentStatus") = "HOLD" Then
									'Response.Write("<td sorttable_customkey='" & 0 & "'>" & "Hold<br>") 
									elapsedMinutes = ServiceCallElapsedMinutes(rs.Fields("MemoNumber"))
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
								End If
							End If
							%>
							</td>
							<td>
							<%
							set fs = CreateObject("Scripting.FileSystemObject")
							Pth =  "../../clientfiles/" & trim(MUV_Read("ClientID")) & "/SvcMemoPics/" & rs("MemoNumber") & "-1.jpg"
							Pth2 =  "../../clientfiles/" & trim(MUV_Read("ClientID")) & "/SvcMemoPics/" & rs("MemoNumber") & "-1.jpeg"
							If fs.FileExists(Server.MapPath(Pth)) or fs.FileExists(Server.MapPath(Pth2)) Then
								%>X<%
							End If
							%>
							</td>

							<% If rs.Fields("RecordSubType") = "CLOSE" Then 
								
								'----------------------------
								'Service Signature Check
								'----------------------------
								set fs = CreateObject("Scripting.FileSystemObject")
								Pth =  "../../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & rs("MemoNumber") & ".png"
								
								If fs.FileExists(Server.MapPath(Pth)) Then
									hasServiceSignature = True
								Else
									hasServiceSignature = False
								End If
													
								'Response.Write(Pth)
								
								'***************************************************************************************************
								'Display signature file, if any exist in the signaturesave directory
								''Check for the existance of a thumbnail image in the directory, otherwise, size the image with CSS
								'***************************************************************************************************
				
								Pth =  "../../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & rs("MemoNumber") & ".png"
								PthThumb =  "../../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & rs("MemoNumber") & "-thumb.png"
			
								SignaturePathNameFull = BaseURL & "clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & rs("MemoNumber") & ".png"
								SignaturePathNameThumb = BaseURL & "clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & rs("MemoNumber") & "-thumb.png"
								
								If hasServiceSignature = True Then
									
									If fs.FileExists(Server.MapPath(PthThumb)) Then
								    	%><td align="left"><a href="<%= SignaturePathNameFull %>" target="_blank" style="border:0px;"><img src="<%= SignaturePathNameThumb %>" alt="Ticket <%= rs("MemoNumber") %> Signature" class="signature-width"></a></td><%
								    Else
								    	%><td align="left"><a href="<%= SignaturePathNameFull %>" target="_blank" style="border:0px;"><img src="<%= SignaturePathNameFull %>" alt="Ticket <%= rs("MemoNumber") %> Signature" class="signature-width"></a></td><%
								    End If
								    
								Else
									 %><td align="left">No Signature</td><%
								End If
								
								
							End If
							set fs=nothing
							%>
						
							<td><%= rs.Fields("SubmissionSource") %></td>
							</tr>
							<!-- eof table line !-->
						<%
				
						End If
						
						rs.movenext	
		   			Loop %>
				</tbody>
			</table>
		</div>
	</div>
	
	<%
	End IF
	cnn8.close
	set rs = nothing
	set cnn8 = nothing	
	%>

</div>