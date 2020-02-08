 <div role="tabpanel" class="tab-pane fade in" id="ServiceElapsed"> 
			<br>
			<div class="row">
			 	<div class="col-lg-12">
				 	<p>
						 <a href="addAlertServiceElapsed.asp">
					    	<button type="button" class="btn btn-success">Add New Service Alert</button>
						</a>
				 	</p>
			 	</div>
			</div>
			<br>

			<div class="table-responsive">
            <table class="table table-striped table-condensed table-hover sortable">
				<thead>
					<tr>
						<th>Alert Name</th>
						<th>Condition</th>
						<th>Time</th>
						<th class="sorttable_nosort">Email To</th>
						<th class="sorttable_nosort">Email Addtl</th>
						<th class="sorttable_nosort">Text To</th>
						<th class="sorttable_nosort">Text Addtl</th>
						<th class="sorttable_nosort">Enabled</th>
						<th class="sorttable_nosort">Delete </th>
						<th>Created<br>By</th>
					</tr>
				</thead>
				
				<tbody>
      
					<%
	
					SQL = "SELECT * FROM SC_Alerts Where AlertType='ServiceElapsed' order by AlertName"

					Set cnn8 = Server.CreateObject("ADODB.Connection")
					cnn8.open (Session("ClientCnnString"))
					Set rs = Server.CreateObject("ADODB.Recordset")
					rs.CursorLocation = 3 
					Set rs = cnn8.Execute(SQL)
		
					If not rs.EOF Then

						Do While Not rs.EOF
		
				        %>
						<!-- table line !-->
							<tr>
								<td>
									<a href='editAlertServiceElapsed.asp?a=<%= rs.Fields("InternalAlertRecNumber")%>'><%= rs.Fields("AlertName")%></a></td>
									<td><%
										Select Case rs.Fields("Condition")
											Case "NotDispatched"
												Response.Write("Ticket has not been dispatched")
											Case "NoACK"
												Response.Write("No dispatch ACK received")
											Case "NoOnSite"
												Response.Write("No tech on-site")
											Case "OpenTooLong"
												Response.Write("Ticket has been open")
											Case "RedispatchTooLong"
												Response.Write("Ticket has been awaiting RE-dispatch")
											Case "AnyStage"
												Response.Write("Ticket has been idle in any stage")
											Case "Declined"
												Response.Write("Dispatch has been declined by tech")
											Case "ARHold"
												Response.Write("When a ticket has been on A/R Hold longer than")
											Case "GPHold"
												Response.Write("When a ticket has been on GP Hold longer than")
												
										End Select
									%>
									</td>		
									<td>
										<%
										If rs("NBMinutes") mod 60 = 0 Then
											Response.Write(rs("NBMinutes")/60 & "hour(s)")
										Else
											Response.Write(rs("NBMinutes") & " minutes")
										End If
										%>
									</td>
									<td>
									<%
									UserListToWrite = ""
									If Not IsNull(rs.Fields("EmailToUserNos")) Then
										If rs.Fields("EmailToUserNos") <> "" And rs.Fields("EmailToUserNos") <> "0" Then
											UserNoList = Split(rs.Fields("EmailToUserNos"),",")
											For x = 0 To UBound(UserNoList)
												UserListToWrite = UserListToWrite  & GetUserFirstAndLastNameByUserNo(UserNoList(x)) & "<br>"
											Next
											UserListToWrite  = Left(UserListToWrite,Len(UserListToWrite)-4) ' Strip last <br>
											Response.Write(UserListToWrite)
										End If
									End If
									%>
									</td>
									<td>
									<%
									If rs.Fields("AdditionalEmails") <> "" Then
										If Ucase(rs.Fields("Condition")) = "LOG" Then
											If rs.Fields("NBIncludeLog") = vbTrue Then
												Response.Write("<strong>(log incuded)</strong><br>")
											Else
												Response.Write("<strong>(log not incuded)</strong><br>")
											End If
										End If
										Response.Write(Replace(rs.Fields("AdditionalEmails"),";","<br>"))
									End If
									%>
									</td>					
									<td>
									<%
									UserListToWrite = ""
									If Not IsNull(rs.Fields("TextToUserNos")) Then
										If rs.Fields("TextToUserNos") <> "" And rs.Fields("TextToUserNos") <> "0" Then
											UserNoList = Split(rs.Fields("TextToUserNos"),",")
											For x = 0 To UBound(UserNoList)
												UserListToWrite = UserListToWrite  & GetUserFirstAndLastNameByUserNo(UserNoList(x)) & "<br>"
											Next
											UserListToWrite  = Left(UserListToWrite,Len(UserListToWrite)-4) ' Strip last <br>
											Response.Write(UserListToWrite)
										End If
									End If
									%>
									</td>
									<td><%= Replace(rs.Fields("AdditionalText"),";","<br>")%></td>
								<td>
									<% If rs.Fields("Enabled") = True Then %>
										<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
									<% Else %>
										<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
									<% End If%>
								</td>
						   		<td><a href='deleteAlertQues.asp?a=<%=rs.Fields("InternalAlertRecNumber")%>&tab=ServiceElapsed'><i class="fas fa-trash-alt"></i></a></td>
								<td><%=GetUserDisplayNameByUserNo(rs.Fields("CreatedByUserNo")) %></td>
						   	</tr>
							<%
							rs.movenext
						loop
					End If
					set rs = Nothing
					cnn8.close
					set cnn8 = Nothing
		            %>
		            
				</tbody>
			</table>
 		</div>
 </div>
		