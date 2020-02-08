
				<div role="tabpanel" class="tab-pane fade in"  id="ServiceNumTicks">

				<br>
				<div class="row">
				 	<div class="col-lg-12">
					 	<p>
							 <a href="addAlertServiceNumTick.asp">
						    	<button type="button" class="btn btn-success">Add New Service Alert</button>
							</a>
					 	</p>
				 	</div>
				</div>
				<br>
					
				<div class="table-responsive">
		            <table    class="table table-striped table-condensed table-hover sortable">
		              <thead>
		                <tr>
		                  <th>Alert Name</th>
		                  <th>Field</th>
		                  <th>Value</th>
		                  <th># Tickets</th>
		                  <th># Days</th>
		                  <th>Alert To</th>
		                  <th class="sorttable_nosort">Enabled</th>
		                  <th class="sorttable_nosort">Delete</th>
		                  <th>Created<br>By</th>
		                </tr>
		              </thead>
		              <tbody>
              
						<%
			
						SQL = "SELECT * FROM SC_Alerts Where AlertType='ServiceNumTick' order by AlertName"
		
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
										<a href='editAlertServiceNumTick.asp?a=<%= rs.Fields("InternalAlertRecNumber")%>'><%= rs.Fields("AlertName")%></a></td>
										<td><%= rs.Fields("ReferenceField")%></td>					
										<td><%= rs.Fields("ReferenceValue")%></td>					
										<td><%= rs.Fields("NumberOfTickets")%></td>					
										<td><%= rs.Fields("NumberOfDays")%></td>
										<td><%= rs.Fields("SendAlertTo")%></td>
									<td>
										<% If rs.Fields("Enabled") = True Then %>
											<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
										<% Else %>
											<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
										<% End If%>
									</td>
							   		<td><a href='deleteAlertQues.asp?a=<%=rs.Fields("InternalAlertRecNumber")%>&tab=ServiceNumTicks'><i class="fas fa-trash-alt"></i></a></td>
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
 				
 				</div></div>
							
  				