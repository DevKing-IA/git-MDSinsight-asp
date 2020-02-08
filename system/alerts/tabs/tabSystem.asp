<%If userIsAdmin(Session("UserNo")) Then %> 
			<div role="tabpanel" class="tab-pane fade in active" id="System"> 
				<br>
				<div class="row">
				 	<div class="col-lg-12">
					 	<p>
							 <a href="addAlertSystem.asp">
						    	<button type="button" class="btn btn-success">Add New System Alert</button>
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
		                  <th>Condition</th>
		                  <th class="sorttable_nosort">Email To</th>
		                  <th class="sorttable_nosort">Email Addtl</th>
		                  <th class="sorttable_nosort">Text To</th>
		                  <th class="sorttable_nosort">Text Addtl</th>
		                  <th class="sorttable_nosort">Enabled</th>
		                  <th class="sorttable_nosort">Delete</th>
		                  <th>Created<br>By</th>
		                </tr>
		              </thead>
		              <tbody>
	          
						<%
			
						SQL = "SELECT * FROM SC_Alerts Where AlertType='System' order by AlertName"
		
						Set cnn8 = Server.CreateObject("ADODB.Connection")
						cnn8.open (Session("ClientCnnString"))
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.CursorLocation = 3 
						Set rs = cnn8.Execute(SQL)

						If not rs.EOF Then
		
							Do While Not rs.EOF
							
								ShowLine = True
								
								'Is this a prospecting alert & for whatever reason prospecting module is not enabled
								If rs.Fields("Condition") = "ProspectNoNextActivity" Then
									If MUV_Read("prospectingModuleOn") <> "Enabled" Then ShowLine = False
								End If
					        	
					        	
					        	If ShowLine = True Then%>
								<!-- table line !-->
								<tr>
									<td>
										<a href='editAlertSystem.asp?a=<%= rs.Fields("InternalAlertRecNumber")%>'><%= rs.Fields("AlertName")%></a></td>
										<td><%
											Select Case rs.Fields("Condition")
												Case "BackendNoStart"
													Response.Write("Backend data import did not start")
												Case "BackendRunTooLong"
													Response.Write("Backend data import running too long")
												Case "BackendStarted"
													Response.Write("Backend data import started")
												Case "BackendFinished"
													Response.Write("Backend data import finished")
												Case "RebuildNotRun"
													Response.Write("Daily data rebuild did not run")
												Case "RebuildRunTooLong"
													Response.Write("Daily data rebuild running too long")
												Case "RebuildStarted"
													Response.Write("Daily data rebuild started")
												Case "RebuildFinished"
													Response.Write("Daily data rebuild finished")
												Case "DBoardNotRun"
													Response.Write("Nightly delivery board did not run")
												Case "DBoardRunTooLong"
													Response.Write("Nightly delivery board running too long")
												Case "DBoardSkipped"
													Response.Write("Nightly delivery board update skipped the update")
												Case "DBoardFinished"
													Response.Write("Nightly delivery board update finished")
												Case "DBoardOnDemandRun"
													Response.Write("Delivery board update on demand was run")
												Case "AutoCompJSONNotRun"
													Response.Write("Auto-complete JSON file not rebuilt")
												Case "HistOldInvoice"
													Response.Write("Most recent history invoice too old")
												Case "RouteFileEmpty"
													Response.Write("Route file empty")
												Case "ProspectNoNextActivity"
													Response.Write("Prospect found with no next activity")

											End Select
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
							   		<td><a href='deleteAlertQues.asp?a=<%=rs.Fields("InternalAlertRecNumber")%>&tab=System'><i class="fas fa-trash-alt"></i></a></td>
									<td><%=GetUserDisplayNameByUserNo(rs.Fields("CreatedByUserNo")) %></td>
							   	</tr>
								<% End If ' This is the endif for showline variable
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
<%End IF %>