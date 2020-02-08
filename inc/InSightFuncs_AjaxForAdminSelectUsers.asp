<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
action = Request("action")

Select Case action
    Case "GetContentForSalesRepUserIDs"
        GetContentForSalesRepUserIDs()
    Case "GetContentForProspSnapshotUserNos"
        GetContentForProspSnapshotUserNos()
    Case "GetContentForAPIDailyActivityReportUserIDs"
        GetContentForAPIDailyActivityReportUserIDs()
End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



Sub GetContentForSalesRepUserIDs()

	SQL = "SELECT ProspSnapshotSalesRepDisplayUserNos FROM Settings_Global "
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
        ProspSnapshotSalesRepDisplayUserNos = rs("ProspSnapshotSalesRepDisplayUserNos")
    End If
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	%>
	<!-- list of all users -->
							<div class="col-lg-4 line-full">
								<p>Master Sales Rep List</p>
								<select multiple class="form-control multi-select" id="lstAllSalesRepUserIDs" name="lstAllSalesRepUserIDs">
									<%	'Get list of all sales reps not currently selected
										
									Set cnnUserList = Server.CreateObject("ADODB.Connection")
									cnnUserList.open Session("ClientCnnString")
									
									SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 AND "
									SQLUserList = SQLUserList & " (userType = 'Admin' "
									SQLUserList = SQLUserList & " OR userType = 'CSR' "
									SQLUserList = SQLUserList & " OR userType = 'CSR Manager' "
									SQLUserList = SQLUserList & " OR userType = 'Inside Sales' "
									SQLUserList = SQLUserList & " OR userType = 'Inside Sales Manager' "
									SQLUserList = SQLUserList & " OR userType = 'Outside Sales' "
									SQLUserList = SQLUserList & " OR userType = 'Outside Sales Manager' "
									SQLUserList = SQLUserList & " OR userType = 'Finance' "
									SQLUserList = SQLUserList & " OR userType = 'Telemarketing') "
									
									If ProspSnapshotSalesRepDisplayUserNos <> "" Then
										SQLUserList = SQLUserList & " AND tblUsers.UserNo NOT IN (" & ProspSnapshotSalesRepDisplayUserNos & ")"
									End If
									
									SQLUserList = SQLUserList & " ORDER BY userFirstName,userLastName"
									
									'Response.Write(SQLUserList & "<br><br>")
									
									Set rsUserList = Server.CreateObject("ADODB.Recordset")
									rsUserList.CursorLocation = 3 
									Set rsUserList = cnnUserList.Execute(SQLUserList)
									
									If Not rsUserList.EOF Then
										Do While Not rsUserList.EOF
										
											FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName")
											Response.Write("<option value='" & rsUserList("UserNo") & "'>" & FullName & "</option>")
									
											rsUserList.MoveNext
										Loop
									End If
					
									Set rsUserList = Nothing
									cnnUserList.Close
									Set cnnUserList = Nothing
									
									%>
								</select>
							</div>
							<!-- eof list of all users -->
                            	                                
                    	
	                        <!-- add / remove -->
	                        <div class="col-lg-3 line-full" style="text-align:center">
	                            <a href="javascript:void(0)" onclick="javascript:listbox_addsalesrep()"><button type="button" class="btn btn-success" style="margin-bottom:10px;margin-top:80px;width:95px;">Add<br>Sales Rep <i class="fa fa-arrow-right" aria-hidden="true"></i></button></a>
								<a href="javascript:void(0)" onclick="javascript:listbox_removesalesrep()"><button type="button" class="btn btn-danger"><i class="fa fa-arrow-left" aria-hidden="true"></i> Remove<br>Sales Rep </button></a>
	                        </div>
	                        <!-- eof add / remove -->
					            
                                
                            	
								<!-- list of Selected users -->
								<div class="col-lg-4 line-full">
									<p>Include These Reps In Report</p>
									<select multiple class="form-control multi-select" id="lstSelectedSalesRepUserIDs" name="lstSelectedSalesRepUserIDs">
									<%	'Get list of all users currently selected
									If ProspSnapshotSalesRepDisplayUserNos <> "" Then
										
										Set cnnUserList = Server.CreateObject("ADODB.Connection")
										cnnUserList.open Session("ClientCnnString")
	
										SQLUserList = "SELECT * FROM tblUsers WHERE (userArchived <> 1 AND "
										SQLUserList = SQLUserList & " userType = 'Admin' "
										SQLUserList = SQLUserList & " OR userType = 'CSR' "
										SQLUserList = SQLUserList & " OR userType = 'CSR Manager' "
										SQLUserList = SQLUserList & " OR userType = 'Inside Sales' "
										SQLUserList = SQLUserList & " OR userType = 'Inside Sales Manager' "
										SQLUserList = SQLUserList & " OR userType = 'Outside Sales' "
										SQLUserList = SQLUserList & " OR userType = 'Outside Sales Manager' "
										SQLUserList = SQLUserList & " OR userType = 'Finance' "
										SQLUserList = SQLUserList & " OR userType = 'Telemarketing') "
										SQLUserList = SQLUserList & " AND tblUsers.UserNo IN (" & ProspSnapshotSalesRepDisplayUserNos & ")"
										SQLUserList = SQLUserList & " ORDER BY userFirstName,userLastName"
										
										Set rsUserList = Server.CreateObject("ADODB.Recordset")
										rsUserList.CursorLocation = 3 
										Set rsUserList = cnnUserList.Execute(SQLUserList)
										
										If Not rsUserList.EOF Then
											Do While Not rsUserList.EOF
											
												FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName")
												Response.Write("<option value='" & rsUserList("UserNo") & "' selected>" & FullName & "</option>")
										
												rsUserList.MoveNext
											Loop
										End If
						
										Set rsUserList = Nothing
										cnnUserList.Close
										Set cnnUserList = Nothing
										
									End If%>
								</select>
								</div>
							<!-- eof list of Selected users-->

<%	
End Sub

Sub GetContentForProspSnapshotUserNos()
SQL = "SELECT ProspSnapshotUserNos FROM Settings_Global "
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
        ProspSnapshotUserNos = rs("ProspSnapshotUserNos")
    End If
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
%>
<!-- list of all users -->
							<div class="col-lg-4 line-full">
								<p>Master User List</p>
								<select multiple class="form-control multi-select" id="lstAllUserIDs" name="lstAllUserIDs">
									<%	'Get list of all users not currently selected
										'Dont include sales managers because they are already
										'handles by the checkboxes
										
									Set cnnUserList = Server.CreateObject("ADODB.Connection")
									cnnUserList.open Session("ClientCnnString")

									SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 AND "
									SQLUserList = SQLUserList & "UserType <> 'Inside Sales Manager' AND "
									SQLUserList = SQLUserList & "UserType <> 'Outside Sales Manager'"
									
									If ProspSnapshotUserNos <> "" Then
										SQLUserList = SQLUserList & " AND tblUsers.UserNo NOT IN (" & ProspSnapshotUserNos & ")"
									End If
									
									SQLUserList = SQLUserList & " ORDER BY userFirstName,userLastName"
									
									Set rsUserList = Server.CreateObject("ADODB.Recordset")
									rsUserList.CursorLocation = 3 
									Set rsUserList = cnnUserList.Execute(SQLUserList)
									
									If Not rsUserList.EOF Then
										Do While Not rsUserList.EOF
										
											FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName")
											Response.Write("<option value='" & rsUserList("UserNo") & "'>" & FullName & "</option>")
									
											rsUserList.MoveNext
										Loop
									End If
					
									Set rsUserList = Nothing
									cnnUserList.Close
									Set cnnUserList = Nothing
									
									%>
								</select>
							</div>
							<!-- eof list of all users -->
                            	                                
                    	
	                        <!-- add / remove -->
	                        <div class="col-lg-3 line-full" style="text-align:center">
	                            <a href="javascript:void(0)" onclick="javascript:listbox_adduser()"><button type="button" class="btn btn-success" style="margin-bottom:10px;margin-top:40px;width:95px;">Add<br>User <i class="fa fa-arrow-right" aria-hidden="true"></i></button></a>
								<a href="javascript:void(0)" onclick="javascript:listbox_removeuser()"><button type="button" class="btn btn-danger"><i class="fa fa-arrow-left" aria-hidden="true"></i> Remove<br>User </button></a>
	                        </div>
	                        <!-- eof add / remove -->
					            
                                
                            	
								<!-- list of Selected users -->
								<div class="col-lg-4 line-full">
									<p>Send Report To</p>
									<select multiple class="form-control multi-select" id="lstSelectedUserIDs" name="lstSelectedUserIDs">
									<%	'Get list of all users currently selected
									If ProspSnapshotUserNos <> "" Then
										
										Set cnnUserList = Server.CreateObject("ADODB.Connection")
										cnnUserList.open Session("ClientCnnString")
	
										SQLUserList = "SELECT * FROM tblUsers WHERE "
										SQLUserList = SQLUserList & "tblUsers.UserNo IN (" & ProspSnapshotUserNos & ")"
										SQLUserList = SQLUserList & " ORDER BY userFirstName,userLastName"
										
										Set rsUserList = Server.CreateObject("ADODB.Recordset")
										rsUserList.CursorLocation = 3 
										Set rsUserList = cnnUserList.Execute(SQLUserList)
										
										If Not rsUserList.EOF Then
											Do While Not rsUserList.EOF
											
												FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName")
												Response.Write("<option value='" & rsUserList("UserNo") & "' selected>" & FullName & "</option>")
										
												rsUserList.MoveNext
											Loop
										End If
						
										Set rsUserList = Nothing
										cnnUserList.Close
										Set cnnUserList = Nothing
										
									End If%>
								</select>
								</div>
								<!-- eof list of Selected users-->
<%
End Sub

Sub GetContentForAPIDailyActivityReportUserIDs()
    SQL = "SELECT APIDailyActivityReportUserNos FROM Settings_Global "
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
        APIDailyActivityReportUserNos = rs("APIDailyActivityReportUserNos")
    End If
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
%>
<!-- list of all users -->
					<div class="col-lg-4 line-full">
						<p>Master User List</p>
						<select multiple class="form-control multi-select" id="lstAPIDailyActivityReportUserIDs" name="lstAPIDailyActivityReportUserIDs">
							<%	'Get list of all users not currently selected
								
							Set cnnUserList = Server.CreateObject("ADODB.Connection")
							cnnUserList.open Session("ClientCnnString")

							SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 "
							
							If APIDailyActivityReportUserNos <> "" Then
								SQLUserList = SQLUserList & " AND tblUsers.UserNo NOT IN (" & APIDailyActivityReportUserNos & ")"
							End If
							
							SQLUserList = SQLUserList & " ORDER BY userFirstName,userLastName"
							
							Set rsUserList = Server.CreateObject("ADODB.Recordset")
							rsUserList.CursorLocation = 3 
							Set rsUserList = cnnUserList.Execute(SQLUserList)
							
							If Not rsUserList.EOF Then
								Do While Not rsUserList.EOF
								
									FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName")
									Response.Write("<option value='" & rsUserList("UserNo") & "'>" & FullName & "</option>")
							
									rsUserList.MoveNext
								Loop
							End If
			
							Set rsUserList = Nothing
							cnnUserList.Close
							Set cnnUserList = Nothing
							
							%>
						</select>
					</div>
					<!-- eof list of all users -->
                    	
                        <!-- add / remove -->
                        <div class="col-lg-3 line-full" style="text-align:center">
                            <a href="javascript:void(0)" onclick="javascript:listbox_addAPIReportUser()"><button type="button" class="btn btn-success" style="margin-bottom:10px;margin-top:40px;">Add User <i class="fa fa-arrow-right" aria-hidden="true"></i></button></a>
							<a href="javascript:void(0)" onclick="javascript:listbox_removeAPIReportUser()"><button type="button" class="btn btn-danger"><i class="fa fa-arrow-left" aria-hidden="true"></i> Remove User </button></a>
                        </div>
                        <!-- eof add / remove -->
                    	
						<!-- list of Selected users -->
						<div class="col-lg-4 line-full">
							<p>Send Report To</p>
							<select multiple class="form-control multi-select" id="lstSelectedAPIDailyActivityReportUserIDs" name="lstSelectedAPIDailyActivityReportUserIDs">
							<%	'Get list of all users currently selected
							If APIDailyActivityReportUserNos <> "" Then
								
								Set cnnUserList = Server.CreateObject("ADODB.Connection")
								cnnUserList.open Session("ClientCnnString")

								SQLUserList = "SELECT * FROM tblUsers WHERE "
								SQLUserList = SQLUserList & "tblUsers.UserNo IN (" & APIDailyActivityReportUserNos & ")"
								SQLUserList = SQLUserList & " ORDER BY userFirstName,userLastName"
								
								Set rsUserList = Server.CreateObject("ADODB.Recordset")
								rsUserList.CursorLocation = 3 
								Set rsUserList = cnnUserList.Execute(SQLUserList)
								
								If Not rsUserList.EOF Then
									Do While Not rsUserList.EOF
									
										FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName")
										Response.Write("<option value='" & rsUserList("UserNo") & "' selected>" & FullName & "</option>")
								
										rsUserList.MoveNext
									Loop
								End If
				
								Set rsUserList = Nothing
								cnnUserList.Close
								Set cnnUserList = Nothing
								
							End If%>
						</select>
						</div>
                        <%
End Sub



'********************************************************************************************************************************************************

'END ALL AJAX MODAL SUBROUTINES AND FUNCTIONS

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

%>