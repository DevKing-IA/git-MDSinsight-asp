<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->

<% Response.CacheControl = "no-cache, no-store, must-revalidate" %>


<% IF Session("LoginEmailSentTo") <> "" Then 

msg = "An email has been sent to " & Session("LoginEmailSentTo")

 %>
		<SCRIPT LANGUAGE="JavaScript">
		swal('<%= msg %>');
		</SCRIPT>     

<%
	'Session.Contents.Remove "LoginEmailSentTo"
	Session("LoginEmailSentTo") = ""
	
End If
%>

<% IF Session("LoginTextSentTo") <> "" Then 

msg = "A text has been sent to " & Session("LoginTextSentTo") 
%>

		<SCRIPT LANGUAGE="JavaScript">
		swal('<%= msg %>');
		</SCRIPT>     

<%
	'Session.Contents.Remove "LoginTextSentTo"
	Session("LoginTextSentTo") = ""
	
	
End If%>


<% 

ActiveTab = Request.QueryString("tab")

'Count the users
AdminUsers = 0
CSRUsers = 0
CSRManagerUsers = 0
FieldServiceUsers = 0
FinanceUsers = 0
FinanceManagerUsers = 0
ServiceManagerUsers = 0
ArchivedUsers = 0
InsideSalesUsers = 0
OutsideSalesUsers = 0
TelemarketingUsers = 0 
DriverUsers = 0
SQL = "SELECT userType, Count(userType) AS UTypeCOunt FROM tblUsers Where userArchived <> 1 Group By UserType"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	Do While Not rs.EOF
		Select Case ucase(rs("userType"))
			Case "ADMIN"
				AdminUsers = rs("UTypeCOunt")
			Case "CSR"
				CSRUsers = rs("UTypeCOunt")			
			Case "CSR MANAGER"
				CSRManagerUsers = rs("UTypeCOunt")			
			Case "REPAIR"
				RepairUsers = rs("UTypeCOunt")
			Case "FIELD SERVICE"
				FieldServiceUsers = rs("UTypeCOunt")			
			Case "FINANCE"
				FinanceUsers = rs("UTypeCOunt")			
			Case "FINANCE MANAGER"
				FinanceManagerUsers = rs("UTypeCOunt")			
			Case "SERVICE MANAGER"
				ServiceManagerUsers = rs("UTypeCOunt")		
			Case "INSIDE SALES"
				InsideSalesUsers = rs("UTypeCOunt")	
			Case "OUTSIDE SALES"
				OutsideSalesUsers = rs("UTypeCOunt")		
			Case "TELEMARKETING"
				TelemarketingUsers = rs("UTypeCOunt")
			Case "DRIVER"
				DriverUsers = rs("UTypeCOunt")									
		End Select
		rs.movenext
	Loop
End If
Set rs = Nothing
cnn8.Close
Set cnn8 = Nothing
SQL = "SELECT Count(userNo) AS ArchUserCount FROM tblUsers Where userArchived = 1"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then ArchivedUsers = rs("ArchUserCount")
Set rs = Nothing
cnn8.Close
Set cnn8 = Nothing
%>


 <style type="text/css">
 	.email-table{
		width:46%;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
}

.nav-tabs>li>a{
	background: #f5f5f5;
	border: 1px solid #ccc;
	color: #000;
}

.nav-tabs>li>a:hover{
	border: 1px solid #ccc;
}

.nav-tabs>li.active>a, .nav-tabs>li.active>a:focus, .nav-tabs>li.active>a:hover{
	color: #000;
	border: 1px solid #ccc;
}
 </style>

<!--- eof on/off scripts !-->
<div class="searchable">

<h1 class="page-header"><i class="fa fa-users"></i> Manage users</h1>

<div class="row">
	<div class="col-lg-5">

<div class="row">
 	<div class="col-lg-3">
	 	<p>
 			<a href="adduser.asp?tab=<%= ActiveTab %>"><button type="button" class="btn btn-success">Add New User</button></a>
	 	</p>
 	</div>


<div class="col-lg-6">
		<div class="input-group"> <span class="input-group-addon">Find User</span>
		    <input id="filter" type="text" class="form-control " placeholder="Type here...">
		</div>
	</div>
</div>

</div>
</div>

	<!-- tabs start here !-->
	<div class="row">
		<div class="col-lg-12">
		
			<!-- tabs navigation !-->
			<ul class="nav nav-tabs" role="tablist">
		        <% AllUserCount = CSRUsers+CSRManagerUsers+FieldServiceUsers+ServiceManagerUsers+FinanceUsers+FinanceManagerUsers+AdminUsers+InsideSalesUsers+OutsideSalesUsers+TelemarketingUsers+DriverUsers%>
   		        <li role="presentation" <% If ActiveTab = "" OR ActiveTab = "AllUsers" Then Response.write("class='active'") %>><a href="#AllUsers" aria-controls="tab3" role="tab" data-toggle="tab">All Users (<%=AllUserCount%>)</a></li>			
			    <% If MUV_READ("FILTERTRAX") <> "1" Then %>
			    	<li role="presentation" <% If ActiveTab="CustomerService" Then Response.write("class='active'") %>><a href="#CustomerService" aria-controls="manage" role="tab" data-toggle="tab"><%=GetTerm("Customer Service")%> (<%=CSRUsers+CSRManagerUsers%>)</a></li>
				<% End If %>			    	
   			    <li role="presentation" <% If ActiveTab = "Service2" Then Response.write("class='active'") %>><a href="#Service2" aria-controls="tab3" role="tab" data-toggle="tab"><%=GetTerm("Service")%> (<%=FieldServiceUsers+ServiceManagerUsers%>)</a></li>
			    <% If MUV_READ("FILTERTRAX") <> "1" Then %>
			    	<li role="presentation" <% If ActiveTab = "Accounting" Then Response.write("class='active'") %>><a href="#Accounting" aria-controls="tab3" role="tab" data-toggle="tab"><%=GetTerm("Accounting")%> (<%=FinanceUsers+FinanceManagerUsers%>)</a></li>
			    <% End If %>	
		        <li role="presentation" <% If ActiveTab = "Admins" Then Response.write("class='active'") %>><a href="#Admins" aria-controls="tab3" role="tab" data-toggle="tab"><%=GetTerm("Admins")%> (<%=AdminUsers%>)</a></li>
		        <% If MUV_READ("FILTERTRAX") <> "1" Then %>
   		        	<li role="presentation" <% If ActiveTab = "InsideSales" Then Response.write("class='active'") %>><a href="#InsideSales" aria-controls="tab3" role="tab" data-toggle="tab">Inside Sales (<%=InsideSalesUsers%>)</a></li>
	   		        <li role="presentation" <% If ActiveTab = "OutsideSales" Then Response.write("class='active'") %>><a href="#OutsideSales" aria-controls="tab3" role="tab" data-toggle="tab">Outside Sales (<%=OutsideSalesUsers%>)</a></li>
	   		        <li role="presentation" <% If ActiveTab = "Telemarketing" Then Response.write("class='active'") %>><a href="#Telemarketing" aria-controls="tab3" role="tab" data-toggle="tab">Telemarketing (<%=TelemarketingUsers%>)</a></li>
	   		        <li role="presentation" <% If ActiveTab = "Drivers" Then Response.write("class='active'") %>><a href="#Drivers" aria-controls="tab3" role="tab" data-toggle="tab">Drivers (<%=DriverUsers%>)</a></li>
	   		    <% End If %>
   		        <li role="presentation" <% If ActiveTab = "Archived" Then Response.write("class='active'") %>><a href="#Archived" aria-controls="tab3" role="tab" data-toggle="tab">Archived (<%=ArchivedUsers%>)</a></li>
  			</ul>
			<!-- eof tabs navigation !-->
			
			<!-- tabs content !-->
			<div class="tab-content">


		

<!--  All Users Tab !-->
<div role="tabpanel" class="tab-pane fade in active" id="AllUsers"> 
					
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th>User Name</th>
                  <th>Email Address</th>
                  <th>User Type</th>
                  <th>Last Login</th>
                  <th class="sorttable_nosort">Enabled</th>
                  <th class="sorttable_nosort">Send<br>Credentials</th>
                  <th class="sorttable_nosort">Audit<br>Trail</th>
                  <th class="sorttable_nosort">Archive</th>
                </tr>
              </thead>
              <tbody>
              
				<%
			

					SQL = "SELECT * FROM tblUsers Where userArchived <> 1 Order By userLastName"
		
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
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=AllUsers'><%= rs.Fields("userLastName")%>,&nbsp;<%= rs.Fields("userFirstName")%></a></td>
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=AllUsers'><%= rs.Fields("userEmail")%></a></td>					
						<td class="example"><%= rs.Fields("userType") %></td>
						<td>
						<% If IsNull(rs.Fields("userLastLogin")) Then %>
							Never
						<% Else %>
							<%= rs.Fields("userLastLogin") %>
						<% End If %>
						</td>
						<td>
						<div class="example">
						<% If rs.Fields("userEnabled") = True Then %>
							<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
						<% Else %>
							<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
						<% End If%>
						</div></td><td>
						<div class="example">
							<div>
								<a href='send_login_email.asp?userno=<%= rs.Fields("userNO") %>&tab=AllUsers' >Email</a>
								<% If Not IsNull(rs.Fields("userCellNumber")) Then %>
									<% If len(trim((rs.Fields("userCellNumber")))) > 0 Then %>
										&nbsp;&nbsp;<a href='send_login_text.asp?userno=<%= rs.Fields("userNO") %>&tab=AllUsers' >Text</a>							
									<% End If %>
								<% End If %>
							</div>
						</div></td>
						<td>
							<div class="example">
									    <a href='../../Reports\AuditTrail_OneUser.asp?unam=<%=rs.Fields("UserDisplayName")%>&tab=AllUsers' target='_blank'>
									    <img src='../../img/AuditTrailImage.png' height="35" title='View audit trail for <%=rs.Fields("UserDisplayName")%>' </a>
							</div>
						</td>
						<td>
						<% If rs.Fields("UserNo") <> Session("userno") Then%>
							<a href='archiveUserQues.asp?un=<%=rs.Fields("UserNo")%>&tab=AllUsers'><i class='fa fa-archive' ></i></a>
						<%Else%>
							&nbsp;
						<%End if%>
						</td>
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
<!-- eof All Users Tab !-->
		
				
<!-- Customer Service Tab !-->
<div role="tabpanel" class="tab-pane fade" id="CustomerService"> 
				
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th>User Name</th>
                  <th>Email Address</th>
                  <th>User Type</th>
                  <th>Last Login</th>
                  <th class="sorttable_nosort"><%=GetTerm("CSR")%></th>
                  <th class="sorttable_nosort"><%=GetTerm("CSR Manager")%></th>
                  <th class="sorttable_nosort">Enabled</th>
                  <th class="sorttable_nosort">Send<br>Credentials</th>
                  <th class="sorttable_nosort">Audit<br>Trail</th>
                </tr>
              </thead>
              <tbody>
              
				<%
			
				SQL = "SELECT * FROM tblUsers where (UserType = 'CSR' or UserType = 'CSR Manager') AND userArchived <> 1 order by userLastName"
		
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
						<td><a href='edituser.asp?tab=CustomerService&uno=<%= rs.Fields("userNo")%>'><%= rs.Fields("userLastName")%>,&nbsp;<%= rs.Fields("userFirstName")%></a></td>
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=CustomerService'><%= rs.Fields("userEmail")%></a></td>					
						<td class="example"><%= rs.Fields("userType") %></td>
						<td>
						<% If IsNull(rs.Fields("userLastLogin")) Then %>
							Never
						<% Else %>
							<%= rs.Fields("userLastLogin") %>
						<% End If %>
						</td>
						
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "CSR" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
	
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "CSR Manager" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
	
						<td>
						<div class="example">
						<% If rs.Fields("userEnabled") = True Then %>
							<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
						<% Else %>
							<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
						<% End If%>
						</div></td><td>
						<div class="example">
							<div>
								<a href='send_login_email.asp?userno=<%= rs.Fields("userNO") %>&tab=CustomerService'>Email</a>
								<% If Not IsNull(rs.Fields("userCellNumber")) Then %>
									<% If len(trim((rs.Fields("userCellNumber")))) > 0 Then %>
										&nbsp;&nbsp;<a href='send_login_text.asp?userno=<%= rs.Fields("userNO") %>&tab=CustomerService'>Text</a>							
									<% End If %>
								<% End If %>
							</div>
						</div></td>
						
						<td>
						<div class="example">
								    <a href='../../Reports\AuditTrail_OneUser.asp?unam=<%=rs.Fields("UserDisplayName")%>&tab=CustomerService' target='_blank'>
								    <img src='../../img/AuditTrailImage.png' height="35" title='View audit trail for <%=rs.Fields("UserDisplayName")%>' </a>
						</div></td>
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
<!-- eof Customer Service Tab !-->

				
<!--  Service Tab !-->
<div role="tabpanel" class="tab-pane fade" id="Service2"> 
					
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th>User Name</th>
                  <th>Email Address</th>
                  <th>User Type</th>
                  <th>Last Login</th>
                  <th class="sorttable_nosort"><%=GetTerm("Field Service")%></th>
                  <th class="sorttable_nosort"><%=GetTerm("Service Manager")%></th>
                  <th class="sorttable_nosort">Enabled</th>
                  <th class="sorttable_nosort">Send<br>Credentials</th>
                  <th class="sorttable_nosort">Audit<br>Trail</th>
                </tr>
              </thead>
              <tbody>
              
				<%
			
				SQL = "SELECT * FROM tblUsers where (UserType = 'Field Service' or UserType = 'Repair' or UserType = 'Service Manager') AND userArchived <> 1 order by userLastName"
		
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
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=Service2'><%= rs.Fields("userLastName")%>,&nbsp;<%= rs.Fields("userFirstName")%></a></td>
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=Service2'><%= rs.Fields("userEmail")%></a></td>					
						<td class="example"><%= rs.Fields("userType") %></td>
						<td>
						<% If IsNull(rs.Fields("userLastLogin")) Then %>
							Never
						<% Else %>
							<%= rs.Fields("userLastLogin") %>
						<% End If %>
						</td>
						
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "Field Service" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
	
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "Service Manager" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
	
						<td>
						<div class="example">
						<% If rs.Fields("userEnabled") = True Then %>
							<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
						<% Else %>
							<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
						<% End If%>
						</div></td><td>
						<div class="example">
							<div>
								<a href='send_login_email.asp?userno=<%= rs.Fields("userNO") %>&tab=Service2' >Email</a>
								<% If Not IsNull(rs.Fields("userCellNumber")) Then %>
									<% If len(trim((rs.Fields("userCellNumber")))) > 0 Then %>
										&nbsp;&nbsp;<a href='send_login_text.asp?userno=<%= rs.Fields("userNO") %>' >Text</a>							
									<% End If %>
								<% End If %>
							</div>
						</div></td>
						<td>
						<div class="example">
								    <a href='../../Reports\AuditTrail_OneUser.asp?unam=<%=rs.Fields("UserDisplayName")%>&tab=Service2' target='_blank'>
								    <img src='../../img/AuditTrailImage.png' height="35" title='View audit trail for <%=rs.Fields("UserDisplayName")%>' </a>
						</div></td>

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
<!-- eof Service Tab !-->

				
<!--  Accounting Tab !-->
<div role="tabpanel" class="tab-pane fade" id="Accounting"> 
					
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th>User Name</th>
                  <th>Email Address</th>
                  <th>User Type</th>
                  <th>Last Login</th>
                  <th class="sorttable_nosort"><%=GetTerm("Accounting")%></th>
                  <th class="sorttable_nosort"><%=GetTerm("Accounting Manager")%></th>
                  <th class="sorttable_nosort">Enabled</th>
                  <th class="sorttable_nosort">Send<br>Credentials</th>
                  <th class="sorttable_nosort">Audit<br>Trail</th>
                </tr>
              </thead>
              <tbody>
              
				<%
			
				SQL = "SELECT * FROM tblUsers where (UserType = 'Finance' or UserType = 'Finance Manager') AND userArchived <> 1  order by userLastName"
		
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
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=Accounting'><%= rs.Fields("userLastName")%>,&nbsp;<%= rs.Fields("userFirstName")%></a></td>
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=Accounting'><%= rs.Fields("userEmail")%></a></td>					
						<td class="example"><%= rs.Fields("userType") %></td>
						<td>
						<% If IsNull(rs.Fields("userLastLogin")) Then %>
							Never
						<% Else %>
							<%= rs.Fields("userLastLogin") %>
						<% End If %>
						</td>
						
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "Finance" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
	
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "Finance Manager" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
	
						<td>
						<div class="example">
						<% If rs.Fields("userEnabled") = True Then %>
							<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
						<% Else %>
							<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
						<% End If%>
						</div></td><td>
						<div class="example">
							<div>
								<a href='send_login_email.asp?userno=<%= rs.Fields("userNO") %>&tab=Accounting' >Email</a>
								<% If Not IsNull(rs.Fields("userCellNumber")) Then %>
									<% If len(trim((rs.Fields("userCellNumber")))) > 0 Then %>
										&nbsp;&nbsp;<a href='send_login_text.asp?userno=<%= rs.Fields("userNO") %>&tab=Accounting' >Text</a>							
									<% End If %>
								<% End If %>
							</div>
						</div></td>
						<td>
							<div class="example">
									    <a href='../../Reports\AuditTrail_OneUser.asp?unam=<%=rs.Fields("UserDisplayName")%>&tab=Accounting' target='_blank'>
									    <img src='../../img/AuditTrailImage.png' height="35" title='View audit trail for <%=rs.Fields("UserDisplayName")%>' </a>
							</div>
						</td>
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
<!-- eof Accounting Tab !-->


				
<!--  Admins Tab !-->
<div role="tabpanel" class="tab-pane fade" id="Admins"> 
					
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th>User Name</th>
                  <th>Email Address</th>
                  <th>User Type</th>
                  <th>Last Login</th>
                  <th class="sorttable_nosort">Enabled</th>
                  <th class="sorttable_nosort">Send<br>Credentials</th>
                  <th class="sorttable_nosort">Audit<br>Trail</th>
                </tr>
              </thead>
              <tbody>
              
				<%
			
				SQL = "SELECT * FROM tblUsers where UserType = 'Admin'  AND userArchived <> 1 order by userLastName"
		
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
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=Admins'><%= rs.Fields("userLastName")%>,&nbsp;<%= rs.Fields("userFirstName")%></a></td>
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=Admins'><%= rs.Fields("userEmail")%></a></td>					
						<td class="example"><%= rs.Fields("userType") %></td>
						<td>
						<% If IsNull(rs.Fields("userLastLogin")) Then %>
							Never
						<% Else %>
							<%= rs.Fields("userLastLogin") %>
						<% End If %>
						</td>
						<td>
						<div class="example">
						<% If rs.Fields("userEnabled") = True Then %>
							<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
						<% Else %>
							<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
						<% End If%>
						</div></td><td>
						<div class="example">
							<div>
								<a href='send_login_email.asp?userno=<%= rs.Fields("userNO") %>&tab=Admins' >Email</a>
								<% If Not IsNull(rs.Fields("userCellNumber")) Then %>
									<% If len(trim((rs.Fields("userCellNumber")))) > 0 Then %>
										&nbsp;&nbsp;<a href='send_login_text.asp?userno=<%= rs.Fields("userNO") %>&tab=Admins' >Text</a>							
									<% End If %>
								<% End If %>
							</div>
						</div></td>
						<td>
							<div class="example">
									    <a href='../../Reports\AuditTrail_OneUser.asp?unam=<%=rs.Fields("UserDisplayName")%>&tab=Admins' target='_blank'>
									    <img src='../../img/AuditTrailImage.png' height="35" title='View audit trail for <%=rs.Fields("UserDisplayName")%>' </a>
							</div>
						</td>
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
<!-- eof Admins Tab !-->
				

				
<!--  Inside Sales Tab !-->
<div role="tabpanel" class="tab-pane fade" id="InsideSales"> 
					
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th>User Name</th>
                  <th>Email Address</th>
                  <th>User Type</th>
                  <th>Last Login</th>
                  <th class="sorttable_nosort"><%=GetTerm("Inside Sales")%></th>
                  <th class="sorttable_nosort"><%=GetTerm("Inside Sales Manager")%></th>
                  <th class="sorttable_nosort">Enabled</th>
                  <th class="sorttable_nosort">Send<br>Credentials</th>
                  <th class="sorttable_nosort">Audit<br>Trail</th>
                </tr>
              </thead>
              <tbody>
              
				<%
			
				SQL = "SELECT * FROM tblUsers where (UserType = 'Inside Sales' or UserType = 'Inside Sales Manager') AND userArchived <> 1 order by userLastName"
		
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
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=InsideSales'><%= rs.Fields("userLastName")%>,&nbsp;<%= rs.Fields("userFirstName")%></a></td>
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=InsideSales'><%= rs.Fields("userEmail")%></a></td>					
						<td class="example"><%= rs.Fields("userType") %></td>
						<td>
						<% If IsNull(rs.Fields("userLastLogin")) Then %>
							Never
						<% Else %>
							<%= rs.Fields("userLastLogin") %>
						<% End If %>
						</td>
						<td>
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "Inside Sales" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
	
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "Inside Sales Manager" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
						<td>
						<div class="example">
						<% If rs.Fields("userEnabled") = True Then %>
							<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
						<% Else %>
							<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
						<% End If%>
						</div></td><td>
						<div class="example">
							<div>
								<a href='send_login_email.asp?userno=<%= rs.Fields("userNO") %>&tab=InsideSales' >Email</a>
								<% If Not IsNull(rs.Fields("userCellNumber")) Then %>
									<% If len(trim((rs.Fields("userCellNumber")))) > 0 Then %>
										&nbsp;&nbsp;<a href='send_login_text.asp?userno=<%= rs.Fields("userNO") %>&tab=InsideSales' >Text</a>							
									<% End If %>
								<% End If %>
							</div>
						</div></td>
						<td>
							<div class="example">
									    <a href='../../Reports\AuditTrail_OneUser.asp?unam=<%=rs.Fields("UserDisplayName")%>&tab=InsideSales' target='_blank'>
									    <img src='../../img/AuditTrailImage.png' height="35" title='View audit trail for <%=rs.Fields("UserDisplayName")%>' </a>
							</div>
						</td>
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
<!-- eof Inside Sales Tab !-->

				
<!--  Outside Sales Tab !-->
<div role="tabpanel" class="tab-pane fade" id="OutsideSales"> 
					
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th>User Name</th>
                  <th>Email Address</th>
                  <th>User Type</th>
                  <th>Last Login</th>
                  <th class="sorttable_nosort"><%=GetTerm("Outside Sales")%></th>
                  <th class="sorttable_nosort"><%=GetTerm("Outside Sales Manager")%></th>
                  <th class="sorttable_nosort">Enabled</th>
                  <th class="sorttable_nosort">Send<br>Credentials</th>
                  <th class="sorttable_nosort">Audit<br>Trail</th>
                </tr>
              </thead>
              <tbody>
              
				<%
			
				SQL = "SELECT * FROM tblUsers where (UserType = 'Outside Sales' or UserType = 'Outside Sales Manager')  AND userArchived <> 1 order by userLastName"
		
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
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=OutsideSales'><%= rs.Fields("userLastName")%>,&nbsp;<%= rs.Fields("userFirstName")%></a></td>
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=OutsideSales'><%= rs.Fields("userEmail")%></a></td>					
						<td class="example"><%= rs.Fields("userType") %></td>
						<td>
						<% If IsNull(rs.Fields("userLastLogin")) Then %>
							Never
						<% Else %>
							<%= rs.Fields("userLastLogin") %>
						<% End If %>
						</td>
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "Outside Sales" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
	
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "Outside Sales Manager" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
	
						<td>
						
						<td>
						<div class="example">
						<% If rs.Fields("userEnabled") = True Then %>
							<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
						<% Else %>
							<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
						<% End If%>
						</div></td><td>
						<div class="example">
							<div>
								<a href='send_login_email.asp?userno=<%= rs.Fields("userNO") %>&tab=OutsideSales' >Email</a>
								<% If Not IsNull(rs.Fields("userCellNumber")) Then %>
									<% If len(trim((rs.Fields("userCellNumber")))) > 0 Then %>
										&nbsp;&nbsp;<a href='send_login_text.asp?userno=<%= rs.Fields("userNO") %>&tab=OutsideSales' >Text</a>							
									<% End If %>
								<% End If %>
							</div>
						</div></td>
						<td>
							<div class="example">
									    <a href='../../Reports\AuditTrail_OneUser.asp?unam=<%=rs.Fields("UserDisplayName")%>&tab=OutsideSales' target='_blank'>
									    <img src='../../img/AuditTrailImage.png' height="35" title='View audit trail for <%=rs.Fields("UserDisplayName")%>' </a>
							</div>
						</td>
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
<!-- eof Outside Sales Tab !-->


				
<!--  Telemarketing Tab !-->
<div role="tabpanel" class="tab-pane fade" id="Telemarketing"> 
					
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th>User Name</th>
                  <th>Email Address</th>
                  <th>User Type</th>
                  <th>Last Login</th>
                  <th class="sorttable_nosort">Enabled</th>
                  <th class="sorttable_nosort">Send<br>Credentials</th>
                  <th class="sorttable_nosort">Audit<br>Trail</th>
                </tr>
              </thead>
              <tbody>
              
				<%
			
				SQL = "SELECT * FROM tblUsers where UserType = 'Telemarketing'  AND userArchived <> 1 order by userLastName"
		
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
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=Telemarketing'><%= rs.Fields("userLastName")%>,&nbsp;<%= rs.Fields("userFirstName")%></a></td>
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=Telemarketing'><%= rs.Fields("userEmail")%></a></td>					
						<td class="example"><%= rs.Fields("userType") %></td>
						<td>
						<% If IsNull(rs.Fields("userLastLogin")) Then %>
							Never
						<% Else %>
							<%= rs.Fields("userLastLogin") %>
						<% End If %>
						</td>
						<td>
						<div class="example">
						<% If rs.Fields("userEnabled") = True Then %>
							<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
						<% Else %>
							<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
						<% End If%>
						</div></td><td>
						<div class="example">
							<div>
								<a href='send_login_email.asp?userno=<%= rs.Fields("userNO") %>&tab=Telemarketing' >Email</a>
								<% If Not IsNull(rs.Fields("userCellNumber")) Then %>
									<% If len(trim((rs.Fields("userCellNumber")))) > 0 Then %>
										&nbsp;&nbsp;<a href='send_login_text.asp?userno=<%= rs.Fields("userNO") %>&tab=Telemarketing' >Text</a>							
									<% End If %>
								<% End If %>
							</div>
						</div></td>
						<td>
							<div class="example">
									    <a href='../../Reports\AuditTrail_OneUser.asp?unam=<%=rs.Fields("UserDisplayName")%>&tab=Telemarketing' target='_blank'>
									    <img src='../../img/AuditTrailImage.png' height="35" title='View audit trail for <%=rs.Fields("UserDisplayName")%>' </a>
							</div>
						</td>
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
<!-- eof Telemarketing Tab !-->

				
<!--  Drivers Tab !-->
<div role="tabpanel" class="tab-pane fade" id="Drivers"> 
					
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th>User Name</th>
                  <th>Email Address</th>
                  <th>User Type</th>
                  <th>Last Login</th>
                  <th class="sorttable_nosort"><%=GetTerm("Driver")%></th>
                  <th class="sorttable_nosort"><%=GetTerm("Route Manager")%></th>
                  <th class="sorttable_nosort">Enabled</th>
                  <th class="sorttable_nosort">Send<br>Credentials</th>
                  <th class="sorttable_nosort">Audit<br>Trail</th>
                </tr>
              </thead>
              <tbody>
              
				<%
			
				SQL = "SELECT * FROM tblUsers where (UserType = 'Driver' OR UserType='Route Manager') AND userArchived <> 1 order by userLastName"
		
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
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=Drivers'><%= rs.Fields("userLastName")%>,&nbsp;<%= rs.Fields("userFirstName")%></a></td>
						<td><a href='edituser.asp?uno=<%= rs.Fields("userNo")%>&tab=Drivers'><%= rs.Fields("userEmail")%></a></td>					
						<td class="example"><%= rs.Fields("userType") %></td>
						<td>
						<% If IsNull(rs.Fields("userLastLogin")) Then %>
							Never
						<% Else %>
							<%= rs.Fields("userLastLogin") %>
						<% End If %>
						</td>
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "Driver" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
	
						<td>
							<div class="example">
							<% If rs.Fields("userType") = "Route Manager" Then %>
								<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
							<% Else %>
								<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
							<% End If%>
							</div>
						</td>
						
						<td>
						<div class="example">
						<% If rs.Fields("userEnabled") = True Then %>
							<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
						<% Else %>
							<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
						<% End If%>
						</div></td><td>
						<div class="example">
							<div>
								<a href='send_login_email.asp?userno=<%= rs.Fields("userNO") %>&tab=Drivers' >Email</a>
								<% If Not IsNull(rs.Fields("userCellNumber")) Then %>
									<% If len(trim((rs.Fields("userCellNumber")))) > 0 Then %>
										&nbsp;&nbsp;<a href='send_login_text.asp?userno=<%= rs.Fields("userNO") %>&tab=Drivers' >Text</a>							
									<% End If %>
								<% End If %>
							</div>
						</div></td>
						<td>
							<div class="example">
									    <a href='../../Reports\AuditTrail_OneUser.asp?unam=<%=rs.Fields("UserDisplayName")%>&tab=Drivers' target='_blank'>
									    <img src='../../img/AuditTrailImage.png' height="35" title='View audit trail for <%=rs.Fields("UserDisplayName")%>'></a>
							</div>
						</td>
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
<!-- eof Drivers Tab !-->

				
<!-- Archived Tab !-->
<div role="tabpanel" class="tab-pane fade" id="Archived"> 
					
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th>User Name</th>
                  <th>Email Address</th>
                  <th>User Type</th>
                  <th>Last Login</th>
                  <th class="sorttable_nosort">Enabled</th>
                  <th class="sorttable_nosort">Audit<br>Trail</th>
                  <th class="sorttable_nosort">Reactivate</th>
                </tr>
              </thead>
              <tbody>
              
				<%
			
				SQL = "SELECT * FROM tblUsers WHERE userArchived = 1 Order By userLastName"
		
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
						<td><%= rs.Fields("userLastName")%>,&nbsp;<%= rs.Fields("userFirstName")%></td>
						<td><%= rs.Fields("userEmail")%></td>					
						<td class="example"><%= rs.Fields("userType") %></td>
						<td>
						<% If IsNull(rs.Fields("userLastLogin")) Then %>
							Never
						<% Else %>
							<%= rs.Fields("userLastLogin") %>
						<% End If %>
						</td>
						<td>
						<div class="example">
						<% If rs.Fields("userEnabled") = True Then %>
							<input type="checkbox" checked data-toggle="toggle" data-size="mini" disabled >
						<% Else %>
							<input type="checkbox" data-toggle="toggle" data-size="mini" disabled >					
						<% End If%>
						</div></td>
						<td>
							<div class="example">
									    <a href='../../Reports\AuditTrail_OneUser.asp?unam=<%=rs.Fields("UserDisplayName")%>&tab=Archived' target='_blank'>
									    <img src='../../img/AuditTrailImage.png' height="35" title='View audit trail for <%=rs.Fields("UserDisplayName")%>' </a>
							</div>
						</td>
						<td>
							<a href='reactivateUserQues.asp?un=<%=rs.Fields("UserNo")%>&tab=Archived'><i class='fa fa-undo' ></i></a>
						</td>
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
<!-- eof Archived  Tab !-->

				
				
			</div>
			<!-- eof tabs content !-->
		
		
		</div>
	</div>
	<!-- tabs end here !-->
    

</div>
<!-- eof row !-->    
</div>

<!--#include file="../../inc/footer-main.asp"-->