<!--#include file="../inc/header.asp"-->




<%
'************************
'Read Settings_Screens
'************************
SQL = "SELECT * from Settings_Screens where ScreenNumber = 1000 AND UserNo = " & Session("userNo")
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rsInsight2 = Server.CreateObject("ADODB.Recordset")
rsInsight2.CursorLocation = 3 
Set rsInsight2= cnn8.Execute(SQL)
If NOT rsInsight2.EOF Then
	SelectedUserDisplayNames = rsInsight2("ScreenSpecificData2")
	NumberOfDays = cint(rsInsight2("ScreenSpecificData1"))
	FservCloseOnly = rsInsight2("ScreenSpecificData3")
	If FservCloseOnly <> "1" AND FservCloseOnly <> "0" then FservCloseOnly = "0"
Else
	' We must do an insert here with initial default values - all users - 10 days
	If Session("AdminPrivelages") = True Then
		SQL = "Select userDisplayName from tblUsers WHERE userArchived <> 1 order by userDisplayName"
	Elseif userIsServiceManager(Session("userno")) Then
		SQL = "SELECT userDisplayName FROM tblUsers where userType = 'Field Service' and userArchived <> 1 order by userDisplayName"			
	Elseif  userIsCSRManager(Session("userno")) Then 
		SQL = "SELECT userDisplayName FROM tblUsers where userType = 'CSR' and userArchived <> 1 order by userDisplayName"
	Elseif  userIsFinanceManager(Session("userno")) Then 
		SQL = "SELECT userDisplayName FROM tblUsers where userType = 'Finance' and userArchived <> 1 order by userDisplayName"
	End If 
'	response.write(SQL)
	Set rsInsight1 = Server.CreateObject("ADODB.Recordset")
	rsInsight1.CursorLocation = 3 
	Set rsInsight1 = cnn8.Execute(SQL)
	If not rsInsight1.Eof Then
		UserList=","
		Do
			UserList = UserList & rsInsight1("userDisplayName") & "," ' Dont strip the trailing comma, we use it later
			rsInsight1.movenext
		Loop until rsInsight1.eof
	End IF
	SQL = "Insert Into Settings_Screens (ScreenNumber,UserNo,ScreenSpecificData1,ScreenSpecificData2,ScreenSpecificData3) Values "
	SQL = SQL & "(1000," & Session("UserNo") & ",'10','" & UserList & "','0')"
	Set rsInsight1 = Server.CreateObject("ADODB.Recordset")
	rsInsight1.CursorLocation = 3 
	Set rsInsight1 = cnn8.Execute(SQL)
	SelectedUserDisplayNames = UserList 
	NumberOfDays = 10
	chkFservCloseOnly = 0
End If
'****************************
'End Read Settings_Screens
'****************************
%>

<!-- on/off scripts !-->
 
 <style type="text/css">
 	.email-table{
		width:46%;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
}

.genline-up{
	margin-bottom: 20px;
}

.btn-cancel-save{
	display: inline-block;
	width: 44%;
} 

.form-chart{
	display: inline-block;
	width:auto;
	margin-left: 10px;
}
 </style>

<!--- eof on/off scripts !-->

<h1 class="page-header"><i class="fa fa-users"></i> User Activity Chart Settings</h1>

	<form method="POST" action="user_activity_chart_setup_submit.asp" name="frmactivitychart">		    
	
	<div class="row">

	<div class="row genline-up">
		
 		<div class="col-lg-2 alertbutton">
			<a href="<%= BaseURL %>main/default.asp">
	    		<button type="button" class="btn btn-default btn-cancel-save">&lsaquo; Cancel </button>
			</a>
			<button type="submit" class="btn btn-primary btn-cancel-save"><i class="far fa-save"></i> Save</button>
 		</div>
 		   
			<!-- drop down !-->
			<div class="col-lg-4">
			<label>Number of days to chart</label>
				<select class="form-control form-chart" name="selNumberOfDays" id="selNumberOfDays">
					<% For x = 1 to 30
					IF x = NumberOfDays Then 
						Response.Write("<option selected>" & x & "</option>")
					Else
						Response.Write("<option>" & x & "</option>")
					End If 
					Next %>
				</select>
			</div>
			<!-- eof drop down !-->
			

			<%
			If Session("AdminPrivelages") = True or userIsServiceManager(Session("userno")) Then
				Response.Write("<div>")
				Response.Write("<input type='checkbox' class='check' id='chkFservCloseOnly' name='chkFservCloseOnly'")
				If FservCloseOnly = "1" Then Response.Write(" checked ")
				Response.Write("> Only chart closed tickets for field techs")
				Response.Write("</div>")
			End If
			%>

			
			
	</div>
</div>
	
<!-- row !-->
<div class="row">
 	<div class="col-lg-12">
 	
    

 
    	
        <div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th>User Name</th>
                  <th>Email Address</th>
                  <th class="sorttable_nosort">Include In Chart</th>
                </tr>
              </thead>
              
              <tbody>
              
			<%
			If Session("AdminPrivelages") = True Then
				SQL = "SELECT userNo,userFirstName,userLastName,userEmail,userDisplayName FROM tblUsers WHERE userArchived <> 1 order by userFirstName"
			Elseif userIsServiceManager(Session("userno")) Then
				SQL = "SELECT userNo,userFirstName,userLastName,userEmail,userDisplayName FROM tblUsers where userType = 'Field Service' and userArchived <> 1 order by userFirstName"			
			Elseif  userIsCSRManager(Session("userno")) Then 
				SQL = "SELECT userNo,userFirstName,userLastName,userEmail,userDisplayName FROM tblUsers where userType = 'CSR' and userArchived <> 1 order by userFirstName"
			Elseif  userIsFinanceManager(Session("userno")) Then 
				SQL = "SELECT userNo,userFirstName,userLastName,userEmail,userDisplayName FROM tblUsers where userType = 'Finance' and userArchived <> 1 order by userFirstName"

			End If 
	
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

					<td><%= rs.Fields("userFirstName")%>&nbsp;<%= rs.Fields("userLastName")%></td>
					<td><%= rs.Fields("userEmail")%></td>					
					<td>
					<div class="example">
					<%'Setup var to find this display name
					ItemToFind = "," & rs.Fields("userDisplayName") & ","
					 If Instr(SelectedUserDisplayNames,ItemToFind) <> 0 Then %>
						<input type="checkbox" checked data-toggle="toggle" data-size="mini" name='chk<%=rs("userNo")%>' id='chk<%=rs("userNo")%>'>
					<% Else %>
						<input type="checkbox" data-toggle="toggle" data-size="mini" name='chk<%=rs("userNo")%>' id='chk<%=rs("userNo")%>'>
					<% End If%>
					</div>
					</td>
					</tr>
					<!-- eof table line !-->
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
			<div class="row">
				<div class="row genline-up">
					<div class="col-lg-12 alertbutton">
						<a href="<%= BaseURL %>main/default.asp">
				    		<button type="button" class="btn btn-default">&lsaquo; Cancel </button>
						</a>
						<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
					   </div>	
				</div>

   			</div>
    </div>	

</div>   
 </form>
<!-- eof row !-->    

<!--#include file="../inc/footer-main.asp"-->