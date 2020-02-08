<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/Insightfuncs_Prospecting.asp"-->

<%
'************************
'Read Settings_Screens
'************************
SQL = "SELECT * FROM Settings_Screens WHERE ScreenNumber = 1100 AND UserNo = " & Session("userNo")
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rsInsight2 = Server.CreateObject("ADODB.Recordset")
rsInsight2.CursorLocation = 3 
Set rsInsight2= cnn8.Execute(SQL)
If NOT rsInsight2.EOF Then
	SelectedUserNumbers = rsInsight2("ScreenSpecificData1")
Else	
	SelectedUserNumbers = ""
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

<h1 class="page-header"><i class="fa fa-users"></i> Select Sales Reps To Display in Dashboard Reports</h1>

	<form method="POST" action="dashboard_salesrep_display_setup_submit.asp" name="frmDashboardSalesReps">		    
	
	<div class="row">

	<div class="row genline-up">
		
 		<div class="col-lg-2 alertbutton">
			<a href="dashboard_salesrep_display_setup.asp">
	    		<button type="button" class="btn btn-default btn-cancel-save">&lsaquo; Cancel </button>
			</a>
			<button type="submit" class="btn btn-primary btn-cancel-save"><i class="far fa-save"></i> Save</button>
 		</div>
			
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
                  <th class="sorttable_nosort">Include As Sales Rep In Dashboard</th>
                </tr>
              </thead>
              
              <tbody>
              
			<%
			SQL = "SELECT userNo,userFirstName,userLastName,userEmail,userDisplayName FROM tblUsers"
			SQL = SQL & " WHERE userType = 'Admin' OR userType = 'CSR'  OR userType = 'CSR Manager' OR userType = 'Inside Sales' OR userType = 'Inside Sales Manager'"
			SQL = SQL & " OR userType = 'Outside Sales' OR userType = 'Outside Sales Manager' OR userType = 'Finance' OR userType = 'Telemarketing'"
			SQL = SQL & " AND userArchived <> 1 ORDER BY userFirstName"
	
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

					<td><%= rs("userFirstName")%>&nbsp;<%= rs("userLastName")%></td>
					<td><%= rs("userEmail")%></td>					
					<td>
					<div class="example">
					<%
					'Setup var to find this user number
					ItemToFind = "," & rs("userNo") & ","
					
					 If Instr(SelectedUserNumbers,ItemToFind) <> 0 Then %>
						<input type="checkbox" checked data-toggle="toggle" data-size="mini" name="chk<%=rs("userNo")%>" id="chk<%=rs("userNo")%>">
					<% Else %>
						<input type="checkbox" data-toggle="toggle" data-size="mini" name="chk<%=rs("userNo")%>" id="chk<%=rs("userNo")%>">
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
						<a href="dashboard_salesrep_display_setup.asp">
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

<!--#include file="../../inc/footer-main.asp"-->