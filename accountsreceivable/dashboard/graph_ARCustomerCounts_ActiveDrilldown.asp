<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InSightFuncs_BizIntel.asp"--> 
<%
Server.ScriptTimeout = 900000 'Default value

monthToAnalyze = Request.QueryString("m")
yearToAnalyze = Request.QueryString("y")

CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: New/Active Customer Accounts by Month"
%>

  
<style>
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
		content: " \25B4\25BE" 
	}
	
	table.sortable thead {
		color:#222;
		font-weight: bold;
		cursor: pointer;
	}
	
	.column-header{
		font-size: 1em;
		vertical-align: top !important;
		text-align: center;
		background: #3B579D;
		color:#fff;
	}	
		
	.page-header {
		padding-bottom:20px;
	}

</style>

<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>

<script type="text/javascript">

	$(document).ready(function() {
	
	    $('#tableSuperSum').DataTable({
	        scrollY: 500,
	        scrollCollapse: true,
	        paging: false,
	        order: [ 2, 'asc' ]
	    	}
	    );
	});
</script>



<h3 class="page-header"><i class="fa fa-dollar"></i> New Customer Accounts (Marked Active) for <%= monthToAnalyze %>/<%= yearToAnalyze %>

<a href="<%= BaseURL %>accountsreceivable/dashboard.asp" class="pull-right"><button type="button" class="btn btn-primary">Back To <%= GetTerm("Accounts Receivable") %> Dashboard</button></a>

</h3>


<!-- row !-->
<div class="row">

<div class="container-fluid">
    <div class="row">
           <table id="tableSuperSum" class="display  compact" style="width:100%;">
              <thead>
                  <tr>	
					<th class="sorttable numeric column-header"><br>Acct</th>
            		<th class="sorttable column-header"><br>Client</th>
            		<th class="sorttable numeric column-header"><br>Active/Install Date</th>
            		<th class="sorttable column-header"><br><%= GetTerm("Primary Salesman") %></th>
            		<th class="sorttable column-header"><br><%= GetTerm("Secondary Salesman") %></th>
            		<th class="sorttable column-header"><br>Referral Code</th>
				</tr>
              </thead>
              
		<tbody>	

		<%	
			
		Set cnnActiveARCustomer = Server.CreateObject("ADODB.Connection")
		cnnActiveARCustomer.open (Session("ClientCnnString"))
		Set rsActiveARCustomer = Server.CreateObject("ADODB.Recordset")
		rsActiveARCustomer.CursorLocation = 3
		
				
		SQLActiveARCustomer = "SELECT * FROM AR_Customer WHERE DATEPART(month, InstallDate) = '" & monthToAnalyze & "' AND "
		SQLActiveARCustomer = SQLActiveARCustomer & " DATEPART(year, InstallDate) = '" & yearToAnalyze & "' "
		
	
		Set rsActiveARCustomer = cnnActiveARCustomer.Execute(SQLActiveARCustomer)
	
		If Not rsActiveARCustomer.EOF Then
			
			Do While Not rsActiveARCustomer.EOF
			
				CustID = rsActiveARCustomer("CustNum")
				CustName = GetCustNameByCustNum(CustID)
				ActiveDate = rsActiveARCustomer("InstallDate")
				PrimarySalesman = rsActiveARCustomer("Salesman")
				SecondarySalesman = rsActiveARCustomer("SecondarySalesman")
				ReferralCode = rsActiveARCustomer("ReferalCode")
				ReferralName = GetReferralNameByCode(ReferralCode)
				PrimarySalesPerson = GetSalesmanNameBySlsmnSequence(PrimarySalesman)
				SecondarySalesPerson = GetSalesmanNameBySlsmnSequence(SecondarySalesman)
			
				%>
				<tr>
			    <td><a href="<%= BaseURL %>bizintel/tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=<%= CustID %>&ZDC=0&VB=3Periods&oon=new" target="_blank"><%= CustID %></a></td>
			    <td><a href="<%= BaseURL %>bizintel/tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=<%= CustID %>&ZDC=0&VB=3Periods&oon=new" target="_blank"><%= CustName %></a></td>
	   		    <td align="center"><%= FormatDateTime(ActiveDate,2) %></td>
			    
			    <% If Instr(PrimarySalesPerson ," ") <> 0 Then %>
					<td align="center"><%= Left(PrimarySalesPerson,Instr(PrimarySalesPerson ," ")+1) %></td>
				<% Else %>
					<td align="center"><%= PrimarySalesPerson %></td>
				<% End If %>
				
				<% If Instr(SecondarySalesPerson," ") <> 0 Then %>
					<td align="center"><%= Left(SecondarySalesPerson,Instr(SecondarySalesPerson," ")+1) %></td>
				<% Else %>
					<td align="center"><%= SecondarySalesPerson %></td>
				<% End If %>

				<td align="center"><%= ReferralName %></td>

                
			    </tr>
					    
		 		<%
			
			rsActiveARCustomer.movenext
					
			Loop
			
		End If
	
	
	cnnActiveARCustomer.Close
	Set cnnActiveARCustomer = Nothing
	Set rsActiveARCustomer = Nothing
	%>


	</tbody>
	</table>
	</div>


    </div>
         
          </div>
<!-- eof responsive tables !-->



<!-- eof row !-->

<!-- row !-->
<div class="row">
   <%
'Response.Write("<div class='col-lg-12'><h3>" & "Total Customers Listed:" & TotalCustsReported  & "</h3></div>")
%>

<div class="col-lg-12"><hr></div>
</div>
<!-- eof row !-->

<!-- eof row !-->
<!--#include file="../../inc/footer-main.asp"-->
