<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->



 
 <style type="text/css">
 
	table {
	    vertical-align: middle;
	    display: table-cell;
	}
	
	tr {
	    vertical-align: middle;
	}
	
 	.email-table{
		width:46%;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
	}
	
	
	.container{
		max-width:1100px;
		margin:0 auto;
	}

 </style>

<!--- eof on/off scripts !-->

<!-- <h1 class="page-header"><i class="fa fa-map-marker" aria-hidden="true"></i> Customer Mapping Tool - Account By First Letter Selection</h1> -->
<h1 class="page-header"><i class="fa fa-map-marker" aria-hidden="true"></i>Customer Mapping Table</h1>


	
	<div class="container">
	  <h2><i class="fa fa-bolt" aria-hidden="true"></i> Letter Quick Select</h2>
	  <p>Select the letter to list/map all customer account names that begin with that letter:</p>
	  <div class="btn-toolbar" style="width:1200px; margin-bottom:20px; margin-top:20px;">
	    <div class="btn-group btn-group-lg">
	      <% for i = asc("A") to asc("Z") %>
	      	<a href="customerAccountMappingScreen.asp?letter=<%= chr(i) %>&i=<%= Request.QueryString("i") %>"><button class="btn btn-default"><%= chr(i) %></button></a>
	      <% next %>
	    </div>
	  </div>
	</div>
	
	
	
	
	<div class="container">
	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover table-bordered sortable">
              <thead>
                <tr>
                  <th>Account Name Begins With</th>
                  <th>Total Accounts</th>
                  <th>Partner Accounts Defined</th>
                  <th class="sorttable_nosort">Manage/Map Accounts</th>
                </tr>
              </thead>
              <tbody class='searchable'>
              
				<%
				
				InternalRecordIdentifier = Request.QueryString("i")

				SQL9 = "SELECT COUNT(CustNum) as TotalCustomerCount FROM AR_Customer WHERE AcctStatus = 'A'" 
				Set cnn9 = Server.CreateObject("ADODB.Connection")
				cnn9.open (Session("ClientCnnString"))
				Set rs9 = Server.CreateObject("ADODB.Recordset")
				rs9.CursorLocation = 3 
				Set rs9 = cnn9.Execute(SQL9)
				If not rs9.EOF Then
					TotalCustomerCount = rs9("TotalCustomerCount")
				Else
					TotalCustomerCount = 0
				End If
				
				SQL9 = "SELECT COUNT(partnerCustID) as TotalEquivalentPartnerCustCount FROM AR_CustomerMapping WHERE partnerRecID = " & InternalRecordIdentifier
				Set rs9 = cnn9.Execute(SQL9)
				If not rs9.EOF Then
					TotalEquivalentPartnerCustCount = rs9("TotalEquivalentPartnerCustCount")
				Else
					TotalEquivalentPartnerCustCount = 0
				End If
				set rs9 = Nothing
				cnn9.close
				set cnn9 = Nothing
						

				%>              
				<tr>
					<td align="center" style="padding-top:30px;padding-bottom:30px;">
					<a href="customerAccountMappingScreen.asp?letter=all&i=<%= InternalRecordIdentifier %>">ALL CUSTOMERS</a></td>
					<td align="center" style="padding-top:30px;padding-bottom:30px;">
					<a href="customerAccountMappingScreen.asp?letter=all&i=<%= InternalRecordIdentifier %>"><%= TotalCustomerCount %></a></td>
					<td align="center" style="padding-top:30px;padding-bottom:30px;">
					<a href="customerAccountMappingScreen.asp?letter=all&i=<%= InternalRecordIdentifier %>"><%= TotalEquivalentPartnerCustCount %></a></td>	
					<td align="center" style="padding-top:20px;padding-bottom:20px;">
					<a href="customerAccountMappingScreen.asp?letter=all&i=<%= InternalRecordIdentifier %>"><button type="button" class="btn btn-primary">MAP ALL CUSTOMER ACCOUNTS <i class="fa fa-arrow-circle-o-right" aria-hidden="true"></i></button></a></td>
			   	</tr>
             
				<% for i = asc("A") to asc("Z") %>	
			
				<%
					
					SQL9 = "SELECT COUNT(CustNum) as TotalCustomerCount FROM AR_Customer WHERE LEFT(Name,1) = '" & chr(i) & "' AND AcctStatus = 'A'"
					Set cnn9 = Server.CreateObject("ADODB.Connection")
					cnn9.open (Session("ClientCnnString"))
					Set rs9 = Server.CreateObject("ADODB.Recordset")
					rs9.CursorLocation = 3 
					Set rs9 = cnn9.Execute(SQL9)
					If not rs9.EOF Then
						TotalCustomerCount = rs9("TotalCustomerCount")
					Else
						TotalCustomerCount = 0
					End If
					
					SQL9 = "SELECT COUNT(partnerCustID) as TotalEquivalentPartnerCustCount, "
					SQL9 = SQL9 & " Count(ourCustID) as OurCustCOunt FROM AR_CustomerMapping where ourCustID in "
					SQL9 = SQL9 & " (select custnum from  AR_Customer "
					SQL9 = SQL9 & " WHERE LEFT(AR_Customer.Name,1) = '" & chr(i) & "' AND AR_Customer.AcctStatus = 'A') "
										
					Set rs9 = cnn9.Execute(SQL9)
					'Response.Write(SQL9)
					If not rs9.EOF Then
						TotalEquivalentPartnerCustCount = rs9("TotalEquivalentPartnerCustCount")
					Else
						TotalEquivalentPartnerCustCount = 0
					End If
						
			        %>
					<!-- table line !-->
					<tr>
						<td align="center">
						<a href="customerAccountMappingScreen.asp?letter=<%= chr(i) %>&i=<%= InternalRecordIdentifier %>"><h2><%= chr(i) %></h2></a></td>
						<td align="center" style="padding-top:30px;">
						<a href="customerAccountMappingScreen.asp?letter=<%= chr(i) %>&i=<%= InternalRecordIdentifier %>"><%= TotalCustomerCount %></a></td>
						<td align="center" style="padding-top:30px;">
						<a href="customerAccountMappingScreen.asp?letter=<%= chr(i) %>&i=<%= InternalRecordIdentifier %>"><%= TotalEquivalentPartnerCustCount %></a></td>	
						<td align="center" style="padding-top:20px;">
						<a href="customerAccountMappingScreen.asp?letter=<%= chr(i) %>&i=<%= InternalRecordIdentifier %>"><button type="button" class="btn btn-success">MAP "<%= chr(i) %>" Accounts  <i class="fa fa-arrow-circle-o-right" aria-hidden="true"></i></button></a></td>
				   	</tr>
					<%
				Next

				set rs9 = Nothing
				cnn9.close
				set cnn9 = Nothing
				
	            %>
			</tbody>
		</table>
	</div>
 
		</div>

</div>
<!-- eof row !-->
								

<!--#include file="../../inc/footer-main.asp"-->