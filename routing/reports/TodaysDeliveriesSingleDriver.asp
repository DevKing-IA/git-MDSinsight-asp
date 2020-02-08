<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Routing.asp"-->
<%
TruckNum = Request.QuerySTring("trk")
CreateAuditLogEntry GetTerm("Routing") & " Report",GetTerm("Routing") & " Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Todays Deliveries For Driver: " & GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(TruckNum))) 
%>

<style>

	.table-responsive {
		overflow-x:hidden;
	}

	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
	    content: " \25B4\25BE" 
	}
	table.sortable thead {
	    color:#222;
	    font-weight: bold;
	    cursor: pointer;
	}
	
	#PleaseWaitPanel{
	position: fixed;
	left: 470px;
	top: 275px;
	width: 975px;
	height: 300px;
	z-index: 9999;
	background-color: #fff;
	opacity:1.0;
	text-align:center;
	}   
	
	.container-center{
		max-width:800px;
		margin:0 auto;
	} 
	
	.page-header{
		width:100%;
		text-align:center;
	}
	
	.sortable-right{
		text-align:right;
	}
	
	.sortable-center{
		text-align:center;
	}
</style>

<script type="text/javascript">
	$(document).ready(function() {
	    $("#PleaseWaitPanel").hide();
	});
</script>


<h3 class="page-header"><i class="fa fa-file-text-o"></i> Today's Deliveries for <%= GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(TruckNum))) %> &nbsp;&nbsp;
	<a href="<%= BaseURL %>routing/reports/TodaysDeliveriesByDriver.asp"><button type="button" class="btn btn-primary">Back To Driver List</button></a></h3>

<div class="row">
<div class="container-center">

	<div class="table-responsive">


    	<table id="tableDeliveries" class="food_planner sortable table table-striped table-condensed table-hover">
	   		<thead>
				<tr>
					<% If DelBoardDontUseStopSequencing() = False Then %>
						<th class="sortable sortable-right">Sequence #</th>
					<% End If %>
					<th class="sortable"><%=GetTerm("Customer")%> #</th>
                    <th class="sortable">Name</th>
					<th class="sortable sortable-right">Invoice #</th>
					<th class="sortable sortable-right">Value</th>
				</tr>
			</thead>

			<tbody class="searchable">
					
				<%
	
				Set cnn9 = Server.CreateObject("ADODB.Connection")
				cnn9.open (Session("ClientCnnString"))
				Set rsDeliveries = Server.CreateObject("ADODB.Recordset")
				rsDeliveries.CursorLocation = 3 
	
				SQL9 = "SELECT * FROM RT_DeliveryBoard WHERE TruckNumber = '" & TruckNum  & "' ORDER BY SequenceNumber"

				Set rsDeliveries = cnn9.Execute(SQL9)
					
				GrandTot_Stops = 0
				GrandTot_Invoices = 0
				GrandTot_Value = 0
																	
				Do While not rsDeliveries.Eof

					If rsDeliveries("AMorPM") = "AM" Then
						Response.Write("<tr style='border: 2px solid red;'>")
					Else
						Response.Write("<tr>")
					End If
						
					If DelBoardDontUseStopSequencing() = False Then						
						Response.Write("<td class='sortable-center'>" & rsDeliveries("SequenceNumber") & "</td>")
					End If
					Response.Write("<td>" & rsDeliveries("CustNum") & "</td>")
					Response.Write("<td>" & GetCustNameByCustNum(rsDeliveries("CustNum")) &  "</td>")
					Response.Write("<td class='sortable-center'>" & rsDeliveries("IvsNum") & "</td>")
					Response.Write("<td class='sortable-right'>" & FormatCurrency(rsDeliveries("Value")) & "</td>")
						
					Response.Write("</tr>")
					
					rsDeliveries.Movenext
				Loop
				
				Set rsDeliveries = Nothing
				cnn9.Close
				Set cnn9 = Nothing
				
				'Print the totals
				'Response.Write("<tr>")
				
				'Response.Write("<td><strong>Totals</strong></td>")	
				'Response.Write("<td><strong>" & GetNumberOfCustomersByTruckNumber(TruckNum) & " Stops</strong></td>")
				'Response.Write("<td><strong>" & GetNumberOfInvoicesByTruckNumber(TruckNum) & " Invoices</strong></td>")
				'Response.Write("<td><strong>" & FormatCurrency(GetValueOfDeliveriesByTruckNumber(TruckNum)) & "</strong></td>")
				'Response.Write("<td> &nbsp; </td>")
						
				'Response.Write("</tr>")

				%>
			</tbody>
	    </table>
        
        
        <table id="tableDeliveries" class="table table-striped table-condensed table-hover">
        	<tfoot>
            	<tr>
                	
                    <td width="30%"><strong>Totals</strong></td>
                    <td width="32%" align="left"><strong><%=GetNumberOfCustomersByTruckNumber(TruckNum) %> Stops</strong></td>
                    <td width="25%" align="right"><strong><%=GetNumberOfInvoicesByTruckNumber(TruckNum) %> Invoices</strong></td>
                    <td width="25%" align="right"><strong><%=FormatCurrency(GetValueOfDeliveriesByTruckNumber(TruckNum)) %></strong></td>
                </tr>
            </tfoot>
        </table>
        
        
	</div>
    </div>
</div>

<!-- row !-->
<div class="row">
	<div class="col-lg-12"><hr></div>
</div>
<!-- eof row !-->

<!-- row !-->
<div class="row">
</div>
<!-- eof row !-->

<!--#include file="../../inc/footer-main.asp"-->