<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Routing.asp"-->

<%
TruckNum = Request.Form("txtTruckNumber")
DateToRetrieve = Request.Form("txtDriverDate")
CreateAuditLogEntry GetTerm("Routing") & " Report",GetTerm("Routing") & " Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Historical Deliveries For Driver: " & GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(TruckNum))) & " for date " & DateToRetrieve 
%>

<style>

	.table-responsive {
		overflow-x:hidden;
	}
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
	    content: " \25B4\25BE" 
	}
	table.sorttable thead {
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
		max-width:925px;
		margin:0 auto;
		
	}
		
	.page-header{
		width:100%;
		text-align:center;
	}
	
	.sorttable-right{
		text-align:right;
	}
	
	.sorttable-center{
		text-align:center;
	}
	
	.sorttable-nosort{
	    content: "";
	    background: none !important;
	}
	
	.delivered{
	    color:green;
	}
	.notdelivered{
	    color:red;
	}
	
</style>

<script type="text/javascript">
	$(document).ready(function() {
	    $("#PleaseWaitPanel").hide();
	});
</script>

<h3 class="page-header"><i class="fa fa-file-text-o"></i> Historical Deliveries for <%= GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(TruckNum))) %> on <%= DateToRetrieve  %>&nbsp;&nbsp;
	<a href="<%= BaseURL %>routing/reports/HistoricalDeliveriesByDriver.asp?date=<%= DateToRetrieve %>"><button type="button" class="btn btn-primary">Back To Driver List</button></a></h3>

<div class="row">
<div class="container-center">

	<div class="table-responsive">


    	<table id="tableDeliveries" class="food_planner sortable table table-striped table-condensed table-hover">
	   		<thead>
				<tr>
					<% If DelBoardDontUseStopSequencing() = False Then %>
						<th class="sortable sortable-center" width="10%">Sequence #</th>
					<% End If %>
					<th class="sortable" width="10%"><%=GetTerm("Customer")%> #</th>
                    <th class="sortable" width="23%">Name</th>
					<th class="sortable sortable-center" width="15%">Invoice #</th>
					<th class="sortable sortable-right" width="15%">Value</th>
					<th class="sortable sorttable_nosort" width="15%">Status</th>
					<th class="sortable sorttable_nosort" width="15%">Time</th>
				</tr>
			</thead>

			<tbody class="searchable">
					
				<%
	
				Set cnn9 = Server.CreateObject("ADODB.Connection")
				cnn9.open (Session("ClientCnnString"))
				Set rsDeliveries = Server.CreateObject("ADODB.Recordset")
				rsDeliveries.CursorLocation = 3 
	
				SQL9 = SQL9 & "SELECT * FROM RT_DeliveryBoardHistory "
				SQL9 = SQL9 & " WHERE (Year(LastDeliveryStatusChange) = " & Year(DateToRetrieve) & " AND "
				SQL9 = SQL9 & " Month(LastDeliveryStatusChange) = " & Month(DateToRetrieve) & " AND "
				SQL9 = SQL9 & " Day(LastDeliveryStatusChange) = " & Day(DateToRetrieve) & ") AND "
				SQL9 = SQL9 & " TruckNumber = '" & TruckNum & "' " 
				SQL9 = SQL9 & " ORDER BY LastDeliveryStatusChange"
				
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
					Response.Write("<td class='sortable-center'>" & rsDeliveries("CustNum") & "</td>")
					Response.Write("<td class='sortable-center'>" &  rsDeliveries("CustName") &  "</td>")
					Response.Write("<td class='sortable-center'>" & rsDeliveries("IvsNum") & "</td>")
					Response.Write("<td class='sortable-right'>" & FormatCurrency(rsDeliveries("Value")) & "</td>")
					If rsDeliveries("DeliveryStatus") = "Delivered" Then
						Response.Write("<td class='sorttable-nosort'><span class='alert-success'>" & rsDeliveries("DeliveryStatus") & "</span></td>")
					ElseIf rsDeliveries("DeliveryStatus") = "No Delivery" Then
						Response.Write("<td class='sorttable-nosort'><span class='alert-danger'>" & rsDeliveries("DeliveryStatus") & "</span></td>")
					Else
						Response.Write("<td class='sorttable-nosort'>" & rsDeliveries("DeliveryStatus") & "</td>")					
					End If
					Response.Write("<td class='sorttable_nosort'>" & FormatDateTime(rsDeliveries("LastDeliveryStatusChange"),3) & "</td>")

						
					Response.Write("</tr>")
					
					rsDeliveries.Movenext
				Loop
				
				Set rsDeliveries = Nothing
				cnn9.Close
				Set cnn9 = Nothing
				
				%>
			</tbody>
	    </table>
        
        
        <table id="tableDeliveries" class="table table-striped table-condensed table-hover">
        	<tfoot>
            	<tr>
                	
                    <td width="30%"><strong>Totals</strong></td>
                    <td width="10%" align="left"><strong><%=GetNumberOfCustomersByTruckNumberHistorical(TruckNum,DateToRetrieve) %> Stops</strong></td>
                    <td width="10%" align="right"><strong><%=GetNumberOfInvoicesByTruckNumberHistorical(TruckNum,DateToRetrieve) %> Invoices</strong></td>
                    <td width="10%" align="right"><strong><%=FormatCurrency(GetValueOfDeliveriesByTruckNumberHistorical(TruckNum,DateToRetrieve)) %></strong></td>
                    <td width="10%" align="right">&nbsp;</td>
                    <td width="10%" align="right">&nbsp;</td>
                    <td width="10%" align="right">&nbsp;</td>
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