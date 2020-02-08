<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Routing.asp"-->
<%
CreateAuditLogEntry GetTerm("Routing") & " Report",GetTerm("Routing") & " Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Todays Deliveries By Driver"
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


<h3 class="page-header"><i class="fa fa-file-text-o"></i> Today's Deliveries By Driver for <%=FormatDateTime(Now(),1) %> &nbsp;&nbsp;
	<a href="<%= BaseURL %>routing/reports.asp"><button type="button" class="btn btn-primary">Back To <%= GetTerm("Routing") %> Reports List</button></a></h3>

<div class="row">
<div class="container-center">
	<div class="table-responsive">


    	<table id="tableDeliveries" class="sortable table table-striped table-condensed table-hover">
	   		<thead>
				<tr>
					<th width="40%">Driver</th> 
					<th class="sortable-center" width="20%"># Stops<br>(# <%=GetTerm("Customers")%>)</th>
					<th class="sortable-center" width="20%"># Invoices</th>
					<th class="sortable-right" width="20%">Value</th>
				</tr>
			</thead>

			<tbody>
					
				<%
	
				Set cnn9 = Server.CreateObject("ADODB.Connection")
				cnn9.open (Session("ClientCnnString"))
				Set rsDeliveries = Server.CreateObject("ADODB.Recordset")
				rsDeliveries.CursorLocation = 3 
	
				SQL9 = "SELECT Distinct TruckNumber FROM RT_DeliveryBoard ORDER BY TruckNumber"

				Set rsDeliveries = cnn9.Execute(SQL9)
					
				GrandTot_Stops = 0
				GrandTot_Invoices = 0
				GrandTot_Value = 0
																	
				Do While not rsDeliveries.Eof

					Response.Write("<tr>")
						
				    Response.write("<td><a href='TodaysDeliveriesSingleDriver.asp?trk=" & rsDeliveries("TruckNumber") &  "'>" & GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(rsDeliveries("TruckNumber"))))  & "</a></td>")
					Response.Write("<td align='center'>" & GetNumberOfCustomersByTruckNumber(rsDeliveries("TruckNumber")) & "</td>")
					Response.Write("<td align='center'>" & GetNumberOfInvoicesByTruckNumber(rsDeliveries("TruckNumber")) & "</td>")
					Response.Write("<td align='right'>" & FormatCurrency(GetValueOfDeliveriesByTruckNumber(rsDeliveries("TruckNumber"))) & "</td>")
						
					Response.Write("</tr>")
					
					GrandTot_Stops = GrandTot_Stops + GetNumberOfCustomersByTruckNumber(rsDeliveries("TruckNumber"))
					GrandTot_Invoices = GrandTot_Invoices + GetNumberOfInvoicesByTruckNumber(rsDeliveries("TruckNumber"))
					GrandTot_Value = GrandTot_Value + GetValueOfDeliveriesByTruckNumber(rsDeliveries("TruckNumber"))
		
					rsDeliveries.Movenext
				Loop
				
				Set rsDeliveries = Nothing
				cnn9.Close
				Set cnn9 = Nothing
				
				'Print the totals
				
				'Response.Write("</tbody>")
				
				'Response.Write("<tfoot>")
 				'Response.Write("<tr>")
					
				'Response.Write("<td><strong>Totals</strong></td>")
				'Response.Write("<td><strong>" & GrandTot_Stops & " Stops</strong></td>")
				'Response.Write("<td><strong>" & GrandTot_Invoices & " Invoices</strong></td>")
				'Response.Write("<td><strong>" & FormatCurrency(GrandTot_Value) & "</strong></td>")
				
				'Response.Write("</tr>")	
				'Response.Write("</tfoot>")		
 
				%>
			 
         </tbody>   
	    </table>
        
        <table id="tableDeliveries" class="table table-striped table-condensed table-hover">
        	<tfoot>
            	<tr>
                	
                    <td width="40%"><strong>Totals</strong></td>
                    <td width="20%" align="center"><strong><%=GrandTot_Stops %> Stops</strong></td>
                    <td width="20%" align="center"><strong><%=GrandTot_Invoices %> Invoices</strong></td>
                    <td width="20%" align="right"><strong><%=FormatCurrency(GrandTot_Value) %></strong></td>
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

 

<!--#include file="../../inc/footer-main.asp"-->