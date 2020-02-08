<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Routing.asp"-->
<%
CreateAuditLogEntry GetTerm("Routing") & " Report",GetTerm("Routing") & " Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Todays Deliveries By Driver"
%>
 
<style>
	table.sorttable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
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
		max-width:800px;
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
</style>

<!-- datepicker for historical delivery board !-->
<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<link href="<%= baseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet" type="text/css">
<script src="<%= baseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.js" type="text/javascript"></script>
<!-- end datepicker for historical delivery board !-->


<%
				
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs9 = Server.CreateObject("ADODB.Recordset")
rs9.CursorLocation = 3 

DateToRetrieve = Request.QueryString("date")

If DateToRetrieve = "" Then
	SQL8 = "SELECT MAX (LastDeliveryStatusChange) AS Expr1 FROM RT_DeliveryBoardHistory"
	Set rs8 = cnn8.Execute(SQL8)
	If NOT rs8.Eof Then 
		If rs8("Expr1") <> "" AND NOT IsNull(rs8("Expr1")) AND NOT IsEmpty(rs8("Expr1")) Then
			DateToRetrieve = formatDateTime(rs8("Expr1"),2)
		Else
			DateToRetrieve = ""
		End If
	End If
	
End If

%>

<script>
   $(document).ready(function(){	
   
        $('#datepicker1').datetimepicker({
        	format: 'MM/DD/YYYY',
        	useCurrent: false,
        	//defaultDate: moment(momentString),
        	maxDate:moment().add(-1, 'days')
        });
        
		$("#datepicker1").on("dp.change", function (e) {
	    	selectedDate = $("#datepicker1").find("input").val();
	        location.href = 'HistoricalDeliveriesByDriver.asp?date=' + selectedDate;
        });	  
        
		$('.driverButton').click(function() {
				    
    		var truckID = $(this).attr('id');
    		$('#txtTruckNumber').val(truckID);
    		
    		document.frmSetDriverHistoricalDate.submit();
		    
		}); 
		
		 $("#PleaseWaitPanel").hide(); 
            
      });     
      
</script>


<h3 class="page-header"><i class="fa fa-file-text-o"></i> Past Deliveries By Driver for <%= DateToRetrieve %> &nbsp;&nbsp;</h3>

<form action="HistoricalDeliveriesSingleDriver.asp" method="POST" name="frmSetDriverHistoricalDate" id="frmSetDriverHistoricalDate">

<div class="row">
<div class="container-center">
<div class="col-lg-12">
     <!-- datepicker !-->
     <div class="col-lg-8">
		<!-- Bootstrap datepicker for filtering leads by date -->
        <div class="form-group">
            <div class="input-group date" id="datepicker1">
                <input type="text" class="form-control" name="txtDriverDate" id="txtDriverDate" value="<%= DateToRetrieve %>">
                <input type="hidden" name="txtTruckNumber" id="txtTruckNumber">
                <span class="input-group-addon">
                    <span class="glyphicon glyphicon-calendar"></span>
                </span>
            </div>
        </div>	
      </div>
    <!-- eof datepicker !-->

	<div class="col-lg-2">
	<a href="<%= BaseURL %>routing/reports.asp"><button type="button" class="btn btn-primary">Back To <%= GetTerm("Routing") %> Reports List</button></a>
	</div>
</div>
</div>
</div>

<div class="row">
<div class="container-center">
	<div class="table-responsive">


    	<table id="tableDeliveries" class="sorttable table table-striped table-condensed table-hover">
	   		<thead>
				<tr>
					<th width="30%">Driver</th> 
					<th class="sorttable-center" width="20%"># Stops<br>(# <%=GetTerm("Customers")%>)</th>
					<th class="sorttable-center" width="20%"># Invoices</th>
					<th class="sorttable-right" width="20%">Value</th>
					<th class="sorttable_nosort" width="15%">&nbsp;</th>
				</tr>
			</thead>

			<tbody>
					
				<%
	
				Set cnn9 = Server.CreateObject("ADODB.Connection")
				cnn9.open (Session("ClientCnnString"))
				Set rsDeliveries = Server.CreateObject("ADODB.Recordset")
				rsDeliveries.CursorLocation = 3 
	
				If DateToRetrieve <> "" Then
				
					SQL9 = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoardHistory "
					SQL9 = SQL9 & " WHERE Year(LastDeliveryStatusChange) = " & Year(DateToRetrieve) & " AND "
					SQL9 = SQL9 & " Month(LastDeliveryStatusChange) = " & Month(DateToRetrieve) & " AND "
					SQL9 = SQL9 & " Day(LastDeliveryStatusChange) = " & Day(DateToRetrieve)
	
	
					Set rsDeliveries = cnn9.Execute(SQL9)
						
					GrandTot_Stops = 0
					GrandTot_Invoices = 0
					GrandTot_Value = 0
					
					If NOT rsDeliveries.EOF Then
																		
						Do While not rsDeliveries.Eof
		
							Response.Write("<tr>")
								
						    Response.write("<td>" & GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(rsDeliveries("TruckNumber"))))  & "</td>")
							Response.Write("<td align='center'>" & GetNumberOfCustomersByTruckNumberHistorical(rsDeliveries("TruckNumber"),DateToRetrieve) & "</td>")
							Response.Write("<td align='center'>" & GetNumberOfInvoicesByTruckNumberHistorical(rsDeliveries("TruckNumber"),DateToRetrieve) & "</td>")
							Response.Write("<td align='right'>" & FormatCurrency(GetValueOfDeliveriesByTruckNumberHistorical(rsDeliveries("TruckNumber"),DateToRetrieve),2) & "</td>")
							%><td align="right"><button type="button" class="btn btn-success driverButton" id="<%= rsDeliveries("TruckNumber") %>">View By Selected Date</button></td><%
								
							Response.Write("</tr>")
							
							GrandTot_Stops = GrandTot_Stops + GetNumberOfCustomersByTruckNumberHistorical(rsDeliveries("TruckNumber"),DateToRetrieve)
							GrandTot_Invoices = GrandTot_Invoices + GetNumberOfInvoicesByTruckNumberHistorical(rsDeliveries("TruckNumber"),DateToRetrieve)
							GrandTot_Value = GrandTot_Value + GetValueOfDeliveriesByTruckNumberHistorical(rsDeliveries("TruckNumber"),DateToRetrieve)
				
							rsDeliveries.Movenext
						Loop
					Else
						Response.Write("<tr><td align='center' colpsan='5' width='100%'>No Historical Deliveries To Show</td></tr>")
					End If
						
				Else
					Response.Write("<tr><td align='center' colpsan='5' width='100%'>No Historical Deliveries To Show</td></tr>")
				End If
				
				Set rsDeliveries = Nothing
				cnn9.Close
				Set cnn9 = Nothing
 
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
                    <td>&nbsp;</td>
                </tr>
            </tfoot>
        </table>
	</div>
    </div>
</div>

</form>

<!-- row !-->
<div class="row">
	<div class="col-lg-12"><hr></div>
</div>
<!-- eof row !-->

 

<!--#include file="../../inc/footer-main.asp"-->