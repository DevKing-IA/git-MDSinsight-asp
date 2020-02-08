<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Routing.asp"-->
<%
CreateAuditLogEntry GetTerm("Routing") & " Report",GetTerm("Routing") & " Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Historical Deliveries By Driver"
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

<script type="text/javascript">
	$(document).ready(function() {
	    $("#PleaseWaitPanel").hide();
	});
</script>


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
	        location.href = 'OutOfSequenceDeliveriesByDriver.asp?date=' + selectedDate;
        });	  
        
		$('.driverButton').click(function() {
				    
    		var truckID = $(this).attr('id');
    		$('#txtTruckNumber').val(truckID);
    		
    		document.frmSetDriverOutOfSequenceDate.submit();
		    
		});  
            
      });     
      
</script>


<h3 class="page-header"><i class="fa fa-file-text-o"></i> Out of Sequence Deliveries By Driver for <%= DateToRetrieve %></h3>

<form action="OutOfSequenceDeliveriesSingleDriver.asp" method="POST" name="frmSetDriverOutOfSequenceDate" id="frmSetDriverOutOfSequenceDate">

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
					<th width="40%">Driver</th>
					<th class="sorttable-center" width="15%"># of Stops</th> 
					<th class="sorttable-center" width="15%"># Out<br>of Sequence</th>
					<th class="sorttable-center" width="15%">% Out<br>of Sequence</th>
					<th class="sorttable_nosort" width="15%">&nbsp;</th>
				</tr>
			</thead>

			<tbody>

				<%
				
				If DateToRetrieve <> "" Then
					
					SQL9 = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoardHistory "
					SQL9 = SQL9 & " WHERE Year(LastDeliveryStatusChange) = " & Year(DateToRetrieve) & " AND "
					SQL9 = SQL9 & " Month(LastDeliveryStatusChange) = " & Month(DateToRetrieve) & " AND "
					SQL9 = SQL9 & " Day(LastDeliveryStatusChange) = " & Day(DateToRetrieve)
					
					Set rs9 = cnn8.Execute(SQL9)
					
					If NOT rs9.EOF Then
										
						Do While not rs9.Eof
						
							on error resume next
							SQL8 = "DROP TABLE zOutOfSequenceReport_" & Trim(Session("userNo"))
							Set rs8 = cnn8.Execute(SQL8)
							on error goto 0
							
							
							SQL8 = "CREATE TABLE zOutOfSequenceReport_" & Trim(Session("userNo")) 
							SQL8 = SQL8 & "("
							SQL8 = SQL8 & "                [TruckNumber] [varchar](50) NULL, "
							SQL8 = SQL8 & "                [IvsNum] [varchar](50) NULL, "
							SQL8 = SQL8 & "                [SequenceNumber] [int] NULL, "
							SQL8 = SQL8 & "                [ActualDeliverySequence] [int] IDENTITY(1,1) NOT NULL "
							SQL8 = SQL8 & ") ON [PRIMARY] "
							
							Set rs8 = cnn8.Execute(SQL8)
						
							SQL8 = "INSERT INTO zOutOfSequenceReport_" & Trim(Session("userNo")) & "("
							SQL8 = SQL8 & "TruckNumber, IvsNum, SequenceNumber ) "
							SQL8 = SQL8 & "SELECT TruckNumber, IvsNum, SequenceNumber FROM RT_DeliveryBoardHistory "
							SQL8 = SQL8 & " WHERE (Year(LastDeliveryStatusChange) = " & Year(DateToRetrieve) & " AND "
							SQL8 = SQL8 & " Month(LastDeliveryStatusChange) = " & Month(DateToRetrieve) & " AND "
							SQL8 = SQL8 & " Day(LastDeliveryStatusChange) = " & Day(DateToRetrieve) & ") AND "
							SQL8 = SQL8 & " TruckNumber = '" & rs9("TruckNumber") & "' " 
							SQL8 = SQL8 & " ORDER BY LastDeliveryStatusChange"
							
							Set rs8 = cnn8.Execute(SQL8)
		
							TotalNumberOfStops = GetNumberOfInvoicesByTruckNumberHistorical(rs9("TruckNumber"),DateToRetrieve)
							NumberStopsOutOfSequence = GetNumberOutOfSequenceByTruckNumber(rs9("TruckNumber"))
							NumberStopsInSequence = TotalNumberOfStops - NumberStopsOutOfSequence
							PercentStopsOutOfSequence = formatNumber(((NumberStopsOutOfSequence / TotalNumberOfStops) * 100),2)	
		
							Response.Write("<tr>")
						    Response.write("<td>" & GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(rs9("TruckNumber"))))  & "</td>")
							Response.Write("<td align='center'>" & TotalNumberOfStops & "</td>")
							Response.Write("<td align='center'>" & NumberStopsOutOfSequence & "</td>")
							Response.Write("<td align='center'>" & PercentStopsOutOfSequence & "%</td>")
							Response.Write("<td align='right'><button type='button' class='btn btn-success driverButton' id='" & rs9("TruckNumber") & "'>View By Selected Date</button></td>")
							Response.Write("</tr>")
				
							rs9.Movenext
						Loop
						
					Else
						Response.Write("<tr><td align='center' colpsan='5' width='100%'>No Out Of Sequence Deliveries To Show</td></tr>")
					End If
						
				Else
					Response.Write("<tr><td align='center'colpsan='5 width='100%'>No Out Of Sequence Deliveries To Show</td></tr>")
				End If
				
				Set rs8 = Nothing
				Set rs9 = Nothing
				cnn8.Close
				Set cnn8 = Nothing


				 
				%>
			 
         </tbody>   
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