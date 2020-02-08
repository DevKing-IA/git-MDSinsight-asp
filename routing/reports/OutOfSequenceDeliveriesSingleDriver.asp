<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Routing.asp"-->
<%
TruckNum = Request.Form("txtTruckNumber")
DateToRetrieve = Request.Form("txtDriverDate")

CreateAuditLogEntry GetTerm("Routing") & " Report",GetTerm("Routing") & " Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Out Of Sequence Deliveries For Driver: " & GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(TruckNum))) & " for date " & DateToRetrieve 
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
		max-width:925px;
		margin:0 auto;
		overflow:hidden;
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


<h3 class="page-header"><i class="fa fa-file-text-o"></i> Out of Sequence Deliveries for <%= GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(TruckNum))) %> on <%= DateToRetrieve  %>&nbsp;&nbsp;
	<a href="<%= BaseURL %>routing/reports/OutOfSequenceDeliveriesByDriver.asp?date=<%= DateToRetrieve %>"><button type="button" class="btn btn-primary">Back To Driver List</button></a></h3>


<div class="row">
<div class="container-center">

	<div class="table-responsive">


    	<table id="tableDeliveries" class="food_planner sortable table table-condensed table-hover">
	   		<thead>
				<tr>
					<th class="sortable"><%=GetTerm("Customer")%> #</th>
                    <th class="sortable">Name</th>
					<th class="sortable sortable-right">Invoice #</th>
					<th class="sortable sortable-right">Assigned Stop #</th>
					<th class="sortable sortable-right">Actual Stop #</th>
				</tr>
			</thead>

			<tbody>
					
				<%
	
				Set cnn9 = Server.CreateObject("ADODB.Connection")
				cnn9.open (Session("ClientCnnString"))
				Set rsDeliveries = Server.CreateObject("ADODB.Recordset")
				rsDeliveries.CursorLocation = 3 
	
				on error resume next
				SQL9 = "DROP TABLE zOutOfSequenceReport_" & Trim(Session("userNo"))
				Set rsDeliveries = cnn9.Execute(SQL9)
				on error goto 0
								
				SQL9 = "CREATE TABLE zOutOfSequenceReport_" & Trim(Session("userNo")) 
				SQL9 = SQL9 & "("
				SQL9 = SQL9 & "                [TruckNumber] [varchar](50) NULL, "
				SQL9 = SQL9 & "                [IvsNum] [varchar](50) NULL, "
				SQL9 = SQL9 & "                [SequenceNumber] [int] NULL, "
				SQL9 = SQL9 & "                [ActualDeliverySequence] [int] IDENTITY(1,1) NOT NULL, "
				SQL9 = SQL9 & "                [AMorPM] [varchar](50) NULL "
				SQL9 = SQL9 & ") ON [PRIMARY] "
				
				Set rsDeliveries = cnn9.Execute(SQL9)
				
				SQL9 = "INSERT INTO zOutOfSequenceReport_" & Trim(Session("userNo")) & "("
				SQL9 = SQL9 & "TruckNumber, IvsNum, SequenceNumber, AMorPM) "
				SQL9 = SQL9 & "SELECT TruckNumber, IvsNum, SequenceNumber, AMorPM FROM RT_DeliveryBoardHistory "
				SQL9 = SQL9 & " WHERE (Year(LastDeliveryStatusChange) = " & Year(DateToRetrieve) & " AND "
				SQL9 = SQL9 & " Month(LastDeliveryStatusChange) = " & Month(DateToRetrieve) & " AND "
				SQL9 = SQL9 & " Day(LastDeliveryStatusChange) = " & Day(DateToRetrieve) & ") AND "
				SQL9 = SQL9 & " TruckNumber = '" & TruckNum & "' " 
				SQL9 = SQL9 & " ORDER BY LastDeliveryStatusChange"
				
				Set rsDeliveries = cnn9.Execute(SQL9)

				SQL9 = "SELECT * FROM zOutOfSequenceReport_" & Trim(Session("userNo")) &" ORDER BY SequenceNumber"
				Set rsDeliveries = cnn9.Execute(SQL9)
				
				outofSequenceCount = 0
				totalStopsCount = 0
																	
				Do While not rsDeliveries.Eof
				
					totalStopsCount = totalStopsCount + 1
					
					If (rsDeliveries("SequenceNumber") <> rsDeliveries("ActualDeliverySequence")) AND rsDeliveries("AMorPM") = "AM" Then
						Response.Write("<tr style='background-color:#dfa8a8; border: 2px solid red;'>")
						outofSequenceCount = outofSequenceCount + 1

					ElseIf (rsDeliveries("SequenceNumber") = rsDeliveries("ActualDeliverySequence")) AND rsDeliveries("AMorPM") = "AM" Then
						Response.Write("<tr style='border: 2px solid red;'>")	
						
					ElseIf (rsDeliveries("SequenceNumber") <> rsDeliveries("ActualDeliverySequence")) AND rsDeliveries("AMorPM") <> "AM" Then
						Response.Write("<tr style='background-color:#f2dede;'>")
						outofSequenceCount = outofSequenceCount + 1	
								
					Else
						Response.Write("<tr>")
					End If
					Response.Write("<td>" & GetCustNumberByInvoiceNumDelBoardHistory(rsDeliveries("IvsNum")) & "</td>")
					Response.Write("<td>" & GetCustNameByCustNum(GetCustNumberByInvoiceNumDelBoardHistory(rsDeliveries("IvsNum"))) &  "</td>")
					Response.Write("<td class='sortable-center'>" & rsDeliveries("IvsNum") & "</td>")
					Response.Write("<td class='sortable-center'>" & rsDeliveries("SequenceNumber") & "</td>")
					Response.Write("<td class='sortable-center'>" & rsDeliveries("ActualDeliverySequence") & "</td>")
					Response.Write("</tr>")
		
					rsDeliveries.Movenext
				Loop
				
				Set rsDeliveries = Nothing
				cnn9.Close
				Set cnn9 = Nothing
				
				percentOutOfSequence = formatNumber(((outofSequenceCount / totalStopsCount) * 100),2)
				%>
			</tbody>
	    </table>
        
        
        <table id="tableDeliveries" class="table table-striped table-condensed table-hover">
        	<tfoot>
            	<tr>
                	
                    <td width="25%"><strong>Totals</strong></td>
                    <td width="15%" align="left"><strong><%= totalStopsCount %> Stops</strong></td>
                    <td width="30%" align="right"><strong><%= outofSequenceCount %> Stops Out of Sequence</strong></td>
                    <td width="30%" align="right"><strong><%= percentOutOfSequence %>% Out of Sequence</strong></td>
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