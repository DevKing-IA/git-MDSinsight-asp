<% 

 'Read delivery board settings
SQL = "SELECT * FROM Settings_Global"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
		DelBoardNextStopColor = rs("DelBoardNextStopColor")
		DelBoardScheduledColor = rs("DelBoardScheduledColor")	
		DelBoardCompletedColor = rs("DelBoardCompletedColor")				
		DelBoardSkippedColor = rs("DelBoardSkippedColor")		
		DelBoardProfitDollars = rs("DelBoardProfitDollars")			
		DelBoardAtOrAboveProfitColor = rs("DelBoardAtOrAboveProfitColor")			
		DelBoardBelowProfitColor = rs("DelBoardBelowProfitColor")			
		DelBoardUserAlertColor = rs("DelBoardUserAlertColor")
		DelBoardAMColor = rs("DelBoardAMColor")
		DelBoardPriorityColor = rs("DelBoardPriorityColor")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing
If DelBoardNextStopColor = "" Then DelBoardNextStopColor = "#FFA500"
If IsNull(DelBoardNextStopColor) Then DelBoardNextStopColor = "#FFA500"
If DelBoardScheduledColor = "" Then DelBoardScheduledColor = "#F6F6F6"
If IsNull(DelBoardScheduledColor) Then DelBoardScheduledColor = "#F6F6F6"
If DelBoardCompletedColor = "" Then DelBoardCompletedColor = "#D8F9D1"
If IsNull(DelBoardCompletedColor) Then DelBoardCompletedColor = "#D8F9D1"
If DelBoardSkippedColor = "" Then DelBoardSkippedColor = "#FCB3B3"
If IsNull(DelBoardSkippedColor) Then DelBoardSkippedColor = "#FCB3B3"
If DelBoardAtOrAboveProfitColor = "" Then DelBoardAtOrAboveProfitColor = "#D8F9D1"
If IsNull(DelBoardAtOrAboveProfitColor) Then DelBoardAtOrAboveProfitColor = "#D8F9D1"
If DelBoardBelowProfitColor = "" Then DelBoardBelowProfitColor = "#FCB3B3"
If IsNull(DelBoardBelowProfitColor) Then DelBoardBelowProfitColor = "#FCB3B3"
If DelBoardUserAlertColor = "" Then DelBoardUserAlertColor = "#FFA500"
If IsNull(DelBoardUserAlertColor) Then DelBoardUserAlertColor = "#FFA500"
If DelBoardAMColor = "" Then DelBoardAMColor = "#000000"
If IsNull(DelBoardAMColor) Then DelBoardAMColor = "#000000"
If DelBoardPriorityColor = "" Then DelBoardPriorityColor = "#000000"
If IsNull(DelBoardPriorityColor) Then DelBoardPriorityColor = "#000000"
%>

<!--#include file="../inc/header-deliveryboard-historical.asp"-->
<!--#include file="../inc/jquery_table_search.asp"-->
<!--#include file="../inc/InSightFuncs_Routing.asp"-->

<link href="deliveryBoardHistorical.css" rel="stylesheet"> 

<style type="text/css">

	
	.tr-completed{
		<% Response.Write("background:" & DelBoardCompletedColor & ";") %>
	}
	
	.tr-nodelivery{
		<% Response.Write("background:" & DelBoardSkippedColor & ";") %>
	}
	
	.tr-nextstop{
		<% Response.Write("background:" & DelBoardNextStopColor & ";") %>
	}
		
	.tr-scheduled{
		<% Response.Write("background:" & DelBoardScheduledColor & ";") %>
	}

	.tr-scheduled-top{
		<% Response.Write("border-top: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}
	
	.tr-scheduled-bottom{
		<% Response.Write("border-bottom: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}
	
	.AM-border-top{
		<% Response.Write("border-top: 3px solid " & DelBoardAMColor & ";") %>
		<% Response.Write("border-left: 3px solid " & DelBoardAMColor & ";") %>
		<% Response.Write("border-right: 3px solid " & DelBoardAMColor & ";") %>
	}
	
	.AM-border-bottom{
		<% Response.Write("border-bottom: 3px solid " & DelBoardAMColor & ";") %>
		<% Response.Write("border-left: 3px solid " & DelBoardAMColor & ";") %>
		<% Response.Write("border-right: 3px solid " & DelBoardAMColor & ";") %>
	}
	
	.Priority-border-top{
		<% Response.Write("border-top: 3px solid " & DelBoardPriorityColor & ";") %>
		<% Response.Write("border-left: 3px solid " & DelBoardPriorityColor & ";") %>
		<% Response.Write("border-right: 3px solid " & DelBoardPriorityColor & ";") %>
	}
	
	.Priority-border-bottom{
		<% Response.Write("border-bottom: 3px solid " & DelBoardPriorityColor & ";") %>
		<% Response.Write("border-left: 3px solid " & DelBoardPriorityColor & ";") %>
		<% Response.Write("border-right: 3px solid " & DelBoardPriorityColor & ";") %>
	}
	
	
	.tr-user-alert{
		<% Response.Write("background:" & DelBoardUserAlertColor & ";") %>
	}
	
	.tr-user-alert-top{
		<% Response.Write("border-top: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}
	
	.tr-user-alert-bottom{
		<% Response.Write("border-bottom: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}

</style>


<script type="text/javascript">
	$(document).ready(function() {		
			
		$('#myDeliveryBoardCompletedOrSkippedModal').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myInvoiceNumber = $(e.relatedTarget).data('invoice-number');
		    var myCustomerName = $(e.relatedTarget).data('customer-name');	
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtInvoiceNumber"]').val(myInvoiceNumber);
		    	    
		    var $modal = $(this);
	
    		$modal.find('#myDeliveryBoardCompletedOrSkippedLabel').html("Delivery Information for " + myCustomerName + " - Invoice #" + myInvoiceNumber);
    		
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForRoutingModals.asp",
				cache: false,
				data: "action=GetHistoricalContentForCompletedOrSkippedInfoModal&myInvoiceNumber=" + encodeURIComponent(myInvoiceNumber),
				success: function(response)
				 {
	               	 $modal.find('#deliveryBoardCompletedOrSkippedModalContent').html(response);
	             },
	             failure: function(response)
				 {
				   $modal.find('#deliveryBoardCompletedOrSkippedModalContent').html("Failed");
	             }
			});
		    
		});
		
	
	});
</script>

<script type="text/javascript" src="<%= BaseURL %>js/doublescroll/jquery.doubleScroll.js"></script>

<script type="text/javascript">
    $(document).ready(function(){

       $('.double-scroll').doubleScroll({
       		resetOnWindowResize: true
       	});
    });
</script>



<div class="horizontal-layout">
<h1 class="page-header"><i class="fa fa-truck"></i> Currently Viewing Historical Delivery Board for <%= DateToRetrieve %></h1>
<div class="double-scroll">
<table>
	<tr>
		<td>
			<div class='list-boxes' id='sortableList'>

			<%
			dim fs,t,truckorder
			truckorder=0
			set fs=Server.CreateObject("Scripting.FileSystemObject")
			filename = Server.MapPath(".")&"\truckorder\"&Session("Userno")&".txt"
			if fs.FileExists(filename) then
				set t=fs.OpenTextFile(filename,1,false)
				truckorder=t.ReadLine
				t.close
			end if
			
			Dim truckorder_arr
			truckorder_arr=split(truckorder,",")
			
			For i=0 to Ubound(truckorder_arr)
				truckorder_arr(i) = "'" & truckorder_arr(i) & "'"
			Next
			

			GridColumn = 1
			Set cnn_DeliveryBoardSum = Server.CreateObject("ADODB.Connection")
			cnn_DeliveryBoardSum.open (Session("ClientCnnString"))
			Set rs_DeliveryBoardSum = Server.CreateObject("ADODB.Recordset")
			rs_DeliveryBoardSum.CursorLocation = 3 
			
			'Write ordered trucks
			For Each TruckNumber In truckorder_arr
				SQL_DeliveryBoardSum = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoardHistory WHERE TruckNumber = " & TruckNumber & " AND DeliveryDate = '" & DateToRetrieve   & "'"
				'Response.write(SQL_DeliveryBoardSum)
				Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL_DeliveryBoardSum)
				If not rs_DeliveryBoardSum.EOF Then
					Do While Not rs_DeliveryBoardSum.Eof
						
						DriverUserNo = Trim(GetUserNumberByTruckNumber(rs_DeliveryBoardSum("TruckNumber")))
					
						If userIsArchived(DriverUserNo) = False AND userIsEnabled(DriverUserNo) = True Then
							Call TruckNumberWrite(rs_DeliveryBoardSum("TruckNumber"), GridColumn) 
						End If
						
						rs_DeliveryBoardSum.Movenext
					Loop
				End If
			Next

			'Lets get all the trucks not in order
			SQL_DeliveryBoardSum = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoardHistory WHERE TruckNumber NOT IN ('" & Replace(truckorder,",","','") & "') AND DeliveryDate = '" & DateToRetrieve & "' ORDER BY TruckNumber"
			Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL_DeliveryBoardSum)
			
			If not rs_DeliveryBoardSum.EOF Then
				Do While Not rs_DeliveryBoardSum.Eof
				
					DriverUserNo = Trim(GetUserNumberByTruckNumber(rs_DeliveryBoardSum("TruckNumber")))
				
					If userIsArchived(DriverUserNo) = False AND userIsEnabled(DriverUserNo) = True Then
						Call TruckNumberWrite(rs_DeliveryBoardSum("TruckNumber"), GridColumn) 
					End If
					
					rs_DeliveryBoardSum.Movenext
				Loop
			End If
			%> 
			</div>

 </tr></table>

</div>
</div>


<%



Sub TruckNumberWrite(TruckNumber, GridColumn)

		Response.Write("<td class='item col-lg-cust' TruckNumber='"&TruckNumber&"'>")
		Response.Write("<div class='item-box'>")
		Response.Write("<div class='scrollable-title' style='position: relative;'><strong>Truck " & TruckNumber & "<br>" & GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(TruckNumber))) & "</strong><a class='btn-move' href='#' style='position: absolute; top: 5px; right: 7px;'><i class='fa fa-arrows'></i></a></div>")
		%> 

	        <div class='table-responsive scrollable-table'>
		        <% Response.Write("<table id='truck" & TruckNumber & "' name='truck" & TruckNumber & "' class='food_planner table table-condensed sortable clickable '>") %>
					<thead>
			        	<tr>
			        		<th class='sorttable_nosort'>Invoice</th>
			        		<th class='sorttable_nosort'><%=GetTerm("Customer")%></th>.
			        	</tr>
			        </thead>
			        <tbody class='searchable'>
			        	<%'Get all the tickets for this truck
						Set cnn_Tickets = Server.CreateObject("ADODB.Connection")
						cnn_Tickets.open (Session("ClientCnnString"))
						Set rs_DeliveryBoardDet = Server.CreateObject("ADODB.Recordset")
						rs_DeliveryBoardDet.CursorLocation = 3 
						SQL_Tickets = "SELECT * FROM RT_DeliveryBoardHistory "
						SQL_Tickets = SQL_Tickets & "WHERE TruckNumber = '" & TruckNumber  & "' "
						
						SQL_Tickets = SQL_Tickets & "AND DeliveryDate = '" & DateToRetrieve & "' "
						
                        SQL_Tickets = SQL_Tickets & "Order By SequenceNumber, CustNum" 
                        Set rs_DeliveryBoardDet = cnn_Tickets.Execute(SQL_Tickets)
						If not rs_DeliveryBoardDet.Eof Then

							NumLines = 0
							Do While not rs_DeliveryBoardDet.Eof
							
								trclass = ""
								TipText = ""
								PriorityDelivery = rs_DeliveryBoardDet("Priority")
								
								'Write first table row
								'**********************
								If rs_DeliveryBoardDet("DeliveryStatus") = "Delivered" Then
								
									If rs_DeliveryBoardDet("AMorPM") = "AM" Then
									
										%><tr class="tr-completed AM-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_DeliveryBoardDet("IvsNum") %>" data-customer-name="<%= rs_DeliveryBoardDet("CustName") %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
									
									Else
									
										If PriorityDelivery = 1 Then
											%><tr class="tr-completed Priority-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_DeliveryBoardDet("IvsNum") %>" data-customer-name="<%= rs_DeliveryBoardDet("CustName") %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
										Else
											%><tr class="tr-completed tr-scheduled-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_DeliveryBoardDet("IvsNum") %>" data-customer-name="<%= rs_DeliveryBoardDet("CustName") %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
										End If
									
									End If
									
								ElseIf rs_DeliveryBoardDet("DeliveryStatus") = "No Delivery" Then
								
									If rs_DeliveryBoardDet("AMorPM") = "AM" Then

										%><tr class="tr-nodelivery  AM-border-to" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_DeliveryBoardDet("IvsNum") %>" data-customer-name="<%= rs_DeliveryBoardDet("CustName") %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
									
									Else
								
										If PriorityDelivery = 1 Then
											%><tr class="tr-nodelivery  Priority-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_DeliveryBoardDet("IvsNum") %>" data-customer-name="<%= rs_DeliveryBoardDet("CustName") %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
										Else
											%><tr class="tr-nodelivery tr-scheduled-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_DeliveryBoardDet("IvsNum") %>" data-customer-name="<%= rs_DeliveryBoardDet("CustName") %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
										End If

									End If
									
								Else
								
									If rs_DeliveryBoardDet("AMorPM") = "AM" Then
									
										%><tr class="tr-scheduled AM-border-top" data-invoice-number="<%= rs_DeliveryBoardDet("IvsNum") %>" data-customer-name="<%= rs_DeliveryBoardDet("CustName") %>"><%
									
									Else
									
										If PriorityDelivery = 1 Then
											%><tr class="tr-scheduled Priority-border-top" data-invoice-number="<%= rs_DeliveryBoardDet("IvsNum") %>" data-customer-name="<%= rs_DeliveryBoardDet("CustName") %>"><%
										Else
											%><tr class="tr-scheduled tr-scheduled-top" data-invoice-number="<%= rs_DeliveryBoardDet("IvsNum") %>" data-customer-name="<%= rs_DeliveryBoardDet("CustName") %>"><%
										End If
									
									End If
								End If
								
								Response.Write(trclass)

								If TipText <> "" Then
									Response.Write("<td>" & rs_DeliveryBoardDet("IvsNum") & "</td>")
									Response.Write("<td><div class='alarm-bell'>" & rs_DeliveryBoardDet("CustNum") & "<span class='alert-pop-up'>" & TipText  & "</span></div></td></tr>")
								Else
									Response.Write("<td>" & rs_DeliveryBoardDet("IvsNum") & "</td>")
									Response.Write("<td>" & rs_DeliveryBoardDet("CustNum") & "</td></tr>")
								End If

																	
								'Write second table row
								'**********************
								If rs_DeliveryBoardDet("DeliveryStatus") = "Delivered" Then
								
									If rs_DeliveryBoardDet("AMorPM") = "AM" Then
									
										trclass = "<tr class='tr-completed AM-border-bottom'>"
									
									Else
									
										If PriorityDelivery = 1 Then
											trclass = "<tr class='tr-completed Priority-border-bottom'>"
										Else
											trclass = "<tr class='tr-completed tr-scheduled-bottom'>"
										End If

									End If
									
									
								ElseIf rs_DeliveryBoardDet("DeliveryStatus") = "No Delivery" Then
								
									If rs_DeliveryBoardDet("AMorPM") = "AM" Then
									
										trclass = "<tr class='tr-nodelivery AM-border-bottom'>"
										
									Else
									
										If PriorityDelivery = 1 Then
											trclass = "<tr class='tr-nodelivery Priority-border-bottom''>"
										Else
											trclass = "<tr class='tr-nodelivery tr-scheduled-bottom'>"
										End If

									End If
									
								Else
								
									If rs_DeliveryBoardDet("AMorPM") = "AM" Then
									
										trclass = "<tr class='tr-scheduled AM-border-bottom'>"
										
									Else

										If PriorityDelivery = 1 Then
											trclass = "<tr class='tr-scheduled Priority-border-bottom''>"
										Else
											trclass = "<tr class='tr-scheduled tr-scheduled-bottom'>"
										End If
										
									End If
									
								End If
								
								Response.Write(trclass)
								If len(rs_DeliveryBoardDet("CustName")) > 19 then Cnam = left(rs_DeliveryBoardDet("CustName"),19) Else Cnam = rs_DeliveryBoardDet("CustName")
								Response.Write("<td colspan='2'>" & Cnam)
								Response.Write("<span class='alarm-bell'>")
								Response.Write("</span>")
								Response.Write("</td></tr>")
			
								
								rs_DeliveryBoardDet.movenext
								NumLines = NumLines + 1
								
							Loop
							
							'Make all boxes even
							If NumLines < MaxNumberOfDeliveries() Then
								For x = 1 to MaxNumberOfDeliveries() - NumLines
									Response.Write("<tr ><td>&nbsp;</td></tr>")
									Response.Write("<tr ><td>&nbsp;</td></tr>")
								Next
							End IF
							
						End IF%>
                        </td>
			        </tbody>
		        </table>
	        </div>
            </div>
        <%Response.Write("</div>")
		GridColumn = GridColumn +1
End Sub 

Set rs_DeliveryBoardSum = Nothing
cnn_DeliveryBoardSum.Close
Set cnn_DeliveryBoardSum = Nothing
%>	


<script type="text/javascript">
	function setSortable() {
		$("#sortableList").sortable({ placeholder: "ui-state-highlight item col-lg-1 col-lg-cust", handle: ".btn-move", scrollSensitivity: 40, scrollSpeed: 60, update: function (event, ui) { saveSelection(); } });
		$("#sortableList").disableSelection();
	}
	function saveSelection() {
		var list = "";
		try {
			var sep = "";
			$("#sortableList .item").each(function () {
				list += "" + sep + $(this).attr("TruckNumber");
				sep = ",";
			});
		}
		catch (ex) {
			alert(ex);
			return;
		}
		
		var url = "truckorder/save.asp";
		var jsondata = {};
		jsondata.truckorder = list;
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {
			}
		});
	}
	$(function () {
		setSortable();
	});
</script>


<!-- same height titles !-->
<script type="text/javascript" src="<%= BaseURL %>js/grids.js"></script>

<script type="text/javascript">
	jQuery(function($) {
		$('.scrollable-title').responsiveEqualHeightGrid();	
 	});
</script>
    <!-- eof same height titles !-->
    

<!-- **************************************************************************************************************************** -->
<!-- MODAL FOR DELIVERY ALERTS BEGINS HERE !-->
<!-- **************************************************************************************************************************** -->


<div class="modal fade" id="myDeliveryBoardCompletedOrSkippedModal" tabindex="-1" role="dialog" aria-labelledby="myDeliveryBoardCompletedOrSkippedLabel">

	<div class="modal-dialog" role="document">
						
		<div class="modal-content">
	    
			<!-- modal header !-->
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="myDeliveryBoardCompletedOrSkippedLabel"></h4>
			</div>
			<!-- eof modal header !-->
	  
			<!-- modal body !-->
			<div class="modal-body">
			
				<input type="hidden" name="txtInvoiceNumber" id="txtInvoiceNumber" value="">
					<div id="deliveryBoardCompletedOrSkippedModalContent">
						<!-- Content for the modal will be generated and written here -->
						<!-- Content generated by Sub GetContentForCompletedOrSkippedInfoModal() in InsightFuncs_AjaxForRoutingModals.asp -->
					</div>
			</div>

		</div>
		<!-- eof modal content !-->
</div>
<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->


<!-- **************************************************************************************************************************** -->
<!-- MODAL FOR DELIVERY ALERTS ENDS HERE !-->
<!-- **************************************************************************************************************************** -->


<!--#include file="../inc/footer-deliveryBoard.asp"-->
