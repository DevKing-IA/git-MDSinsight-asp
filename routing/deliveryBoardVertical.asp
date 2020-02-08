<% 'Read delivery board settings
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
<!--#include file="../inc/header-deliveryboard.asp"-->
<!--#include file="../inc/InSightFuncs_Routing.asp"-->

<link href="deliveryBoardVertical.css" rel="stylesheet"> 
 

<script type="text/javascript">

	$(document).ready(function() {

		$('.btn-toggle').click(function() {
		
			  driverUserNo = $(this).attr("id");
			
			  if ($(this).find('.btn-nag-on').size() > 0) {
			  				    
				    $("#" + driverUserNo + "ON").removeClass('btn-nag-on');
				    $("#" + driverUserNo + "ON").addClass('btn-default');
				    $("#" + driverUserNo + "ON").addClass('active');
				    
				    $("#" + driverUserNo + "OFF").removeClass('btn-default');
				    $("#" + driverUserNo + "OFF").removeClass('active');
				    $("#" + driverUserNo + "OFF").addClass('btn-nag-off');
				    				    
				    $("#" + driverUserNo + "ON").html("ON");
				    $("#" + driverUserNo + "OFF").html("NAG OFF");	
				    
			    	$.ajax({
						type:"POST",
						url: "../inc/InSightFuncs_AjaxForRoutingModals.asp",
						cache: false,
						data: "action=TurnOnNagAlertsForDeliveryBoardDriver&driverUserNo=" + encodeURIComponent(driverUserNo),
						success: function(response)
						 {
			             }
					});
			  }
			  else {

				    $("#" + driverUserNo + "OFF").removeClass('btn-nag-off');
				    $("#" + driverUserNo + "OFF").addClass('btn-default');
				    $("#" + driverUserNo + "OFF").addClass('active');
				    
				    $("#" + driverUserNo + "ON").removeClass('btn-default');
				    $("#" + driverUserNo + "ON").removeClass('active');
				    $("#" + driverUserNo + "ON").addClass('btn-nag-on');
				    
				    $("#" + driverUserNo + "ON").html("NAG ON");
				    $("#" + driverUserNo + "OFF").html("OFF");			    
				    
			    	$.ajax({
						type:"POST",
						url: "../inc/InSightFuncs_AjaxForRoutingModals.asp",
						cache: false,
						data: "action=TurnOffNagAlertsForDeliveryBoardDriver&driverUserNo=" + encodeURIComponent(driverUserNo),
						success: function(response)
						 {
			             }
					});
			  }
			  
		});	



			
		$('#deliveryBoardInvoiceOptionsModal').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myInvoiceNumber = $(e.relatedTarget).data('invoice-number');
		    var myCustomerName = $(e.relatedTarget).data('customer-name');	
		    var myCustID = $(e.relatedTarget).data('customer-id');
		    var myTruckNumber = $(e.relatedTarget).data('truck-number');
		    
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtInvoiceNumber"]').val(myInvoiceNumber);
		    $(e.currentTarget).find('input[name="txtTruckNumber"]').val(myTruckNumber);
		    $(e.currentTarget).find('input[name="txtCustID"]').val(myCustID);
		    	    
		    var $modal = $(this);
	
    		$modal.find('#deliveryBoardLabel').html("Delivery Options For " + myCustomerName + " - Invoice  #" + myInvoiceNumber);
    		
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForRoutingModals.asp",
				cache: false,
				data: "action=GetContentForDeliveryBoardOptionsModal&returnPage=routing/deliveryBoardVertical.asp&invoiceNum=" + encodeURIComponent(myInvoiceNumber) + "&custID=" + encodeURIComponent(myCustID) + "&truckNum=" + encodeURIComponent(myTruckNumber),
				success: function(response)
				 {
	             	$modal.find('#deliveryBoardInvoiceOptionsModalContent').html(response);
	             },
	             failure: function(response)
				 {
				 	$modal.find('#deliveryBoardInvoiceOptionsModalContent').html("Failed");
	             }
			});
		    
		});
		

			
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
				data: "action=GetContentForCompletedOrSkippedInfoModal&myInvoiceNumber=" + encodeURIComponent(myInvoiceNumber),
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



<!-- DYNAMIC FORM !-->
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

	.col-lg-2-border{
		border:1px solid #666;
		padding-top:5px;
		padding-bottom:5px;
		padding-right:0px;
		padding-left:5px;
		margin-right:-1px;
		margin-bottom:-1px;
		font-size:11px;
		min-height: 85px;
	}

	.col-lg-2-border-AM{
		<% Response.Write("-webkit-box-shadow:inset 0px 0px 0px 5px " & DelBoardAMColor & ";") %>
		<% Response.Write("-moz-box-shadow:inset 0px 0px 0px 5px " & DelBoardAMColor & ";") %>
		<% Response.Write("box-shadow:inset 0px 0px 0px 5px " & DelBoardAMColor & ";") %>
		padding-top:5px;
		padding-bottom:5px;
		padding-right:0px;
		padding-left:5px;
		margin-right:-1px;
		margin-bottom:-1px;
		font-size:11px;
		min-height: 85px;
	}
	
	
	.col-lg-2-border-Priority{
		<% Response.Write("-webkit-box-shadow:inset 0px 0px 0px 5px " & DelBoardPriorityColor & ";") %>
		<% Response.Write("-moz-box-shadow:inset 0px 0px 0px 5px " & DelBoardPriorityColor & ";") %>
		<% Response.Write("box-shadow:inset 0px 0px 0px 5px " & DelBoardPriorityColor & ";") %>		
		padding-top:5px;
		padding-bottom:5px;
		padding-right:0px;
		padding-left:5px;
		margin-right:-1px;
		margin-bottom:-1px;
		font-size:11px;
		min-height: 85px;
	}
		
	.tr-user-alert{
		<% Response.Write("background:" & DelBoardUserAlertColor & ";") %>
	}
	
	.tr-user-alert-top{
		<% Response.Write("border: 1px solid #000000;") %>
 	}
	 
	 .tr-border-line{
		 border-bottom: 1px solid #ccc;
	 }
	 
	 .tr-border-line-red{
		 border-bottom: 1px solid #999;
	 }
	

	</style>

    <!-- icons !-->
    <script src="https://use.fontawesome.com/b99adf8d86.js"></script>
       
<!-- truck drivers list starts here !-->
<div class="container-fluid container-fluid-trucks">


<div class="form-group" style="margin-left:-10px;">
	<form id="live-search" action="" class="styled" method="post">
	    <div class="input-group input-group-lg">
	        <span class="input-group-addon"><i class="fa fa-truck"></i></span>
	        <div class="icon-addon addon-lg">
	            <input type="text" placeholder="Search By Customer, Invoice #, Etc." class="form-control" id="filter" style="border-top-right-radius:6px;border-bottom-right-radius:6px;">
	            <label for="email" class="fa fa-search" rel="tooltip" title="email"></label>
	        </div>
	    </div>
    </form>
</div>    

<div class="row">
 
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
		SQL_DeliveryBoardSum = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard WHERE TruckNumber = " & TruckNumber
		Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL_DeliveryBoardSum)
		If not rs_DeliveryBoardSum.EOF Then
		
			CurrentColumn = 0
		
			Do While Not rs_DeliveryBoardSum.Eof
			
				If CurrentColumn = 0 Then 
					%><div class="row" style="margin-right:0px !important; margin-left :0px !important; font-size:14px;"><%
				End If
				
				If DelBoardIgnoreThisRoute(rs_DeliveryBoardSum("TruckNumber")) <> True Then 
				
					DriverUserNo = Trim(GetUserNumberByTruckNumber(rs_DeliveryBoardSum("TruckNumber")))
				
					If userIsArchived(DriverUserNo) = False AND userIsEnabled(DriverUserNo) = True Then

						%><div class="col-lg-6"><%
						Call TruckNumberWrite(rs_DeliveryBoardSum("TruckNumber"), GridColumn)
						%></div><% 
						CurrentColumn = CurrentColumn + 1
						
					End If
					
				End If

				If CurrentColumn = 2 Then 
					%></div><%
				End If
				
				rs_DeliveryBoardSum.Movenext
			Loop
		End If
	Next
	
	'Lets get all the trucks not in order
	SQL_DeliveryBoardSum = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard WHERE TruckNumber NOT IN ('" & Replace(truckorder,",","','") & "')  ORDER BY TruckNumber"
	Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL_DeliveryBoardSum)
	
	If not rs_DeliveryBoardSum.EOF Then
	
		CurrentColumn = 0
		
		Do While Not rs_DeliveryBoardSum.Eof
	
			If CurrentColumn = 0 Then 
				%><div class="row" style="margin-right:0px !important; margin-left :0px !important; font-size:14px;"><%
			End If
					
			If DelBoardIgnoreThisRoute(rs_DeliveryBoardSum("TruckNumber")) <> True Then 
			
				DriverUserNo = Trim(GetUserNumberByTruckNumber(rs_DeliveryBoardSum("TruckNumber")))
			
				If userIsArchived(DriverUserNo) = False AND userIsEnabled(DriverUserNo) = True Then

					%><div class="col-lg-6"><%
					Call TruckNumberWrite(rs_DeliveryBoardSum("TruckNumber"), GridColumn)
					%></div><% 
					CurrentColumn = CurrentColumn + 1
					
				End If
				
			End If
			
			If CurrentColumn = 2 Then 
				%></div><%
			End If
			
			
			rs_DeliveryBoardSum.Movenext
		Loop
	End If
%> 

</div>

<%
Sub TruckNumberWrite(TruckNumber, GridColumn)


		max=25
		min=5
		Randomize
		TotalStops = Int((max-min+1)*Rnd+min)
		RemainingStops = Int((max-min+1)*Rnd+min)
		
		If TotalStops < RemainingStops Then
			Do While TotalStops < RemainingStops
			    RemainingStops = Int((max-min+1)*Rnd+min)
			Loop
		End If

		TotalStops = GetTotalStopsByTruckNumber(TruckNumber)
		RemainingStops = GetRemainingStopsByTruckNumber(TruckNumber)
		CurrentStop = Abs(RemainingStops - TotalStops)
	
		If TotalStops > 0 Then
			PercentComplete = Round(((TotalStops - RemainingStops) / TotalStops) * 100)
		Else
			PercentComplete = 0
		End If
		
		Response.Write("<div class='col-lg-12 col-lg-hide' TruckNumber='" & TruckNumber & "'>") %>
		
         <button class="btn-truck" role="button" data-toggle="collapse" href="#<%=TruckNumber%>" aria-expanded="false" aria-controls="<%=TruckNumber%>">
         

         	<div class="col-lg-12">
         		<div class="col-lg-6">Route: <%= TruckNumber %>,  <%= GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(TruckNumber))) %></div>
         		<div class="col-lg-6"><span class="stopinfo">Total Stops: <%= TotalStops %>, Remaining: <%= RemainingStops %></span></div>
         	</div>
         	<% If TotalStops > 0 Then %>
         	<div class="col-lg-12">
	         	<div class="progress">
	                <div class="progress-bar progress-bar-success" role="progressbar" aria-valuenow="40" aria-valuemin="0" aria-valuemax="100" style="width: <%= PercentComplete %>%">
	                    <span class="sr-only"><%= PercentComplete %>% Complete (success)</span>
	                </div>
	                <span class="progress-completed"><%= PercentComplete %>%</span>
	            </div>				
			</div>	
			<% End If %>
         </button>
         
         <div class='collapse' id="<%=TruckNumber%>"> 
			<div class="well">
			
				<%	
				'**************************************************************************
				'show nag alerts to admins or route managers only
				'**************************************************************************
				
				If userIsAdmin(Session("userNo")) OR userIsRouteManager(Session("userNo")) Then
				%>

	  				<div class="row" style="margin-left:2px;margin-bottom:15px;">
	  					<%
	  					DriverUserNo = Trim(GetUserNumberByTruckNumber(TruckNumber))
			
						If DriverUserNo <> "*Not Found*" Then
							If DriverNumberHasNagAlerts(DriverUserNo) = True Then
								buttonClassGreen = "btn btn-xs btn-nag-on" 
								buttonClassRed= "btn btn-xs btn-default active"
							Else 
								buttonClassGreen = "btn btn-xs btn-default active" 
								buttonClassRed = "btn btn-xs btn-nag-off"
							End If
							%>	  
							<% If buttonClassGreen = "btn btn-xs btn-nag-on" Then %>
								  <div class="btn-group btn-toggle" id="<%= DriverUserNo %>">
								    <button class="<%= buttonClassGreen %>" id="<%= DriverUserNo %>ON">NAG ON</button>
								    <button class="<%= buttonClassRed %>" id="<%= DriverUserNo %>OFF">TOFF</button>
								  </div>
							<% Else %>
								  <div class="btn-group btn-toggle" id="<%= DriverUserNo %>">
								    <button class="<%= buttonClassGreen %>" id="<%= DriverUserNo %>ON">ON</button>
								    <button class="<%= buttonClassRed %>" id="<%= DriverUserNo %>OFF">NAG OFF</button>
								  </div>						
							<% End If %>
							<%
						End If %>
	  				</div>
  				<% End If %>
  
		        <% Response.Write("<div id='truck" & TruckNumber & "' name='truck" & TruckNumber & "'") %>
                
 			        	<%'Get all the tickets for this truck
						Set cnn_Tickets = Server.CreateObject("ADODB.Connection")
						cnn_Tickets.open (Session("ClientCnnString"))
						Set rs_DeliveryBoardDet = Server.CreateObject("ADODB.Recordset")
						rs_DeliveryBoardDet.CursorLocation = 3 
						SQL_Tickets = "SELECT * FROM RT_DeliveryBoard "
						SQL_Tickets = SQL_Tickets & "WHERE TruckNumber = '" & TruckNumber  & "' "
						If DelBoardDontUseStopSequencing() = False Then
	                        SQL_Tickets = SQL_Tickets & "Order By SequenceNumber, CustNum" 
	                    Else
   	                        SQL_Tickets = SQL_Tickets & "Order By CustNum" 
	                    End If

                        Set rs_DeliveryBoardDet = cnn_Tickets.Execute(SQL_Tickets)
						If not rs_DeliveryBoardDet.Eof Then

							NumLines = 0
							Do While not rs_DeliveryBoardDet.Eof
							
								trclass = ""
								TipText = ""
								
								PriorityDelivery = rs_DeliveryBoardDet("Priority")
								InvoiceNumber = rs_DeliveryBoardDet("IvsNum")
								CustName = rs_DeliveryBoardDet("CustName")
								CustID = rs_DeliveryBoardDet("CustNum")
								TruckNumber = rs_DeliveryBoardSum("TruckNumber")
								AMorPM = rs_DeliveryBoardDet("AMorPM")
								DeliveryStatus = rs_DeliveryBoardDet("DeliveryStatus")
								
								If CustID = GetNextCustomerStopByTruck(TruckNumber) Then
									
									If AMorPM = "AM" Then
									
										If DeliveryAlertSet(InvoiceNumber,Session("UserNo")) Then
											%><div class="tr-user-alert col-lg-2 col-lg-2-border-AM" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											TipText = "Alert when " &  DeliveryAlertCondition(InvoiceNumber,Session("UserNo")) 
										Else
											%><div class="tr-nextstop col-lg-2 col-lg-2-border-AM" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
										End If
										 
										Response.Write("<div class='col-lg-6' data-searchvalue='" & InvoiceNumber & "'>" & InvoiceNumber & "</div>")
										Response.Write("<div class='col-lg-6' data-searchvalue='" & CustID & "'>" & CustID & "</div>")
 										
										If len(CustName) > 19 then
											Response.Write("<div class='col-lg-12' data-searchvalue='" & left(CustName,19) & "'>" & left(CustName,19)) 
										Else
											Response.Write("<div class='col-lg-6' data-searchvalue='" & CustName & "'>" & CustName) 
										End If
										 
										Response.Write("</div>")
										
										If AMorPM = "AM" Then
											Response.Write("</div>")
										End If
										
									Else
										%> <%
										If DeliveryAlertSet(InvoiceNumber,Session("UserNo")) Then
											If PriorityDelivery = 1 Then
												%><div class="tr-user-alert tr-scheduled-top col-lg-2 col-lg-2-border-Priority" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											Else
												%><div class="tr-user-alert tr-scheduled-top col-lg-2 col-lg-2-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											End If
											TipText = "Alert when " &  DeliveryAlertCondition(InvoiceNumber,Session("UserNo")) 
										Else
											If PriorityDelivery = 1 Then
												%><div class="tr-nextstop tr-scheduled-top col-lg-2 col-lg-2-border-Priority" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											Else
												%><div class="tr-nextstop tr-scheduled-top col-lg-2 col-lg-2-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											End If
										End If
									
										
										 
										Response.Write("<div class='col-lg-6' data-searchvalue='" & InvoiceNumber & "'>" & InvoiceNumber & "</div>")
										Response.Write("<div class='col-lg-6'' data-searchvalue='" & CustID & "'>" & CustID & "</div>")
 						
										If len(CustName) > 19 then
											Response.Write("<div class='col-lg-12' data-searchvalue='" & left(CustName,19) & "'>" & left(CustName,19))
										Else
											Response.Write("<div class='col-lg-12' data-searchvalue='" & CustName & "'>" & CustName)
										End If
										 
										Response.Write("</div>")
										Response.Write("</div>")
									End If
									
								Else
								
									'Write first table row
									'**********************
									If DeliveryStatus = "Delivered" Then
										If AMorPM = "AM" Then
											%><div class="tr-completed  col-lg-2 col-lg-2-border-AM" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
										Else
											If PriorityDelivery = 1 Then
												%><div class="tr-completed tr-scheduled-top col-lg-2 col-lg-2-border-Priority" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											Else
												%><div class="tr-completed tr-scheduled-top col-lg-2 col-lg-2-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											End If
										End IF
									ElseIf DeliveryStatus = "No Delivery" Then
										If AMorPM = "AM" Then
											%><div class="tr-nodelivery   col-lg-2 col-lg-2-border-AM" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
										Else
											If PriorityDelivery = 1 Then
												%><div class="tr-nodelivery tr-scheduled-top col-lg-2 col-lg-2-border-Priority" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											Else
												%><div class="tr-nodelivery tr-scheduled-top col-lg-2 col-lg-2-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											End If
										End If
									Else
										If AMorPM = "AM" Then
											If DeliveryAlertSet(InvoiceNumber,Session("UserNo")) Then
												%><div class="tr-user-alert  col-lg-2 col-lg-2-border-AM" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
												TipText = "Alert when " & DeliveryAlertCondition(InvoiceNumber,Session("UserNo")) 
											Else
												If PriorityDelivery = 1 Then
													%><div class="tr-scheduled col-lg-2 col-lg-2-border-Priority" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;" id="<%= uniqueID %>"><%
												Else
													%><div class="tr-scheduled col-lg-2 col-lg-2-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;" id="<%= uniqueID %>"><%
												End If
											End If
										Else
											If DeliveryAlertSet(InvoiceNumber,Session("UserNo")) Then
												If PriorityDelivery = 1 Then
													%><div class="tr-user-alert tr-user-alert-top col-lg-2 col-lg-2-border-Priority" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;" id="<%= uniqueID %>"><%
												Else
													%><div class="tr-user-alert tr-user-alert-top col-lg-2 col-lg-2-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;" id="<%= uniqueID %>"><%
												End If
												TipText = "Alert when " & DeliveryAlertCondition(InvoiceNumber,Session("UserNo")) 
											Else
												If PriorityDelivery = 1 Then
													%><div class="tr-scheduled tr-scheduled-top col-lg-2 col-lg-2-border-Priority" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;" id="<%= uniqueID %>"><%
												Else
													%><div class="tr-scheduled tr-scheduled-top col-lg-2 col-lg-2-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;" id="<%= uniqueID %>"><%
												End If
											End If
										End If
									End If
									

 
									If TipText <> "" Then
									
										If GetLastInvoiceMarkedByTruckNumber(rs_DeliveryBoardSum("TruckNumber")) = InvoiceNumber Then
										
											Response.Write("<div class='col-lg-6' data-searchvalue='" & InvoiceNumber & "'>" & InvoiceNumber & "<i class='fa fa-star' aria-hidden='true'></i></div>")
										Else
											Response.Write("<div class='col-lg-6' data-searchvalue='" & InvoiceNumber & "'>" & InvoiceNumber & "</div>")
										End If
										
										Response.Write("<div class='col-lg-6' data-searchvalue='" & CustID & "'><div class='alarm-bell'>" & CustID & "<span class='alert-pop-up'>" & TipText  & "</span></div></div>")
										
										If len(CustName) > 19 then
											Response.Write("<div class='col-lg-12' data-searchvalue='" & left(CustName,19) & "'>" & left(CustName,19))
										Else
											Response.Write("<div class='col-lg-12' data-searchvalue='" & CustName & "'>" & CustName)
										End If
										
										Response.Write("</div>")
										Response.Write("</div>")

									Else
										If GetLastInvoiceMarkedByTruckNumber(rs_DeliveryBoardSum("TruckNumber")) = InvoiceNumber Then
											Response.Write("<div class='col-lg-6' data-searchvalue='" & InvoiceNumber & "'>" & InvoiceNumber & "<i class='fa fa-star' aria-hidden='true'></i></div>")
										Else
											Response.Write("<div class='col-lg-6' data-searchvalue='" & InvoiceNumber & "'>" & InvoiceNumber & "</div>")
										End If
										
										Response.Write("<div class='col-lg-6' data-searchvalue='" & CustID & "'>" & CustID & "</div>")										
 						
										If len(CustName) > 19 then
											Response.Write("<div class='col-lg-12' data-searchvalue='" & left(CustName,19) & "'>" & left(CustName,19))
										Else
											Response.Write("<div class='col-lg-12' data-searchvalue='" & CustName & "'>" & CustName)
										End If
										
										Response.Write("</div>")
										Response.Write("</div>")
										
									End If


																		
									'Write second table row
									'**********************
									If DeliveryStatus = "Delivered" Then
										If AMorPM = "AM" Then
											trclass = "<div class='tr-completed AM-border-bottom'>"
										Else
											If PriorityDelivery = 1 Then
												trclass = "<tr class='tr-completed Priority-border-bottom'>"
											Else
												trclass = "<tr class='tr-completed tr-scheduled-bottom'>"
											End If
										End IF
									ElseIf DeliveryStatus = "No Delivery" Then
										If AMorPM = "AM" Then
											trclass = "<div class='tr-nodelivery AM-border-bottom'>"
										Else
											If PriorityDelivery = 1 Then
												trclass = "<tr class='tr-nodelivery Priority-border-bottom''>"
											Else
												trclass = "<tr class='tr-nodelivery tr-scheduled-bottom'>"
											End If
										End If
									Else
										If AMorPM = "AM" Then
											trclass = "<div class='tr-scheduled AM-border-bottom'>"
										Else
											If PriorityDelivery = 1 Then
												trclass = "<tr class='tr-scheduled Priority-border-bottom'>"
											Else
												trclass = "<tr class='tr-scheduled tr-scheduled-bottom'>"
											End If
										End If
									End If
									
									If len(CustName) > 19 then Cnam = left(CustName,19) Else Cnam = CustName
										
  								End If
			
								
								rs_DeliveryBoardDet.movenext
								NumLines = NumLines + 1
								
							Loop
							
							'Make all boxes even
							If NumLines < MaxNumberOfDeliveries() Then
								For x = 1 to MaxNumberOfDeliveries() - NumLines
									'Response.Write("<tr ><td>&nbsp;</td></tr>")
									'Response.Write("<tr ><td>&nbsp;</td></tr>")
								Next
							End IF
							
						End If
 	      
        
        Response.Write("</div></div></div>")
 		GridColumn = GridColumn +1
End Sub 

Set rs_DeliveryBoardSum = Nothing
cnn_DeliveryBoardSum.Close
Set cnn_DeliveryBoardSum = Nothing
%>	

</div></div></div></div>

 <script type="text/javascript">
 	$(document).ready(function(){
 	
 	    	$("#filter").keyup(function(e){
		
				// Retrieve the input field text and reset the count to zero
	        	var filter = $(this).val(), count = 0;

	        	 // Search through the content
	        	$(".col-lg-hide").each(function(){
	 
		           // If the list item does not contain the text phrase fade it out
		           if ($(this).text().search(new RegExp(filter, "i")) < 0){
		                $(this).fadeOut();
						$("button").attr("aria-expanded","true");
		 				$('.col-lg-12 .collapse').collapse('show');
			        	$(".col-lg-2").each(function(){
					        if ($(this).text().search(new RegExp(filter, "i")) >= 0){
					           	$(this).addClass("findElement");	
					        }
					        else {
					        	$(this).removeClass("findElement");
					        }
					    });
						
		           }
		            
		          // Show the list item if the phrase matches and increase the count by 1  
		          else if ($(this).text().search(new RegExp(filter, "i")) >= 0) {
	                $(this).show();
	                count++;
					$("button").attr("aria-expanded","true");
	 				$('.col-lg-12 .collapse').collapse('show');
		        	$(".col-lg-2").each(function(){
				        if ($(this).text().search(new RegExp(filter, "i")) > 0){
					        $(this).addClass("findElement");
				        }
				        else if ($(this).text().search(new RegExp(filter, "i")) == 0){
				        	$(this).removeClass("findElement");
				        }
				    });
					 	 } 
				
							 
		 		  else {
	                $(this).show();
	                count++;
					$("button").attr("aria-expanded","false");
					$('.col-lg-12 .collapse').collapse('hide');
		        	$(".col-lg-2").each(function(){
				        if ($(this).text().search(new RegExp(filter, "i")) == 0){
				           		$(this).removeClass("findElement");
				        }
				    }); 

	
	
				}
				
				if ($(this).text().search(new RegExp(filter, "i")) == 0) {
	                $(this).fadeIn();
	                count++;
					$("button").attr("aria-expanded","false");
	 				$('.col-lg-12 .collapse').collapse('hide');
				 }	
				 
           });	
	});	
		
});

		
 
 </script>
<!-- eof same page search !-->

 

 <!-- same height titles  !-->
    <script type="text/javascript" src="<%= BaseURL %>js/grids.js"></script>

<script type="text/javascript">
	jQuery(function($) {
		$('.col-lg-2-border').responsiveEqualHeightGrid();	
 	});
</script>
    <!-- eof same height titles !-->


  
<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS BEGIN HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="deliveryBoardCommonModals.asp"-->

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS END HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="../inc/footer-deliveryBoard-vertical.asp"-->
