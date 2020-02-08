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
<!--#include file="../inc/jquery_table_search.asp"-->
<!--#include file="../inc/InSightFuncs_Routing.asp"-->

<script type="text/javascript" src="<%= BaseURL %>js/doublescroll/jquery.doubleScroll.js"></script>

<link href="deliveryBoardPlanner.css" rel="stylesheet"> 

<style>
	
	.li-completed{
		<% Response.Write("background:" & DelBoardCompletedColor & " !important;") %>
	}
	
	.li-nodelivery{
		<% Response.Write("background:" & DelBoardSkippedColor & " !important;") %>
	}
	
	.li-nextstop{
		<% Response.Write("background:" & DelBoardNextStopColor & " !important;") %>
	}
		
	.li-scheduled{
		<% Response.Write("background:" & DelBoardScheduledColor & ";") %>
		border: 1px solid #e5e5e5;
	}

	.AM-border{
		<% Response.Write("border: 3px solid " & DelBoardAMColor & ";") %>
	}

	.Priority-border{
		<% Response.Write("border: 3px solid " & DelBoardPriorityColor & ";") %>
	}

	.li-user-alert{
		<% Response.Write("background:" & DelBoardUserAlertColor & ";") %>
	}
	
	.li-user-alert{
		border: 1px solid #000000;
	}
	 
	 .li-border-line{
		 border-bottom: 1px solid #ccc;
	 }
	 
	 .li-border-line-red{
		 border-bottom: 1px solid #999;
	 }
	
	 

</style> 
  
<script type="text/javascript">


	$(document).ready(function() {


		$('.double-scroll').doubleScroll({
		  scrollCss: {                
		    'overflow-x': 'auto',
		    'overflow-y': 'hidden'
		  },
		  contentCss: {
		    'overflow-x': 'auto',
		    'overflow-y': 'hidden'
		  },
		});     
		
		
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
						data: "action=TurnOffNagAlertsForDeliveryBoardDriver&driverUserNo=" + encodeURIComponent(driverUserNo),
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
						data: "action=TurnOnNagAlertsForDeliveryBoardDriver&driverUserNo=" + encodeURIComponent(driverUserNo),
						success: function(response)
						 {
			             }
					});
			  }
			  
		});	
       		
		// Sortable and connectable lists with visual helper
		$('#delboardScheduler .sortable-list').sortable({
			connectWith: '#delboardScheduler .sortable-list',
			placeholder: 'placeholder',
			containment: '#delboardScheduler',
			items: '> :not(.nodragorsort)',
						   
			   start: function(e, ui) {
			        // creates a temporary attribute on the element with the old index
			        $(this).attr('data-previndex', ui.item.index());
			        $(this).attr('data-previndex', ui.item.index());

			    },
		       receive: function(e, ui) {
			        var invoice = ui.item.attr('data-inv-number');
			        alert(invoice);
				    
			    	$.ajax({
						type:"POST",
						url: "../inc/InSightFuncs_AjaxForRoutingModals.asp",
						cache: false,
						data: "action=CheckDeliveryStatus&invoiceNo=" + encodeURIComponent(invoice),
						success: function(response)
						 {
						 	if (response == '' || response == null) {
						 	
						    	$.ajax({
								type:"POST",
								url: "../inc/InSightFuncs_AjaxForRoutingModals.asp",
								cache: false,
								data: "action=CheckDeliveryIsNextStop&invoiceNo=" + encodeURIComponent(invoice),
								success: function(response)
								 {
								 	if (response == "True") {
								 		ui.sender.sortable("cancel");
										swal({
										    title: "Planning Error",
										    text: "This invoice has already been marked by the driver as the next stop and cannot be modified.",
										    confirmButtonColor: "#337ab7",
										    confirmButtonText: 'OK'
										});
								 	}
								 	else {
								 		//alert("Not Delivered and Not The Next Stop You may proceed.");

								 	}
					             }
								});
	
						 	}
						 	else{
						 	
						 		ui.sender.sortable("cancel");
								swal({
								    title: "Planning Error",
								    text: "This invoice has already been marked by the driver and cannot be modified.",
								    confirmButtonColor: "#337ab7",
								    confirmButtonText: 'OK'
								});

						 	}
			             }
					});       
		       },			    
			   update: function(e, ui) {
			   
			        // gets the new and old index then removes the temporary attribute
			        var newIndex = ui.item.index();
			        var oldIndex = $(this).attr('data-previndex');
			        
			        var invoiceNo = ui.item.attr('data-inv-number');
			        var destinationSeqNo = ui.item.index();
			        var destinationTruck = ui.item.closest('ul').attr('id');
			        
			        destinationSeqNo = destinationSeqNo + 1
			        
			        //if (destinationSeqNo == 0) {
			        	//destinationSeqNo = 1;
			        //}
			        
			        alert("New Index: " + destinationSeqNo);
			        alert("New Truck Number: " + ui.item.closest('ul').attr('id'));
			          			        
			        $(this).removeAttr('data-previndex');
			        
			    	$.ajax({
						type:"POST",
						url: "../inc/InSightFuncs_AjaxForRoutingModals.asp",
						cache: false,
						data: "action=ChangeDeliveryPlanningBoard&invoiceNo=" + encodeURIComponent(invoiceNo) +"&destinationTruck=" + encodeURIComponent(destinationTruck) + "&destinationSeqNo=" + encodeURIComponent(destinationSeqNo),
						success: function(response)
						 {
							//swal({
							    //title: "Planning Success",
							    //text: "Invoice Successfully Moved.",
							    //confirmButtonColor: "#337ab7",
							    //confirmButtonText: 'OK'
							//});
			             }
					});
			        
			    }		
		});
			
	      $('#save-button').click(function(event) {  
	         $('#print-order').html("");
	         //$('.sortable-list').each(function(index) { 
	         
	         
	         
	 			$('.sortable-item').each(function(index) {
	 			
	 			var truck = $(this).data('truck-number');
	 			var inv = $(this).data('inv-number');
	            //$('#print-order').append($(this).html()+"<br/>");
	            if (truck !== '') {
	            	$('#print-order').append("Truck: " + truck + " Inv:" + inv + "<br/>");
	            }
	
	         });
	      });
	      
	      
      		$('#deliveryBoardSetAlertModal').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myInvoiceNumber = $(e.relatedTarget).data('invoice-number');
		    var myCustomerName = $(e.relatedTarget).data('customer-name');	
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtInvoiceNumber"]').val(myInvoiceNumber);
		    	    
		    var $modal = $(this);
	
    		$modal.find('#myDeliveryBoardLabel').html("Delivery Alert for " + myCustomerName + " - Invoice #" + myInvoiceNumber);
    		
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForRoutingModals.asp",
				cache: false,
				data: "action=GetContentForDeliveryBoardAlertModal&myInvoiceNumber=" + encodeURIComponent(myInvoiceNumber),
				success: function(response)
				 {
	               	 $modal.find('#deliveryBoardModalContent').html(response);
	             },
	             failure: function(response)
				 {
				   $modal.find('#deliveryBoardModalContent').html("Failed");
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

<div class="delboard-section" id="ex-1-3">
<div class="delboard-section-header">
<h3 class="delboard-h3 delboard-title">Delivery Board Planner</h3>
<div class="delboard-description">Drag and drop deliveries within and between trucks. A visual helper is displayed indicating where the delivery will be positioned if dropped.</div>
</div>


<%
MaxNumberOfDeliveriesToday = MaxNumberOfDeliveries()

'********************************************************************************************************************
'READ TRUCK FILE TO SEE IF A SORT ORDER HAS BEEN SET BY THE USER
'********************************************************************************************************************

trucksContainedInTextFile = False
dim fs,t,truckorder
truckorder=0
set fs=Server.CreateObject("Scripting.FileSystemObject")
filename = Server.MapPath(".")&"\truckorder\"&Session("Userno")&".txt"
If fs.FileExists(filename) Then
	set t=fs.OpenTextFile(filename,1,false)
	truckorder=t.ReadLine
	t.close
	trucksContainedInTextFile = True
End If

Dim truckorder_arr
truckorder_arr=split(truckorder,",")

%>

<%
	
TruckCount = 0
totalColumns = 0

Set cnn_DeliveryBoardTotalColumnCount = Server.CreateObject("ADODB.Connection")
cnn_DeliveryBoardTotalColumnCount.open (Session("ClientCnnString"))
Set rs_DeliveryBoardTotalColumnCount = Server.CreateObject("ADODB.Recordset")
rs_DeliveryBoardTotalColumnCount.CursorLocation = 3 

If trucksContainedInTextFile = True Then

	For Each TruckNumberInFile In truckorder_arr
	
		SQL_DeliveryBoardTotalColumnCount = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard WHERE TruckNumber = '" & TruckNumberInFile & "'"
		Set rs_DeliveryBoardTotalColumnCount = cnn_DeliveryBoardTotalColumnCount.Execute(SQL_DeliveryBoardTotalColumnCount)
		
		If not rs_DeliveryBoardTotalColumnCount.EOF Then
			
			Do While NOT rs_DeliveryBoardTotalColumnCount.EOF

				If DelBoardIgnoreThisRoute(rs_DeliveryBoardTotalColumnCount("TruckNumber")) <> True Then
					TruckCount = TruckCount + 1
				End If
				rs_DeliveryBoardTotalColumnCount.MoveNext
			Loop 
		End If		
	Next
	
Else

	SQL_DeliveryBoardTotalColumnCount = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard"
	Set rs_DeliveryBoardTotalColumnCount = cnn_DeliveryBoardTotalColumnCount.Execute(SQL_DeliveryBoardTotalColumnCount)
	
	If not rs_DeliveryBoardTotalColumnCount.EOF Then
		
		Do While NOT rs_DeliveryBoardTotalColumnCount.EOF
			If DelBoardIgnoreThisRoute(rs_DeliveryBoardTotalColumnCount("TruckNumber")) <> True Then
				TruckCount = TruckCount + 1
			End If
			rs_DeliveryBoardTotalColumnCount.MoveNext
		Loop
	End If
End If


containerWidth = 175 * TruckCount

%>

<div class="horizontal-layout">
<div class="double-scroll">
 <table>
    <tbody>
        <tr>
           <td>
            <div class="delboard-section-content" id="delboardContent">
			<div id="delboardScheduler" style="width:<%= containerWidth %>px !important;">  
    		<!-- <div id="delboardScheduler"> -->
<%
Set cnn_DeliveryBoardDisplay = Server.CreateObject("ADODB.Connection")
cnn_DeliveryBoardDisplay.open (Session("ClientCnnString"))
Set rs_DeliveryBoardDisplay = Server.CreateObject("ADODB.Recordset")
rs_DeliveryBoardDisplay.CursorLocation = 3 


%>
<ul id="sortableTrucks" class="flex-container-truck flex-row" style="margin-left: -30px;">
<%	
		

'********************************************************************************************************************
'IF TRUCK FILE EXISTS, WRITE TRUCKS OUT OF TRUCK FILE
'********************************************************************************************************************
If trucksContainedInTextFile = True Then
	
	truckCount = 1
	'Write ordered trucks
	For Each TruckNumberInFile In truckorder_arr
	
		SQL_DeliveryBoardDisplay = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard WHERE TruckNumber = '" & TruckNumberInFile & "'"
		Set rs_DeliveryBoardDisplay = cnn_DeliveryBoardDisplay.Execute(SQL_DeliveryBoardDisplay)
		
		If not rs_DeliveryBoardDisplay.EOF Then
	
			Do While NOT rs_DeliveryBoardDisplay.EOF 
			
				If DelBoardIgnoreThisRoute(rs_DeliveryBoardDisplay("TruckNumber")) <> True Then
				
					TruckNumber = rs_DeliveryBoardDisplay("TruckNumber")
					
					DriverUserNo = Trim(GetUserNumberByTruckNumber(TruckNumber))
				
					If userIsArchived(DriverUserNo) = False AND userIsEnabled(DriverUserNo) = True Then
					
						%>
		                <li data-TruckNumber="<%= TruckNumber %>" class="singletruck">
							<ul class="column left sortableTruck">
								<li>
								
									<% Call TruckNumberWrite(TruckNumber) %>
									<% rs_DeliveryBoardDisplay.Movenext %>
							    </li>
					    	</ul>
		        		</li>
						<%
					End If
				Else
					rs_DeliveryBoardDisplay.Movenext	
				End If
			
			truckCount = truckCount + 1
			Loop
		End If
		
	Next

End If

'********************************************************************************************************************
'ELSE IF TRUCK FILE DOES NOT EXIT, WRITE TRUCKS OUT OF RT_DeliveryBoard DIRECTLY
'********************************************************************************************************************

If trucksContainedInTextFile = False Then
	
	truckCount = 1
	'Lets get all the trucks not in order
	SQL_DeliveryBoardDisplay = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard ORDER BY TruckNumber"
	Set rs_DeliveryBoardDisplay = cnn_DeliveryBoardDisplay.Execute(SQL_DeliveryBoardDisplay)
	
	If not rs_DeliveryBoardDisplay.EOF Then

		Do While NOT rs_DeliveryBoardDisplay.EOF 
		
			If DelBoardIgnoreThisRoute(rs_DeliveryBoardDisplay("TruckNumber")) <> True Then
			
				TruckNumber = rs_DeliveryBoardDisplay("TruckNumber")
				
				DriverUserNo = Trim(GetUserNumberByTruckNumber(TruckNumber))
			
				If userIsArchived(DriverUserNo) = False AND userIsEnabled(DriverUserNo) = True Then
				
					%>
	                <li data-TruckNumber="<%= TruckNumber %>" class="singletruck">
						<ul class="column left sortableTruck">
							<li>
							
								<% Call TruckNumberWrite(TruckNumber) %>
								<% rs_DeliveryBoardDisplay.Movenext %>
						    </li>
				    	</ul>
	        		</li>
					<%
				End If
			Else
				rs_DeliveryBoardDisplay.Movenext	
			End If
		
		truckCount = truckCount + 1
		Loop
	End If

End If

%>
</ul>


	
<div class="clearer">&nbsp;</div>
</div></div>
</tr>
</tbody>
</table>

<input type="submit" value="Submit" id="save-button"/>
<div id="print-order">
</div>

</div>

</div>
</div>

<!-- **************************************************************************************************************************** -->
<!-- BEGIN FUNCTION THAT WRITES A SINGLE TRUCK TO THE DELIVERY BOARD !-->
<!-- **************************************************************************************************************************** -->

<% Sub TruckNumberWrite(TruckNumber) %>

	<h4 class="delboard-h4">

	<span class="trucknumber">Route: <%= TruckNumber %></span>

	<span class="drivername"><%= GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(TruckNumber))) %></span>
	
	<a class="btn-move" id="moveTruck" href="#" style="position: relative; top:-55px; float: right;"><i class="fa fa-arrows"></i></a>

	<%
	'**************************************************************************
	'show nag alerts to admins or route managers only
	'**************************************************************************
	
	If userIsAdmin(Session("userNo")) OR userIsRouteManager(Session("userNo")) Then
	
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
				    <button class="<%= buttonClassRed %>" id="<%= DriverUserNo %>OFF">OFF</button>
				  </div>
			<% Else %>
				  <div class="btn-group btn-toggle" id="<%= DriverUserNo %>">
				    <button class="<%= buttonClassGreen %>" id="<%= DriverUserNo %>ON">ON</button>
				    <button class="<%= buttonClassRed %>" id="<%= DriverUserNo %>OFF">NAG OFF</button>
				  </div>						
			<% End If %>
			<%
		Else
			%><div class="btn-group btn-toggle">Nag Alerts Off</div><%
		End If
	End If
	'**************************************************************************

	%> 
	</h4>
	
	<ul class="sortable-list" id="<%= TruckNumber %>">
	
	<%'Get all the tickets for this truck
		
	Set cnn_DeliveryTickets = Server.CreateObject("ADODB.Connection")
	cnn_DeliveryTickets.open (Session("ClientCnnString"))
	Set rs_DeliveryBoardDeliveryTickets = Server.CreateObject("ADODB.Recordset")
	rs_DeliveryBoardDeliveryTickets.CursorLocation = 3 
	SQL_DeliveryTickets = "SELECT * FROM RT_DeliveryBoard "
	SQL_DeliveryTickets = SQL_DeliveryTickets & "WHERE TruckNumber = '" & TruckNumber  & "' "
	
	If DelBoardDontUseStopSequencing() = False Then
	    SQL_DeliveryTickets = SQL_DeliveryTickets & "Order By SequenceNumber, CustNum" 
	Else
	    SQL_DeliveryTickets = SQL_DeliveryTickets & "Order By CustNum" 
	End If
	
	Set rs_DeliveryBoardDeliveryTickets = cnn_DeliveryTickets.Execute(SQL_DeliveryTickets)
	    
	If NOT rs_DeliveryBoardDeliveryTickets.EOF Then
	
		NumLines = 0
		
		Do While NOT rs_DeliveryBoardDeliveryTickets.EOF
							
			TipText = ""
	
			CustomerName = rs_DeliveryBoardDeliveryTickets("CustName")
			IvsNum = rs_DeliveryBoardDeliveryTickets("IvsNum") 
			CustNum = rs_DeliveryBoardDeliveryTickets("CustNum")
			DeliveryStatus = rs_DeliveryBoardDeliveryTickets("DeliveryStatus")
			AMorPM = rs_DeliveryBoardDeliveryTickets("AMorPM") 
			PriorityDelivery = rs_DeliveryBoardDeliveryTickets("Priority")
				
			If CustNum = GetNextCustomerStopByTruck(TruckNumber) Then
				
				If AMorPM = "AM" Then
				
					If DeliveryAlertSet(IvsNum,Session("UserNo")) Then
						%>
						<li class="sortable-item li-user-alert AM-border" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardSetAlertModal" data-tooltip="true" data-title="Set Delivery Alert" style="cursor:pointer;">
						<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
							<i class="fa fa-star" aria-hidden="true"></i>
						<% End If %>
						<%= IvsNum %>
						<%
						TipText = "Alert when " &  DeliveryAlertCondition(IvsNum,Session("UserNo"))  
						%>
						<span class="custid"><div class="alarm-bell"><%= CustNum %><span class="alert-pop-up"><%= TipText  %></span></div></span><br>
						<span class="customer"><%= CustomerName %></span></a></li>
						<%
					Else 
						%>
						<li class="sortable-item li-nextstop AM-border nodragorsort" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardSetAlertModal" data-tooltip="true" data-title="Set Delivery Alert" style="cursor:pointer;">
						<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
							<i class="fa fa-star" aria-hidden="true"></i>
						<% End If %>
						<%= IvsNum %>
						<span class="custid"><%= CustNum %></span><br>
						<span class="customer"><%= CustomerName %></span></a></li>
						<%
					End If
	
				Else
					If DeliveryAlertSet(IvsNum,Session("UserNo")) Then
					
						If PriorityDelivery = 1 Then %>
							<li class="sortable-item li-user-alert li-scheduled Priority-border" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<% Else %>
							<li class="sortable-item li-user-alert li-scheduled" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<% End If %>
						
						<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardSetAlertModal" data-tooltip="true" data-title="Set Delivery Alert" style="cursor:pointer;">
						<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
							<i class="fa fa-star" aria-hidden="true"></i>
						<% End If %>
						<%= IvsNum %>
						<%
						TipText = "Alert when " &  DeliveryAlertCondition(IvsNum,Session("UserNo"))  
						%>
						<span class="custid"><div class="alarm-bell"><%= CustNum %><span class="alert-pop-up"><%= TipText %></span></div></span><br>
						<span class="customer"><%= CustomerName %></span></a></li>
						<%
					Else
					
						
						If PriorityDelivery = 1 Then %>
							<li class="sortable-item li-nextstop li-scheduled nodragorsort Priority-border" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<% Else %>
							<li class="sortable-item li-nextstop li-scheduled nodragorsort" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<% End If %>
						
						
						<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardSetAlertModal" data-tooltip="true" data-title="Set Delivery Alert" style="cursor:pointer;">
						<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
							<i class="fa fa-star" aria-hidden="true"></i>
						<% End If %>
						<%= IvsNum %>
						<span class="custid"><%= CustNum %></span><br>
						<span class="customer"><%= CustomerName %></span></a></li>
						<%
					End If
					
				End If
				
			Else
		
				If DeliveryStatus = "Delivered" Then
				
					If AMorPM= "AM" Then
						%>
						<li class="sortable-item li-completed AM-border nodragorsort" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;">
						<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
							<i class="fa fa-star" aria-hidden="true"></i>
						<% End If %>
						<%= IvsNum %>
						<span class="custid"><%= CustNum %></span><br>
						<span class="customer"><%= CustomerName %></span></a></li>
					<% Else 
					
						If PriorityDelivery = 1 Then %>
							<li class="sortable-item li-completed li-scheduled nodragorsort Priority-border" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<% Else %>
							<li class="sortable-item li-completed li-scheduled nodragorsort" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<% End If %>
					
						<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;">
						<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
							<i class="fa fa-star" aria-hidden="true"></i>
						<% End If %>
						<%= IvsNum %>
						<span class="custid"><%= CustNum %></span><br>
						<span class="customer"><%= CustomerName %></span></a></li>
						<%
					End If
					
				ElseIf DeliveryStatus = "No Delivery" Then
				
					If AMorPM = "AM" Then
						%>
						<li class="sortable-item li-nodelivery AM-border nodragorsort" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;">
						<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
							<i class="fa fa-star" aria-hidden="true"></i>
						<% End If %>
						<%= IvsNum %>
						<span class="custid"><%= CustNum %></span><br>
						<span class="customer"><%= CustomerName %></span></a></li>
						<%
					Else
						
						If PriorityDelivery = 1 Then %>
							<li class="sortable-item li-nodelivery li-scheduled nodragorsort Priority-border" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<% Else %>
							<li class="sortable-item li-nodelivery li-scheduled nodragorsort" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
						<% End If %>
						
						<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;">
						<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
							<i class="fa fa-star" aria-hidden="true"></i>
						<% End If %>
						<%= IvsNum %>
						<span class="custid"><%= CustNum %></span><br>
						<span class="customer"><%= CustomerName %></span></a></li>
						<%
					End If
				Else
					If AMorPM = "AM" Then
						If DeliveryAlertSet(IvsNum,Session("UserNo")) Then
							%>
							<li class="sortable-item li-user-alert AM-border" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
							<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardSetAlertModal" data-tooltip="true" data-title="Set Delivery Alert" style="cursor:pointer;">
							<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
								<i class="fa fa-star" aria-hidden="true"></i>
							<% End If %>
							<%= IvsNum %>
							<%
							TipText = "Alert when " & DeliveryAlertCondition(IvsNum,Session("UserNo")) 
							%>
							<span class="custid"><div class="alarm-bell"><%= CustNum %><span class="alert-pop-up"><%= TipText  %></span></div></span><br>
							<span class="customer"><%= CustomerName %></span></a></li>
							<%
						Else
							%>
							<li class="sortable-item li-scheduled AM-border" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
							<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardSetAlertModal" data-tooltip="true" data-title="Set Delivery Alert" style="cursor:pointer;">
							<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
								<i class="fa fa-star" aria-hidden="true"></i>
							<% End If %>
							<%= IvsNum %>
							<span class="custid"><%= CustNum %></span><br>
							<span class="customer"><%= CustomerName %></span></a></li>
							<%
						End If
					Else
						If DeliveryAlertSet(IvsNum,Session("UserNo")) Then
							
							
							If PriorityDelivery = 1 Then %>
								<li class="sortable-item li-user-alert Priority-border" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
							<% Else %>
								<li class="sortable-item li-user-alert" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
							<% End If %>
							
							<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardSetAlertModal" data-tooltip="true" data-title="Set Delivery Alert" style="cursor:pointer;">
							<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
								<i class="fa fa-star" aria-hidden="true"></i>
							<% End If %>
							<%= IvsNum %>
							<%
							TipText = "Alert when " & DeliveryAlertCondition(IvsNum,Session("UserNo")) 
							%>
							<span class="custid"><div class="alarm-bell"><%= CustNum %><span class="alert-pop-up"><%= TipText  %></span></div></span><br>
							<span class="customer"><%= CustomerName %></span></a></li>
							<%
						Else
							
							
							If PriorityDelivery = 1 Then %>
								<li class="sortable-item li-user-scheduled Priority-border" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
							<% Else %>
								<li class="sortable-item li-user-scheduled" id="<%= IvsNum %>" data-truck-number="<%= TruckNumber %>" data-inv-number="<%= IvsNum %>">
							<% End If %>
							
							<a data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= IvsNum %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardSetAlertModal" data-tooltip="true" data-title="Set Delivery Alert" style="cursor:pointer;">
							<% If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = IvsNum Then %>
								<i class="fa fa-star" aria-hidden="true"></i>
							<% End If %>
							<%= IvsNum %>
							<span class="custid"><%= CustNum %></span><br>
							<span class="customer"><%= CustomerName %></span></a></li>
							<%
						End If
					End If
				End If
		End If
				
		rs_DeliveryBoardDeliveryTickets.MoveNext
		NumLines = NumLines + 1
		Loop
		
		
		'Make all truck containers the same height
		If NumLines < MaxNumberOfDeliveriesToday Then
			For x = 1 to MaxNumberOfDeliveriesToday - NumLines
			%>
			<li class="sortable-item " data-truck-number="" data-inv-number="" ><a href="#">&nbsp;</a></li>
			<%	
			Next
		End If
	
	End If
		
%>					
		</ul>

<% End Sub %>
<!-- **************************************************************************************************************************** -->
<!-- END TRUCK WRITE FUNCTION !-->
<!-- **************************************************************************************************************************** -->



<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS BEGIN HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="deliveryBoardCommonModals.asp"-->

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS END HERE !-->
<!-- **************************************************************************************************************************** -->


<script type="text/javascript">
	function setSortable() {
		//$("#sortableList").sortable({ placeholder: "ui-state-highlight item col-lg-1 col-lg-cust", handle: ".delboard-h4", scrollSensitivity: 40, scrollSpeed: 60, update: function (event, ui) { saveSelection(); } });
		//$("#sortableList").disableSelection();
		
		 
        $('#sortableTrucks').sortable({
		    cursor: "move",
		    placeholder: 'placeholder-truck',
		    opacity: 0.7,
		    handle: ".btn-move",
		    scroll: true, 
		    scrollSensitivity: 100,
		    revert: true,
		    update: function (event, ui) { saveSelection(); }
		});	
		
	}
	function saveSelection() {
		var list = "";
		try {
			var sep = "";
			$("#sortableTrucks .singletruck").each(function () {
				list += "" + sep + $(this).attr("data-TruckNumber");
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


<!--#include file="../inc/footer-deliveryBoard.asp"-->
