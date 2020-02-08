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
		DelBoardInProgressColor = rs("DelBoardInProgressColor")			
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
If DelBoardPriorityColor = "" Then DelBoardPriorityColor = "#FF0000"
If IsNull(DelBoardPriorityColor) Then DelBoardPriorityColor = "#FF0000"
%>
<!--#include file="../inc/header-deliveryboard.asp"-->
<!--#include file="../inc/jquery_table_search.asp"-->
<!--#include file="../inc/InSightFuncs_Routing.asp"-->

<link href="deliveryBoard.css" rel="stylesheet"> 

<style>

	.tr-inprogress{
		<% Response.Write("background:" & DelBoardInProgressColor & ";") %>
	}
	
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
				data: "action=GetContentForDeliveryBoardOptionsModal&returnPage=routing/deliveryBoard.asp&invoiceNum=" + encodeURIComponent(myInvoiceNumber) + "&custID=" + encodeURIComponent(myCustID) + "&truckNum=" + encodeURIComponent(myTruckNumber),
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


<!--- eof on/off scripts !-->
 <script type="text/javascript" src="<%= BaseURL %>js/doublescroll/jquery.doubleScroll.js"></script>


<script type="text/javascript">
    $(document).ready(function(){

       $('.double-scroll').doubleScroll({
       		resetOnWindowResize: true
       	});
    });
</script>

<div class="horizontal-layout">
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
				SQL_DeliveryBoardSum = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard WHERE TruckNumber = " & TruckNumber
				'esponse.write(SQL_DeliveryBoardSum)
				Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL_DeliveryBoardSum)
				If not rs_DeliveryBoardSum.EOF Then
					Do While Not rs_DeliveryBoardSum.Eof
						If DelBoardIgnoreThisRoute(rs_DeliveryBoardSum("TruckNumber")) <> True Then 
						
							DriverUserNo = Trim(GetUserNumberByTruckNumber(rs_DeliveryBoardSum("TruckNumber")))
						
							If userIsArchived(DriverUserNo) = False AND userIsEnabled(DriverUserNo) = True Then
								Call TruckNumberWrite(rs_DeliveryBoardSum("TruckNumber"), GridColumn) 
							End If
							
						End If
						rs_DeliveryBoardSum.Movenext
					Loop
				End If
			Next

		'Lets get all the trucks not in order
		SQL_DeliveryBoardSum = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard WHERE TruckNumber NOT IN ('" & Replace(truckorder,",","','") & "')  ORDER BY TruckNumber"
		Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL_DeliveryBoardSum)
		
		If not rs_DeliveryBoardSum.EOF Then
			Do While Not rs_DeliveryBoardSum.Eof

				If DelBoardIgnoreThisRoute(rs_DeliveryBoardSum("TruckNumber")) <> True Then 
				
					DriverUserNo = Trim(GetUserNumberByTruckNumber(rs_DeliveryBoardSum("TruckNumber")))
				
					If userIsArchived(DriverUserNo) = False AND userIsEnabled(DriverUserNo) = True Then
						Call TruckNumberWrite(rs_DeliveryBoardSum("TruckNumber"), GridColumn) 
					End If

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
Sub TruckNumberWrite(TruckNumber, GridColumn) %>
		<td class='item col-lg-cust' TruckNumber='<%= TruckNumber %>'>
		<div class='item-box'>
		<div class='scrollable-title' style='position: relative;'>
		
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
					'First check to see if nags are off entirely for this user
					NagsON = False
					
					SQLUsers = "SELECT * FROM tblUsers Where UserNo = " & DriverUserNo 
					
					Set cnn_Users = Server.CreateObject("ADODB.Connection")
					cnn_Users.open (Session("ClientCnnString"))
					Set rsUsers = Server.CreateObject("ADODB.Recordset")
					rsUsers.CursorLocation = 3 
					'Response.write(SQLUsers)
					Set rsUsers = cnn_Users.Execute(SQLUsers)

					'ANY YES CONDITION TURNS THE BUTTON ON
					If Not rsUsers.EOF Then

						If rsUsers("userNextStopNagMessageOverride") = "Yes" Then NagsON = True
						If rsUsers("userNoActivityNagMessageOverride") = "Yes" Then NagsON = True
						
						If NagsON = False Then' only check if not already on
						
							If rsUsers("userNextStopNagMessageOverride") = "Use Global" or rsUsers("userNoActivityNagMessageOverride") = "Use Global" Then
							
								SQLGlobal = "SELECT * FROM Settings_Global "
								Set rsGlobal = Server.CreateObject("ADODB.Recordset")
								rsGlobal.CursorLocation = 3 
								Set rsGlobal = cnn_Users.Execute(SQLGlobal)
		
								If Not rsGlobal.EOF Then
									NoAct = rsGlobal("NoActivityNagMessageONOFF")
									NextSt = rsGlobal("NextStopNagMessageONOFF")
								End If
							
								Set rsGlobal = Nothing
							End If
							
							If NextSt  = 1 Then NagsON = True
							If NoAct = 1 Then NagsON = True
							
						End If
					
					End If
					
					Set rsUsers = Nothing
					cnn_Users.Close
					Set cnn_Users = Nothing
				End If

				
				If DriverUserNo <> "*Not Found*" Then
				
					If NagsOn = True Then
				
						If  DriverInNagSkipTable(DriverUserNo,"routingNoNextStop") = False AND DriverInNagSkipTable(DriverUserNo,"routingNoActivity") = False Then

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
						<% End If 
					
					Else
						%><div class="btn-group btn-toggle">Nags Off</div><%
					End If

				Else
					%><div class="btn-group btn-toggle">No User Setup</div><%
				End If
			End If
			'**************************************************************************
		
			%> 
			
		</div>

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
								
								'Response.write("PRIORITY: " & rs_DeliveryBoardDet("Priority") & "<br><br")
								
								PriorityDelivery = rs_DeliveryBoardDet("Priority")
								InvoiceNumber = rs_DeliveryBoardDet("IvsNum")
								CustName = rs_DeliveryBoardDet("CustName")
								CustID = rs_DeliveryBoardDet("CustNum")
								TruckNumber = rs_DeliveryBoardSum("TruckNumber")
								AMorPM = rs_DeliveryBoardDet("AMorPM")
								DeliveryStatus = rs_DeliveryBoardDet("DeliveryStatus")
							
								If CustID = GetNextCustomerStopByTruck(TruckNumber) Then
									
									If AMorPM = "AM" Then
									
										If DeliveryAlertSet(InvoiceNumber,Session("UserNo")) Then %>
											
											<tr class="tr-user-alert AM-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
											<%
											TipText = "Alert when " &  DeliveryAlertCondition(InvoiceNumber,Session("UserNo")) 
											
										Else
										
										
											If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then %>
											
												<tr class="tr-inprogress AM-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
												
											<% Else %>
											
												<tr class="tr-nextstop AM-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
												
											<% End If %>
										
										<%
										End If
										
										Response.Write("<td>" & InvoiceNumber & "</td>")
										Response.Write("<td>" & CustID & "</td></tr>")
										
										
										If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then
										
											Response.Write("<tr class='tr-inprogress AM-border-bottom'>")
											
										Else
										
											Response.Write("<tr class='tr-nextstop AM-border-bottom'>")
											
										End If

																				
										If len(CustName) > 19 then
											Response.Write("<td colspan='2'>" & left(CustName,19)) 
										Else
											Response.Write("<td colspan='2'>" & CustName) 
										End If
										
										Response.Write("</td></tr>")
										
									Else
									
										If PriorityDelivery = 1 Then
											%><tr class="tr-nextstop Priority-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
										Else
											%><tr class="tr-nextstop tr-scheduled-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
										End If
									
										
										If DeliveryAlertSet(InvoiceNumber,Session("UserNo")) Then
										
											If PriorityDelivery = 1 Then
												%><tr class="tr-user-alert Priority-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											Else
												%><tr class="tr-user-alert tr-scheduled-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%											
											End If
											
											TipText = "Alert when " &  DeliveryAlertCondition(InvoiceNumber,Session("UserNo")) 
											
											
										'''This is the ELSE for If DeliveryAlertSet(InvoiceNumber,Session("UserNo")) Then
										Else
										
											If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then %>
											
												<% If PriorityDelivery = 1 Then %>
													<tr class="tr-inprogress Priority-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-inprogress tr-scheduled-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
												<% End If %>
												
											<% Else %>
											
												<% If PriorityDelivery = 1 Then %>
													<tr class="tr-nextstop Priority-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-nextstop tr-scheduled-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">												
												<% End If %>
												
											<% End If
											
										End If
										
										Response.Write("<td>" & InvoiceNumber & "</td>")
										Response.Write("<td>" & CustID & "</td></tr>")
										
										If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then
										
											If PriorityDelivery = 1 Then
												Response.Write("<tr class='tr-inprogress Priority-border-bottom>")
											Else
												Response.Write("<tr class='tr-inprogress tr-scheduled-bottom'>")
											End If
											
										Else
										
											If PriorityDelivery = 1 Then
												Response.Write("<tr class='tr-nextstop Priority-border-bottom'>")
											Else
												Response.Write("<tr class='tr-nextstop tr-scheduled-bottom'>")
											End If
												
										End If
										
										If len(CustName) > 19 then
											Response.Write("<td colspan='2'>" & left(CustName,19))
										Else
											Response.Write("<td colspan='2'>" & CustName)
										End If
										Response.Write("</td></tr>")
									End If
									
								Else
								
									'Write first table row
									'**********************
									If DeliveryStatus = "Delivered" Then
									
									
										If AMorPM = "AM" Then
										
											%><tr class="tr-completed AM-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
										
										Else
										
											If PriorityDelivery = 1 Then
												%><tr class="tr-completed Priority-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
											Else
												%><tr class="tr-completed tr-scheduled-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
											End If
											
										End If
										
									ElseIf DeliveryStatus = "No Delivery" Then


										If AMorPM = "AM" Then
										
											%><tr class="tr-nodelivery AM-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
										
										Else
										
											If PriorityDelivery = 1 Then
												%><tr class="tr-nodelivery Priority-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
											Else
												%><tr class="tr-nodelivery tr-scheduled-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Info" style="cursor:pointer;"><%
											End If
											
										End If
										
									Else
										If AMorPM = "AM" Then
										
											If DeliveryAlertSet(InvoiceNumber,Session("UserNo")) Then
											
												%><tr class="tr-user-alert AM-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
														
												TipText = "Alert when " & DeliveryAlertCondition(InvoiceNumber,Session("UserNo")) 
												
											Else
											
												If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then %>
													<tr class="tr-inprogress AM-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-scheduled AM-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
												<%	
												End If
												
											End If
											
										Else
										
											If DeliveryAlertSet(InvoiceNumber,Session("UserNo")) Then

												If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then 
												
													If PriorityDelivery = 1 Then
														%><tr class="tr-user-alert Priority-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
													Else
														%><tr class="tr-user-alert tr-user-alert-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
													End If
												
													TipText = "Alert when " & DeliveryAlertCondition(InvoiceNumber,Session("UserNo")) 
													
												End If

											Else
											
												If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then 
												
													If PriorityDelivery = 1 Then
														%><tr class="tr-inprogress Priority-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
													Else
														%><tr class="tr-inprogress tr-scheduled-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
													End If
												
												Else
												
													If PriorityDelivery = 1 Then
														%><tr class="tr-scheduled Priority-border-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
													Else
														%><tr class="tr-scheduled tr-scheduled-top" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
													End If
												
												End If													
																																			
											End If
										End If
									End If
									
									Response.Write(trclass)

									If TipText <> "" Then
										If GetLastInvoiceMarkedByTruckNumber(rs_DeliveryBoardSum("TruckNumber")) = InvoiceNumber Then
											Response.Write("<td>" & InvoiceNumber & "<i class='fa fa-star' aria-hidden='true'></i></td>")
										Else
											Response.Write("<td>" & InvoiceNumber & "</td>")
										End If
										Response.Write("<td><div class='alarm-bell'>" & CustID & "<span class='alert-pop-up'>" & TipText  & "</span></div></td></tr>")
									Else
										If GetLastInvoiceMarkedByTruckNumber(rs_DeliveryBoardSum("TruckNumber")) = InvoiceNumber Then
											Response.Write("<td>" & InvoiceNumber & "<i class='fa fa-star' aria-hidden='true'></i></td>")
										Else
											Response.Write("<td>" & InvoiceNumber & "</td>")
										End If
										Response.Write("<td>" & CustID & "</td></tr>")
									End If


																		
									'Write second table row
									'**********************
									If DeliveryStatus = "Delivered" Then
										
										If AMorPM = "AM" Then
										
											trclass = "<tr class='tr-completed AM-border-bottom'>"
											
										Else
										
											If PriorityDelivery = 1 Then
												trclass = "<tr class='tr-completed Priority-border-bottom'>"
											Else
												trclass = "<tr class='tr-completed tr-scheduled-bottom'>"
											End If
											
										End If
										
									ElseIf DeliveryStatus = "No Delivery" Then


										If AMorPM = "AM" Then
										
											trclass = "<tr class='tr-nodelivery AM-border-bottom'>"
											
										Else
										
											If PriorityDelivery = 1 Then
												trclass = "<tr class='tr-nodelivery Priority-border-bottom''>"
											Else
												trclass = "<tr class='tr-nodelivery tr-scheduled-bottom'>"
											End If
											
										End If

									Else
										If AMorPM = "AM" Then

											If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then
												trclass = "<tr class='tr-inprogress AM-border-bottom'>"
											Else
												trclass = "<tr class='tr-scheduled AM-border-bottom'>"
											End If

										Else

											If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then
											
												If PriorityDelivery = 1 Then
													trclass = "<tr class='tr-inprogress Priority-border-bottom'>"
												Else
													trclass = "<tr class='tr-inprogress tr-scheduled-bottom'>"
												End If
												
											Else
											
												If PriorityDelivery = 1 Then
													trclass = "<tr class='tr-scheduled Priority-border-bottom'>"
												Else
													trclass = "<tr class='tr-scheduled tr-scheduled-bottom'>"
												End If
												
											End If

										End If
									End If
									
									Response.Write(trclass)
									
									
									If len(CustName) > 19 then Cnam = left(CustName,19) Else Cnam = CustName
									Response.Write("<td colspan='2'>" & Cnam)
									Response.Write("<span class='alarm-bell'>")
									Response.Write("</span>")
									Response.Write("</td></tr>")
 								End If
			
								
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
<!-- tooltip JS !-->
<script type="text/javascript">
	$(function () {
	  $('[data-toggle="tooltip"]').tooltip()
	})
</script>
<!-- eof tooltip JS !-->


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
<!-- MODALS FOR DELIVERY ALERTS BEGIN HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="deliveryBoardCommonModals.asp"-->

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS END HERE !-->
<!-- **************************************************************************************************************************** -->


<!--#include file="../inc/footer-deliveryBoard.asp"-->
