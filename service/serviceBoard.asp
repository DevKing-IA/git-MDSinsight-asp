<!--#include file="../inc/header-serviceboard.asp"-->
<!--#include file="../inc/jquery_table_search.asp"-->
<!--#include file="../inc/InSightFuncs_Service.asp"-->
<!--#include file="../inc/InSightFuncs_Routing.asp"-->
<!--#include file="../inc/InSightFuncs_AR_AP.asp"-->
<script type="text/javascript">
	$(document).ready(function() {

		$('.btn-toggle').click(function() {
		
			  technicianUserNo = $(this).attr("id");

			  if ($(this).find('.btn-nag-on').size() > 0) {
			  				    
				    $("#" + technicianUserNo + "ON").removeClass('btn-nag-on');
				    $("#" + technicianUserNo + "ON").addClass('btn-default');
				    $("#" + technicianUserNo + "ON").addClass('active');
				    
				    $("#" + technicianUserNo + "OFF").removeClass('btn-default');
				    $("#" + technicianUserNo + "OFF").removeClass('active');
				    $("#" + technicianUserNo + "OFF").addClass('btn-nag-off');
				    				    
				    $("#" + technicianUserNo + "ON").html("ON");
				    $("#" + technicianUserNo + "OFF").html("NAG OFF");	
				    
			    	$.ajax({
						type:"POST",
						url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
						cache: false,
						data: "action=TurnOnNagAlertsForFieldServiceTechnicianKiosk&technicianUserNo=" + encodeURIComponent(technicianUserNo),
						success: function(response)
						 {
			             }
					});
			  }
			  else {

				    $("#" + technicianUserNo + "OFF").removeClass('btn-nag-off');
				    $("#" + technicianUserNo + "OFF").addClass('btn-default');
				    $("#" + technicianUserNo + "OFF").addClass('active');
				    
				    $("#" + technicianUserNo + "ON").removeClass('btn-default');
				    $("#" + technicianUserNo + "ON").removeClass('active');
				    $("#" + technicianUserNo + "ON").addClass('btn-nag-on');
				    
				    $("#" + technicianUserNo + "ON").html("NAG ON");
				    $("#" + technicianUserNo + "OFF").html("OFF");			    
				    
			    	$.ajax({
						type:"POST",
						url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
						cache: false,
						data: "action=TurnOffNagAlertsForFieldServiceTechnicianKiosk&technicianUserNo=" + encodeURIComponent(technicianUserNo),
						success: function(response)
						 {
			             }
					});
			  }
			  
		});	
		

			
		$('#serviceBoardTicketOptionsModal').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myTicketNumber = $(e.relatedTarget).data('invoice-number');
		    var myCustomerName = $(e.relatedTarget).data('customer-name');	
		    var myCustID = $(e.relatedTarget).data('customer-id');
		    var myUserNo = $(e.relatedTarget).data('user-no');
		    
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtTicketNumber"]').val(myTicketNumber);
		    $(e.currentTarget).find('input[name="txtCustID"]').val(myCustID);
		    $(e.currentTarget).find('input[name="txtUserNo"]').val(myUserNo);
		    	    
		    var $modal = $(this);
	
    		$modal.find('#serviceBoardLabel').html("Service Options For " + myCustomerName + " - Ticket  #" + myTicketNumber);
    		
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetContentForServiceBoardTicketOptionsModal&returnURL=service/serviceBoard.asp&memoNum=" + encodeURIComponent(myTicketNumber) + "&custID=" + encodeURIComponent(myCustID) + "&userNo=" + encodeURIComponent(myUserNo),
				success: function(response)
				 {
	             	$modal.find('#ServiceBoardTicketOptionsModalContent').html(response);
	             },
	             failure: function(response)
				 {
				 	$modal.find('#ServiceBoardTicketOptionsModalContent').html("Failed");
	             }
			});
		    
		});
		
			
		$('#serviceBoardSetAlertModal').on('show.bs.modal', function(e) {
		
			//close the service ticket options modal where we came from
			$('#serviceBoardTicketOptionsModal').modal('hide');
		    	
		    //get data-id attribute of the clicked prospect
		    var myInvoiceNumber = $(e.relatedTarget).data('invoice-number');
		    var myCustomerName = $(e.relatedTarget).data('customer-name');	
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtInvoiceNumber"]').val(myInvoiceNumber);
		    	    
		    var $modal = $(this);
	
    		$modal.find('#serviceBoardSetAlertModalLabel').html("Delivery Information for " + myCustomerName + " - Invoice #" + myInvoiceNumber);
		});




		$('#serviceBoardXferModal').on('show.bs.modal', function(e) {
		
			//close the service ticket options modal where we came from
			$('#serviceBoardTicketOptionsModal').modal('hide');
		
		    //get data-id attribute of the clicked service ticket
		    var myTicketNumber = $(e.relatedTarget).data('invoice-number');
		    var myCustomerName = $(e.relatedTarget).data('customer-name');	
		    var myCustID = $(e.relatedTarget).data('customer-id');
		    var myUserNo = $(e.relatedTarget).data('user-no');
		    
		    //populate the textboxes with the id of the clicked service ticket
		    $(e.currentTarget).find('input[name="txtTicketNumber"]').val(myTicketNumber);
		    $(e.currentTarget).find('input[name="txtCustID"]').val(myCustID);
		    $(e.currentTarget).find('input[name="txtUserNo"]').val(myUserNo);
		    	    
		    var $modal = $(this);
	    		
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetTitleForServiceBoardTransferRedispatchModal&memoNum=" + encodeURIComponent(myTicketNumber) + "&custID=" + encodeURIComponent(myCustID) + "&userNo=" + encodeURIComponent(myUserNo),
				success: function(response)
				 {
	               	 $modal.find('#ServiceBoardXferModalTitle').html(response);
	             },
	             failure: function(response)
				 {
				 	$modal.find('#ServiceBoardXferModalTitle').html("Failed");
	             }
			});
    		
    		
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetContentForServiceBoardTransferRedispatchModal&memoNum=" + encodeURIComponent(myTicketNumber) + "&custID=" + encodeURIComponent(myCustID) + "&userNo=" + encodeURIComponent(myUserNo),
				success: function(response)
				 {
	             	$modal.find('#ServiceBoardXferModalContent').html(response);
	             },
	             failure: function(response)
				 {
				 	$modal.find('#ServiceBoardXferModalContent').html("Failed");
	             }
			});
		    
		});
		
	});
</script>


<!-- DYNAMIC FORM !-->
<style type="text/css">

	body{
		overflow-x:hidden;
	}
	
	.wrapper {
	 	margin-top:25px;
	}
	
   .delivery-status{
	   margin-top: 0px;
	   color: #fff;
    }
       
	.bs-example-modal-lg-customize .row{
		margin-bottom: 10px;
	 	width: 100%;
		overflow: hidden;
	}
	
	.bs-example-modal-lg-customize .left-column{
		background: #eaeaea;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}
	
	.bs-example-modal-lg-customize .left-column h4{
		margin-top: 0px;
	}
	
	.bs-example-modal-lg-customize .right-column{
		background: #fff;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}
	
	.bs-example-modal-lg-customize .row{
		margin-bottom: 10px;
	 	width: 100%;
		overflow: hidden;
	}
	
	.bs-example-modal-lg-customize .left-column{
		background: #eaeaea;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}

	.bs-example-modal-lg-customize .left-column h4{
		margin-top: 0px;
	}
	
	.bs-example-modal-lg-customize .right-column{
		background: #fff;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}
	

	.modal-body{
		font-size:14px;
	}
	
	.modal-body label{
		font-weight:bold;
		padding-top:10px;
	}
	
	.modal-body .row-line{
		width:100%;
		display:inline-block;
		margin:0px 0px 10px 0px;
	}
	
	.modal-body .row-line .multiselect,.textarea{
		min-height:110px;
		max-height:110px;
		margin-bottom:5px;
	}
	
	.modal-body .row-line .right{
		text-align:right;
	}
  
   .delivery-status h2{
	   margin:8px 0px 0px 0px;
	   line-height:1;
   }
	
	table thead a{
		color: #000;
	}

	
	
	.tr-awaiting-dispatch{
		<% Response.Write("background:" & FSBoardKioskGlobalColorAwaitingDispatch & ";") %>
	}
	

	.tr-closed{
		<% Response.Write("background:" & FSBoardKioskGlobalColorClosed & ";") %>
	}
		
	.tr-enroute{
		<% Response.Write("background:" & FSBoardKioskGlobalColorEnRoute & ";") %>
	}
	
	.tr-onsite{
		<% Response.Write("background:" & FSBoardKioskGlobalColorOnSite & ";") %>
	}
		
	.tr-redo-swap{
		<% Response.Write("background:" & FSBoardKioskGlobalColorRedoSwap & ";") %>
	}

	.tr-redo-waitforparts{
		<% Response.Write("background:" & FSBoardKioskGlobalColorRedoWaitForParts & ";") %>
	}

	.tr-redo-followup{
		<% Response.Write("background:" & FSBoardKioskGlobalColorRedoFollowUp & ";") %>
	}

	.tr-redo-unabletowork{
		<% Response.Write("background:" & FSBoardKioskGlobalColorRedoUnableToWork & ";") %>
	}

	.tr-awaiting-acknowledgement{
		<% Response.Write("background:" & FSBoardKioskGlobalColorAwaitingAcknowledgement & ";") %>
	}

	.tr-dispatch-acknowledged{
		<% Response.Write("background:" & FSBoardKioskGlobalColorDispatchAcknowledged& ";") %>
	}

	.tr-declined{
		<% Response.Write("background:" & FSBoardKioskGlobalColorDispatchDeclined& ";") %>
	}

	.tr-awaiting-dispatch-top{
		<% Response.Write("border-top: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}
	
	.tr-awaiting-dispatch-bottom{
		<% Response.Write("border-bottom: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
		<% Response.Write("border-bottom-style: none;") %>
	}
	
	.Urgent-border-top{
		<% Response.Write("border-top: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
		<% Response.Write("border-left: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
		<% Response.Write("border-right: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
		<% Response.Write("border-bottom-style: none;") %>
	}
	
	.Urgent-border-bottom{
		<% Response.Write("border-bottom: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
		<% Response.Write("border-left: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
		<% Response.Write("border-right: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
	}
	
	 .tr-border-line{
		 border-bottom: 1px solid #ccc;
	 }
	 
	 .tr-border-line-red{
		 border-bottom: 1px solid #999;
	 }
	
	.row{
		/*font-size:12px;*/
	}

	.row-line{
		margin-bottom: 25px;
		margin-top: 30px;
		/*font-size:12px;*/
	}
	
			
	.table-condensed>tbody>tr>td, .table-condensed>tbody>tr>th, .table-condensed>tfoot>tr>td, .table-condensed>tfoot>tr>th, .table-condensed>thead>tr>td, .table-condensed>thead>tr>th{
		padding: 2px;
	}
	 
	.scrollable-table{
	 	overflow: hidden;
		border: 1px solid #ccc;
	 	font-size: 9px;
	 	border-bottom-left-radius: 5px;
	 	border-bottom-right-radius: 5px;
	 	
	}

	.scrollable-title{
		border: 1px solid #ccc;
		padding: 10px;
		margin-bottom: -1px;
		background: #DCE6E9;
		font-size: 12px;
		border-top-left-radius: 5px;
		border-top-right-radius: 5px;
	}
	
	.scrollable-title strong{
		width:100%;
		display:block;
		white-space:normal;
	}
	 
	  
	[class^="col-"]{
		padding:2px;
	}
	   
	.col-lg-cust{
	   /* width:7%; */
	   display:inline-block;
	   vertical-align:top;
	}
	
	.table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
		border:0px;
	}
	
	#sortableList .item{
		padding:2px;
	}
	
	.item-box{
		margin:2px;
		float:left;
		width:135px;
	}
	
	   
	.ui-state-highlight.item{height: 100px;}
	 
	.list-boxes{
		/*margin-left: 230px;*/
	}
	 
	.horizontal-layout{
		/*float:left;*/
		width:100%;
		margin-top:20px;
		padding-bottom:40px;
		/*margin-left:-60px;
		margin-left: -475px;*/
	}
	
	.double-scroll{
		width:100%;
	}
		
	.trucknumber{
		width: 100%;
	    display: block;
	    white-space: normal;
	    height:25px;
	    font-weight:bold;
	}  
	
	
	.drivername{
		width: 100%;
	    display: block;
	    white-space: normal;
	    text-transform:uppercase;
	    line-height:12px;
	    height:30px;
	    font-weight:bold;
	    color:#00008B;
	}
		
	 .nag-on-off span{
	 	font-size: 10px;
	 	font-weight: normal;
	 	background: #459e44;
	 	color: #fff;
	 	display: inline-block;
	 	padding: 3px;
	 	border-radius: 2px;
 	 	position: absolute;
	 	top:7px;
	 	right:7px;
 	 }
	
	.btn-nag-on {
	  background-color:#449d44;
	  color:#FFF;
	}
	.btn-nag-off {
	  background-color:#ac2925;
	  color:#FFF;
	}
	
	.btn-nag-on:not(.active){
	  background-color:#449d44;
	  color:#FFF;
	}
	.btn-nag-off:not(.active){
	  background-color:#ac2925;
	  color:#FFF;
	}	

	.tooltip-wrapper {
	  display: inline-block; /* display: block works as well */
	  margin: 50px; /* make some space so the tooltip is visible */
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

</style>

<script type="text/javascript" src="<%= BaseURL %>js/doublescroll/jquery.doubleScroll.js"></script>


<script type="text/javascript">
	$(document).ready(function() {
	
	   $("#PleaseWaitPanel").hide();
	    
       $('.double-scroll').doubleScroll({
       		resetOnWindowResize: true
       	});
	    
	});
</script>

<!-- content area !-->
<div class="wrapper">
<%
	Response.Write("<div id=""PleaseWaitPanel"">")
	Response.Write("<br><br>Loading Today's Service Board<br><br>This may take up to a full minute, please wait...<br><br>")
	Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
	Response.Write("</div>")
	Response.Flush()
%>
<div class="horizontal-layout">
	<div class="double-scroll">
		<table align="center">
			<tr>
		    	<td>
		 			<div class='list-boxes' id='sortableList'>
					<%

					Set cnnUserRegionsForServiceBoard = Server.CreateObject("ADODB.Connection")
					cnnUserRegionsForServiceBoard.open (Session("ClientCnnString"))
					Set rsUserRegionsForServiceBoard = Server.CreateObject("ADODB.Recordset")
					rsUserRegionsForServiceBoard.CursorLocation = 3 
					
					SQLUserRegionsForServiceBoard = "SELECT UserRegionsToViewService FROM tblUsers WHERE UserNo = " & Session("UserNo")
					Set rsUserRegionsForServiceBoard = cnnUserRegionsForServiceBoard.Execute(SQLUserRegionsForServiceBoard)
				
					If IsNull(rsUserRegionsForServiceBoard("UserRegionsToViewService")) Then 
						UserRegionList  = ""
					Else
						UserRegionList = rsUserRegionsForServiceBoard("UserRegionsToViewService")
					End If
					
					set rsUserRegionsForServiceBoard = Nothing
					cnnUserRegionsForServiceBoard.close
					set cnnUserRegionsForServiceBoard = Nothing

					
					Set cnn_FSBoardSum = Server.CreateObject("ADODB.Connection")
					cnn_FSBoardSum.open (Session("ClientCnnString"))
					Set rs_FSBoardSum = Server.CreateObject("ADODB.Recordset")
					Set rs_FSBoardSumForRegions = Server.CreateObject("ADODB.Recordset")
					rs_FSBoardSum.CursorLocation = 3 
					Set rs_FSBoardSum = cnn_FSBoardSum.Execute(SQL)
					
					'**************************************************************************************************					
					'SQL STMT to return only the field service technicians that have service call today
					'**************************************************************************************************
					
					SQL_FSBoardSum = "SELECT DISTINCT UserNoOfServiceTech FROM FS_ServiceMemosDetail "
					SQL_FSBoardSum = SQL_FSBoardSum & "WHERE MemoNumber IN "
                    SQL_FSBoardSum = SQL_FSBoardSum & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE "
                    SQL_FSBoardSum = SQL_FSBoardSum & "(CurrentStatus = 'OPEN') OR "
                    SQL_FSBoardSum = SQL_FSBoardSum & "(CurrentStatus = 'CLOSE' AND "
                    
                    SQL_FSBoardSum = SQL_FSBoardSum & "YEAR(RecordCreatedDateTime) = YEAR(GetDate()) AND "
                    SQL_FSBoardSum = SQL_FSBoardSum & "MONTH(RecordCreatedDateTime) = MONTH(GetDate()) AND "
                    SQL_FSBoardSum = SQL_FSBoardSum & "DAY(RecordCreatedDateTime) = DAY(GetDate()) "
                    SQL_FSBoardSum = SQL_FSBoardSum & ")) AND UserNoOfServiceTech <> ''"
					SQL_FSBoardSum = SQL_FSBoardSum & " ORDER BY UserNoOfServiceTech"
					
					'Response.write(SQL_FSBoardSum & "<br><br>")
					

					Set rs_FSBoardSum = cnn_FSBoardSum.Execute(SQL_FSBoardSum)
					
					If not rs_FSBoardSum.EOF Then
						Do While Not rs_FSBoardSum.Eof

							If userIsArchived(rs_FSBoardSum("UserNoOfServiceTech")) = False AND userIsEnabled(rs_FSBoardSum("UserNoOfServiceTech")) = True Then	
												
								TechHasANYTicketsInRegion = False
						
								SQL_FSBoardSumForRegions = "SELECT DISTINCT CUstNum AS AccountNumber, MemoNumber FROM FS_ServiceMemosDetail "
								SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "WHERE MemoNumber IN "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "(CurrentStatus = 'OPEN') OR "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "(CurrentStatus = 'CLOSE' AND "
			                    
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "YEAR(RecordCreatedDateTime) = YEAR(GetDate()) AND "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "MONTH(RecordCreatedDateTime) = MONTH(GetDate()) AND "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "DAY(RecordCreatedDateTime) = DAY(GetDate()) "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & ")) AND UserNoOfServiceTech  = " & rs_FSBoardSum("UserNoOfServiceTech")
	
'Response.Write(SQL_FSBoardSumForRegions & "<br>")
			                    
			                    Set rs_FSBoardSumForRegions = cnn_FSBoardSum.Execute(SQL_FSBoardSumForRegions)
			                    
								If not rs_FSBoardSumForRegions.EOF Then
								
									Do While NOT rs_FSBoardSumForRegions.EOF
									
										If GetServiceTicketCurrentStage(rs_FSBoardSumForRegions("MemoNumber")) <> "Received" Then 
									
											If LastTechUserNo(rs_FSBoardSumForRegions("MemoNumber")) = rs_FSBoardSum("UserNoOfServiceTech") Then
										
												If UserRegionList <> "" Then
												
													CustRegion = GetCustRegionIntRecIDByCustID(rs_FSBoardSumForRegions("AccountNumber"))
												
													RegionArray = Split(UserRegionList,",")
													
													For x = 0 to Ubound(RegionArray)
														If cint(RegionArray(x)) = cint(CustRegion) Then
															TechHasANYTicketsInRegion = True
															Exit For
														End IF
													Next
													
												End If
		
											End If
										
										End If
	
										If TechHasANYTicketsInRegion = True Then Exit Do
									
										rs_FSBoardSumForRegions.MoveNext
									Loop
								
								End IF
			                    
								If UserRegionList = "" Then TechHasANYTicketsInRegion = True
								
								'Only Write Route If The User Is Not Archived and Not Disabled
								If TechHasANYTicketsInRegion = True Then 
									Call TruckNumberWrite(rs_FSBoardSum("UserNoOfServiceTech")) 
								End IF
								
							End If
							
							rs_FSBoardSum.Movenext
						Loop
					End If
					%> 
				</div>
			</tr>
		</table>
	</div>
</div>

<%
Sub TruckNumberWrite(UserNoOfServiceTech) %>
		<td class='item col-lg-cust' TruckNumber='<%= UserNoOfServiceTech %>'>
		<div class='item-box'>
		<div class='scrollable-title' style='position: relative;'>
		
			<span class="trucknumber"><i class="fa fa-wrench" aria-hidden="true"></i>&nbsp;<%= UserNoOfServiceTech %></span>
		
			<span class="drivername"><%= GetUserDisplayNameByUserNo(UserNoOfServiceTech) %></span>

			<%
			'**************************************************************************
			'show nag alerts to admins or route managers only
			'**************************************************************************

			If UserNoOfServiceTech <> "*Not Found*" Then
			
				'First check to see if nags are off entirely for this user
				NagsON = False
				
				SQLUsers = "SELECT * FROM tblUsers Where UserNo = " & UserNoOfServiceTech 
				
				Set cnn_Users = Server.CreateObject("ADODB.Connection")
				cnn_Users.open (Session("ClientCnnString"))
				Set rsUsers = Server.CreateObject("ADODB.Recordset")
				rsUsers.CursorLocation = 3 
				'Response.write(SQLUsers)
				Set rsUsers = cnn_Users.Execute(SQLUsers)

				'ANY YES CONDITION TURNS THE BUTTON ON
				If Not rsUsers.EOF Then

					If rsUsers("userNextStopNagMessageOverride_FS") = "Yes" Then NagsON = True 
					If rsUsers("userNoActivityNagMessageOverride_FS") = "Yes" Then NagsON = True

					
					If NagsON = False Then' only check if not already on
					
						If rsUsers("userNextStopNagMessageOverride_FS") = "Use Global" or rsUsers("userNoActivityNagMessageOverride_FS") = "Use Global" Then
						
							SQLGlobal = "SELECT * FROM Settings_Global "
							Set rsGlobal = Server.CreateObject("ADODB.Recordset")
							rsGlobal.CursorLocation = 3 
							Set rsGlobal = cnn_Users.Execute(SQLGlobal)
	
							If Not rsGlobal.EOF Then
								NoAct = rsGlobal("NoActivityNagMessageONOFF_FS")
								NextSt = rsGlobal("NextStopNagMessageONOFF_FS")
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
			

			
			If UserNoOfServiceTech <> "*Not Found*" Then
			
				If NagsOn = True Then
			
					If  DriverInNagSkipTable(UserNoOfServiceTech,"fsNoNextStop") = False AND DriverInNagSkipTable(UserNoOfServiceTech,"fsNoActivity") = False Then
						buttonClassGreen = "btn btn-xs btn-nag-on" 
						buttonClassRed= "btn btn-xs btn-default active"
					Else 
						buttonClassGreen = "btn btn-xs btn-default active" 
						buttonClassRed = "btn btn-xs btn-nag-off"
					End If
					%>	  
					<% If buttonClassGreen = "btn btn-xs btn-nag-on" Then %>
						  <div class="btn-group btn-toggle" id="<%= UserNoOfServiceTech %>">
						    <button class="<%= buttonClassGreen %>" id="<%= UserNoOfServiceTech %>ON">NAG ON</button>
						    <button class="<%= buttonClassRed %>" id="<%= UserNoOfServiceTech %>OFF">OFF</button>
						  </div>
					<% Else %>
						  <div class="btn-group btn-toggle" id="<%= UserNoOfServiceTech %>">
						    <button class="<%= buttonClassGreen %>" id="<%= UserNoOfServiceTech %>ON">ON</button>
						    <button class="<%= buttonClassRed %>" id="<%= UserNoOfServiceTech %>OFF">NAG OFF</button>
						  </div>						
					<% End If 
				
				Else
					%><div class="btn-group btn-toggle">Nags Off</div><%
				End If

			Else
				%><div class="btn-group btn-toggle">No User Setup</div><%
			End If

			'**************************************************************************
		
			%> 
						
		</div>

	        <div class='table-responsive scrollable-table'>
		        <% Response.Write("<table id='truck" & UserNoOfServiceTech & "' name='truck" & UserNoOfServiceTech & "' class='food_planner table table-condensed clickable'>") %>
					<thead>
			        	<tr>
			        		<th class='sorttable_nosort'>Ticket #</th>
			        		<th class='sorttable_nosort'><%=GetTerm("Customer")%></th>.
			        	</tr>
			        </thead>
			        <tbody class='searchable'>
			        	<%'Get all the tickets for this truck
			        	
						Set cnn_Tickets = Server.CreateObject("ADODB.Connection")
						cnn_Tickets.open (Session("ClientCnnString"))
						Set rs_FSBoardDet = Server.CreateObject("ADODB.Recordset")
						rs_FSBoardDet.CursorLocation = 3 
						   	                    
						SQL_Tickets = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemosDetail "
						SQL_Tickets = SQL_Tickets & "WHERE (MemoNumber IN "
                        SQL_Tickets = SQL_Tickets & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE "
                        SQL_Tickets = SQL_Tickets & "(CurrentStatus = 'OPEN') OR "
                        SQL_Tickets = SQL_Tickets & "(CurrentStatus = 'CLOSE' AND "
	                    SQL_Tickets = SQL_Tickets & "YEAR(RecordCreatedDateTime) = YEAR(GetDate()) AND "
	                    SQL_Tickets = SQL_Tickets & "MONTH(RecordCreatedDateTime) = MONTH(GetDate()) AND "
	                    SQL_Tickets = SQL_Tickets & "DAY(RecordCreatedDateTime) = DAY(GetDate())) "
                        SQL_Tickets = SQL_Tickets & ")) AND (UserNoOfServiceTech = " & UserNoOfServiceTech &")"
   	                    
   	                   ' Response.Write(SQL_Tickets )
   	                    
                        Set rs_FSBoardDet = cnn_Tickets.Execute(SQL_Tickets)
						If not rs_FSBoardDet.Eof Then

							NumLines = 0
							Do While not rs_FSBoardDet.Eof		
							
								ShowThisRec = True
											
								
								CustID = GetServiceTicketCust(rs_FSBoardDet("MemoNumber"))
								
								If UserRegionList <> "" Then
								
									CustRegion = GetCustRegionIntRecIDByCustID(CustID)
									ShowThisRec = False
									
									RegionArray = Split(UserRegionList,",")
									
									For x = 0 to Ubound(RegionArray)
										If cint(RegionArray(x)) = cint(CustRegion) Then
											ShowThisRec = True
											Exit For
										End IF
									Next
								End If
			
								'Response.Write("UserRegionList " & UserRegionList )
								
								If ShowThisRec = True AND GetServiceTicketCurrentStage(rs_FSBoardDet("MemoNumber")) <> "Received" Then 
								
									If LastTechUserNo(rs_FSBoardDet("MemoNumber")) = UserNoOfServiceTech Then
									
										ServiceTicketCurrentStage = GetServiceTicketCurrentStage(rs_FSBoardDet("MemoNumber"))
										ServiceTicketCurrentStatus = GetServiceTicketStatus(rs_FSBoardDet("MemoNumber"))
									
										'Write first table row
										'**********************
										
										'CustID = GetServiceTicketCust(rs_FSBoardDet("MemoNumber"))
																		
										If len(GetCustNameByCustNum(GetServiceTicketCust(rs_FSBoardDet("MemoNumber")))) > 19 then 
											Cnam = left(GetCustNameByCustNum(GetServiceTicketCust(rs_FSBoardDet("MemoNumber"))),19) 
										Else 
											Cnam = GetCustNameByCustNum(GetServiceTicketCust(rs_FSBoardDet("MemoNumber")))
										End If
										
										
										If ServiceTicketCurrentStatus = "CLOSE" OR ServiceTicketCurrentStatus = "CANCEL" Then
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-closed Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>" data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-closed Urgent-border-top">
												<% End If 
											Else
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-closed tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-closed tr-awaiting-dispatch-top">
												<% End If
											End If
										
										End If
										
																				
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage <> "En Route"_
										AND ServiceTicketCurrentStage <> "On Site"_
										AND ServiceTicketCurrentStage <> "Dispatched"_
										AND ServiceTicketCurrentStage <> "Dispatch Acknowledged"_
										AND ServiceTicketCurrentStage <> "Dispatch Declined" Then
										
											
											If AwaitingRedispatch(rs_FSBoardDet("MemoNumber")) <> True Then
												If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
													%>
													<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
														<tr class="tr-awaiting-dispatch Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
													<% Else %>
														<tr class="tr-awaiting-dispatch Urgent-border-top">
													<% End If
												Else
													%>
													<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
														<tr class="tr-awaiting-dispatch tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
													<% Else %>
														<tr class="tr-awaiting-dispatch tr-awaiting-dispatch-top">
													<% End If
												End If
											Else
												If ServiceTicketCurrentStage = "Swap" Then
													className = "tr-redo-swap"
												ElseIf ServiceTicketCurrentStage = "Wait for parts" Then
													className = "tr-redo-waitforparts"
												ElseIf ServiceTicketCurrentStage = "Follow Up" Then
													className = "tr-redo-followup"
												ElseIf ServiceTicketCurrentStage = "Unable To Work" Then
													className = "tr-redo-unabletowork"
												End If

												If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
													%>
													<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
														<tr class="<%= className %> Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
													<% Else %>
														<tr class="<%= className %> Urgent-border-top">
													<% End If	
												Else
													%>
													<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
														<tr class="<%= className %> tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
													<% Else %>
														<tr class="<%= className %> tr-awaiting-dispatch-top">
													<% End If
												End If
											End If
										End If
										
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "En Route" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-enroute Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-enroute Urgent-border-top">
												<% End If
											Else
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-enroute tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-enroute tr-awaiting-dispatch-top">
												<% End If
											End If
										
										End If
		
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "On Site" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-onsite Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-onsite Urgent-border-top">
												<% End If
											Else
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-onsite tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-onsite tr-awaiting-dispatch-top">
												<% End If
											End If
										
										End If
										
										
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "Dispatched" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-dispatched Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-dispatched Urgent-border-top">
												<% End If
											Else
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-dispatched tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-dispatched tr-awaiting-dispatch-top">
												<% End If
											End If
										
										End If
										
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "Dispatch Acknowledged" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-acknowledged Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-acknowledged Urgent-border-top">
												<% End If 
											Else
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-acknowledged tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-acknowledged tr-awaiting-dispatch-top">
												<% End If
											End If
										
										End If

										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "Dispatch Declined" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-declined Urgent-border-top" style="cursor:pointer;"><%
											Else
												%><tr class="tr-declined tr-awaiting-dispatch-top" style="cursor:pointer;"><%
											End If
										
										End If
		
										%>
										
											<td><%= rs_FSBoardDet("MemoNumber") %></td>
											<td>
											<% If ServiceTicketCurrentStatus = "CLOSE" And ServiceTicketCurrentStage = "On Site" Then %>
												<span class="tooltip-button" data-toggle="tooltip" data-placement="bottom" title="<%= GetServiceTicketSTAGEDateTime(rs_FSBoardDet("MemoNumber"),ServiceTicketCurrentStage) %>">CLOSED</span>
											<% Else %>
												<span class="tooltip-button" data-toggle="tooltip" data-placement="bottom" title="<%= GetServiceTicketSTAGEDateTime(rs_FSBoardDet("MemoNumber"),ServiceTicketCurrentStage) %>"><%= ServiceTicketCurrentStage %></span>
											<% End If %>
											
											</td>
										
										</tr>
										
										<%
		
																				
										'Write second table row
										'**********************
										
										If ServiceTicketCurrentStatus = "CLOSE" OR ServiceTicketCurrentStatus = "CANCEL" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-closed Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-closed Urgent-border-bottom">
												<% End If
											Else
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-closed tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-closed tr-awaiting-dispatch-bottom">
												<% End If
											End If
										
										End If
										
										
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage <> "En Route"_
										AND ServiceTicketCurrentStage <> "On Site"_
										AND ServiceTicketCurrentStage <> "Dispatched"_
										AND ServiceTicketCurrentStage <> "Dispatch Acknowledged"_
										AND ServiceTicketCurrentStage <> "Dispatch Declined" Then
										
											
											If AwaitingRedispatch(rs_FSBoardDet("MemoNumber")) <> True Then
												If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
													%>
													<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
														<tr class="tr-awaiting-dispatch Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
													<% Else %>
														<tr class="tr-awaiting-dispatch Urgent-border-bottom">
													<% End If
												Else
													%>
													<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
														<tr class="tr-awaiting-dispatch tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
													<% Else %>
														<tr class="tr-awaiting-dispatch tr-awaiting-dispatch-bottom">
													<% End If
												End If
											Else
											
												If ServiceTicketCurrentStage = "Swap" Then
													className = "tr-redo-swap"
												ElseIf ServiceTicketCurrentStage = "Wait for parts" Then
													className = "tr-redo-waitforparts"
												ElseIf ServiceTicketCurrentStage = "Follow Up" Then
													className = "tr-redo-followup"
												ElseIf ServiceTicketCurrentStage = "Unable To Work" Then
													className = "tr-redo-unabletowork"
												End If

												If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
													%><tr class="<%= className %> Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
												Else
													%><tr class="<%= className %> tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
												End If
											
											End If
										End If
										
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "En Route" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-enroute Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-enroute Urgent-border-bottom">
												<% End If
											Else
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-enroute tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-enroute tr-awaiting-dispatch-bottom">
												<% End If
											End If
										
										End If
		
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "On Site" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-onsite Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-onsite Urgent-border-bottom">
												<% End If
											Else
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-onsite tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-onsite tr-awaiting-dispatch-bottom">
												<% End If
											End If
										
										End If
										
		
		
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "Dispatched" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-dispatched Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-dispatched Urgent-border-bottom">
												<% End If 
											Else
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-dispatched tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-dispatched tr-awaiting-dispatch-bottom">
												<% End If
											End If
										
										End If
		
		
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "Dispatch Acknowledged" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-acknowledged Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-acknowledged Urgent-border-bottom">
												<% End If
											Else
												%>
												<% If userCanAccessServiceActionsModalButton(Session("UserNo")) = true Then %>
													<tr class="tr-acknowledged tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;">
												<% Else %>
													<tr class="tr-acknowledged tr-awaiting-dispatch-bottom">
												<% End If
											End If
										
										End If

										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage ="Dispatch Declined" Then

											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-declined Urgent-border-bottom" style="cursor:pointer;"><%
											Else
												%><tr class="tr-declined tr-awaiting-dispatch-bottom" style="cursor:pointer;"><%
											End If
										
										End If
		
										%>
										
											<td colspan="2"><span class="tooltip-button" data-toggle="tooltip" data-placement="bottom" title="Account #<%= CustID %>"><%= Cnam %></span></td>
										
										</tr>
										
	 							<%
	 								End If ' last tech user no
	 							End If ' status not received
								rs_FSBoardDet.movenext
								NumLines = NumLines + 1
								
							Loop
							
						End IF%>
                        </td>
			        </tbody>
		        </table>
	        </div>
            </div>
        <%Response.Write("</div>")
End Sub 

Set rs_FSBoardSum = Nothing
cnn_FSBoardSum.Close
Set cnn_FSBoardSum = Nothing
%>	

<!-- tooltip JS !-->
<script type="text/javascript">
	$(function () {
	  $('[data-toggle="tooltip"]').tooltip()
	})
</script>
<!-- eof tooltip JS !-->


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

<!--#include file="serviceBoardCommonModals.asp"-->

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS END HERE !-->
<!-- **************************************************************************************************************************** -->


<!--#include file="../inc/footer-deliveryBoard.asp"-->
