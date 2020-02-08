<!--#include file="settings.asp"-->
<!--#include file="mail.asp"-->
<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<!--#include file="InSightFuncs_Routing.asp"-->
<!--#include file="InSightFuncs_Service.asp"-->
<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
 
'Sub TurnOnNagAlertsForFieldServiceTechnicianKiosk()
'Sub TurnOffNagAlertsForFieldServiceTechnicianKiosk()
'Sub GetContentForServiceBoardTicketOptionsModal()
'Sub GetTitleForServiceBoardTransferRedispatchModal()
'Sub GetContentForServiceBoardTransferRedispatchModal()
'Sub ToggleTicketAsUrgentFromServiceBoardModal()
'Sub ToggleTicketAsUrgentFromServiceBoardModalAndSendText()
'Sub GetContentForServiceTicketNotesModal()
'Sub GetTitleForServiceBoardChangeTypeModal()
'Sub GetContentForServiceBoardChangeTypeModal()
'Sub GetTitleForServiceBoardDispatchModal()
'Sub GetContentForServiceBoardDispatchModal()
'Sub SaveAddNewFilter()
'Sub GetContentForEditFilterModal()
'Sub SaveEditExistingFilter()
'Sub CheckForDuplicateFilterIDNewFilter()
'Sub CheckForDuplicateFilterUPCCodeNewFilter()
'Sub CheckForDuplicateFilterIDExistingFilter()
'Sub CheckForDuplicateFilterUPCCodeExistingFilter()
'Sub GetContentForDeleteFilterModal()
'Sub GetContentForCreateServiceTicketModal()
'Sub SetRegionFilterListByUserForServiceBoard()
'Sub GetContentForEditServiceTicketModal()
'Sub GetContentForViewOpenClosedServiceTicketModal()
'Sub GetContentForEditServiceTicketModalTitle()
'Sub GetNextPageForLinkedKiosks()
'***************************************************
'End List of all the AJAX functions & subs
'***************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'ALL AJAX MODAL SUBROUTINES AND FUNCTIONS BELOW THIS AREA

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

action = Request("action")

Select Case action
	Case "TurnOnNagAlertsForFieldServiceTechnicianKiosk"
		TurnOnNagAlertsForFieldServiceTechnicianKiosk()
	Case "TurnOffNagAlertsForFieldServiceTechnicianKiosk"
		TurnOffNagAlertsForFieldServiceTechnicianKiosk()
	Case "GetContentForServiceBoardTicketOptionsModal"
		GetContentForServiceBoardTicketOptionsModal()	
	Case "GetTitleForServiceBoardTransferRedispatchModal"
		GetTitleForServiceBoardTransferRedispatchModal()
	Case "GetContentForServiceBoardTransferRedispatchModal"
		GetContentForServiceBoardTransferRedispatchModal()
	Case "ToggleTicketAsUrgentFromServiceBoardModal"
		ToggleTicketAsUrgentFromServiceBoardModal()
	Case "ToggleTicketAsUrgentFromServiceBoardModalAndSendText"
		ToggleTicketAsUrgentFromServiceBoardModalAndSendText()
	Case "GetContentForServiceTicketNotesModal"
		GetContentForServiceTicketNotesModal()
	Case "GetContentForServiceBoardChangeTypeModal"
		GetContentForServiceBoardChangeTypeModal
	Case "GetTitleForServiceBoardChangeTypeModal"
		GetTitleForServiceBoardChangeTypeModal
	Case "GetContentForServiceBoardDispatchModal"
		GetContentForServiceBoardDispatchModal
	Case "GetTitleForServiceBoardDispatchModal"
		GetTitleForServiceBoardDispatchModal
	Case "SetCurrentServiceTabInMUVWrite"
		SetCurrentServiceTabInMUVWrite()
	Case "SaveAddNewFilter"
		SaveAddNewFilter()
	Case "GetContentForEditFilterModal"
		GetContentForEditFilterModal()
	Case "SaveEditExistingFilter"
		SaveEditExistingFilter()
	Case "CheckForDuplicateFilterIDNewFilter"
		CheckForDuplicateFilterIDNewFilter()
	Case "CheckForDuplicateFilterUPCCodeNewFilter"
		CheckForDuplicateFilterUPCCodeNewFilter()	
	Case "CheckForDuplicateFilterIDExistingFilter"
		CheckForDuplicateFilterIDExistingFilter()
	Case "CheckForDuplicateFilterUPCCodeExistingFilter"
		CheckForDuplicateFilterUPCCodeExistingFilter()	
	Case "GetContentForDeleteFilterModal"
		GetContentForDeleteFilterModal()
	Case "GetContentForCreateServiceTicketModal"
		GetContentForCreateServiceTicketModal()	
	Case "SetRegionFilterListByUserForServiceBoard"
		SetRegionFilterListByUserForServiceBoard()
	Case "GetContentForEditServiceTicketModal"
		GetContentForEditServiceTicketModal()
	Case "GetContentForViewOpenClosedServiceTicketModal"
		GetContentForViewOpenClosedServiceTicketModal()
	Case "GetContentForEditServiceTicketModalTitle"
		GetContentForEditServiceTicketModalTitle()
	Case "GetNextPageForLinkedKiosks"
		GetNextPageForLinkedKiosks()
End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub SetCurrentServiceTabInMUVWrite() 

	selectedTab = Request.Form("selectedTab")
	
	dummy = MUV_Write("selectedServiceTab",	selectedTab)
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub TurnOffNagAlertsForFieldServiceTechnicianKiosk() 

	technicianUserNo = Request.Form("technicianUserNo")
	
	'***************************************************************************************
	'Delete nag types of fsNoActivity and fsNoNextStop for this user number
	'***************************************************************************************

	Set cnnFieldServiceNagAlert = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceNagAlert.open (Session("ClientCnnString"))
	Set rsFieldServiceNagAlert = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceNagAlert.CursorLocation = 3 
	
	SQLFieldServiceNagAlert = "DELETE FROM SC_NagSkipUsers WHERE UserNo = " & technicianUserNo & " AND NagType = 'fsNoActivity'"
	
	Response.write(SQLFieldServiceNagAlert)
	
	Set rsFieldServiceNagAlert = cnnFieldServiceNagAlert.Execute(SQLFieldServiceNagAlert)
	
	SQLFieldServiceNagAlert = "DELETE FROM SC_NagSkipUsers WHERE UserNo = " & technicianUserNo & " AND NagType = 'fsNoNextStop'"
	Set rsFieldServiceNagAlert = cnnFieldServiceNagAlert.Execute(SQLFieldServiceNagAlert)

		
	set rsFieldServiceNagAlert = Nothing
	cnnFieldServiceNagAlert.close
	set cnnFieldServiceNagAlert = Nothing
	

	Description = GetUserDisplayNameByUserNo(0) & " turned OFF nag alerts today for " & GetUserDisplayNameByUserNo(technicianUserNo)
	CreateAuditLogEntry "Nag Alerts Turned Off For " & GetTerm("Field Service") & " Technician","Nag Alerts Turned Off For " & GetTerm("Field Service") & " Technician","Minor",0,Description
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub TurnOnNagAlertsForFieldServiceTechnicianKiosk() 

	technicianUserNo = Request.Form("technicianUserNo")
	
	'***************************************************************************************
	'Get nag type values for passed driver user number
	'***************************************************************************************

	Set cnnFieldServiceNagAlert = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceNagAlert.open (Session("ClientCnnString"))
	Set rsFieldServiceNagAlert = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceNagAlert.CursorLocation = 3 
	
	Set cnnFieldServiceNagAlertUpdateInsert = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceNagAlertUpdateInsert.open (Session("ClientCnnString"))
	Set rsFieldServiceNagAlertUpdateInsert = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceNagAlertUpdateInsert.CursorLocation = 3 
	
	
	SQLFieldServiceNagAlert = "SELECT * FROM SC_NagSkipUsers WHERE UserNo = " & technicianUserNo & " AND NagType = 'fsNoActivity'"
	response.write(SQLFieldServiceNagAlert)
	Set rsFieldServiceNagAlert = cnnFieldServiceNagAlert.Execute(SQLFieldServiceNagAlert)
	
	If rsFieldServiceNagAlert.EOF THEN
		SQLFieldServiceNagAlertUpdateInsert = "INSERT INTO SC_NagSkipUsers (UserNo, NagType) VALUES (" & technicianUserNo & ",'fsNoActivity')"
		Set rsFieldServiceNagAlertUpdateInsert = cnnFieldServiceNagAlertUpdateInsert.Execute(SQLFieldServiceNagAlertUpdateInsert)
	End If	

	
	SQLFieldServiceNagAlert = "SELECT * FROM SC_NagSkipUsers WHERE UserNo = " & technicianUserNo & " AND NagType = 'fsNoNextStop'"
	Set rsFieldServiceNagAlert = cnnFieldServiceNagAlert.Execute(SQLFieldServiceNagAlert)
	
	If rsFieldServiceNagAlert.EOF THEN
		SQLFieldServiceNagAlertUpdateInsert = "INSERT INTO SC_NagSkipUsers (UserNo, NagType) VALUES (" & technicianUserNo & ",'fsNoNextStop')"
		Set rsFieldServiceNagAlertUpdateInsert = cnnFieldServiceNagAlertUpdateInsert.Execute(SQLFieldServiceNagAlertUpdateInsert)
	End If	
	
	set rsFieldServiceNagAlertUpdateInsert = Nothing
	cnnFieldServiceNagAlertUpdateInsert.close
	set cnnFieldServiceNagAlertUpdateInsert = Nothing
		
	set rsFieldServiceNagAlert = Nothing
	cnnFieldServiceNagAlert.close
	set cnnFieldServiceNagAlert = Nothing

	Description = GetUserDisplayNameByUserNo(0) & " turned ON nag alerts today for " & GetUserDisplayNameByUserNo(technicianUserNo)
	CreateAuditLogEntry "Nag Alerts Turned On For " & GetTerm("Field Service") & " Technician","Nag Alerts Turned On For " & GetTerm("Field Service") & " Technician","Minor",0,Description
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************





'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForServiceBoardTicketOptionsModal() 

	TicketNumber = Request.Form("memoNum")
	CustID = Request.Form("custID")
	UserNoOfServiceTech = Request.Form("userNo")
	ReturnURL = Request.Form("returnURL")
	
	ServiceTicketCurrentStage = UCASE(GetServiceTicketCurrentStage(TicketNumber))
	ServiceTicketCurrentStatus = UCASE(GetServiceTicketStatus(TicketNumber))


	'***********************************************
	'Determine if the ticket is urgent and being set to NOT urgent
	'or not urgent and being set to URGENT
	'***********************************************

	SQLToggleUrgent = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & TicketNumber & "'"
	Set cnnToggleUrgent = Server.CreateObject("ADODB.Connection")
	cnnToggleUrgent.open (Session("ClientCnnString"))
	
	Set rsToggleUrgent = Server.CreateObject("ADODB.Recordset")
	rsToggleUrgent.CursorLocation = 3 
	Set rsToggleUrgent = cnnToggleUrgent.Execute(SQLToggleUrgent)
	
	If NOT rsToggleUrgent.EOF Then
		'All records should be the same so just check the 1st one
		If rsToggleUrgent("Urgent") <> 1 Then
			CurrentlyUrgent = false
		Else
			CurrentlyUrgent = true
		End If
	Else
		CurrentlyUrgent = false
	End If		
	
	set rsToggleUrgent = Nothing
	cnnToggleUrgent.Close
	Set cnnToggleUrgent = Nothing
	
%>

	<script type="text/javascript">
	
		$(document).ready(function() {
	
			$('#btnToggleUrgentNoTexting').on('click', function(e) {

			    //get data-id attribute of the clicked service ticket
			    var ticketNumber = $("#txtTicketNum").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val();
			    
				//close the service ticket options modal where we came from
				$('#serviceBoardTicketOptionsModal').modal('hide');
				
				//turn off the automatic page refresh
				$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			    		    		
		    	$.ajax({
					type:"POST",
					url: BaseURL + "inc/InSightFuncs_AjaxForServiceModals.asp",
					data: "action=ToggleTicketAsUrgentFromServiceBoardModal&returnURL=" + encodeURIComponent(ReturnURL) + "&memoNum=" + encodeURIComponent(ticketNumber),
					success: function(response)
					 {
						 if (ReturnURL.indexOf("fieldservicekiosknopaging") >=0) {
						   	window.location = BaseURL + "directLaunch/kiosks/service/fieldservicekiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
						 }
						 else {
						   location.reload();
						 }					 	
		             }
				});
	    	});	
	    	
	    	

			$('#btnMarkAsNotUrgentWithTexting').on('click', function(e) {
			    
			    //get data-id attribute of the clicked service ticket
			    var ticketNumber = $("#txtTicketNum").val();
			    var ticketCustID = $("#txtCustID").val();
			    var techUserNo = $("#txtUserNo").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val();
	    
				//close the service ticket options modal where we came from
				$('#serviceBoardTicketOptionsModal').modal('hide');
				
				//turn off the automatic page refresh
				$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
				
				
				//open the modal window that asks if you would like to send a text message to the technician
				$("#serviceBoardMarkAsNotUrgentWithTextingModal").modal('show');
				
				
				//if they want to send a text message to the technician,
				//call the appropriate function that will mark the ticket as urgent and send the text message
				$("#modal-btn-yes-send-text").on("click", function(){
					
					$("#serviceBoardMarkAsNotUrgentWithTextingModal").modal('hide');
										
			    	$.ajax({
						type:"POST",
						url: BaseURL + "inc/InSightFuncs_AjaxForServiceModals.asp",
						data: "action=ToggleTicketAsUrgentFromServiceBoardModalAndSendText&returnURL=" + encodeURIComponent(ReturnURL) + "&memoNum=" + encodeURIComponent(ticketNumber) + "&custID=" + encodeURIComponent(ticketCustID) + "&userNo=" + encodeURIComponent(techUserNo),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("fieldservicekiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/service/fieldservicekiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	
			             }
					});
					
				});

				
				//if they do not want to send a text message to the technician,
				//call the appropriate function that will mark the ticket as urgent only
				$("#modal-btn-no-text").on("click", function(){
				
					$("#serviceBoardMarkAsNotUrgentWithTextingModal").modal('hide');
					
			    	$.ajax({
						type:"POST",
						url: BaseURL + "inc/InSightFuncs_AjaxForServiceModals.asp",
						data: "action=ToggleTicketAsUrgentFromServiceBoardModal&returnURL=" + encodeURIComponent(ReturnURL) + "&memoNum=" + encodeURIComponent(ticketNumber),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("fieldservicekiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/service/fieldservicekiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	
			             }
					});
					
				});				
			    		    		
	    	});	
	    	
	


			$('#btnMarkAsUrgentWithTexting').on('click', function(e) {

			    //get data-id attribute of the clicked service ticket
			    var ticketNumber = $("#txtTicketNum").val();
			    var ticketCustID = $("#txtCustID").val();
			    var techUserNo = $("#txtUserNo").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val();
			    
				//close the service ticket options modal where we came from
				$('#serviceBoardTicketOptionsModal').modal('hide');
				
				//turn off the automatic page refresh
				$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
				
				
				//open the modal window that asks if you would like to send a text message to the technician
				$("#serviceBoardMarkAsUrgentWithTextingModal").modal('show');
				
				
				//if they want to send a text message to the technician,
				//call the appropriate function that will mark the ticket as urgent and send the text message
				$("#modal-btn-yes-send-urgent-text").on("click", function(){
					
					$("#serviceBoardMarkAsUrgentWithTextingModal").modal('hide');

			    	$.ajax({
						type:"POST",
						url: BaseURL + "inc/InSightFuncs_AjaxForServiceModals.asp",
						data: "action=ToggleTicketAsUrgentFromServiceBoardModalAndSendText&returnURL=" + encodeURIComponent(ReturnURL) + "&memoNum=" + encodeURIComponent(ticketNumber) + "&custID=" + encodeURIComponent(ticketCustID) + "&userNo=" + encodeURIComponent(techUserNo),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("fieldservicekiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/service/fieldservicekiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	
			             }
					});
					
				});

				
				//if they do not want to send a text message to the technician,
				//call the appropriate function that will mark the ticket as urgent only
				$("#modal-btn-no-urgent-text").on("click", function(){
				
					$("#serviceBoardMarkAsUrgentWithTextingModal").modal('hide');
					
			    	$.ajax({
						type:"POST",
						url: BaseURL + "inc/InSightFuncs_AjaxForServiceModals.asp",
						data: "action=ToggleTicketAsUrgentFromServiceBoardModal&returnURL=" + encodeURIComponent(ReturnURL) + "&memoNum=" + encodeURIComponent(ticketNumber),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("fieldservicekiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/service/fieldservicekiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	

			             }
					});
					
				});				
			    		    		
	    	});	
	    	    		
		});
	</script>

	<input type="hidden" name="txtTicketNum" id="txtTicketNum" value="<%= TicketNumber %>">
	<input type="hidden" name="txtCustID" id="txtCustID" value="<%= CustID %>">
	<input type="hidden" name="txtUserNo" id="txtUserNo" value="<%= UserNoOfServiceTech %>">
	<input type="hidden" name="txtReturnURL" id="txtReturnURL" value="<%= ReturnURL %>">
	<input type="hidden" name="txtBaseURL" id="txtBaseURL" value="<%= BaseURL %>">

	<div class="row">
		<div class="center">
			<%
				If ServiceTicketCurrentStatus = "OPEN" AND (ServiceTicketCurrentStage = "DISPATCHED" OR ServiceTicketCurrentStage = "DISPATCH ACKNOWLEDGED") Then
				
					%>
					
					<div class="col-lg-11">
						<% If CurrentlyUrgent = true Then %>
							<a href="#" class="btn btn-danger btn-lg btn-block btn-huge" id="btnMarkAsNotUrgentWithTexting" data-show="true" href="#" data-invoice-number="<%= TicketNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>" data-target="#serviceBoardMarkAsNotUrgentWithTextingModal" style="cursor:pointer;">Remove Urgent Status</a>
						<% Else %>
							<a href="#" class="btn btn-danger btn-lg btn-block btn-huge" id="btnMarkAsUrgentWithTexting" data-show="true" href="#" data-invoice-number="<%= TicketNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>" data-target="#serviceBoardMarkAsUrgentWithTextingModal" style="cursor:pointer;">Mark Ticket As Urgent</a>
						<% End If %>
					</div>
					
					<%
					
				ElseIf ServiceTicketCurrentStatus = "OPEN" AND (ServiceTicketCurrentStage = "EN ROUTE" OR ServiceTicketCurrentStage = "ON SITE") Then
				
					%>&nbsp;<%
					
				ElseIf ServiceTicketCurrentStatus = "OPEN" AND (ServiceTicketCurrentStage = "DISPTACH DECLINED"_
						OR ServiceTicketCurrentStage = "FOLLOW UP"_
						OR ServiceTicketCurrentStage = "UNABLE TO WORK"_
						OR ServiceTicketCurrentStage = "WAIT FOR PARTS"_
						OR ServiceTicketCurrentStage = "SWAP") Then
					%>
					<div class="col-lg-11">
						<% If CurrentlyUrgent = true Then %>
							<a href="#" class="btn btn-danger btn-lg btn-block btn-huge" id="btnToggleUrgentNoTexting">Remove Urgent Status</a>
						<% Else %>
							<a href="#" class="btn btn-danger btn-lg btn-block btn-huge" id="btnToggleUrgentNoTexting">Mark Ticket As Urgent</a>
						<% End If %>
					</div>
					<%
				End If
			%>	        
	        <div class="col-lg-11">
	            <a href="#" class="btn btn-primary btn-lg btn-block btn-huge" id="btnTransferRedispatch" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= TicketNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>" data-target="#serviceBoardXferModal" data-title="Transfer or Redispatch" style="cursor:pointer;">Transfer or Redispatch</a>
	        </div>
	        <div class="col-lg-11">
	            <a href="#" class="btn btn-warning btn-lg btn-block btn-huge" id="btnSetAlert" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= TicketNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-target="#serviceBoardSetAlertModal" style="cursor:pointer;">Set Alert (Coming Soon)</a>
	        </div>
	        <% If Session("UserNo") <> "" Then %>
		        <% If UserIsServiceManager(Session("UserNo")) = True OR userIsAdmin(Session("UserNo")) = True Then %>
			        <!--<div class="col-lg-11">
			            <a href="editServiceMemo.asp?memo=<%= TicketNumber %>" class="btn btn-info btn-lg btn-block btn-huge" id="btnCloseCancelEarly">Close or Cancel Early</a>
			        </div>-->
			    <% End If %>
			<% End If %>
	        <div class="col-lg-11">
	            <a href="#" class="btn btn-success btn-lg btn-block btn-huge" id="btnRequestETA" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= TicketNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-target="#serviceBoardRequestETAModal" style="cursor:pointer;">Request an ETA (Coming Soon)</a>
	        </div>
	        <% If FilterChangeModuleOn() Then %>
		        <!--<div class="col-lg-11">
		            <a href="#" class="btn btn-info btn-lg btn-block btn-huge" id="btnChangeType" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= TicketNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-target="#serviceBoardChangeTypeModal" style="cursor:pointer;">Change Ticket Type</a>
		        </div>-->
	        <% End If %>
	     </div>					        
	</div>	
	

<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub ToggleTicketAsUrgentFromServiceBoardModal() 

	MemoNum = Request.Form("memoNum")

	SQLToggleUrgent = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & MemoNum  & "'"
	
	Response.write(SQLToggleUrgent)
	
	Set cnnToggleUrgent = Server.CreateObject("ADODB.Connection")
	cnnToggleUrgent.open (Session("ClientCnnString"))
	
	Set rsToggleUrgent = Server.CreateObject("ADODB.Recordset")
	rsToggleUrgent.CursorLocation = 3 
	Set rsToggleUrgent = cnnToggleUrgent.Execute(SQLToggleUrgent)
	
	If Not rsToggleUrgent.Eof Then
		'All records should be the same so just check the 1st one
		If rsToggleUrgent("Urgent") <> 1 Then 
			Urgent = 1
			CreateAuditLogEntry "Service Ticket Urgency Changed","Service Ticket Urgency Changed","Minor",0,"Service ticket #: " & MemoNum & " - changed to urgent"
		Else
			Urgent = 0
			CreateAuditLogEntry "Service Ticket Urgency Changed","Service Ticket Urgency Changed","Minor",0,"Service ticket #: " & MemoNum & " - changed to not urgent"		
		End If	
		SQL = "UPDATE FS_ServiceMemos Set Urgent = " & Urgent & " WHERE MemoNumber= '" & MemoNum & "'"
		Set rsToggleUrgent = cnnToggleUrgent.Execute(SQL)
		'If we are here, we found details records so the headers should be un-marked, regardless of the Urgency
		'SQL = "UPDATE FS_ServiceMemos Set Urgent = 0 WHERE MemoNumber= '" & MemoNum & "'"
		'Set rsToggleUrgent = cnnToggleUrgent.Execute(SQL)
	
	Else
		'There were no details found which means:
		'1. Advanced dispatch is not on
		'2. The ticket hasn't bee dispatched yet
		'So we mark the header instead
		SQL = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & MemoNum  & "'"
		Set rsToggleUrgent = cnnToggleUrgent.Execute(SQL)
		If Not rsToggleUrgent.Eof Then
			'All records should be the same so just check the 1st one
			If rsToggleUrgent("Urgent") <> 1 Then 
				Urgent = 1
				CreateAuditLogEntry "Service Ticket Urgency Changed","Service Ticket Urgency Changed","Minor",0,"Service ticket #: " & MemoNum & " - changed to urgent"
			Else
				Urgent = 0
				CreateAuditLogEntry "Service Ticket Urgency Changed","Service Ticket Urgency Changed","Minor",0,"Service ticket #: " & MemoNum & " - changed to not urgent"		
			End If	
			SQL = "UPDATE FS_ServiceMemos Set Urgent = " & Urgent & " WHERE MemoNumber= '" & MemoNum & "'"
			Set rsToggleUrgent = cnnToggleUrgent.Execute(SQL)
		End If
	End If
	
	set rsToggleUrgent = Nothing
	cnnToggleUrgent.Close
	Set cnnToggleUrgent = Nothing
	

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************





'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub ToggleTicketAsUrgentFromServiceBoardModalAndSendText() 

	MemoNum = Request.Form("memoNum")
	CustID = Request.Form("custID")
	TechUserNo = Request.Form("userNo")
	ReturnURL = Request.Form("returnURL")

	SQLToggleUrgent = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & MemoNum  & "'"
	Set cnnToggleUrgent = Server.CreateObject("ADODB.Connection")
	cnnToggleUrgent.open (Session("ClientCnnString"))
	
	Set rsToggleUrgent = Server.CreateObject("ADODB.Recordset")
	rsToggleUrgent.CursorLocation = 3 
	Set rsToggleUrgent = cnnToggleUrgent.Execute(SQLToggleUrgent)
	
	If Not rsToggleUrgent.Eof Then
		'All records should be the same so just check the 1st one
		If rsToggleUrgent("Urgent") <> 1 Then 
			Urgent = 1
			CreateAuditLogEntry "Service Ticket Urgency Changed","Service Ticket Urgency Changed","Minor",0,"Service ticket #: " & MemoNum & " - changed to urgent"
		Else
			Urgent = 0
			CreateAuditLogEntry "Service Ticket Urgency Changed","Service Ticket Urgency Changed","Minor",0,"Service ticket #: " & MemoNum & " - changed to not urgent"		
		End If	
		SQL = "UPDATE FS_ServiceMemos Set Urgent = " & Urgent & " WHERE MemoNumber= '" & MemoNum & "'"
		Set rsToggleUrgent = cnnToggleUrgent.Execute(SQL)
		'If we are here, we found details records so the headers should be un-marked, regardless of the Urgency
		'SQL = "UPDATE FS_ServiceMemos Set Urgent = 0 WHERE MemoNumber= '" & MemoNum & "'"
		'Set rsToggleUrgent = cnnToggleUrgent.Execute(SQL)
	
	Else
		'There were no details found which means:
		'1. Advanced dispatch is not on
		'2. The ticket hasn't bee dispatched yet
		'So we mark the header instead
		SQL = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & MemoNum  & "'"
		Set rsToggleUrgent = cnnToggleUrgent.Execute(SQL)
		If Not rsToggleUrgent.Eof Then
			'All records should be the same so just check the 1st one
			If rsToggleUrgent("Urgent") <> 1 Then 
				Urgent = 1
				CreateAuditLogEntry "Service Ticket Urgency Changed","Service Ticket Urgency Changed","Minor",0,"Service ticket #: " & MemoNum & " - changed to urgent"
			Else
				Urgent = 0
				CreateAuditLogEntry "Service Ticket Urgency Changed","Service Ticket Urgency Changed","Minor",0,"Service ticket #: " & MemoNum & " - changed to not urgent"		
			End If	
			SQL = "UPDATE FS_ServiceMemos Set Urgent = " & Urgent & " WHERE MemoNumber= '" & MemoNum & "'"
			Set rsToggleUrgent = cnnToggleUrgent.Execute(SQL)
		End If
	End If
	
	set rsToggleUrgent = Nothing
	cnnToggleUrgent.Close
	Set cnnToggleUrgent = Nothing
	

	'**********************
	'Send text 
	'**********************
	

	If getUserCellNumber(TechUserNo) <> "" Then
		Send_To = getUserCellNumber(TechUserNo)

		URL = BaseURL & "inc/sendtext.php"
		QString = "?n=" & Replace(getUserCellNumber(TechUserNo),"-","")
			
		QString = QString & "&u1=" & EzTextingUserID()
		QString = QString & "&u2=" & EzTextingPassword()
		
		QString = QString & "&t=NOTIFICATION"
		
		
		If InStr(ReturnURL, "FieldServiceKioskNoPaging") Then
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & ReturnURL & "?pp=" & Session("PassPhrase") & "&cl=" & Session("ClientKey") & "&ri=" & Session("RefreshInterval"))
		Else
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & ReturnURL)
		End If	
		
		If Urgent = 1 Then
			'Text message should alert them that ticket is urgent
			
			If GetCustNameByCustNum(CustID) <> "" Then
				txtMSG = "The service ticket for " & EZTexting_Filter1(Replace(GetCustNameByCustNum(CustID),"&"," ")) & " has been marked as URGENT. "
			Else
				txtMSG = "The service ticket for " & GetTerm("Account") & ": " & CustID & " has been marked as URGENT. "
			End If
		Else
			'Text message should alert them that ticket is no longer urgent
			If GetCustNameByCustNum(CustID) <> "" Then
				txtMSG = "The service ticket for " & EZTexting_Filter1(Replace(GetCustNameByCustNum(CustID),"&"," ")) & " is no longer URGENT. "
			Else
				txtMSG = "The service ticket for " & GetTerm("Account") & ": " & CustID & " is no longer URGENT. "
			End If
		End If
		
		
		QString = QString & "&m=" & txtMSG 

		QString = QString & "    Tap the link to see the details for this ticket    "
		QString = QString & Server.URLEncode(baseURL & "directlaunch/service/moreinfo_dispatch_from_urgent_text.asp?t=" & MemoNum & "&u=" & TechUserNo & "&c=" & CustID & "&cl=" & MUV_READ("SERNO"))


		QString = Replace(Qstring," ", "%20")

		Response.Redirect (URL & Qstring)

		Description = "An urgent ticket text message was sent to " & GetUserDisplayNameByUserNo(TechUserNo) & " (" & getUserCellNumber(TechUserNo) & ") at " & NOW()
		CreateAuditLogEntry "Service Ticket System","Urgent ticket text message sent","Minor",0,Description
	Else
		' Could not send dispatch test, no address on file
		emailBody = "Insight was unable to send an urgent text message to " & GetUserDisplayNameByUserNo(TechUserNo) & ". No cell number on file"
		If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com"
		SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send urgent text message",emailBody,GetTerm("Service"),"Missing Cell Number"
		Description = "Insight was unable to send a urgent text message to " & GetUserDisplayNameByUserNo(TechUserNo) & ". No cell number on file"
		CreateAuditLogEntry "Service Ticket System","Unable to send urgent text message","Major",0,Description

	End If

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetTitleForServiceBoardTransferRedispatchModal() 

	MemoNumber = Request.Form("memoNum")
	UserNo = Request.Form("userNo")
	%>
	
	<%= GetUserDisplayNameByUserNo(UserNo) %>&nbsp;-&nbsp;Ticket #: <%= MemoNumber %>&nbsp;-&nbsp;<%= GetTerm("Account") %> #:<%=GetServiceTicketCust(MemoNumber) %>
	
	<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForServiceBoardTransferRedispatchModal() 

	MemoNumber = Request.Form("memoNum")
	UserNo = Request.Form("userNo")
	CustID = Request.Form("custID")
	ReturnURL = Request.Form("returnURL")
	
	%>
	<input type="hidden" id="txtReturnURL" name="txtReturnURL" value="<%= ReturnURL %>">
	<input type="hidden" id="txtServiceTicketNumber" name="txtServiceTicketNumber" value="<%= MemoNumber %>">
	<input type="hidden" id="txtAccountNumber" name="txtAccountNumber" value="<%= GetServiceTicketCust(MemoNumber) %>">
	
	<!-- field techs !-->
	<div class="col-lg-12">
		<p align="left"><label>Select <%= GetTerm("Field Service Tech") %> to reassign this ticket to</label></p>
		<select name="selFieldTech" id="selFieldTech" multiple="multiple" class="form-control the-select" style="height:200px;">
			<%	
			SQLTransferRedispatchModal = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".tblUsers WHERE UserNo <> " & UserNo & " AND (userType = 'Field Service' OR userType = 'Service Manager') and userArchived <> 1 Order By UserType,userDisplayName"
			
			Set cnnTransferRedispatch = Server.CreateObject("ADODB.Connection")
			cnnTransferRedispatch.open (Session("ClientCnnString"))
			Set rsTransferRedispatch = Server.CreateObject("ADODB.Recordset")
			rsTransferRedispatch.CursorLocation = 3 
			Set rsTransferRedispatch = cnnTransferRedispatch.Execute(SQLTransferRedispatchModal)

			If not rsTransferRedispatch.EOF Then

				Do While Not rsTransferRedispatch.EOF
					userFirstName = rsTransferRedispatch("userFirstName")
					userLastName = rsTransferRedispatch("userLastName")
					userDisplayName = rsTransferRedispatch("userDisplayName")
					userEmail = rsTransferRedispatch("userEmail")
					userNo = rsTransferRedispatch("UserNo")
					
					%><option value="<%= userNo %>"><%= userFirstName %>&nbsp;<%= userLastName %></option><%
					
					rsTransferRedispatch.MoveNext
				Loop

			End If
			
			Set rsTransferRedispatch = Nothing
			cnnTransferRedispatch.Close
			Set cnnTransferRedispatch = Nothing
			%>
			<option value="0">UN-DISPATCH</option>
		</select>
	 </div>
	<!-- eof field techs !-->

	<!-- checkboxes !-->
	<div class="col-lg-12">
		<div class="checkbox">
			<label>
   				<input type="checkbox" name="chkSendEmail" id="chkSendEmail" <%If Instr(FSDefaultNotificationMethod(),"Email") <> 0 Then Response.Write(" checked")%>>
   				<strong>Send email</strong>
			</label>
		</div>
		<div class="checkbox">
			<label>
			    <input type="checkbox" name="chkSendText" id="chkSendText" <%If Instr(FSDefaultNotificationMethod(),"Text") <> 0 Then Response.Write(" checked")%>>
			    <strong>Send text message</strong>
			</label>
		</div>
	</div>
	<!-- eof checkboxes !-->
		
<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetTitleForServiceBoardChangeTypeModal() 

	MemoNumber = Request.Form("memoNum")
	UserNo = Request.Form("userNo")
	%>
	
	Change Ticket Type - Ticket #: <%= MemoNumber %>&nbsp;-&nbsp;<%= GetTerm("Account") %> #:<%=GetServiceTicketCust(MemoNumber) %>
	
	<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForServiceBoardChangeTypeModal() 

	MemoNumber = Request.Form("memoNum")
	CustID = Request.Form("custID")
	ReturnURL = Request.Form("returnURL")
	
	%>
	<input type="hidden" id="txtReturnURL" name="txtReturnURL" value="<%= ReturnURL %>">
	<input type="hidden" id="txtServiceTicketNumber" name="txtServiceTicketNumber" value="<%= MemoNumber %>">
	<input type="hidden" id="txtAccountNumber" name="txtAccountNumber" value="<%= GetServiceTicketCust(MemoNumber) %>">
	
	<!-- field techs !-->
	<div class="col-lg-12">
		<p align="left"><label>Change service ticket type</label></p>
		<%
		'Figure out what type of ticket it currently is
		CurrentTicketType = "Service"
		
		SQLChangeType = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '"  & MemoNumber & "'"
			
		Set cnnChangeType = Server.CreateObject("ADODB.Connection")
		cnnChangeType.open (Session("ClientCnnString"))
		Set rsChangeType = Server.CreateObject("ADODB.Recordset")
		rsChangeType.CursorLocation = 3 
		
		Set rsChangeType = cnnChangeType.Execute(SQLChangeType)

		If not rsChangeType.EOF Then

			If rsChangeType("FilterChange") = 1 Then CurrentTicketType = "Filter Change"
			
				
		End If
			
		Set rsChangeType = Nothing
		cnnChangeType.Close
		Set cnnChangeType = Nothing
		
		%>
	 </div>
	<!-- eof field techs !-->

	<!-- checkboxes !-->
	<div class="col-lg-12">
		<div class="radio">
			<label>
   				<input type="radio" name="optTicketType" id="optServiceTicket" value="Service Ticket"<%If CurrentTicketType = "Service" Then Response.Write(" checked ")%>>
   				<strong>Service Ticket</strong>
			</label>
		</div>
		<% If filterChangeModuleOn() Then %>
			<div class="radio">
				<label>
	   				<input type="radio" name="optTicketType" id="optFilterChange" value="Filter Change"<%If CurrentTicketType = "Filter Change" Then Response.Write(" checked ")%>>
	   				<strong>Filter Change</strong>
				</label>
			</div>
		<% End If %>
	</div>
	<!-- eof checkboxes !-->
<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetTitleForServiceBoardDispatchModal() 

	MemoNumber = Request.Form("memoNum")

	CustID = GetServiceTicketCust(MemoNumber) 
	%>
	Dispatch<br><br><%= GetTerm("Account") %> #:<%=CustID%> - <%=GetCustNameByCustNum(CustID) %>
	<%
	
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForServiceBoardDispatchModal() 

	MemoNumber = Request.Form("memoNum")
	ReturnURL = Request.Form("returnURL")
	CustID = GetServiceTicketCust(MemoNumber) 
	ListOfPossibleTickets = ""
	
	%>
	<input type="hidden" id="txtReturnURL" name="txtReturnURL" value="<%= ReturnURL %>">
	<input type="hidden" id="txtServiceTicketNumber" name="txtServiceTicketNumber" value="<%= MemoNumber %>">
	<input type="hidden" id="txtAccountNumber" name="txtAccountNumber" value="<%= CustID %>">
	
	<%' Lookup ALL OPEN service tickets for this client
	Set rsmodal = Server.CreateObject("ADODB.Recordset")

	Set cnnChangeType = Server.CreateObject("ADODB.Connection")
	cnnChangeType.open (Session("ClientCnnString"))
	Set rsChangeType = Server.CreateObject("ADODB.Recordset")
	rsChangeType.CursorLocation = 3 
	
	SQLChangeType = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND AccountNumber='" & CustID  & "' ORDER BY submissionDateTime DESC"	
	
	Set rsChangeType = cnnChangeType.Execute(SQLChangeType)
	%>
	<div class="col-lg-12">
	<%
	If not rsChangeType.EOF Then %>

		<table class="table table-condensed table-hover large-table">			
			<thead>
			  <tr style="background-color: #EEE;">
			  	<th style="width: 10%;">Ticket#</th>
			  	<th style="width: 5%;">Type</th>
			  	<th style="width: 50%;">Description</th>
			  	<th style="width: 15%;">Stage</th>
			  	<th style="width: 20%;">Select Technician</th>
			  </tr>
			</thead>
			<tbody>
		<%
		Do While Not rsChangeType.EOF
		
			MemoNumber = rsChangeType("MemoNumber") %>	
			<tr>
			<td><%=MemoNumber%></td>
			<td>
			<%If filterChangeModuleOn() Then
				If rsChangeType("FilterChange") = 1 Then
					%><span class="filtercircle"  title="Filter change">F</span><%
				End If
			End If
			If rsChangeType("FilterChange") <> 1 Then
					%><span class="bluecircle"  title="Service Ticket">S</span><%
			End If
			%>
			</td>
			<td>
			<%If filterChangeModuleOn() Then
				If rsChangeType("FilterChange") = 1 Then
					Response.Write("Filter Change")
				Else
					Response.Write(rsChangeType("ProblemDescription"))
				End If
			Else
				Response.Write(rsChangeType("ProblemDescription"))
			End IF%>
			</td>
			<td>
				<%
				GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsChangeType("MemoNumber"))
				If rsChangeType("RecordSubType") <> "HOLD" AND (GetServiceTicketCurrentStageVar = "Received" or GetServiceTicketCurrentStageVar = "Released") Then
					Response.Write("<span class='labelAwaitingDispatch'>Awaiting Dispatch</span><br>")
				Else
					If GetServiceTicketCurrentStageVar = "Awaiting Acknowledgement" Then
						Response.Write("<span class='labelAwaitingAcknowledgement'>" & GetServiceTicketCurrentStageVar  & "</span>")
					ElseIf GetServiceTicketCurrentStageVar = "Dispatch Acknowledged" Then
						Response.Write("<span class='labelDispatchAcknowledged'>" & GetServiceTicketCurrentStageVar  & "</span>")
					ElseIf GetServiceTicketCurrentStageVar = "En Route" Then
						Response.Write("<span class='labelEnRoute'>" & GetServiceTicketCurrentStageVar  & "</span>")
					ElseIf GetServiceTicketCurrentStageVar = "On Site" Then
						Response.Write("<span class='labelOnSite'>" & GetServiceTicketCurrentStageVar  & "</span>")
					ElseIf GetServiceTicketCurrentStageVar = "Swap" Then
						Response.Write("<span class='labelSwap'>" & GetServiceTicketCurrentStageVar  & "</span>")
					ElseIf GetServiceTicketCurrentStageVar = "Wait for parts" Then
						Response.Write("<span class='labelWaitForParts'>" & GetServiceTicketCurrentStageVar  & "</span>")
					ElseIf GetServiceTicketCurrentStageVar = "Follow Up" Then
						Response.Write("<span class='labelFollowUp'>" & GetServiceTicketCurrentStageVar  & "</span>")
					ElseIf GetServiceTicketCurrentStageVar = "Unable To Work" Then
						Response.Write("<span class='labelUnableToWork'>" & GetServiceTicketCurrentStageVar  & "</span>")
					Else
						Response.Write("<span class='label-default'>" & GetServiceTicketCurrentStageVar  & "</span>")
					End If
				End If
				%>
			</td>
			<td>
				<%
				' it is currently dispatched
				If GetServiceTicketCurrentStageVar  = "Awaiting Acknowledgement" or GetServiceTicketCurrentStageVar  = "Dispatched" or GetServiceTicketCurrentStageVar  = "Dispatch Acknowledged" or GetServiceTicketCurrentStageVar  = "En Route" or GetServiceTicketCurrentStageVar  = "On Site" Then 
					ticketStageDateTime = GetServiceTicketSTAGEDateTime(rsChangeType("MemoNumber"),GetServiceTicketCurrentStageVar)
					ticketStageHour = Hour(ticketStageDateTime)
					ticketStageMinute = Minute(ticketStageDateTime)
					ticketStageZeroFactor = "0" & ticketStageMinute
					ticketStageAMPM = "AM"
					If ticketStageHour >= 12 then ticketStageAMPM = "PM"
					If ticketStageHour > 12 then ticketStageHour = ticketStageHour - 12
					If ticketStageMinute <= 9 then ticketStageMinute = ticketStageZeroFactor	 
					ticketStageDateTimeDisplay = padDate(MONTH(ticketStageDateTime),2) & "/" & padDate(DAY(ticketStageDateTime),2) & "/" & padDate(RIGHT(YEAR(ticketStageDateTime),2),2)
					Response.Write("<br>" & GetUserDisplayNameByUserNo(GetServiceTicketDispatchedTech(rsChangeType("MemoNumber"))) & "<br>")
					Response.Write(ticketStageDateTimeDisplay & " " & ticketStageHour & ":" & ticketStageMinute & " " & ticketStageAMPM)
				Else
					ListOfPossibleTickets = ListOfPossibleTickets & MemoNumber & "," %>
					<select name='selFieldTech<%=MemoNumber%>' id='selFieldTech<%=MemoNumber%>'>
						<option value='0'>Do Not Dispatch</option>
						<%
						SQLmodal = "SELECT * FROM tblUsers WHERE userArchived <> 1 ORDER BY (CASE WHEN [UserType] ='Field Service' THEN 0 ELSE 1 END) ,userDisplayName"
						
						Set rsmodal = cnnChangeType.Execute(SQLmodal)
		
						If not rsmodal.EOF Then
		
							Do While Not rsmodal.EOF
								userFirstName = rsmodal("userFirstName")
								userLastName = rsmodal("userLastName")
								userNo = rsmodal("UserNo")
								
								%><option value='<%=userNo%>'><%=userFirstName%>&nbsp;<%=userLastName%></option><%
								
								rsmodal.MoveNext
							Loop
		
						End If
						
						%>
					</select>
				<% End If %>
			</td>
			</tr>
				<%
				' If it is a filter change, write a row where we give a little more info
				If filterChangeModuleOn() Then
					If rsChangeType("FilterChange") = 1 Then
					
						'Get & display the individual filter information
						
						Set rsActiveTicket2 = Server.CreateObject("ADODB.Recordset")
						Set rsActiveTicket = Server.CreateObject("ADODB.Recordset")
						SQActiveTicket = "SELECT * FROM FS_ServiceMemosFilterInfo WHERE ServiceTicketID = '" &  rsChangeType("MemoNumber") & "'"
					
						Set rsActiveTicket = cnnChangeType.Execute(SQActiveTicket)
	
						If Not rsActiveTicket.EOF Then
						
							Do While Not rsActiveTicket.EOF
		
								SQActiveTicket = "SELECT * FROM FS_CustomerFilters WHERE InternalRecordIdentifier = " &  rsActiveTicket("CustFilterIntRecId") 
							
								Set rsActiveTicket2 = cnnChangeType.Execute(SQActiveTicket)
								
								'Response.Write(SQActiveTicket)
			
								Response.Write("<tr>")
								Response.Write("<td style='border-top: 0px !important; border-bottom: 0px !important;'>&nbsp;</td>")
								Response.Write("<td colspan='2' style='border-top: 0px !important; border-bottom: 0px !important;'><small>" & GetFilterIDByIntRecID(rsActiveTicket2("FilterIntRecID")) & " - "  & GetFilterDescByIntRecID(rsActiveTicket2("FilterIntRecID")) & "</small></td>")	
								If rsActiveTicket2("Notes") <> "" Then
									Response.Write("<td colspan='3' style='border-top: 0px !important; border-bottom: 0px !important;'><small>(Location: " & rsActiveTicket2("Notes") & ")"& "</td>")
								Else
									Response.Write("<td colspan='3' style='border-top: 0px !important; border-bottom: 0px !important;'>&nbsp;</td>")
								End If
	
								Response.Write("</tr>")
								
								rsActiveTicket.MoveNext
		
							Loop
	
						End If

				End If
			End If
			%>
			<%
			rsChangeType.MoveNext
		Loop

		'******************************************************************
		' Now we have to write the info for any filter changes that are due
		'******************************************************************
		If filterChangeModuleOn() Then
		
				'Remember, it reads the setting FieldServiceDays from tblSetting_Global to determine how many days to use in the evaluation
				SQLPendingFilterChange = "SELECT * FROM Settings_Global"
				Set cnnPendingFilterChange = Server.CreateObject("ADODB.Connection")
				cnnPendingFilterChange.open (Session("ClientCnnString"))
				Set rsPendingFilterChange = Server.CreateObject("ADODB.Recordset")
				rsPendingFilterChange.CursorLocation = 3 
				Set rsPendingFilterChange = cnnPendingFilterChange.Execute(SQLPendingFilterChange)
				If not rsPendingFilterChange.EOF Then FilterChangeDays = rsPendingFilterChange("FilterChangeDays") Else FilterChangeDays = 15
				set rsPendingFilterChange = Nothing
			
	
				SQLPendingFilterChange = "SELECT TOP 1 * , FS_CustomerFilters.InternalRecordIdentifier AS CFIntRecID ,"
				SQLPendingFilterChange = SQLPendingFilterChange & " CASE WHEN FS_CustomerFilters.FrequencyType='D' THEN DATEADD(day, FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
				SQLPendingFilterChange = SQLPendingFilterChange & " WHEN FS_CustomerFilters.FrequencyType='M' THEN DATEADD(day, 28*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
				SQLPendingFilterChange = SQLPendingFilterChange & " WHEN FS_CustomerFilters.FrequencyType='W' THEN DATEADD(day, 7*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
				SQLPendingFilterChange = SQLPendingFilterChange & " ELSE FS_CustomerFilters.LastChangeDateTime END AS nextdate "
				SQLPendingFilterChange = SQLPendingFilterChange & " FROM FS_CustomerFilters WHERE "
				SQLPendingFilterChange = SQLPendingFilterChange & " CustID = '" & CustID  & "' AND "
				SQLPendingFilterChange = SQLPendingFilterChange & " CASE WHEN FS_CustomerFilters.FrequencyType='D' THEN DATEADD(day, FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
				SQLPendingFilterChange = SQLPendingFilterChange & " WHEN FS_CustomerFilters.FrequencyType='M' THEN DATEADD(day, 28*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
				SQLPendingFilterChange = SQLPendingFilterChange & " WHEN FS_CustomerFilters.FrequencyType='W' THEN DATEADD(day, 7*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
				SQLPendingFilterChange = SQLPendingFilterChange & " ELSE FS_CustomerFilters.LastChangeDateTime END"
				SQLPendingFilterChange = SQLPendingFilterChange & " <= DateAdd(day," & FilterChangeDays & ",getdate()) "
				SQLPendingFilterChange = SQLPendingFilterChange & " AND "
				SQLPendingFilterChange = SQLPendingFilterChange & "("
				SQLPendingFilterChange = SQLPendingFilterChange & " FS_CustomerFilters.FilterIntRecID NOT IN "
				SQLPendingFilterChange = SQLPendingFilterChange & "(SELECT FilterIntRecID FROM FS_ServiceMemosFilterInfo WHERE CustID = '" & CustID & "' AND ServiceTicketID IN (Select MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus='OPEN')) "
				SQLPendingFilterChange = SQLPendingFilterChange & ")"
				
				'Response.Write(SQLPendingFilterChange )

				Set rsPendingFilterChange = Server.CreateObject("ADODB.Recordset")
				rsPendingFilterChange.CursorLocation = 3 
				Set rsPendingFilterChange = cnnPendingFilterChange.Execute(SQLPendingFilterChange)
				
				If not rsPendingFilterChange.eof then 
		
					SelectShown = 0
					
					Do While Not rsPendingFilterChange.eof 
					
							%>							
							<tr>
								<td>Pending</td>
								<td>&nbsp;</td>
								<td><%=rsPendingFilterChange("Notes")%></td>
								<td> 
								<%
									DaysTilChange = datediff("d",Date(),rsPendingFilterChange("nextdate"))
									If DaysTilChange < 0 Then ' overdue	%>
										<span class="labelFilterChangeIndicatorAndButtonColorOverDue">Filter change<br>overdue <%=DaysTilChange%> days</span>
									<% Else %>
										<span class="labelFilterChangeIndicatorAndButtonColor">Filter change due<br>in <%=DaysTilChange%> days</span>
									<% End If %>
								</td>

								<td>
									<%
									
									' In this case the MemoNumber is different
									MemoNumber = "F" & rsPendingFilterChange("CFIntRecID")
									ListOfPossibleTickets = ListOfPossibleTickets & MemoNumber & "," %>
									
									<% If SelectShown = 0 Then %>
									
										<select name='selFieldTech<%=MemoNumber%>' id='selFieldTech<%=MemoNumber%>'>
											<option value='0'>Do Not Dispatch</option>
											<%
											
												SQLmodal = "SELECT * FROM tblUsers WHERE userArchived <> 1 ORDER BY (CASE WHEN [UserType] ='Field Service' THEN 0 ELSE 1 END) ,userDisplayName"
												
												Set rsmodal = cnnChangeType.Execute(SQLmodal)
								
												If not rsmodal.EOF Then
								
													Do While Not rsmodal.EOF
														userFirstName = rsmodal("userFirstName")
														userLastName = rsmodal("userLastName")
														userNo = rsmodal("UserNo")
														
														%><option value='<%=userNo%>'><%=userFirstName%>&nbsp;<%=userLastName%></option><%
														
														rsmodal.MoveNext
													Loop
								
												End If
											%>
										</select>
										
										<% SelectShown = 1 %>
										
									<% Else %>
									
										<select name='selFieldTech<%=MemoNumber%>' id='selFieldTech<%=MemoNumber%>'>
											<option value='0'>Do Not Dispatch</option>
											<option value='9999'>Dispatch to same</option>
										</select>
										
									<% End If %>
								</td>
							</tr>
						<%

						rsPendingFilterChange.movenext
					Loop
					
				End If
				
				set rsPendingFilterChange = Nothing
				cnnPendingFilterChange.Close
				set cnnPendingFilterChange = Nothing
	
		End If


		'*********************************************************************
		'EOF Now we have to write the info for any filter changes that are due
		'*********************************************************************
	
		Response.Write("</tbody>")
		Response.Write("</table>")
		
	End If 
	%>
	</div>

<%

If Right(ListOfPossibleTickets,1) = "," Then ListOfPossibleTickets = Left(ListOfPossibleTickets,Len(ListOfPossibleTickets)-1) ' Strip trailing comma

%><input type="hidden" id="txtListOfPossibleTickets" name="txtListOfPossibleTickets" value="<%= ListOfPossibleTickets %>"><%
	
Set rsChangeType = Nothing
Set rsmodal = Nothing
cnnChangeType.Close
Set cnnChangeType= Nothing


End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetContentForServiceTicketNotesModal() 

	MemoNumber = Request.Form("memoNum")
	CustID = Request.Form("custID")
	UserNoOfServiceTech = Request.Form("userNo")
	CustNamePassed = GetCustNameByCustNum(CustID)
	
	'********************************************************
	'CODE HERE TO MARK NOTES AS BEING READ
	
	Call MarkNoteNewForUserServiceTicket(MemoNumber)
	
	'********************************************************
%>			
	
	<%'******************
	' **** Notes Tab ****
	'********************
	%>
	
	<!-- sort table script !-->
	<script src="../../js/sorttable.js"></script>
	<!-- eof sort table script !-->
	
	<!-- modal header !-->
	<div class="modal-header" style="min-height:35px !important;">
		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		<h4 class="modal-title" id="myWebOrdersLabel">Notes For Ticket #<%= MemoNumber %>&nbsp;-&nbsp;<%= GetTerm("Account") %> #<%= CustID %> (<%= CustNamePassed %>) </h4>
	</div>
	<!-- eof modal header !-->
	
	<!-- modal body !-->
	<div class="modal-body" style="max-height:450px;overflow:scroll">
	
	<input type="hidden" name="txtTicketNumberToPass" id="txtTicketNumberToPass" value="<%= MemoNumber %>">
									
	<div id="log">
	
		<p> <button type="button" class="btn btn-success" onclick="ajaxRowNewLogNotes();">New Note</button> </p>
	
			<div class="input-group narrow-results"> <span class="input-group-addon">Search Notes</span>
			    <input id="filter-notes" type="text" class="form-control filter-search-width" placeholder="Type here...">
			</div>
		  <br>
			<div class="table-responsive">
	            <table id="ajaxContainerLogNotesTable" class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
	              <thead>
	                <tr>
	                  <th width="15%">Date</th>
					  <th width="15%">Time</th>
					  <th width="20%">Entered By</th>
					  <th width="40%">Details</th>
	                  <th class="sorttable_nosort text-center" style="width: 80px;">Actions</th>
	                </tr>
	              </thead>
	
				<tbody id="ajaxContainerLogNotes" class='searchable-notes ajax-loading'></tbody>
			</table>
		</div>
	</div>
	<%'**********************
	' **** eof Notes Tab ****
	'************************
	%>
	
	<script>
		$(document).ready(function () { ajaxLoadLogNotes(); });

		function ajaxRowNewLogNotes() {
			var value = {};
			value.id = 0;
			value.Date = "-";
			value.Time = "-";
			value.User = "-";
			value.LogNote = "";
			$('#ajaxRowLogNotes-' + 0 + '').remove();		
			$("#ajaxContainerLogNotes").prepend(ajaxRowHtmlNotes(value));
		}
		function ajaxRowHtmlNotes(value) {
			var btns = '\
						<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'LogNotes\', ' + value.id + ', \'Edit\');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxLoadLogNotes(\'delete\', ' + value.id + ');"><i class="fas fa-trash-alt"></i></a></div>\
						<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadLogNotes(\'save\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'LogNotes\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
					';
			if(value.id==0)
				btns = '\
						<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadLogNotes(\'insert\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'LogNotes\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
					';					
			var html = '\
				<tr id="ajaxRowLogNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>';								
					html += '<td>\
						<div class="visibleRowView">' + value.LogNote + '</div>\
						<div class="visibleRowEdit"><input class="form-control" data-type="LogNote" value="' + value.LogNote.replace(/"/g, '&quot;') + '" /></div>\
					</td>\
					<td class="text-center">'+btns+'</td>\
			   </tr>\
				';				
			return html;
		}
		function ajaxLoadLogNotes(updateAction, updateActionId) {
		
			updateServiceTicketID = $("#txtTicketNumberToPass").val();					
			if (updateAction == "delete" && !confirm("Are you sure you want to delete this service ticket note?")) return;
			$("#ajaxContainerLogNotes").addClass("ajax-loading");
			var url = "ajax/serviceTicketNote_log.asp?i=<%= IntRecID %>&serviceTicketID=" + updateServiceTicketID;
			var jsondata = {};
			jsondata.updateAction = updateAction;
			jsondata.updateActionId = updateActionId;
			jsondata.updateServiceTicketID = updateServiceTicketID;

			if(updateAction=="save" || updateAction=="insert"){
				jsondata.LogNote= $('#ajaxRowLogNotes-' + updateActionId + ' [data-type="LogNote"]').val();
			}
			
			$.ajax({
				type: "POST",
				url: url,
				dataType: "json",
				data: jsondata,
				success: function (data) {					
					//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }				
					var html = "";
					$("#ajaxContainerLogNotesTable").find('tr').each(function(){
						$(this).find('th').eq(1).attr('width', '15%');;
						$(this).find('th').eq(2).attr('width', '15%');;
						$(this).find('th').eq(3).attr('width', '20%');;	
						$(this).find('th').eq(3).attr('width', '40%');;							
					});	
					$.each(data, function (key, value) {
						html += ajaxRowHtmlNotes(value);
					});
					//alert(html);
					$("#ajaxContainerLogNotes").html(html);

					//var newTableObject = document.getElementById("#ajaxContainerLogNotesTable");
					//sorttable.makeSortable(newTableObject);
					
					setTimeout(function(){
						$("#ajaxContainerLogNotes").removeClass("ajax-loading");
					}, 0);
					
				}
			});
		}
	</script>	
	
	 <!-- custom table search !-->
	
	<script>
	
	$(document).ready(function () {
		
	    (function ($) {
	        
	        $('#filter-notes').keyup(function () {
	
	            var rex = new RegExp($(this).val(), 'i');
	            $('.searchable-notes tr').hide();
	            $('.searchable-notes tr').filter(function () {
	                return rex.test($(this).text());
	            }).show();
	        })
	 
	
	    }(jQuery));
	
	});
	</script>
	<!-- eof custom table search !-->
			</div>
<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub SaveAddNewFilter() 


	FilterID = Request.Form("FilterID")
	FilterDescription = Request.Form("FilterDescription")
	FilterCost = Request.Form("FilterCost")
	FilterListPrice = Request.Form("FilterListPrice")
	FilterTaxable = Request.Form("FilterTaxable")
	FilterInventoried = Request.Form("FilterInventoried")
	FilterPickable = Request.Form("FilterPickable")
	FilterUPC = Request.Form("FilterUPC")
	FilterRecordSource = "INSIGHT"
	FilterprodSKU = Request.Form("FilterprodSKU")
	FilterdisplayOrder = Request.Form("FilterdisplayOrder")
	
	'***************************************************************************************
	'Add this filter in IC_FILTERS
	'***************************************************************************************

	Set cnnSaveAddNewFilter = Server.CreateObject("ADODB.Connection")
	cnnSaveAddNewFilter.open (Session("ClientCnnString"))
	Set rsSaveAddNewFilter = Server.CreateObject("ADODB.Recordset")
	rsSaveAddNewFilter.CursorLocation = 3 

	SQLSaveAddNewFilter = "INSERT INTO IC_FILTERS (RecordSource, FilterID, Description, DefaultCost, ListPrice, "
	SQLSaveAddNewFilter = SQLSaveAddNewFilter & " InventoriedItem, PickableItem, Taxable, UPCCode, prodSKU,displayOrder) VALUES "
	SQLSaveAddNewFilter = SQLSaveAddNewFilter & " ('" & FilterRecordSource & "','" & FilterID & "','" & FilterDescription & "'," & FilterCost & "," & FilterListPrice & ","
	SQLSaveAddNewFilter = SQLSaveAddNewFilter & FilterInventoried & "," & FilterPickable & "," & FilterTaxable & ",'" & FilterUPC & "','" & FilterprodSKU & "'," & FilterdisplayOrder & ")"

	Response.write(SQLSaveAddNewFilter)
	
	Set rsSaveAddNewFilter = cnnSaveAddNewFilter.Execute(SQLSaveAddNewFilter)

	Description = GetUserDisplayNameByUserNo(0) & " added a new filter with an ID of <strong>" & FilterID & "<strong>, and description of <strong>" & FilterDescription & "<strong>."
	CreateAuditLogEntry GetTerm("Field Service") & " Filter Added",GetTerm("Field Service") & " Filter Added","Minor",0,Description
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForEditFilterModal() 

	IntRecID = Request.Form("IntRecID")
	
	%>
	<input type="hidden" id="txtIntRecID" name="txtIntRecID" value="<%= IntRecID %>">
	
	<%
	
	SQLEditFilter = "SELECT * FROM IC_FILTERS WHERE InternalRecordIdentifier = "  & IntRecID
		
	Set cnnEditFilter = Server.CreateObject("ADODB.Connection")
	cnnEditFilter.open(Session("ClientCnnString"))
	Set rsEditFilter = Server.CreateObject("ADODB.Recordset")
	rsEditFilter.CursorLocation = 3 
	
	Set rsEditFilter = cnnEditFilter.Execute(SQLEditFilter)

	If not rsEditFilter.EOF Then

		FilterID = rsEditFilter("FilterID")
		Description = rsEditFilter("Description")
		ListPrice = rsEditFilter("ListPrice")
		DefaultCost = rsEditFilter("DefaultCost")
		Taxable = rsEditFilter("Taxable")
		InventoriedItem = rsEditFilter("InventoriedItem")
		PickableItem = rsEditFilter("PickableItem")
		UPCCode = rsEditFilter("UPCCode")	
		prodSKU = rsEditFilter("prodSKU")
		displayOrder = rsEditFilter("displayOrder")
			
	End If
		
	Set rsEditFilter = Nothing
	cnnEditFilter.Close
	Set cnnEditFilter = Nothing
	
	%>
			
	<script type="text/javascript">
		
		$(document).ready(function() {
			
			$("#modalEditExistingFilter #btnEditFilterSave").bind("click",function(e){
			
				var IntRecID = $("#modalEditExistingFilter #txtIntRecID").val();
				var FilterID = $("#modalEditExistingFilter #txtFilterID").val();
				var FilterDescription = $("#modalEditExistingFilter #txtFilterDescription").val();
				var FilterCost = $("#modalEditExistingFilter #txtFilterCost").val();
				var FilterListPrice = $("#modalEditExistingFilter #txtFilterListPrice").val();
				var FilterTaxable = $("#modalEditExistingFilter #selFilterTaxable option:selected").val();
				var FilterInventoried = $("#modalEditExistingFilter #selFilterInventoried option:selected").val();
				var FilterPickable = $("#modalEditExistingFilter #selFilterPickable option:selected").val();
				var FilterUPC = $("#modalEditExistingFilter #txtFilterUPC").val();
				var FilterprodSKU = $("#modalEditExistingFilter #selFilterprodSKU").val();
				var FilterdisplayOrder = $("#modalEditExistingFilter #selFilterdisplayOrder").val();				
				
		
				if (FilterID.length <=0) {
					swal({
						title: 'Error Adding Filter',
						text: 'Please specify a filter ID',
						type: 'error'
					});
					return false;
				}
		
				if (FilterDescription.length <=0) {
					swal({
						title: 'Error Adding Filter',
						text: 'Please specify a filter description',
						type: 'error'
					});
					return false;
				}
						
				if (FilterTaxable.length <=0) {
					swal({
						title: 'Error Adding Filter',
						text: 'Please specify whether the filter is taxable or not.',
						type: 'error'
					});
					return false;
				}
				
				if (FilterInventoried.length <=0) {
					swal({
						title: 'Error Adding Filter',
						text: 'Please specify whether the filter is inventoried or not.',
						type: 'error'
					});
					return false;
				}
		
				if (FilterPickable.length <=0) {
					swal({
						title: 'Error Adding Filter',
						text: 'Please specify whether the filter is pickable or not.',
						type: 'error'
					});
					return false;
				}		
				
		
		    	$.ajax({
					type:"POST",
					url: "../../../../inc/InSightFuncs_AjaxForServiceModals.asp",
					cache: false,
					data: "action=CheckForDuplicateFilterIDExistingFilter&FilterID="+encodeURIComponent(FilterID),
					
					success: function(response)
					 {
						if (response.startsWith("We are sorry")) {				
							swal({
								title: 'Error Editing Existing Filter',
								text: response,
								type: 'error'
							})
							return false;
						} 
						
						else {
						
					    	$.ajax({
								type:"POST",
								url: "../../../../inc/InSightFuncs_AjaxForServiceModals.asp",
								cache: false,
								data: "action=CheckForDuplicateFilterUPCCodeExistingFilter&FilterUPC="+encodeURIComponent(FilterUPC)+"&FilterID="+encodeURIComponent(FilterID),
								
								success: function(response)
								 {
									if (response.startsWith("We are sorry")) {				
										swal({
											title: 'Error Editing Existing Filter',
											text: response,
											type: 'error'
										})
										return false;
									} 
									
									else {
										
										var IntRecID2 = $("#modalEditExistingFilter #txtIntRecID").val();
										var FilterID2 = $("#modalEditExistingFilter #txtFilterID").val();
										var FilterDescription2 = $("#modalEditExistingFilter #txtFilterDescription").val();
										var FilterCost2 = $("#modalEditExistingFilter #txtFilterCost").val();
										var FilterListPrice2 = $("#modalEditExistingFilter #txtFilterListPrice").val();
										var FilterTaxable2 = $("#modalEditExistingFilter #selFilterTaxable option:selected").val();
										var FilterInventoried2 = $("#modalEditExistingFilter #selFilterInventoried option:selected").val();
										var FilterPickable2 = $("#modalEditExistingFilter #selFilterPickable option:selected").val();
										var FilterUPC2 = $("#modalEditExistingFilter #txtFilterUPC").val();
										var FilterprodSKU2 = $("#modalEditExistingFilter #selFilterprodSKU").val();
										var FilterdisplayOrder2 = $("#modalEditExistingFilter #selFilterdisplayOrder").val();
										
										
								    	$.ajax({
											type:"POST",
											url: "../../../../inc/InSightFuncs_AjaxForServiceModals.asp",
											cache: false,
											data: "action=SaveEditExistingFilter&IntRecID=" + encodeURIComponent(IntRecID2) + "&FilterID=" + encodeURIComponent(FilterID2) + "&FilterDescription=" + encodeURIComponent(FilterDescription2) + "&FilterCost=" + encodeURIComponent(FilterCost2) + "&FilterListPrice=" + encodeURIComponent(FilterListPrice2) + "&FilterTaxable=" + encodeURIComponent(FilterTaxable2) + "&FilterInventoried=" + encodeURIComponent(FilterInventoried2) + "&FilterPickable=" + encodeURIComponent(FilterPickable2) + "&FilterUPC=" + encodeURIComponent(FilterUPC2) + "&FilterprodSKU=" + encodeURIComponent(FilterprodSKU2)+ "&FilterdisplayOrder=" + encodeURIComponent(FilterdisplayOrder2),

											success: function(response)
											 {
												if (response.startsWith("Error:")) {				
													swal({
														title: 'Error Saving Filter Changes',
														text: response,
														type: 'error'
													})
													return;
												} else {
													$("#frmAddFilter").submit();				
												}
											 },
											failure: function(response)
											 {
												swal({
													title: 'Error Saving Filter Changes',
													text: response,
													type: 'error'
												})
								             }
										});
									}
								 },
							});
		
						}			
					 },
				});
	    	});

	   });	
	   </script>		    
	
	<div class="modal-header">
		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		<h4 class="modal-title" id="modalAddFilterTitle"><i class="fa fa-pencil" aria-hidden="true"></i> Edit Filter <%= FilterID %></h4>
	</div>
	
	<div class="modal-body modalResponsiveTable">

     	<div class="row modalrow">
     	   	<div class="col-lg-4">Filter ID</div>
         	<div class="col-lg-8">
				<input type="text" id="txtFilterID" name="txtFilterID" class="form-control" value="<%= FilterID %>">
			</div>
		</div>

     	<div class="row modalrow">
     	   	<div class="col-lg-4">Filter Description</div>
         	<div class="col-lg-8">
				<textarea class="form-control" rows="4" id="txtFilterDescription" name="txtFilterDescription"><%= Description %></textarea>
			</div>
		</div>
		
     	<div class="row modalrow">
     	   	<div class="col-lg-6">Default Cost $&nbsp;<input type="text" id="txtFilterCost" name="txtFilterCost" class="form-control width75" value="<%= DefaultCost %>"></div>
     	   	<div class="col-lg-6">List Price $&nbsp;<input type="text" id="txtFilterListPrice" name="txtFilterListPrice" class="form-control width75" value="<%= ListPrice %>"></div>				
		</div>

     	<div class="row modalrow">
     	   	<div class="col-lg-2">Taxable?</div>
         	<div class="col-lg-2">
				<select class="form-control" id="selFilterTaxable" name="selFilterTaxable">			
					<option value="1" <% If Taxable = 1 Then Response.Write("selected='selected'")%>>Y</option>
					<option value="0" <% If Taxable = 0 Then Response.Write("selected='selected'")%>>N</option>
				</select>
			</div>	         	
     	   	<div class="col-lg-2">Inventoried?</div>
         	<div class="col-lg-2">
				<select class="form-control" id="selFilterInventoried" name="selFilterInventoried">			
					<option value="1" <% If InventoriedItem = 1 Then Response.Write("selected='selected'")%>>Y</option>
					<option value="0" <% If InventoriedItem = 0 Then Response.Write("selected='selected'")%>>N</option>
				</select>
			</div>
     	   	<div class="col-lg-2">Pickable?</div>
         	<div class="col-lg-2">
				<select class="form-control" id="selFilterPickable" name="selFilterPickable">			
					<option value="1" <% If PickableItem = 1 Then Response.Write("selected='selected'")%>>Y</option>
					<option value="0" <% If PickableItem = 0 Then Response.Write("selected='selected'")%>>N</option>
				</select>
			</div>
		</div>

     	<div class="row modalrow">
     	   	<div class="col-lg-4">UPC Code</div>
         	<div class="col-lg-8">
				<input type="text" id="txtFilterUPC" name="txtFilterUPC" class="form-control width75" value="<%= UPCCode %>">
			</div>
		</div>

     	<div class="row modalrow">
     	   	<div class="col-lg-4">Prod ID</div>
         	<div class="col-lg-8">
	     	   	<select class="form-control" id="selFilterprodSKU" name="selFilterprodSKU">		
	         	   	<option value="">NONE</option>
	         	   	<%
	         	 	Set cnnICFilters = Server.CreateObject("ADODB.Connection")
					cnnICFilters.open (Session("ClientCnnString"))
					Set rsICFilters = Server.CreateObject("ADODB.Recordset")
	
					SQLICFilters = "SELECT prodSKU, prodDescription FROM IC_Product ORDER BY prodSKU"
					Set rsICFilters = cnnICFilters.Execute(SQLICFilters)
					
					If NOT rsICFilters.EOF Then
					
						Do While NOT rsICFilters.EOF
						
							If rsICFilters("prodSKU") = prodSKU Then
								Response.Write("<option selected value='" & rsICFilters("prodSKU") & "'>" & rsICFilters("prodSKU") & " - " & rsICFilters("prodDescription") & "</option>")
							Else
								Response.Write("<option value='" & rsICFilters("prodSKU") & "'>" & rsICFilters("prodSKU") & " - " & rsICFilters("prodDescription") & "</option>")
							End If
						
							rsICFilters.Movenext
						Loop
					
					End If
					
					set rsICFilters = nothing
					cnnICFilters.close
					set cnnICFilters = nothing
	
					%>	
				</select>
			</div>
		</div>

     	<div class="row modalrow">
     	   	<div class="col-lg-4">Display Order</div>
         	<div class="col-lg-2">
	     	   	<select class="form-control" id="selFilterdisplayOrder" name="seldisplayOrder">		
	         	   	<option value="0">0</option>
	         	   	<%
						For x = 1 to 99
							If x = displayOrder Then
								Response.Write("<option selected value='" & x & "'>" & x & "</option>")
							Else
								Response.Write("<option value='" & x & "'>" & x & "</option>")
							End If
						
						Next
					%>	
				</select>
			</div>
		</div>
					
     	<div class="row" style="margin-top:20px">
     	   	<div class="col-lg-6">&nbsp;</div>
         	<div class="col-lg-6 pull-right">
				<button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
				<button type="button" class="btn btn-primary" id="btnEditFilterSave" name="btnEditFilterSave"><i class="fa fa-floppy-o" aria-hidden="true"></i>&nbsp;Save Filter Changes</button>
			</div>
		</div>

	</div>
<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub CheckForDuplicateFilterUPCCodeNewFilter()

	FilterUPC = Request.Form("FilterUPC") 
	
	FilterUPCMessage = "OK"
	SKUList = ""

	If FilterUPC <> "" Then	

		Set rsCheckForDuplicateUPCCode= Server.CreateObject("ADODB.Recordset")
		rsCheckForDuplicateUPCCode.CursorLocation = 3 	
		Set cnnCheckForDuplicateUPCCode = Server.CreateObject("ADODB.Connection")
		cnnCheckForDuplicateUPCCode.open (Session("ClientCnnString"))
		
		'*********************************************************************
		'First Check For Duplicate UPC in IC_Product
		'*********************************************************************
		
		SQLCheckForDuplicateUPCCode = "SELECT * FROM IC_Product WHERE prodUnitUPC = '" & FilterUPC & "' OR prodCaseUPC = '" & FilterUPC & "'"

		Set rsCheckForDuplicateUPCCode = cnnCheckForDuplicateUPCCode.Execute(SQLCheckForDuplicateUPCCode)
		
		If NOT rsCheckForDuplicateUPCCode.EOF Then
		
			Do While NOT rsCheckForDuplicateUPCCode.EOF
			
				prodSKU = rsCheckForDuplicateUPCCode("prodSKU")
				SKUList = SKUList & prodSKU & ","
			
			rsCheckForDuplicateUPCCode.MoveNext
			Loop
			
		End If
		
		'*********************************************************************
		'Then Check For Duplicate UPC in IC_Filters
		'*********************************************************************
		
		SQLCheckForDuplicateUPCCode = "SELECT * FROM IC_Filters WHERE UPCCode = '" & FilterUPC & "'"

		Set rsCheckForDuplicateUPCCode = cnnCheckForDuplicateUPCCode.Execute(SQLCheckForDuplicateUPCCode)
		
		If NOT rsCheckForDuplicateUPCCode.EOF Then
		
			Do While NOT rsCheckForDuplicateUPCCode.EOF
			
				prodSKU = rsCheckForDuplicateUPCCode("FilterID")
				SKUList = SKUList & prodSKU & ","
			
			rsCheckForDuplicateUPCCode.MoveNext
			Loop
			
		End If
		
				
		set rsCheckForDuplicateUPCCode = Nothing
		cnnCheckForDuplicateUPCCode.close
		set cnnCheckForDuplicateUPCCode = Nothing
		
		If SKUList <> "" Then
		
			If SKUList <> "" Then
				If Right(SKUList,1) = "," Then SKUList = Left(SKUList,Len(SKUList)-1) ' Strip trailing comma
			End If
		
			FilterUPCMessage = "We are sorry, but " & FilterUPC & " already exists as a UPC Code for the following SKUs: " & SKUList
		End If
		
	End If
	
	Response.Write(FilterUPCMessage)

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub CheckForDuplicateFilterIDNewFilter()

	FilterID = Request.Form("FilterID") 
	
	FilterSKUMessage = "OK"
	SKUList = ""
	
	If FilterID <> "" Then	

		Set rsCheckForDuplicateFilterIDNewFilter= Server.CreateObject("ADODB.Recordset")
		rsCheckForDuplicateFilterIDNewFilter.CursorLocation = 3 	
		Set cnnCheckForDuplicateFilterIDNewFilter = Server.CreateObject("ADODB.Connection")
		cnnCheckForDuplicateFilterIDNewFilter.open (Session("ClientCnnString"))
		
	
		'*********************************************************************
		'Check For Duplicate ID in IC_Filters
		'*********************************************************************
		
		SQLCheckForDuplicateFilterIDNewFilter = "SELECT * FROM IC_Filters WHERE FilterID = '" & FilterID & "'"

		Set rsCheckForDuplicateFilterIDNewFilter = cnnCheckForDuplicateFilterIDNewFilter.Execute(SQLCheckForDuplicateFilterIDNewFilter)
		
		If NOT rsCheckForDuplicateFilterIDNewFilter.EOF Then
		
			Do While NOT rsCheckForDuplicateFilterIDNewFilter.EOF
			
				prodSKU = rsCheckForDuplicateFilterIDNewFilter("FilterID")
				SKUList = SKUList & prodSKU & ","
			
			rsCheckForDuplicateFilterIDNewFilter.MoveNext
			Loop
			
		End If
		
				
		set rsCheckForDuplicateFilterIDNewFilter = Nothing
		cnnCheckForDuplicateFilterIDNewFilter.close
		set cnnCheckForDuplicateFilterIDNewFilter = Nothing
		
		If SKUList <> "" Then		
			FilterSKUMessage = "We are sorry, but " & FilterID & " already exists as a Filter ID."
		End If
		
	End If
	
	Response.Write(FilterSKUMessage)

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub SaveEditExistingFilter() 

	FilterID = Request.Form("FilterID")
	FilterDescription = Request.Form("FilterDescription")
	FilterCost = Request.Form("FilterCost")
	FilterListPrice = Request.Form("FilterListPrice")
	FilterTaxable = Request.Form("FilterTaxable")
	FilterInventoried = Request.Form("FilterInventoried")
	FilterPickable = Request.Form("FilterPickable")
	FilterUPC = Request.Form("FilterUPC")
	FilterRecordSource = "INSIGHT"
	FilterprodSKU = Request.Form("FilterprodSKU")
	FilterdisplayOrder = Request.Form("FilterdisplayOrder")
			
	Set cnnSaveEditExistingFilter = Server.CreateObject("ADODB.Connection")
	cnnSaveEditExistingFilter.open (Session("ClientCnnString"))
	Set rsSaveEditExistingFilter = Server.CreateObject("ADODB.Recordset")
	rsSaveEditExistingFilter.CursorLocation = 3 

	
	'**********************************************************************
	'Lookup the record as it exists now so we can fill in the audit trail
	'**********************************************************************
	
	SQL = "SELECT * FROM IC_FILTERS where FilterID = '" & FilterID & "'"
		
	Set rsSaveEditExistingFilter = cnnSaveEditExistingFilter.Execute(SQL)
		
	If not rsSaveEditExistingFilter.EOF Then
		IntRecID = rsSaveEditExistingFilter("InternalRecordIdentifier")
		ORIG_RecordSource = rsSaveEditExistingFilter("RecordSource")
		ORIG_FilterID = rsSaveEditExistingFilter("FilterID")
		ORIG_Description = rsSaveEditExistingFilter("Description")
		ORIG_ListPrice = rsSaveEditExistingFilter("ListPrice")
		ORIG_DefaultCost = rsSaveEditExistingFilter("DefaultCost")
		ORIG_Taxable = rsSaveEditExistingFilter("Taxable")
		ORIG_InventoriedItem = rsSaveEditExistingFilter("InventoriedItem")
		ORIG_PickableItem = rsSaveEditExistingFilter("PickableItem")
		ORIG_UPCCode = rsSaveEditExistingFilter("UPCCode")	
		ORIG_prodSKU = rsSaveEditExistingFilter("prodSKU")	
		ORIG_displayOrder = rsSaveEditExistingFilter("displayOrder")		
	End If
	
	'**********************************************************************
	'End Lookup the record as it exists now so we can fill in the audit trail
	'**********************************************************************
	
	
	
	'**********************************************************************
	'Now Update IC_Filters with edited values from modal
	'**********************************************************************
	SQL = "UPDATE IC_FILTERS SET "
	SQL = SQL &  "RecordSource = 'INSIGHT' "
	SQL = SQL &  ", FilterID = '" & FilterID & "' "
	SQL = SQL &  ", Description = '" & FilterDescription & "' "
	SQL = SQL &  ", DefaultCost = '" & FilterCost & "' "
	SQL = SQL &  ", ListPrice = '" & FilterListPrice & "' "
	SQL = SQL &  ", InventoriedItem = '" & FilterInventoried & "' "
	SQL = SQL &  ", PickableItem = '" & FilterPickable & "' "
	SQL = SQL &  ", Taxable = '" & FilterTaxable & "' "
	SQL = SQL &  ", UPCCode = '" & FilterUPC & "' "
	SQL = SQL &  ", prodSKU = '" & FilterprodSKU & "' "
	SQL = SQL &  ", displayOrder = " & FilterdisplayOrder 
	SQL = SQL &  " WHERE InternalRecordIdentifier = " & IntRecID
	
	
	Set rsSaveEditExistingFilter = cnnSaveEditExistingFilter.Execute(SQL)
	
	set rsSaveEditExistingFilter = Nothing
	
	
	'**********************************************************************
	'Create Audit Log Entries For Any Specific Changes Made
	'**********************************************************************
	
	Description = ""

	If Orig_RecordSource<> FilterRecordSource Then
		Description = GetTerm("Service") & " Module Filter Record Source changed from " & Orig_RecordSource & " to " & FilterRecordSource 
		CreateAuditLogEntry GetTerm("Service") & " Module Filter Edited",GetTerm("Service") & " Module Filter Edited","Minor",0,Description
	End If

	If Orig_FilterID <> FilterID Then
		Description = GetTerm("Service") & " Module Filter ID changed from " & Orig_FilterID & " to " & FilterID 
		CreateAuditLogEntry GetTerm("Service") & " Module Filter Edited",GetTerm("Service") & " Module Filter Edited","Minor",0,Description
	End If

	If Orig_Description <> FilterDescription Then
		Description = GetTerm("Service") & " Module Filter Description changed from " & Orig_Description & " to " & FilterDescription 
		CreateAuditLogEntry GetTerm("Service") & " Module Filter Edited",GetTerm("Service") & " Module Filter Edited","Minor",0,Description
	End If

	If Orig_ListPrice <> FilterListPrice Then
		Description = GetTerm("Service") & " Module Filter List Price changed from " & Orig_ListPrice & " to " & FilterListPrice 
		CreateAuditLogEntry GetTerm("Service") & " Module Filter Edited",GetTerm("Service") & " Module Filter Edited","Minor",0,Description
	End If

	If Orig_DefaultCost <> FilterCost Then
		Description = GetTerm("Service") & " Module Filter Default Cost changed from " & Orig_DefaultCost & " to " & FilterCost 
		CreateAuditLogEntry GetTerm("Service") & " Module Filter Edited",GetTerm("Service") & " Module Filter Edited","Minor",0,Description
	End If

	If cInt(FilterTaxable) = 1 then FilterTaxableONOFFMsg = "YES" Else FilterTaxableONOFFMsg = "NO"
	If cInt(Orig_Taxable) = 1 then OrigTaxableONOFFMsg = "YES" Else OrigTaxableONOFFMsg = "NO"
	If Orig_Taxable <> FilterTaxable Then
		Description = GetTerm("Service") & " Module Filter Taxable Item Flag changed from " & OrigTaxableONOFFMsg & " to " & FilterTaxableONOFFMsg
		CreateAuditLogEntry GetTerm("Service") & " Module Filter Edited",GetTerm("Service") & " Module Filter Edited","Minor",0,Description
	End If

	If cInt(FilterInventoried) = 1 then FilterInventoriedONOFFMsg = "YES" Else FilterInventoriedONOFFMsg = "NO"
	If cInt(Orig_InventoriedItem) = 1 then OrigInventoriedItemONOFFMsg = "YES" Else OrigInventoriedItemONOFFMsg = "NO"
	If Orig_InventoriedItem <> FilterInventoried Then
		Description = GetTerm("Service") & " Module Filter Inventoried Item Flag changed from " & OrigInventoriedItemONOFFMsg & " to " & FilterInventoriedONOFFMsg 
		CreateAuditLogEntry GetTerm("Service") & " Module Filter Edited",GetTerm("Service") & " Module Filter Edited","Minor",0,Description
	End If

	If cInt(FilterPickable) = 1 then FilterPickableONOFFMsg = "YES" Else FilterPickableONOFFMsg = "NO"
	If cInt(Orig_PickableItem) = 1 then OrigPickableItemONOFFMsg = "YES" Else OrigPickableItemONOFFMsg = "NO"
	If Orig_PickableItem <> FilterPickable Then
		Description = GetTerm("Service") & " Module Filter Pickable Item Flag changed from " & Orig_OrigPickableItemONOFFMsg & " to " & FilterPickableONOFFMsg 
		CreateAuditLogEntry GetTerm("Service") & " Module Filter Edited",GetTerm("Service") & " Module Filter Edited","Minor",0,Description
	End If
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub CheckForDuplicateFilterUPCCodeExistingFilter()

	FilterUPC = Request.Form("FilterUPC") 
	FilterID = Request.Form("FilterID")
	
	FilterUPCMessage = "OK"
	SKUList = ""

	If FilterUPC <> "" AND FilterID <> "" Then	

		Set rsCheckForDuplicateUPCCode= Server.CreateObject("ADODB.Recordset")
		rsCheckForDuplicateUPCCode.CursorLocation = 3 	
		Set cnnCheckForDuplicateUPCCode = Server.CreateObject("ADODB.Connection")
		cnnCheckForDuplicateUPCCode.open (Session("ClientCnnString"))
		
		'*********************************************************************
		'First Check For Duplicate UPC in IC_Product
		'*********************************************************************
		
		SQLCheckForDuplicateUPCCode = "SELECT * FROM IC_Product WHERE prodUnitUPC = '" & FilterUPC & "' OR prodCaseUPC = '" & FilterUPC & "'"

		Set rsCheckForDuplicateUPCCode = cnnCheckForDuplicateUPCCode.Execute(SQLCheckForDuplicateUPCCode)
		
		If NOT rsCheckForDuplicateUPCCode.EOF Then
		
			Do While NOT rsCheckForDuplicateUPCCode.EOF
			
				prodSKU = rsCheckForDuplicateUPCCode("prodSKU")
				SKUList = SKUList & prodSKU & ","
			
			rsCheckForDuplicateUPCCode.MoveNext
			Loop
			
		End If
		
		'*********************************************************************
		'Then Check For Duplicate UPC in IC_Filters
		'*********************************************************************
		
		SQLCheckForDuplicateUPCCode = "SELECT * FROM IC_Filters WHERE UPCCode = '" & FilterUPC & "' AND FilterID <> '" & FilterID & "'"

		Set rsCheckForDuplicateUPCCode = cnnCheckForDuplicateUPCCode.Execute(SQLCheckForDuplicateUPCCode)
		
		If NOT rsCheckForDuplicateUPCCode.EOF Then
		
			Do While NOT rsCheckForDuplicateUPCCode.EOF
			
				prodSKU = rsCheckForDuplicateUPCCode("FilterID")
				SKUList = SKUList & prodSKU & ","
			
			rsCheckForDuplicateUPCCode.MoveNext
			Loop
			
		End If
		
				
		set rsCheckForDuplicateUPCCode = Nothing
		cnnCheckForDuplicateUPCCode.close
		set cnnCheckForDuplicateUPCCode = Nothing
		
		If SKUList <> "" Then
		
			If SKUList <> "" Then
				If Right(SKUList,1) = "," Then SKUList = Left(SKUList,Len(SKUList)-1) ' Strip trailing comma
			End If
		
			FilterUPCMessage = "We are sorry, but " & FilterUPC & " already exists as a UPC Code for the following SKUs: " & SKUList
		End If
		
	End If
	
	Response.Write(FilterUPCMessage)

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub CheckForDuplicateFilterIDExistingFilter()

	FilterID = Request.Form("FilterID") 
	
	FilterSKUMessage = "OK"
	SKUList = ""
	
	If FilterID <> "" Then	

		Set rsCheckForDuplicateFilterIDNewFilter= Server.CreateObject("ADODB.Recordset")
		rsCheckForDuplicateFilterIDNewFilter.CursorLocation = 3 	
		Set cnnCheckForDuplicateFilterIDNewFilter = Server.CreateObject("ADODB.Connection")
		cnnCheckForDuplicateFilterIDNewFilter.open (Session("ClientCnnString"))
		
		'*********************************************************************
		'Check For Duplicate ID in IC_Filters
		'*********************************************************************
		
		SQLCheckForDuplicateFilterIDNewFilter = "SELECT COUNT(*) as filterCount FROM IC_Filters WHERE FilterID = '" & FilterID & "'"

		Set rsCheckForDuplicateFilterIDNewFilter = cnnCheckForDuplicateFilterIDNewFilter.Execute(SQLCheckForDuplicateFilterIDNewFilter)
		
		skuCount = 0
		
		If NOT rsCheckForDuplicateFilterIDNewFilter.EOF Then
		
			skuCount = rsCheckForDuplicateFilterIDNewFilter("filterCount")
			
		End If
			
		set rsCheckForDuplicateFilterIDNewFilter = Nothing
		cnnCheckForDuplicateFilterIDNewFilter.close
		set cnnCheckForDuplicateFilterIDNewFilter = Nothing
		
		If SKUList <> "" OR skuCount > 1 Then		
			FilterSKUMessage = "We are sorry, but " & FilterID & " already exists as a Filter ID."
		End If
		
	End If
	
	Response.Write(FilterSKUMessage)

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub GetContentForDeleteFilterModal()

	IntRecID = Request.Form("IntRecID")
	
	%>
	<input type="hidden" id="txtIntRecID" name="txtIntRecID" value="<%= IntRecID %>">
	

	<style type="text/css">
	.col-lg-12{
		margin-bottom:20px;
	}
	
	.modal-footer{
		margin-top:15px;
	}
	</style>
	
		
	<div class="modal-header">
		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		<h4 class="modal-title" id="modalDeleteFilterTitle"><i class="fas fa-trash-alt" aria-hidden="true"></i> Replace Filter Before Deletion</h4>
	</div>
	
	<div class="modal-body modalResponsiveTable">
	
		
		<form method="post" action="deleteFilterFromModal.asp" name="frmDeleteFilterFromModal" id="frmDeleteFilterFromModal">
		
			<input type="hidden" name="txtFilterIntRecIDToReplace" id="txtFilterIntRecIDToReplace" value="<%=IntRecID %>">
		
			<div class="row modalrow">
				<p style="margin-left:20px;margin-right:20px;">There are <%= NumberCustomerRecsDefinedForFilterID(IntRecID)%> customers assigned to the filter you are trying to delete. Before this filter can be deleted you must choose a new filter to be assigned to these customers from the list below.</p>
			</div>
		
			<div class="row modalrow">
				<div class="form-group">
					<label class="col-sm-4 control-label">Replace Filter with:</label>
					<div class="col-sm-8">
					  	<select class="form-control" name='selDeleteFilterIntRecIDFromModal' id='selDeleteFilterIntRecIDFromModal'>
						      	<% 'Get all Filters
						      	  	SQL9 = "SELECT * FROM IC_Filters WHERE InternalRecordIdentifier <> " & IntRecID & " ORDER BY FilterID"  ' Select all but the one to delete
		
									Set cnn9 = Server.CreateObject("ADODB.Connection")
									cnn9.open (Session("ClientCnnString"))
									Set rs9 = Server.CreateObject("ADODB.Recordset")
									rs9.CursorLocation = 3 
									Set rs9 = cnn9.Execute(SQL9)
									If not rs9.EOF Then
										Do
											Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("FilterID") & " - " & rs9("Description") & "</option>")
											rs9.movenext
										Loop until rs9.eof
									End If
									set rs9 = Nothing
									cnn9.close
									set cnn9 = Nothing
								%>
						</select>
					</div>
				</div>
			</div>
		
			
	     	<div class="row" style="margin-top:20px">
	     	   	<div class="col-lg-4">&nbsp;</div>
	         	<div class="col-lg-8 pull-right">
					<button type="button" class="btn btn-default" data-dismiss="modal">Cancel Deletion</button>
					<button type="submit" class="btn btn-primary" id="btnDeleteFilterSave" name="btnDeleteFilterSave"><i class="fa fa-repeat" aria-hidden="true"></i>&nbsp;Replace Filter &amp; Delete</button>
				</div>
			</div>
			
		</form>
	</div>
	<%
End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForCreateServiceTicketModal()

	SelectedCustomer = Request.Form("custID")

	%>
	<script type="text/javascript">
	
		$(document).ready(function(){
				
	        $('#datetimepickerWhenProblemStarted').datetimepicker({
	        	useCurrent: true,
	        	maxDate:moment(),
                format: 'MM/DD/YYYY',
                ignoreReadonly: true,
                sideBySide: true,
	
			}); 			
					
		});
		
	</script>
				
		<!-- row !-->		
		<div class="row">
			<div class="col-lg-12">
				<div class="row">
					<!--account number !-->
					<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
						<!--#include file="commonCustomerDisplay.asp"-->				        
					</div>					
					<!-- eof account number !-->
				</div>
			</div>
			<!-- eof row !-->
		</div>


		 <div class="row well">
 
			 <!-- left col !-->
			 <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
			 
 		        <!-- row !-->			
			    <div class="row">

					<!-- Contact Name !-->
					<div class="col-lg-4">
						<strong>Equipment Type*</strong>
						<select class="form-control" name="txtEquipmentType" id="txtEquipmentType">
							<option value="">Please Select Below</option>
							<option value="none listed">My Equipment Is Not Listed</option>
						</select>
					</div>
					<!-- Contact Name !-->
					
					<!-- Contact Phone !-->
					<div class="col-lg-4">
						<strong>Symptom*</strong>
						<select class="form-control" name="txtEquipmentSymptom" id="txtEquipmentSymptom">
							<!-- options generated via on the fly include file -->
						</select>
					</div>
					<!-- Contact Phone !-->

					<!-- Contact Phone !-->
					<div class="col-lg-4">
						<strong>When Problem Started*</strong>
						<div class="col-lg-12">								  	
			                <div class="input-group date" id="datetimepickerWhenProblemStarted">
			                    <input type="text" class="form-control" name="txtDateProblemStarted" id="txtDateProblemStarted">
			                    <span class="input-group-addon">
			                        <span class="glyphicon glyphicon-calendar"></span>
			                    </span>
			                </div>
			             </div>						
					</div>
					<!-- Contact Phone !-->
		    	
				</div>
				<!-- eof row !-->
			 
 
 		        <!-- row !-->			
			    <div class="row">

					<!-- Contact Name !-->
					<div class="col-lg-8">
						<strong>Who Should We See Upon Arrival?*</strong>
						<input type="text" id="txtWhoToContactUponArrival" name="txtWhoToContactUponArrival" class="form-control" value="<%= tmpContact %>">
					</div>
					<!-- Contact Name !-->
					
					<!-- Contact Phone !-->
					<div class="col-lg-4">
						<strong>Contact Phone #*</strong>
						<input type="text" id="txtContactPhone" name="txtContactPhone" class="form-control" value="<%= tmpPhone %>">
					</div>
					<!-- Contact Phone !-->
		    	
				</div>
				<!-- eof row !-->
		    	
				<!-- row !-->			
				<div class="row">
					
					<!-- Contact Phone !-->
					<div class="col-lg-6">
						<strong>Contact Email*</strong>
						<input type="text" id="txtContactEmail" name="txtContactEmail" class="form-control">
					</div>
					<!-- Contact Phone !-->


					<!-- Contact Phone !-->
					<div class="col-lg-6">
						<strong>Problem Location*</strong> <small>(Please include floor # if applicable)</small>
						<input type="text" id="txtFloorSuite" name="txtFloorSuite" class="form-control">
					</div>
					<!-- Contact Phone !-->

				<!-- end row !-->		
				</div>
						
				<!-- right col !-->
				<div class="row">
					<div class="col-lg-8 col-md-8 col-sm-12 col-xs-12">	  
						<!-- Description of problem !-->
						<strong>Please enter a description as completely as possible.</strong>
						<textarea name="txtDescription" id="txtDescription" rows="5" spellcheck="True" class="form-control"></textarea>
						<!-- Description of problem !-->
					</div>
				</div>
				<!-- eof right col !-->
				
			
			</div>
			
		</div>
		<!-- eof main row !-->

	


	<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForEditServiceTicketModalTitle()

	MemoNumber = Request.Form("memo")
	SelectedCustomer = Request.Form("custID")

	%>
	<h4 class="pull-left"><i class="fa fa-wrench"></i> Close or Cancel Service Ticket</h4>
    <div class="alert alert-info" role="alert" style="margin-left:20px;">
    	<strong>Ticket #: <%= MemoNumber %></strong>
    </div>
    <div class="alert alert-warning" role="alert"> 
		<strong>Stage: <%= GetServiceTicketCurrentStage(MemoNumber)%></strong>	        
    </div>
    
	<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForEditServiceTicketModal()

	MemoNumber = Request.Form("memo")
	SelectedCustomer = Request.Form("custID")
	
	SQL = "SELECT * FROM FS_ServiceMemos where MemoNumber = '" & MemoNumber & "'"
		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		ServiceMemoRecNumber = rs("ServiceMemoRecNumber")
		CurrentStatus = rs("CurrentStatus")
		RecordSubType = rs("RecordSubType")
		SubmittedByName = rs("SubmittedByName")
		AccountNumber = rs("AccountNumber")
		Company = rs("Company")
		ProblemLocation = rs("ProblemLocation")
		SubmittedByPhone = rs("SubmittedByPhone")
		SubmittedByEmail = rs("SubmittedByEmail")
		SubmissionDateTime = rs("SubmissionDateTime")
		ProblemDescription = rs("ProblemDescription")
		Mode = rs("Mode")
		SubmissionSource = rs("SubmissionSource")
		UserNoOfServiceTech = rs("UserNoOfServiceTech")
		ReleasedDateTime = rs("ReleasedDateTime")
		ReleasedByUserNo = rs("ReleasedByUserNo")
		ReleasedNotes = rs("ReleasedNotes")
	End If
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	If SubmittedByName = "" Then SubmittedByName = "Not provided"
	If SubmittedByPhone = "" Then SubmittedByPhone = "Not provided"
	If SubmittedByEmail = "" Then SubmittedByEmail = "Not provided"
	If ProblemLocation = "" Then ProblemLocation = "Not provided"
	If ProblemDescription = "" Then ProblemDescription = "Not provided"

	%>
	
	<script type="text/javascript">
	
		//common function to populate selectboxes
		function PopulateSelectBoxes(selectid){
		    $.ajax({
		        type: "POST",
		        url: 'onthefly_SelectBoxes.asp',
		        data: ({ section : selectid, action:'add' }),
		        dataType: "html",
		        success: function(data) {
		            $("#"+selectid).html(data);
		        },
		        error: function() {
		            alert('Error occured');
		        }
		    });	
		}
	
		$(document).ready(function(){
				
	        $('#datetimepickerCloseCancelDate').datetimepicker({
	        	useCurrent: true,
                format: 'MM/DD/YYYY',
                ignoreReadonly: true,
                sideBySide: true,
			}); 
			
			$('#timepicker').timepicker();	
			
			$('#serviceTicketNotesTable').DataTable();	

			$('input[type=radio][name=optCloseOrCancel]').on('change', function() {
			  switch ($(this).val()) {
			    case 'Cancel':
			      $("#problemResolutionRow").hide();
			      break;
			    case 'Close':
			      $("#problemResolutionRow").show();
			      break;
			  }
			});
		
	
			// below code added by nurba
			// 03/15/2019					 
			 PopulateSelectBoxes('txtEquipmentProblem');
			 PopulateSelectBoxes('txtEquipmentResolution');	  
							 
		 	//-------------------------------------------------------------------------------
			// Equipment Problem select box change
		    $("#txtEquipmentProblem").change(function() {
				var val = $( "#txtEquipmentProblem option:selected").val();
				if (val== -1){
					//deselect add new row
					$('#txtEquipmentProblem option[selected="selected"]').each(
						function() {
							$(this).removeAttr('selected');
						}
					);
		
					// mark the first option as selected
					$("#txtEquipmentProblem option:first").attr('selected','selected');
					
					//show modal
					//$('#myProspectingModalEditBusinessCard').modal('hide');
					$('#ONTHEFLYmodalProblemCode').modal('show');
					
				}
			});
			
			
			//Industry modal window submit
			$('#frmAddProblemCode').submit(function(e) {
				
				if ($('#frmAddProblemCode #txtProblemDescription').val()==''){
					 swal("Problem description cannot be blank.");
					return false;
				}
				
				$("#ONTHEFLYmodalProblemCode .btn-primary").html("Saving...");
		        $.ajax({
		            type: "POST",
		            url: "onthefly_ProblemCodes_submit.asp",
		            data: $('#frmAddProblemCode').serialize(),
		            success: function(response) {
						PopulateSelectBoxes('txtEquipmentProblem');
						$("#ONTHEFLYmodalProblemCode .modal-body").html('Problem added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
						
		            },
		            error: function() {
						$("#ONTHEFLYmodalProblemCode .btn-primary").html("Save");
		                //alert('Error adding Problem');
		            }
		        });
		        return false;
		    });
			//-------------------------------------------------------------------------------
			//end nurba
		
		 	//-------------------------------------------------------------------------------
			// Equipment Resolution select box change
		    $("#txtEquipmentResolution").change(function() {
				var val = $( "#txtEquipmentResolution option:selected").val();
				if (val== -1){
					//deselect add new row
					$('#txtEquipmentResolution option[selected="selected"]').each(
						function() {
							$(this).removeAttr('selected');
						}
					);
		
					// mark the first option as selected
					$("#txtEquipmentResolution option:first").attr('selected','selected');
					
					//show modal
					//$('#myProspectingModalEditBusinessCard').modal('hide');
					$('#ONTHEFLYmodalResolutionCode').modal('show');
					
				}
			});
			
			
			//Industry modal window submit
			$('#frmAddResolutionCode').submit(function(e) {
				
				if ($('#frmAddResolutionCode #txtResolutionDescription').val()==''){
					 swal("Resolution description cannot be blank.");
					return false;
				}
				
				$("#ONTHEFLYmodalResolutionCode .btn-primary").html("Saving...");
		        $.ajax({
		            type: "POST",
		            url: "onthefly_ResolutionCodes_submit.asp",
		            data: $('#frmAddResolutionCode').serialize(),
		            success: function(response) {
						PopulateSelectBoxes('txtEquipmentResolution');
						$("#ONTHEFLYmodalResolutionCode .modal-body").html('Resolution added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
						
		            },
		            error: function() {
						$("#ONTHEFLYmodalResolutionCode .btn-primary").html("Save");
		                //alert('Error adding Resolution');
		            }
		        });
		        return false;
		    });
			//-------------------------------------------------------------------------------
			//end nurba
				
		
		});
		
	</script>
				
		<!-- row !-->		
		<div class="row">
			<div class="col-lg-12">
				<!--#include file="commonCustomerDisplayCloseCancelTicket.asp"-->				        
			</div><!-- eof col !-->
		</div><!-- eof row !-->       
	 		        
		<input type="hidden" id="txtMemoNumberCloseCancel" name="txtMemoNumberCloseCancel" value="<%= MemoNumber %>">
	
		 <div class="row well">
	
			 <!-- left col !-->
			 <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
			 
 		        <!-- row !-->			
			    <div class="row">
			    	<!-- Close or Cancel !-->
					<div class="col-lg-5">
						<% If userIsServiceManager(Session("userNo")) or userIsAdmin(Session("userNo")) Then %> 
							<label><input type="radio" name="optCloseOrCancel" id="optClose" value="Close" checked>  Close</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<% 'We need to see if advanced dispatching is on & if so, which stage are we in
								cStage = GetServiceTicketCurrentStage(MemoNumber)
								'Can only be cancelled in the following stages
								If cStage = "Received" or cStage = "Released" or cStage = "Under Review" or cStage = "Dispatched" or cStage = "Dispatch Acknowledged" or cStage = "Dispatch Declined" or cStage = "En Route" Then%>
									<label><input type="radio" name="optCloseOrCancel" id="optCancel" value="Cancel"> Cancel</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<% End If %>
							<br><input type='checkbox' class='check' id='chkDoNotEmail' name='chkDoNotEmail'>&nbsp;<strong class="do-not-send-alert">Do not send a close email to the customer</strong>
						<% Else %>
							<label><input type="radio" name="optCloseOrCancel" id="optCancel" value="Cancel" checked readonly> Cancel</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<br><input type='checkbox' class='check' id='chkDoNotEmail' name='chkDoNotEmail'>&nbsp;<strong class="do-not-send-alert">Do not send a close email to the customer</strong>
						<% End If %>
			        </div>
 			    	<!-- Close or Cancel !-->
 			    	
					<div class="col-lg-4">
						<strong>Close/Cancel Date*</strong>
						<div class="col-lg-12">								  	
			                <div class="input-group date" id="datetimepickerCloseCancelDate">
			                    <input type="text" class="form-control" name="txtCloseCancelDate" id="txtCloseCancelDate" value="<%= Date() %>">
			                    <span class="input-group-addon">
			                        <span class="glyphicon glyphicon-calendar"></span>
			                    </span>
			                </div>
			             </div>						
					</div>
					
					<div class="col-lg-3">
						<strong>Close/Cancel Time*</strong>
						<input type="text" id="timepicker" name="timepicker" class="form-control" value="<%=Hour(Time()) & ":" & Minute(Time()) %>">
					</div>
		    	
				</div>
				<!-- eof row !-->
			 
			 
			 	<% If userIsServiceManager(Session("userNo")) or userIsAdmin(Session("userNo")) Then %>
	 		        <!-- row !-->	
	 		        <!--		
					<div class="row">
						<div class="col-lg-12">
							<strong>Asset Location</strong><br>
							<small>To update the location of an asset, fill in the info below.</small>
						</div>
					</div>
					
					<div class="row">
						<div class="col-lg-6">
							<strong>Asset Tag #</strong>
							<input type="text" id="txtAssetTagNumber" name="txtAssetTagNumber" class="form-control">
						</div>
						
						<div class="col-lg-6">
							<strong>Location</strong>
							<input type="text" id="txtAssetLocation" name="txtAssetLocation" class="form-control">
						</div>
			    	
					</div>-->
					<!-- eof row !-->
				<% End If %>
			 
 		        <!-- row !-->			
				<div class="row" id="problemResolutionRow">
					<div class="col-lg-6">
						<strong>Problem*</strong>
						<select class="form-control" name="txtEquipmentProblem" id="txtEquipmentProblem">
							<!-- options generated via on the fly include file -->
						</select>
					</div>
					
					<div class="col-lg-6">
						<strong>Resolution*</strong>
						<select class="form-control" name="txtEquipmentResolution" id="txtEquipmentResolution">
							<!-- options generated via on the fly include file -->
						</select>
					</div>
				</div>
				<!-- eof row !-->

		
		    	<% If userIsServiceManager(Session("userNo")) or userIsAdmin(Session("userNo")) Then %> 
		    	
	 		        <!-- row !-->			
				    <div class="row">
			    	
				    	<div class="col-lg-6 col-md-6">	
							<label>Field Tech*</label>
							<select name="selFieldTech" id="selFieldTech" class="form-control">
							<option value="">Select Field Tech</option>
								<%	
	
								SQL = "SELECT * FROM tblUsers WHERE userArchived <> 1 Order By userLastName"
								
								Set cnn8 = Server.CreateObject("ADODB.Connection")
								cnn8.open (Session("ClientCnnString"))
								Set rs = Server.CreateObject("ADODB.Recordset")
								rs.CursorLocation = 3 
								Set rs = cnn8.Execute(SQL)
		
								If not rs.EOF Then
		
									Do While Not rs.EOF
										userFirstName = rs("userFirstName")
										userLastName = rs("userLastName")
										userDisplayName = rs("userDisplayName")
										userEmail = rs("userEmail")
										userNo = rs("UserNo")
										
										%><option value='<%=userNo%>'><%=userFirstName%>&nbsp;<%=userLastName%>&nbsp;---<%=userDisplayName%>&nbsp;---<%=userEmail%></option><%
										
										rs.MoveNext
									Loop
		
								End If
								%>
							</select>
						 </div>
						 
						<div class="col-lg-6 col-md-6">	  
							<!-- Description of problem !-->
							<strong>Service Notes*</strong>
							<textarea name="ServiceNotes" id="ServiceNotes" rows="5" spellcheck="True" class="form-control"></textarea>
							<!-- Description of problem !-->
						</div>
					 
					</div>
				 <% End If %>

			</div>
			
		</div>
		<!-- eof main row !-->

		<% MDG_MemoNumber = MemoNumber %>
			
		<div class="row well white">
			<div class="panel-group" id="accordion" role="tablist" aria-multiselectable="true">
			  <div class="panel panel-default">
			    <div class="panel-heading" role="tab" id="headingOne">
			      <h4 class="panel-title">
			        <a role="button" data-toggle="collapse" data-parent="#accordion" href="#collapseFieldServiceNotes" aria-expanded="true" aria-controls="collapseOne">
			          Previous Ticket Notes
			          <span> </span>
			        </a>
			      </h4>
			    </div>
			    <div id="collapseFieldServiceNotes" class="panel-collapse collapse" role="tabpanel" aria-labelledby="headingOne">
			      <div class="panel-body">
					<table id="serviceTicketNotesTable" class="display" style="width:100%;font-size:11px;">
					        <thead>
					            <tr>
					                <th>Date/Time</th>
					                <th>Status/Stage</th>
					                <th>User</th>
					                <th>Notes</th>
					            </tr>
					        </thead>
					        <tbody>
					            <!--#include file="serviceTicketNotesDetailsTable.asp"-->
					        </tbody>
					        <tfoot>
					            <tr>
					                <th width="20%">Date/Time</th>
					                <th width="20%">Status/Stage</th>
					                <th width="15%">User</th>
					                <th width="45%">Notes</th>
					            </tr>
					        </tfoot>
					    </table>	     
			      </div>
			    </div>
			  </div>
			</div>	
		</div>
		<!-- eof main row !-->


	<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForViewOpenClosedServiceTicketModal()

	MemoNumber = Request.Form("memo")
	SelectedCustomer = Request.Form("custID")
	
	'*************
	'OPEN info
	SQL = "SELECT * FROM FS_ServiceMemos where MemoNumber = '" & MemoNumber  & "' And RecordSubType='OPEN'"
		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		OpenServiceMemoRecNumber = rs("ServiceMemoRecNumber")
		OpenCurrentStatus = rs("CurrentStatus")
		OpenRecordSubType = rs("RecordSubType")
		OpenSubmittedByName = rs("SubmittedByName")
		OpenAccountNumber = rs("AccountNumber")
		OpenCompany = rs("Company")
		OpenProblemLocation = rs("ProblemLocation")
		OpenSubmittedByPhone = rs("SubmittedByPhone")
		OpenSubmittedByEmail = rs("SubmittedByEmail")
		OpenSubmissionDateTime = rs("SubmissionDateTime")
		OpenProblemDescription = rs("ProblemDescription")
		OpenMode = rs("Mode")
		OpenSubmissionSource = rs("SubmissionSource")
		OpenUserNoOfServiceTech = rs("UserNoOfServiceTech")
		ReleasedDateTime = rs("ReleasedDateTime")
		ReleasedByUserNo = rs("ReleasedByUserNo")
		ReleasedNotes = rs("ReleasedNotes")
	
	Else
		SQL = "SELECT * FROM FS_ServiceMemos where MemoNumber = '" & MemoNumber  & "' And RecordSubType='HOLD'"
		Set rs = cnn8.Execute(SQL)
		If not rs.EOF Then
			OpenServiceMemoRecNumber = rs("ServiceMemoRecNumber")
			OpenCurrentStatus = rs("CurrentStatus")
			OpenRecordSubType = rs("RecordSubType")
			OpenSubmittedByName = rs("SubmittedByName")
			OpenAccountNumber = rs("AccountNumber")
			OpenCompany = rs("Company")
			OpenProblemLocation = rs("ProblemLocation")
			OpenSubmittedByPhone = rs("SubmittedByPhone")
			OpenSubmittedByEmail = rs("SubmittedByEmail")
			OpenSubmissionDateTime = rs("SubmissionDateTime")
			OpenProblemDescription = rs("ProblemDescription")
			OpenMode = rs("Mode")
			OpenSubmissionSource = rs("SubmissionSource")
			OpenUserNoOfServiceTech = rs("UserNoOfServiceTech")
		End IF
	End If
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	If OpenSubmittedByName = "" Then OpenSubmittedByName = "Not provided"
	If OpenSubmittedByPhone = "" Then OpenSubmittedByPhone = "Not provided"
	If OpenSubmittedByEmail = "" Then OpenSubmittedByEmail = "Not provided"
	If OpenProblemLocation = "" Then OpenProblemLocation = "Not provided"
	If OpenProblemDescription = "" Then OpenProblemDescription = "Not provided"
	

	'*************
	'CloseCancel info
	SQL = "SELECT * FROM FS_ServiceMemos where MemoNumber = '" & MemoNumber  & "' And RecordSubType<>'OPEN'"
		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		CCServiceMemoRecNumber = rs("ServiceMemoRecNumber")
		CCCurrentStatus = rs("CurrentStatus")
		CCRecordSubType = rs("RecordSubType")
		CCSubmittedByName = rs("SubmittedByName")
		CCAccountNumber = rs("AccountNumber")
		CCCompany = rs("Company")
		CCProblemLocation = rs("ProblemLocation")
		CCSubmittedByPhone = rs("SubmittedByPhone")
		CCSubmittedByEmail = rs("SubmittedByEmail")
		CCSubmissionDateTime = rs("SubmissionDateTime")
		CCProblemDescription = rs("ProblemDescription")
		CCMode = rs("Mode")
		CCSubmissionSource = rs("SubmissionSource")
		CCUserNoOfServiceTech = rs("UserNoOfServiceTech")
	End If
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	If CCSubmittedByName = "" Then CCSubmittedByName = "Not provided"
	If CCSubmittedByPhone = "" Then CCSubmittedByPhone = "Not provided"
	If CCSubmittedByEmail = "" Then CCSubmittedByEmail = "Not provided"
	If CCProblemLocation = "" Then CCProblemLocation = "Not provided"
	If CCProblemDescription = "" Then CCProblemDescription = "Not provided"
	
	%>
	
	<script type="text/javascript">
	
		$(document).ready(function(){
							
			$('#serviceTicketNotesTable').DataTable();		
			$('.fancybox').fancybox();
			$('.fancybox-signature').fancybox();
				
		});
		
	</script>
				
		<!-- row !-->		
		<div class="row">
			<div class="col-lg-12">
			    <% SelectedCustomer = OpenAccountNumber %>
				<!--#include file="commonCustomerDisplay.asp"-->				        
			</div><!-- eof col !-->
		</div><!-- eof row !-->

		<div class="well yellow">
			<!-- row !-->		
			<div class="row">
				<div class="col-lg-12">
	     
			        <div class="alert alert-info" role="alert">
			        	<strong>Ticket #: <%= MemoNumber %></strong>
			        </div>
				    
				</div><!-- eof col !-->
			</div><!-- eof row !-->
		        
			<!-- row !-->		
			<div class="row">
				<div class="col-lg-12 the-information">
	     
			         <div class="row">
			         
			        	<!-- Contact Name !-->
				    	<div class="col-lg-3">
				        	<strong>Contact Name</strong>
				        </div>
				    	<!-- Contact Name !-->
		
				    	<!-- Contact Phone !-->
				    	<div class="col-lg-3">
					    	 <strong>Contact Phone</strong>
				        </div>
				    	<!-- Contact Phone !-->
				    	
				    	<!-- Contact Phone !-->
				    	<div class="col-lg-6">
					    	 <strong>Contact Email</strong>
				        </div>
				    	<!-- Contact Phone !-->				    	
		 			    	
					</div>
					
			         <div class="row">
			         
			        	<!-- Contact Name !-->
				    	<div class="col-lg-3">
				        	<%= OpenSubmittedByName %>
				        </div>
				    	<!-- Contact Name !-->
		
				    	<!-- Contact Phone !-->
				    	<div class="col-lg-3">
					    	 <%= OpenSubmittedByPhone %>
				        </div>
				    	<!-- Contact Phone !-->

				    	<!-- Contact Email !-->
				    	<div class="col-lg-6">
					    	 <%= OpenSubmittedByEmail %>
				        </div>
				    	<!-- Contact Email !-->
				    	
					</div>

	

			         <div class="row" style="margin-top:20px">

						<!-- Problem Location !-->
				   		<div class="col-lg-3">
				        	<strong>Problem Location</strong>
				        </div>
				    	<!-- Problem Location !-->
		  					    	
				    	<!-- Description of problem !-->
				    	<div class="col-lg-9">
	 						<strong>Problem Description</strong>
				    	</div>
	 			    	<!-- Description of problem !-->
	 			    	
					</div>
					
			         <div class="row">

						<!-- Problem Location !-->
				   		<div class="col-lg-3">
				        	<%= OpenProblemLocation %>
				        </div>
				    	<!-- Problem Location !-->
		  					    	
				    	<!-- Description of problem !-->
				    	<div class="col-lg-9">
	 						<%= OpenProblemDescription %>
				    	</div>
	 			    	<!-- Description of problem !-->
	 			    	
	 			    	
					</div>
			    
				</div><!-- eof col !-->
			</div><!-- eof row !-->
		</div>	        


    	
    	<!-- Signature !-->

		 <div class="row well">
	
			 <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
			 	<strong>Signature (click to enlarge)</strong>
			 </div>
    	
			<% If GetServiceTicketStatus(MemoNumber) = "CLOSE" Then 

				'----------------------------
				'Service Signature Check
				'----------------------------
				set fs = CreateObject("Scripting.FileSystemObject")
				Pth =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & Trim(MemoNumber) & ".png"
				
				If fs.FileExists(Server.MapPath(Pth)) Then
					hasServiceSignature = True
				Else
					hasServiceSignature = False
				End If
									
				'Response.Write(Pth)
				
				'***************************************************************************************************
				'Display signature file, if any exist in the signaturesave directory
				''Check for the existance of a thumbnail image in the directory, otherwise, size the image with CSS
				'***************************************************************************************************

				Pth =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & Trim(MemoNumber) & ".png"
				PthThumb =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & Trim(MemoNumber) & "-thumb.png"

				SignaturePathNameFull = BaseURL & "clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & Trim(MemoNumber) & ".png"
				SignaturePathNameThumb = BaseURL & "clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & Trim(MemoNumber) & "-thumb.png"
				
				%><div class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><%
				
				If hasServiceSignature = True Then
					If fs.FileExists(Server.MapPath(PthThumb)) Then
				    	%><a class="fancybox" href="<%= SignaturePathNameFull %>" data-fancybox-group="gallerysig" title="Service Ticket #<%= Trim(MemoNumber) %> Signature"><img src="<%= SignaturePathNameThumb %>" alt=""></a><%
				    Else
				    	%><a class="fancybox" href="<%= SignaturePathNameFull %>" data-fancybox-group="gallerysig" title="Service Ticket #<%= Trim(MemoNumber) %> Signature"><img src="<%= SignaturePathNameFull %>" alt=""></a><%
				    End If
				Else
					 %><br>No Signature Entered<%
				End If
				
				%></div><%
			Else
				 %><div class="col-lg-12 col-md-12 col-sm-12 col-xs-12"><br>No Signature Entered - Ticket Still Open</div><%
			End If
				
			set fs=nothing
		%>

    	</div>
    	<!-- Signature !-->

	 		        
		<% If GetServiceTicketStatus(MemoNumber) = "CLOSE" Then
				z=0
				set fs = CreateObject("Scripting.FileSystemObject")
				For x = 1 to 20 ' Only have 3 pics but allow for expansion to 20 
					Pth =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/SvcMemoPics/" & Trim(MemoNumber) & "-" & x & ".png"
					Pth2 =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/SvcMemoPics/" & Trim(MemoNumber) & "-" & x & ".jpg"
					Pth3 =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/SvcMemoPics/" & Trim(MemoNumber) & "-" & x & ".jpeg"

					If fs.FileExists(Server.MapPath(Pth)) or fs.FileExists(Server.MapPath(Pth2)) or fs.FileExists(Server.MapPath(Pth3)) Then%>
						<% If x = 1 Then %>
						 <div class="row well">
					
							 <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
							 	<strong>Photos</strong>
							 </div>
							 
							 <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
								<p class="thumbnail-photos">
						<% End If %>
						
						<%
						If fs.FileExists(Server.MapPath(Pth)) Then %><a class="fancybox" href="<%= Pth %>" data-fancybox-group="gallery" title="Service Ticket #<%= Trim(MemoNumber) %>"><img src="<%= Pth %>" alt="" class="thumbnail"></a><% End If													
						If fs.FileExists(Server.MapPath(Pth2)) Then %><a class="fancybox" href="<%= Pth2 %>" data-fancybox-group="gallery" title="Service Ticket #<%= Trim(MemoNumber) %>"><img src="<%= Pth2 %>" alt="" class="thumbnail"></a><% End If
						If fs.FileExists(Server.MapPath(Pth3)) Then %><a class="fancybox" href="<%= Pth3 %>" data-fancybox-group="gallery" title="Service Ticket #<%= Trim(MemoNumber) %>"><img src="<%= Pth3 %>" alt="" class="thumbnail"></a><% End If													
					End If
					If z = 2 then
						%> <%
						z=0
					Else
						z=z+1 'Three per row
					End If
				Next
				%></p>
				
				        </div>
			    	<!-- Close or Cancel !-->		    	
				</div>
				<!-- eof row !-->
							
				
				<%	
			End If
			set fs=nothing
		%>
		<!-- eof main row !-->

		<% MDG_MemoNumber = MemoNumber %>
		
		<div class="row well white">
 
 			<h4 style="margin-bottom:10px;">Previous Ticket Notes</h4>
 			
			<!-- left col !-->
			<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
			
			<table id="serviceTicketNotesTable" class="display" style="width:100%;font-size:11px;">
			        <thead>
			            <tr>
			                <th>Date/Time</th>
			                <th>Status/Stage</th>
			                <th>User</th>
			                <th>Notes</th>
			            </tr>
			        </thead>
			        <tbody>
			            <!--#include file="serviceTicketNotesDetailsTable.asp"-->
			        </tbody>
			        <tfoot>
			            <tr>
			                <th width="20%">Date/Time</th>
			                <th width="20%">Status/Stage</th>
			                <th width="15%">User</th>
			                <th width="45%">Notes</th>
			            </tr>
			        </tfoot>
			    </table>	

			</div>
			
		</div>
		<!-- eof main row !-->

		<!-- lightbox JS !-->
		<div id="lightbox" class="modal fade" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel" aria-hidden="true">
		    <div class="modal-dialog">
		        <button type="button" class="close hidden" data-dismiss="modal" aria-hidden="true">×</button>
		        <div class="modal-content">
		            <div class="modal-body">
		                <img src="" alt="" />
		            </div>
		        </div>
		    </div>
		</div>

		<script>
			$(document).ready(function() {
			    var $lightbox = $('#lightbox');
			    
			    $('[data-target="#lightbox"]').on('click', function(event) {
			        var $img = $(this).find('img'), 
			            src = $img.attr('src'),
			            alt = $img.attr('alt'),
			            css = {
			                'maxWidth': $(window).width() - 100,
			                'maxHeight': $(window).height() - 100
			            };
			    
			        $lightbox.find('.close').addClass('hidden');
			        $lightbox.find('img').attr('src', src);
			        $lightbox.find('img').attr('alt', alt);
			        $lightbox.find('img').css(css);
			    });
	    
			    $lightbox.on('shown.bs.modal', function (e) {
			        var $img = $lightbox.find('img');
			            
			        $lightbox.find('.modal-dialog').css({'width': $img.width()});
			        $lightbox.find('.close').removeClass('hidden');
			    });
			});
							
		</script>
		<!-- eof lightbox JS !-->

	<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub SetRegionFilterListByUserForServiceBoard() 

	UserNo = Request.Form("UserNo")
	RegionList = Request.Form("RegionsToView")
	

	Set cnnSetRegionFilterListByUserForServiceBoard = Server.CreateObject("ADODB.Connection")
	cnnSetRegionFilterListByUserForServiceBoard.open (Session("ClientCnnString"))
	Set rsSetRegionFilterListByUserForServiceBoard = Server.CreateObject("ADODB.Recordset")
	rsSetRegionFilterListByUserForServiceBoard.CursorLocation = 3 
	
	SQLSetRegionFilterListByUserForServiceBoard = "UPDATE tblUsers SET UserRegionsToViewService = '" & RegionList & "' WHERE UserNo = " & UserNo
	Set rsSetRegionFilterListByUserForServiceBoard = cnnSetRegionFilterListByUserForServiceBoard.Execute(SQLSetRegionFilterListByUserForServiceBoard)

		
	set rsSetRegionFilterListByUserForServiceBoard = Nothing
	cnnSetRegionFilterListByUserForServiceBoard.close
	set cnnSetRegionFilterListByUserForServiceBoard = Nothing
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetNextPageForLinkedKiosks() 

	CurrentKioskURL = Request.Form("url")
	
	'*********************************************************
	'Write function to get all region numbers for CDC here
	'*********************************************************
	
		'We know we are on CDC, if we just displayed the US Coffee Field Service Kiosk
		
		If InStr(CurrentKioskURL, "FieldServiceKioskNoPaging_linked_version.asp?pp=Hicksville&cl=1230") Then
		
			Session("ClientCnnString") = "Driver={SQL Server};Server=66.201.99.15;Database=CDC;Uid=cdcinsight;Pwd=2oobr04dw4y;"
		
			SQLRegionList = "SELECT * FROM AR_Regions WHERE UseRegionForServiceTickets = 1 ORDER BY InternalRecordIdentifier ASC"
			
			Set cnnRegionList = Server.CreateObject("ADODB.Connection")
			cnnRegionList.open (Session("ClientCnnString"))
			Set rsRegionList = Server.CreateObject("ADODB.Recordset")
			rsRegionList.CursorLocation = 3 
			Set rsRegionList = cnnRegionList.Execute(SQLRegionList)
			
			regionNumberList = ""
				
			If Not rsRegionList.EOF Then
				Do While Not rsRegionList.EOF
					regionNumber = rsRegionList("InternalRecordIdentifier")
					regionNumberList = regionNumberList & regionNumber & ","
					rsRegionList.MoveNext
				Loop
			End If
			
			If regionNumberList <> "" Then
				regionNumberList = left(regionNumberList,len(regionNumberList)-1)
				regionNumberListArray = Split(regionNumberList,",")
				firstRegionToDisplay = regionNumberListArray(0)
			End If
				
			set rsRegionList = Nothing
			cnnRegionList.close
			set cnnRegionList = Nothing
			
		End If

	'*********************************************************
	
	If InStr(CurrentKioskURL, "DeliveryBoardKioskNoPaging_linked_version.asp?pp=ButterCup&cl=1071") Then
		NextKioskURL = BaseURL & "directLaunch/kiosks/service/FieldServiceKioskNoPaging_linked_version.asp?pp=ButterCup&cl=1071&ri=15"
		
	ElseIf InStr(CurrentKioskURL, "FieldServiceKioskNoPaging_linked_version.asp?pp=ButterCup&cl=1071") Then
		NextKioskURL = BaseURL & "directLaunch/kiosks/routing/DeliveryBoardKioskNoPaging_linked_version.asp?pp=Hicksville&cl=1230&ri=15"
		
	ElseIf InStr(CurrentKioskURL, "DeliveryBoardKioskNoPaging_linked_version.asp?pp=Hicksville&cl=1230") Then
		NextKioskURL = BaseURL & "directLaunch/kiosks/service/FieldServiceKioskNoPaging_linked_version.asp?pp=Hicksville&cl=1230&ri=15"
		
	ElseIf InStr(CurrentKioskURL, "FieldServiceKioskNoPaging_linked_version.asp?pp=Hicksville&cl=1230") Then
	
		If regionNumberList = "" Then
			NextKioskURL = BaseURL & "directLaunch/kiosks/service/FieldServiceKioskNoPaging_linked_version.asp?pp=Fiddle&cl=1106&ri=15"
		Else
			NextKioskURL = BaseURL & "directLaunch/kiosks/service/FieldServiceKioskNoPaging_linked_version.asp?pp=Fiddle&cl=1106&ri=15&rgn=" & regionNumberList & "&rgnc=" & firstRegionToDisplay
		End If
		
	ElseIf InStr(CurrentKioskURL, "FieldServiceKioskNoPaging_linked_version.asp?pp=Fiddle&cl=1106") Then
		NextKioskURL = BaseURL & "directLaunch/kiosks/routing/DeliveryBoardKioskNoPaging_linked_version.asp?pp=ButterCup&cl=1071&ri=15"
		
	Else 
		NextKioskURL = ""
	End If
	
	Response.Write(NextKioskURL)
	
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'END ALL AJAX MODAL SUBROUTINES AND FUNCTIONS

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

%>