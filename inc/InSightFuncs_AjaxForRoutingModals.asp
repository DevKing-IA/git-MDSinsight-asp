<!--#include file="subsandfuncs.asp"-->
<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Routing.asp"-->
<!--#include file="InSightFuncs_Prospecting.asp"-->
<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
 
'Sub TurnOnNagAlertsForDeliveryBoardDriver()
'Sub TurnOffNagAlertsForDeliveryBoardDriver()
'Sub TurnOnNagAlertsForDeliveryBoardDriverKiosk()
'Sub TurnOffNagAlertsForDeliveryBoardDriverKiosk()
'Sub CheckDeliveryStatus()
'Sub CheckDeliveryIsNextStop()
'Sub ChangeDeliveryPlanningBoard()

'Sub GetContentForDeliveryBoardOptionsModal()
'Sub GetContentForCompletedOrSkippedInfoModal()
'Sub GetHistoricalContentForCompletedOrSkippedInfoModal()

'Sub MarkDeliveryAsPriorityFromDeliveryBoardModalAndSendText()
'Sub MarkDeliveryAsPriorityFromDeliveryBoardModal()
'Sub RemovePriorityFromDeliveryBoardModalAndSendText()
'Sub RemovePriorityFromDeliveryBoardModal()

'Sub ToggleDeliveryAsPriorityFromDeliveryBoardModal()

'Sub GetContentForAddDeliveryBoardAlertModal()
'Sub GetContentForEditDeliveryBoardAlertModal()

'Sub AddAlertFromDeliveryBoardModal()
'Sub EditAlertFromDeliveryBoardModal()
'Sub DeleteAlertFromDeliveryBoardModal()

'Sub MarkAsAMDeliveryFromDeliveryBoardModalAndSendText()
'Sub MarkAsAMDeliveryFromDeliveryBoardModal()
'Sub RemoveAMDeliveryFromDeliveryBoardModalAndSendText()
'Sub RemoveAMDeliveryFromDeliveryBoardModal()

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

	Case "GetContentForAddDeliveryBoardAlertModal" 
		GetContentForAddDeliveryBoardAlertModal()	
	Case "GetContentForEditDeliveryBoardAlertModal" 
		GetContentForEditDeliveryBoardAlertModal()	
	Case "AddAlertFromDeliveryBoardModal" 
		AddAlertFromDeliveryBoardModal()
	Case "EditAlertFromDeliveryBoardModal" 
		EditAlertFromDeliveryBoardModal()
	Case "DeleteAlertFromDeliveryBoardModal"
		DeleteAlertFromDeliveryBoardModal()
		
	Case "GetContentForCompletedOrSkippedInfoModal" 
		GetContentForCompletedOrSkippedInfoModal()
	Case "GetHistoricalContentForCompletedOrSkippedInfoModal" 
		GetHistoricalContentForCompletedOrSkippedInfoModal()
		
	Case "TurnOffNagAlertsForDeliveryBoardDriver"
		TurnOffNagAlertsForDeliveryBoardDriver()	
	Case "TurnOnNagAlertsForDeliveryBoardDriver"
		TurnOnNagAlertsForDeliveryBoardDriver()	
	Case "TurnOffNagAlertsForDeliveryBoardDriverKiosk"
		TurnOffNagAlertsForDeliveryBoardDriverKiosk()	
	Case "TurnOnNagAlertsForDeliveryBoardDriverKiosk"
		TurnOnNagAlertsForDeliveryBoardDriverKiosk()
			
	Case "CheckDeliveryStatus"
		CheckDeliveryStatus()
	Case "CheckDeliveryIsNextStop"
		CheckDeliveryIsNextStop()
	Case "ChangeDeliveryPlanningBoard"
		ChangeDeliveryPlanningBoard()	
		
	Case "GetContentForDeliveryBoardOptionsModal"
		GetContentForDeliveryBoardOptionsModal()
	Case "MarkDeliveryAsPriorityFromDeliveryBoardModalAndSendText"
		MarkDeliveryAsPriorityFromDeliveryBoardModalAndSendText()
	Case "MarkDeliveryAsPriorityFromDeliveryBoardModal"
		MarkDeliveryAsPriorityFromDeliveryBoardModal()
	Case "RemovePriorityFromDeliveryBoardModalAndSendText"
		RemovePriorityFromDeliveryBoardModalAndSendText()
	Case "RemovePriorityFromDeliveryBoardModal"
		RemovePriorityFromDeliveryBoardModal()
	Case "ToggleDeliveryAsPriorityFromDeliveryBoardModal"
		ToggleDeliveryAsPriorityFromDeliveryBoardModal()
	Case "MarkAsAMDeliveryFromDeliveryBoardModalAndSendText"
		MarkAsAMDeliveryFromDeliveryBoardModalAndSendText()
	Case "MarkAsAMDeliveryFromDeliveryBoardModal"
		MarkAsAMDeliveryFromDeliveryBoardModal()
	Case "RemoveAMDeliveryFromDeliveryBoardModalAndSendText"
		RemoveAMDeliveryFromDeliveryBoardModalAndSendText()
	Case "RemoveAMDeliveryFromDeliveryBoardModal"
		RemoveAMDeliveryFromDeliveryBoardModal()		
End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub ChangeDeliveryPlanningBoard() 

	invoiceNo = Request.Form("invoiceNo")
	destinationTruck = Request.Form("destinationTruck")
	destinationSeqNo = Request.Form("destinationSeqNo ")

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub CheckDeliveryStatus() 

	invoiceNo = Request.Form("invoiceNo")
	deliveryStatus = GetDeliveryStatusByInvoice(invoiceNo)
	Response.Write(deliveryStatus)

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub CheckDeliveryIsNextStop() 

	invoiceNo = Request.Form("invoiceNo")
	deliveryIsNextStop = InvoiceIsNextStop(invoiceNo)
	Response.Write(deliveryIsNextStop)

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub TurnOffNagAlertsForDeliveryBoardDriver() 

	driverUserNo = Request.Form("driverUserNo")
	
	'***************************************************************************************
	'Delete nag types of routingNoActivity and routingNoNextStop for this user number
	'***************************************************************************************

	Set cnnDeliveryNagAlert = Server.CreateObject("ADODB.Connection")
	cnnDeliveryNagAlert.open (Session("ClientCnnString"))
	Set rsDeliveryNagAlert = Server.CreateObject("ADODB.Recordset")
	rsDeliveryNagAlert.CursorLocation = 3 
	
	SQLDeliveryNagAlert = "DELETE FROM SC_NagSkipUsers WHERE UserNo = " & driverUserNo & " AND NagType = 'routingNoActivity'"
	
	Response.write(SQLDeliveryNagAlert)
	
	Set rsDeliveryNagAlert = cnnDeliveryNagAlert.Execute(SQLDeliveryNagAlert)
	
	SQLDeliveryNagAlert = "DELETE FROM SC_NagSkipUsers WHERE UserNo = " & driverUserNo & " AND NagType = 'routingNoNextStop'"
	Set rsDeliveryNagAlert = cnnDeliveryNagAlert.Execute(SQLDeliveryNagAlert)

		
	set rsDeliveryNagAlert = Nothing
	cnnDeliveryNagAlert.close
	set cnnDeliveryNagAlert = Nothing
	

	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " turned OFF nag alerts today for " & GetUserDisplayNameByUserNo(driverUserNo)
	CreateAuditLogEntry "Nag Alerts Turned Off For Driver","Nag Alerts Turned Off For Driver","Minor",0,Description
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub TurnOnNagAlertsForDeliveryBoardDriver() 

	driverUserNo = Request.Form("driverUserNo")
	
	'***************************************************************************************
	'Get nag type values for passed driver user number
	'***************************************************************************************

	Set cnnDeliveryNagAlert = Server.CreateObject("ADODB.Connection")
	cnnDeliveryNagAlert.open (Session("ClientCnnString"))
	Set rsDeliveryNagAlert = Server.CreateObject("ADODB.Recordset")
	rsDeliveryNagAlert.CursorLocation = 3 
	
	Set cnnDeliveryNagAlertUpdateInsert = Server.CreateObject("ADODB.Connection")
	cnnDeliveryNagAlertUpdateInsert.open (Session("ClientCnnString"))
	Set rsDeliveryNagAlertUpdateInsert = Server.CreateObject("ADODB.Recordset")
	rsDeliveryNagAlertUpdateInsert.CursorLocation = 3 
	
	
	SQLDeliveryNagAlert = "SELECT * FROM SC_NagSkipUsers WHERE UserNo = " & driverUserNo & " AND NagType = 'routingNoActivity'"
	response.write(SQLDeliveryNagAlert)
	Set rsDeliveryNagAlert = cnnDeliveryNagAlert.Execute(SQLDeliveryNagAlert)
	
	If rsDeliveryNagAlert.EOF THEN
		SQLDeliveryNagAlertUpdateInsert = "INSERT INTO SC_NagSkipUsers (UserNo, NagType) VALUES (" & driverUserNo & ",'routingNoActivity')"
		Set rsDeliveryNagAlertUpdateInsert = cnnDeliveryNagAlertUpdateInsert.Execute(SQLDeliveryNagAlertUpdateInsert)
	End If	

	
	SQLDeliveryNagAlert = "SELECT * FROM SC_NagSkipUsers WHERE UserNo = " & driverUserNo & " AND NagType = 'routingNoNextStop'"
	Set rsDeliveryNagAlert = cnnDeliveryNagAlert.Execute(SQLDeliveryNagAlert)
	
	If rsDeliveryNagAlert.EOF THEN
		SQLDeliveryNagAlertUpdateInsert = "INSERT INTO SC_NagSkipUsers (UserNo, NagType) VALUES (" & driverUserNo & ",'routingNoNextStop')"
		Set rsDeliveryNagAlertUpdateInsert = cnnDeliveryNagAlertUpdateInsert.Execute(SQLDeliveryNagAlertUpdateInsert)
	End If	
	
	set rsDeliveryNagAlertUpdateInsert = Nothing
	cnnDeliveryNagAlertUpdateInsert.close
	set cnnDeliveryNagAlertUpdateInsert = Nothing
		
	set rsDeliveryNagAlert = Nothing
	cnnDeliveryNagAlert.close
	set cnnDeliveryNagAlert = Nothing

	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " turned ON nag alerts today for " & GetUserDisplayNameByUserNo(driverUserNo)
	CreateAuditLogEntry "Nag Alerts Turned On For Driver","Nag Alerts Turned On For Driver","Minor",0,Description
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub TurnOffNagAlertsForDeliveryBoardDriverKiosk() 

	driverUserNo = Request.Form("driverUserNo")
	
	'***************************************************************************************
	'Delete nag types of routingNoActivity and routingNoNextStop for this user number
	'***************************************************************************************

	Set cnnDeliveryNagAlert = Server.CreateObject("ADODB.Connection")
	cnnDeliveryNagAlert.open (Session("ClientCnnString"))
	Set rsDeliveryNagAlert = Server.CreateObject("ADODB.Recordset")
	rsDeliveryNagAlert.CursorLocation = 3 
	
	SQLDeliveryNagAlert = "DELETE FROM SC_NagSkipUsers WHERE UserNo = " & driverUserNo & " AND NagType = 'routingNoActivity'"
	
	Response.write(SQLDeliveryNagAlert)
	
	Set rsDeliveryNagAlert = cnnDeliveryNagAlert.Execute(SQLDeliveryNagAlert)
	
	SQLDeliveryNagAlert = "DELETE FROM SC_NagSkipUsers WHERE UserNo = " & driverUserNo & " AND NagType = 'routingNoNextStop'"
	Set rsDeliveryNagAlert = cnnDeliveryNagAlert.Execute(SQLDeliveryNagAlert)

		
	set rsDeliveryNagAlert = Nothing
	cnnDeliveryNagAlert.close
	set cnnDeliveryNagAlert = Nothing
	

	Description = GetUserDisplayNameByUserNo(0) & " turned OFF nag alerts today for " & GetUserDisplayNameByUserNo(driverUserNo)
	CreateAuditLogEntry "Nag Alerts Turned Off For Driver","Nag Alerts Turned Off For Driver","Minor",0,Description
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub TurnOnNagAlertsForDeliveryBoardDriverKiosk() 

	driverUserNo = Request.Form("driverUserNo")
	
	'***************************************************************************************
	'Get nag type values for passed driver user number
	'***************************************************************************************

	Set cnnDeliveryNagAlert = Server.CreateObject("ADODB.Connection")
	cnnDeliveryNagAlert.open (Session("ClientCnnString"))
	Set rsDeliveryNagAlert = Server.CreateObject("ADODB.Recordset")
	rsDeliveryNagAlert.CursorLocation = 3 
	
	Set cnnDeliveryNagAlertUpdateInsert = Server.CreateObject("ADODB.Connection")
	cnnDeliveryNagAlertUpdateInsert.open (Session("ClientCnnString"))
	Set rsDeliveryNagAlertUpdateInsert = Server.CreateObject("ADODB.Recordset")
	rsDeliveryNagAlertUpdateInsert.CursorLocation = 3 
	
	
	SQLDeliveryNagAlert = "SELECT * FROM SC_NagSkipUsers WHERE UserNo = " & driverUserNo & " AND NagType = 'routingNoActivity'"
	response.write(SQLDeliveryNagAlert)
	Set rsDeliveryNagAlert = cnnDeliveryNagAlert.Execute(SQLDeliveryNagAlert)
	
	If rsDeliveryNagAlert.EOF THEN
		SQLDeliveryNagAlertUpdateInsert = "INSERT INTO SC_NagSkipUsers (UserNo, NagType) VALUES (" & driverUserNo & ",'routingNoActivity')"
		Set rsDeliveryNagAlertUpdateInsert = cnnDeliveryNagAlertUpdateInsert.Execute(SQLDeliveryNagAlertUpdateInsert)
	End If	

	
	SQLDeliveryNagAlert = "SELECT * FROM SC_NagSkipUsers WHERE UserNo = " & driverUserNo & " AND NagType = 'routingNoNextStop'"
	Set rsDeliveryNagAlert = cnnDeliveryNagAlert.Execute(SQLDeliveryNagAlert)
	
	If rsDeliveryNagAlert.EOF THEN
		SQLDeliveryNagAlertUpdateInsert = "INSERT INTO SC_NagSkipUsers (UserNo, NagType) VALUES (" & driverUserNo & ",'routingNoNextStop')"
		Set rsDeliveryNagAlertUpdateInsert = cnnDeliveryNagAlertUpdateInsert.Execute(SQLDeliveryNagAlertUpdateInsert)
	End If	
	
	set rsDeliveryNagAlertUpdateInsert = Nothing
	cnnDeliveryNagAlertUpdateInsert.close
	set cnnDeliveryNagAlertUpdateInsert = Nothing
		
	set rsDeliveryNagAlert = Nothing
	cnnDeliveryNagAlert.close
	set cnnDeliveryNagAlert = Nothing

	Description = GetUserDisplayNameByUserNo(0) & " turned ON nag alerts today for " & GetUserDisplayNameByUserNo(driverUserNo)
	CreateAuditLogEntry "Nag Alerts Turned On For Driver","Nag Alerts Turned On For Driver","Minor",0,Description
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForCompletedOrSkippedInfoModal() 

	InvoiceNumber = Request.Form("myInvoiceNumber")
	
	
	'***************************************************************************************
	'Get values for current invoice/stop
	'***************************************************************************************

	SQLDeliveryAlertEdit = "SELECT * FROM RT_DeliveryBoard WHERE IvsNum = '" & InvoiceNumber & "'"
	
	Set cnnDeliveryAlertEdit = Server.CreateObject("ADODB.Connection")
	cnnDeliveryAlertEdit.open (Session("ClientCnnString"))
	Set rsDeliveryAlertEdit = Server.CreateObject("ADODB.Recordset")
	rsDeliveryAlertEdit.CursorLocation = 3 
	Set rsDeliveryAlertEdit = cnnDeliveryAlertEdit.Execute(SQLDeliveryAlertEdit)
		
	If not rsDeliveryAlertEdit.EOF Then
		DeliveryStatus = rsDeliveryAlertEdit("DeliveryStatus")
		LastDeliveryStatusChange = rsDeliveryAlertEdit("LastDeliveryStatusChange") 
		TruckNumber = rsDeliveryAlertEdit("TruckNumber")
		DriverComments = rsDeliveryAlertEdit("DriverComments")
		ReferenceValue = InvoiceNumber
	End If
	set rsDeliveryAlertEdit = Nothing
	cnnDeliveryAlertEdit.close
	set cnnDeliveryAlertEdit = Nothing
	
	'***************************************************************************************
	
	
%>
		<input type="hidden" name="txtIvsNum" id="txtIvsNum" value="<%= InvoiceNumber %>">
		
		<!-- when line !-->
		<div class="row-line">
	
			<!-- when !-->
			<div class="col-lg-9">
				<%
					If datediff("d",LastDeliveryStatusChange,Now) < 1 Then
											
						If datediff("n",LastDeliveryStatusChange,Now) > 60 then
						
							hours = datediff("n",LastDeliveryStatusChange,Now) \ 60
							minutes = datediff("n",LastDeliveryStatusChange,Now) mod 60
							
							lastUpdate = " " & hours & " hours and " & minutes & " minutes ago."
							
						Else
							lastUpdate = datediff("n",LastDeliveryStatusChange,Now) & " minutes ago."
						End If
						
					End If
				%>
				<% If DeliveryStatus = "No Delivery" Then %>
					<div class="alert alert-danger" role="alert" style="margin-bottom:0px;">
					  <strong>
					  Not Delivered <%= lastUpdate %>
					  <%If DriverComments <> "" Then
					  	Response.Write("<br><br><font color='black'>Driver Comments:</font> " & DriverComments)
					  End If %>
					  </strong> 
					</div>
				<% ElseIf DeliveryStatus = "Delivered" Then %>
					<div class="alert alert-success" role="alert" style="margin-bottom:0px;">
					  <strong>Delivered <%= lastUpdate %>
  					  <%If DriverComments <> "" Then
					  	Response.Write("<br><br><font color='black'>Driver Comments:</font> " & DriverComments)
					  End If %>
					</strong>
					</div>
				<% Else %>
					<label>Stop Information</label>
				<% End If %>
			</div>
			<!-- eof when !-->

        </div>
        <!-- eof when line !-->

		
		<!-- when line !-->
		<div class="row-line">
	
			<!-- condition !-->
			<div class="col-lg-9">
				Marked as <%= DeliveryStatus %> by <%= GetDriverNameByTruckID(TruckNumber) %> at <%= LastDeliveryStatusChange %>.<br>
			</div>
			<!-- eof condition !-->
        </div>
        <!-- eof when line !-->

	
	<!-- eof modal body !-->
      
	<!-- modal footer !-->
    <div class="modal-footer">
		      			      	      
		<!-- close / save !-->
		<div class="col-lg-12">
			<button type="button" class="btn btn-default btn-sm" data-dismiss="modal">Close</button>
		</div>
		<!-- eof close / save !-->

	</div>
	<!-- eof modal footer !-->
	

<%
End Sub



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetHistoricalContentForCompletedOrSkippedInfoModal() 

	InvoiceNumber = Request.Form("myInvoiceNumber")
	
	
	'***************************************************************************************
	'Get values for current invoice/stop
	'***************************************************************************************

	SQLDeliveryAlertEdit = "SELECT * FROM RT_DeliveryBoardHistory WHERE IvsNum = '" & InvoiceNumber & "'"
	
	Set cnnDeliveryAlertEdit = Server.CreateObject("ADODB.Connection")
	cnnDeliveryAlertEdit.open (Session("ClientCnnString"))
	Set rsDeliveryAlertEdit = Server.CreateObject("ADODB.Recordset")
	rsDeliveryAlertEdit.CursorLocation = 3 
	Set rsDeliveryAlertEdit = cnnDeliveryAlertEdit.Execute(SQLDeliveryAlertEdit)
		
	If not rsDeliveryAlertEdit.EOF Then
		DeliveryStatus = rsDeliveryAlertEdit("DeliveryStatus")
		LastDeliveryStatusChange = rsDeliveryAlertEdit("LastDeliveryStatusChange") 
		TruckNumber = rsDeliveryAlertEdit("TruckNumber")
		ReferenceValue = InvoiceNumber
	End If
	set rsDeliveryAlertEdit = Nothing
	cnnDeliveryAlertEdit.close
	set cnnDeliveryAlertEdit = Nothing
	
	'***************************************************************************************
	
	
%>
		<input type="hidden" name="txtIvsNum" id="txtIvsNum" value="<%= InvoiceNumber %>">
		
		<!-- when line !-->
		<div class="row-line">
	
			<!-- when !-->
			<div class="col-lg-9">
				<%
					If datediff("d",LastDeliveryStatusChange,Now) < 1 Then
											
						If datediff("n",LastDeliveryStatusChange,Now) > 60 then
						
							hours = datediff("n",LastDeliveryStatusChange,Now) \ 60
							minutes = datediff("n",LastDeliveryStatusChange,Now) mod 60
							
							lastUpdate = " " & hours & " hours and " & minutes & " minutes ago."
							
						Else
							lastUpdate = datediff("n",LastDeliveryStatusChange,Now) & " minutes ago."
						End If
						
					End If
				%>
				<% If DeliveryStatus = "No Delivery" Then %>
					<div class="alert alert-danger" role="alert" style="margin-bottom:0px;">
					  <strong>Not Delivered <%= lastUpdate %></strong> 
					</div>
				<% ElseIf DeliveryStatus = "Delivered" Then %>
					<div class="alert alert-success" role="alert" style="margin-bottom:0px;">
					  <strong>Delivered <%= lastUpdate %></strong>
					</div>
				<% Else %>
					<label>Stop Information</label>
				<% End If %>
			</div>
			<!-- eof when !-->

        </div>
        <!-- eof when line !-->

		
		<!-- when line !-->
		<div class="row-line">
	
			<!-- condition !-->
			<div class="col-lg-9">
				This invoice was updated to a status of <%= DeliveryStatus %> by <%= GetDriverNameByTruckID(TruckNumber) %> at <%= LastDeliveryStatusChange %>.<br>
			</div>
			<!-- eof condition !-->
        </div>
        <!-- eof when line !-->

	
	<!-- eof modal body !-->
      
	<!-- modal footer !-->
    <div class="modal-footer">
		      			      	      
		<!-- close / save !-->
		<div class="col-lg-12">
			<button type="button" class="btn btn-default btn-sm" data-dismiss="modal">Close</button>
		</div>
		<!-- eof close / save !-->

	</div>
	<!-- eof modal footer !-->
	

<%
End Sub






Sub GetContentForAddDeliveryBoardAlertModal() 

	InvoiceNumber = Request.Form("myInvoiceNumber")
		
%>

	<script type="text/javascript">
	
		$(document).ready(function() {

			$('#btnAddDeliveryAlertSave').on('click', function(e) {
			
			    //get data-id attribute of the clicked alert
			    var invoiceNum = $("#txtIvsNum").val();
			    var condition = $("#selCondition").val();
			    var emailto = $("#selEmailto").val();
			    var addlemails = $("#txtAdditionalEmails").val();
			    var textto = $("#selTextto").val();
			    var addltexts = $("#txtAdditionalTexts").val();
			    				
				//turn off the automatic page refresh
				$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			    		    		
		    	$.ajax({
					type:"POST",
					url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
					data: "action=AddAlertFromDeliveryBoardModal&invoiceNum=" + encodeURIComponent(invoiceNum) + "&condition=" + encodeURIComponent(condition) + "&emailto=" + encodeURIComponent(emailto) + "&addlemails=" + encodeURIComponent(addlemails) + "&textto=" + encodeURIComponent(textto) + "&addltexts=" + encodeURIComponent(addltexts),
					success: function(response)
					 {
					 	location.reload();
		             }
				});
	    	});	
	    	
	    	    		
		});
	</script>

		<input type="hidden" name="txtIvsNum" id="txtIvsNum" value="<%= InvoiceNumber %>">
		
		<!-- when line !-->
		<div class="row-line">
	
			<!-- when !-->
			<div class="col-lg-2">
				<label>When</label>
			</div>
			<!-- eof when !-->
	
			<!-- condition !-->
			<div class="col-lg-6">
				<select class="form-control" name="selCondition" id="selCondition">
					<option value="Stop is completed or skipped">Stop is completed or skipped</option>
					<option value="This becomes the next stop">This becomes the next stop</option>
				</select>
			</div>
			<!-- eof condition !-->
        </div>
        <!-- eof when line !-->


		<!-- email alert line !-->
		<div class="row-line">

			<!-- email alert !-->
			<div class="col-lg-2">
				<label class="right">Send an email alert to</label>
			</div>
			<!-- eof email alert !-->

			<!-- multi select !-->
			<div class="col-lg-4">
				<select class="form-control multiselect" id="selEmailto" name="selEmailto" multiple>
					<option value="0">--- none from here ---</option>
					<option value="<%=Session("UserNo")%>"><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option> 								
					<% 
			      	
					Set cnnDeliveryModal = Server.CreateObject("ADODB.Connection")
					cnnDeliveryModal.open (Session("ClientCnnString"))
					Set rsDeliveryModal = Server.CreateObject("ADODB.Recordset")
							      	
			      	 
		      	  	SQLDeliveryModal = "SELECT UserNo, userFirstName, userLastName FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
		      	  	SQLDeliveryModal = SQLDeliveryModal & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo")
		      	  	SQLDeliveryModal = SQLDeliveryModal & " ORDER BY  userFirstName, userLastName"
	
					rsDeliveryModal.CursorLocation = 3 
					Set rsDeliveryModal = cnnDeliveryModal.Execute(SQLDeliveryModal)
				
					If not rsDeliveryModal.EOF Then
						Do
							FullName = rsDeliveryModal("userFirstName") & " " & rsDeliveryModal("userLastName")
							Response.Write("<option value='" & rsDeliveryModal("UserNo") & "'>" & FullName & "</option>")
							rsDeliveryModal.movenext
							
						Loop until rsDeliveryModal.eof
					End If
					set rsDeliveryModal = Nothing
					cnnDeliveryModal.close
					set cnnDeliveryModal = Nothing
			      	%>
				</select>
				<strong>Use CTRL and SHIFT to make multiple selections</strong>
            </div>
			<!-- eof multi select !-->
    
            <!-- text area !-->
            <div class="col-lg-6">
				<textarea class="form-control textarea" rows="4" id="txtAdditionalEmails" name="txtAdditionalEmails"></textarea>
	            <strong>Separate multiple email addresses with a semicolon</strong>
            </div>
            <!-- eof text area !-->
        </div>
        <!-- eof email alert line !-->


        <!-- text alert line !-->
        <div class="row-line">

        	<!-- text alert !-->
            <div class="col-lg-2">
	            <label class="right">Send a text alert to</label>
            </div>
            <!-- eof text alert !-->
    
            <!-- multi select !-->
            <div class="col-lg-4">
				<select class="form-control multiselect"  id="selTextto" name="selTextto" multiple>
					<option value="0">--- none from here ---</option>
					<option value="<%=Session("UserNo")%>"><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option>								
					<% 


					Set cnnDeliveryModal2 = Server.CreateObject("ADODB.Connection")
					cnnDeliveryModal2.open (Session("ClientCnnString"))
					Set rsDeliveryModal2 = Server.CreateObject("ADODB.Recordset")
			      	 
		      	  	SQLDeliveryModal2 = "SELECT UserNo, userFirstName, userLastName FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
		      	  	SQLDeliveryModal2 = SQLDeliveryModal2 & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo")
		      	  	SQLDeliveryModal2 = SQLDeliveryModal2 & " ORDER BY  userFirstName, userLastName"
		      	  	
		      	  	Response.write(SQLDeliveryModal2)
	
					rsDeliveryModal2.CursorLocation = 3 
					Set rsDeliveryModal2 = cnnDeliveryModal2.Execute(SQLDeliveryModal2)
				
					If not rsDeliveryModal2.EOF Then
						Do
							FullName = rsDeliveryModal2("userFirstName") & " " & rsDeliveryModal2("userLastName")
							If getUserCellNumber(Session("UserNo")) <> "" Then 
								Response.Write("<option value='" & rsDeliveryModal2("UserNo") & "'>" & FullName & "</option>")
							End If
							rsDeliveryModal2.movenext
						Loop until rsDeliveryModal2.eof
					End If
					set rsDeliveryModal2 = Nothing
					cnnDeliveryModal2.close
					set cnnDeliveryModal2 = Nothing
			      	%>
				</select>
				<strong>Use CTRL and SHIFT to make multiple selections</strong>
            </div>
            <!-- eof multi select !-->
    
	     <!-- text area !-->
    	<div class="col-lg-6">
			<textarea class="form-control textarea" rows="4" id="txtAdditionalTexts" name="txtAdditionalTexts"></textarea>
            <strong>Separate multiple phone numbers with a semicolon</strong>
		</div>
        <!-- eof text area !-->
	</div>
	<!-- eof text alert line !-->
	
	<!-- eof modal body !-->
      
	<!-- modal footer !-->
    <div class="modal-footer">
		      
		<!-- delete !-->
		<div class="col-lg-3">
			&nbsp;
		</div>
		<!-- eof delete !-->		
			      
		<!-- alert !-->
		<div class="col-lg-5 bottom-alert">
			*Delivery alerts are one-time events and automatically delete themselves after the alert has been sent. 
		</div>
		<!-- eof alert !-->
	      
		<!-- close / save !-->
		<div class="col-lg-4">
			<button type="button" class="btn btn-default btn-sm" data-dismiss="modal">Close</button>
			<button type="button" id="btnAddDeliveryAlertSave" class="btn btn-primary btn-sm">Save changes</button>
		</div>
		<!-- eof close / save !-->

	</div>
	<!-- eof modal footer !-->

<%
End Sub







Sub GetContentForEditDeliveryBoardAlertModal() 


	InvoiceNumber = Request.Form("myInvoiceNumber")
	
	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

			$('#btnEditDeliveryAlertSave').on('click', function(e) {
			
			    //get data-id attribute of the clicked alert
			    var invoiceNum = $("#txtIvsNum").val();
			    var condition = $("#selCondition").val();
			    var emailto = $("#selEmailto").val();
			    var addlemails = $("#txtAdditionalEmails").val();
			    var textto = $("#selTextto").val();
			    var addltexts = $("#txtAdditionalTexts").val();
			    				
				//turn off the automatic page refresh
				//$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			    		    		
		    	$.ajax({
					type:"POST",
					url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
					data: "action=EditAlertFromDeliveryBoardModal&invoiceNum=" + encodeURIComponent(invoiceNum) + "&condition=" + encodeURIComponent(condition) + "&emailto=" + encodeURIComponent(emailto) + "&addlemails=" + encodeURIComponent(addlemails) + "&textto=" + encodeURIComponent(textto) + "&addltexts=" + encodeURIComponent(addltexts),
					success: function(response)
					 {
					 	location.reload();
		             }
				});
	    	});	
	    	
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing alert
	'***************************************************************************************

	SQLDeliveryAlertEdit = "SELECT * FROM SC_Alerts WHERE AlertType = 'DeliveryBoard' AND "
	SQLDeliveryAlertEdit = SQLDeliveryAlertEdit & " CreatedByUserNo = " & Session("UserNo")
	SQLDeliveryAlertEdit = SQLDeliveryAlertEdit & " AND  ReferenceValue = '" & InvoiceNumber & "'"
	
	Set cnnDeliveryAlertEdit = Server.CreateObject("ADODB.Connection")
	cnnDeliveryAlertEdit.open (Session("ClientCnnString"))
	Set rsDeliveryAlertEdit = Server.CreateObject("ADODB.Recordset")
	rsDeliveryAlertEdit.CursorLocation = 3 
	Set rsDeliveryAlertEdit = cnnDeliveryAlertEdit.Execute(SQLDeliveryAlertEdit)
		
	If not rsDeliveryAlertEdit.EOF Then
		Condition = rsDeliveryAlertEdit("Condition")
		Emailto = rsDeliveryAlertEdit("EmailToUserNos") 
		AdditionalEmails = rsDeliveryAlertEdit("AdditionalEmails")
		Textto = rsDeliveryAlertEdit("TextToUserNos")
		AdditionalTexts = rsDeliveryAlertEdit("AdditionalText")
		
		If IsNull(Emailto) OR IsEmpty(Emailto) OR Emailto = "null" Then
			Emailto = "" 
		End If
		If IsNull(Textto) OR IsEmpty(Textto) OR Textto = "null" Then
			Textto = "" 
		End If
		If IsNull(AdditionalEmails) OR IsEmpty(AdditionalEmails) OR AdditionalEmails = "null" Then
			AdditionalEmails = "" 
		End If
		If IsNull(AdditionalTexts) OR IsEmpty(AdditionalTexts) OR AdditionalTexts = "null" Then
			AdditionalTexts = "" 
		End If
		
		ReferenceValue = InvoiceNumber
	End If
	set rsDeliveryAlertEdit = Nothing
	cnnDeliveryAlertEdit.close
	set cnnDeliveryAlertEdit = Nothing
	
	'***************************************************************************************
%>

		<input type="hidden" name="txtIvsNum" id="txtIvsNum" value="<%= InvoiceNumber %>">
		
		<!-- when line !-->
		<div class="row-line">
	
			<!-- when !-->
			<div class="col-lg-2">
				<label>When</label>
			</div>
			<!-- eof when !-->
	
			<!-- condition !-->
			<div class="col-lg-6">
				<select class="form-control" name="selCondition" id="selCondition">
					<option value="Stop is completed or skipped"<% If Condition = "Stop is completed or skipped" Then Response.Write(" selected ") %>>Stop is completed or skipped</option>
					<option value="This becomes the next stop"<% If Condition = "This becomes the next stop" Then Response.Write(" selected ") %>>This becomes the next stop</option>
				</select>
			</div>
			<!-- eof condition !-->
        </div>
        <!-- eof when line !-->


		<!-- email alert line !-->
		<div class="row-line">

			<!-- email alert !-->
			<div class="col-lg-2">
				<label class="right">Send an email alert to</label>
			</div>
			<!-- eof email alert !-->

			<!-- multi select !-->
			<div class="col-lg-4">
				<select class="form-control multiselect" id="selEmailto" name="selEmailto" multiple>
					<option value="0">--- none from here ---</option>
					<% If UserInList(Session("UserNo"),Emailto) = True Then %>
						<option selected value="<%= Session("UserNo") %>"><%= GetUserFirstAndLastNameByUserNo(Session("UserNo")) %></option>
					<% Else %>
						<option value="<%= Session("UserNo") %>"><%= GetUserFirstAndLastNameByUserNo(Session("UserNo")) %></option> 								
					<% End If
			      	'Users dropdown
			      	
					Set cnnDeliveryModal = Server.CreateObject("ADODB.Connection")
					cnnDeliveryModal.open (Session("ClientCnnString"))
					Set rsDeliveryModal = Server.CreateObject("ADODB.Recordset")
							      	
			      	 
		      	  	SQLDeliveryModal = "SELECT UserNo, userFirstName, userLastName FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
		      	  	SQLDeliveryModal = SQLDeliveryModal & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo")
		      	  	SQLDeliveryModal = SQLDeliveryModal & " ORDER BY  userFirstName, userLastName"
	
					rsDeliveryModal.CursorLocation = 3 
					Set rsDeliveryModal = cnnDeliveryModal.Execute(SQLDeliveryModal)
				
					If not rsDeliveryModal.EOF Then
						Do
							FullName = rsDeliveryModal("userFirstName") & " " & rsDeliveryModal("userLastName")
							If UserInList(rsDeliveryModal("UserNo"),Emailto) = True Then
								Response.Write("<option value='" & rsDeliveryModal("UserNo") & "' selected>" & FullName & "</option>")
							Else
								Response.Write("<option value='" & rsDeliveryModal("UserNo") & "'>" & FullName & "</option>")
							End If
							rsDeliveryModal.movenext
						Loop until rsDeliveryModal.eof
					End If
					set rsDeliveryModal = Nothing
					cnnDeliveryModal.close
					set cnnDeliveryModal = Nothing
			      	%>
				</select>
				<strong>Use CTRL and SHIFT to make multiple selections</strong>
            </div>
			<!-- eof multi select !-->
    
            <!-- text area !-->
            <div class="col-lg-6">
				<textarea class="form-control textarea" rows="4" id="txtAdditionalEmails" name="txtAdditionalEmails"><%= AdditionalEmails %></textarea>
	            <strong>Separate multiple email addresses with a semicolon</strong>
            </div>
            <!-- eof text area !-->
        </div>
        <!-- eof email alert line !-->


        <!-- text alert line !-->
        <div class="row-line">

        	<!-- text alert !-->
            <div class="col-lg-2">
	            <label class="right">Send a text alert to</label>
            </div>
            <!-- eof text alert !-->
    
            <!-- multi select !-->
            <div class="col-lg-4">
				<select class="form-control multiselect"  id="selTextto" name="selTextto" multiple>
					<option value="0">--- none from here ---</option>
					
					<% If Textto <> "" Then %>
						<% If UserInList(Session("UserNo"),Textto) = True Then %>
							<option selected value="<%= Session("UserNo") %>"><%= GetUserFirstAndLastNameByUserNo(Session("UserNo")) %></option>
						<% Else %>
							<option value="<%= Session("UserNo") %>"><%= GetUserFirstAndLastNameByUserNo(Session("UserNo")) %></option>								
						<% End If %>
						
					<% Else %>
						<option value="<%= Session("UserNo") %>"><%= GetUserFirstAndLastNameByUserNo(Session("UserNo")) %></option>
					<% End If %>
					
					<%
			      	'Users dropdown

					Set cnnDeliveryModal2 = Server.CreateObject("ADODB.Connection")
					cnnDeliveryModal2.open (Session("ClientCnnString"))
					Set rsDeliveryModal2 = Server.CreateObject("ADODB.Recordset")
			      	 
		      	  	SQLDeliveryModal2 = "SELECT UserNo, userFirstName, userLastName FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
		      	  	SQLDeliveryModal2 = SQLDeliveryModal2 & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo")
		      	  	SQLDeliveryModal2 = SQLDeliveryModal2 & " ORDER BY  userFirstName, userLastName"
		      	  	
		      	  	'Response.write(SQLDeliveryModal2)
	
					rsDeliveryModal2.CursorLocation = 3 
					Set rsDeliveryModal2 = cnnDeliveryModal2.Execute(SQLDeliveryModal2)
				
					If not rsDeliveryModal2.EOF Then
						Do
							FullName = rsDeliveryModal2("userFirstName") & " " & rsDeliveryModal2("userLastName")
							
							If Textto <> "" Then 
								If UserInList(rsDeliveryModal2("UserNo"),Textto) = True Then
									Response.Write("<option value='" & rsDeliveryModal2("UserNo") & "' selected>" & FullName & "</option>")
								Else
									If getUserCellNumber(Session("UserNo")) <> "" Then 
										Response.Write("<option value='" & rsDeliveryModal2("UserNo") & "'>" & FullName & "</option>")
									End If
								End If
							Else
								If getUserCellNumber(Session("UserNo")) <> "" Then 
									Response.Write("<option value='" & rsDeliveryModal2("UserNo") & "'>" & FullName & "</option>")
								End If
							End If
							rsDeliveryModal2.movenext
						Loop until rsDeliveryModal2.eof
					End If
					set rsDeliveryModal2 = Nothing
					cnnDeliveryModal2.close
					set cnnDeliveryModal2 = Nothing
			      	%>
				</select>
				<strong>Use CTRL and SHIFT to make multiple selections</strong>
            </div>
            <!-- eof multi select !-->
    
	     <!-- text area !-->
    	<div class="col-lg-6">
			<textarea class="form-control textarea" rows="4" id="txtAdditionalTexts" name="txtAdditionalTexts"><%= AdditionalTexts %></textarea>
            <strong>Separate multiple phone numbers with a semicolon</strong>
		</div>
        <!-- eof text area !-->
	</div>
	<!-- eof text alert line !-->
	
	<!-- eof modal body !-->
      
	<!-- modal footer !-->
    <div class="modal-footer">
		      
		<div class="col-lg-3">
			&nbsp;
		</div>
			
		<!-- alert !-->
		<div class="col-lg-5 bottom-alert">
			*Delivery alerts are one-time events and automatically delete themselves after the alert has been sent. 
		</div>
		<!-- eof alert !-->
	      
		<!-- close / save !-->
		<div class="col-lg-4">
			<button type="button" class="btn btn-default btn-sm" data-dismiss="modal">Close</button>
			<button type="button" id="btnEditDeliveryAlertSave" class="btn btn-primary btn-sm">Save changes</button>
		</div>
		<!-- eof close / save !-->

	</div>
	<!-- eof modal footer !-->

<%
End Sub



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForDeliveryBoardOptionsModal() 

	InvoiceNumber = Request.Form("invoiceNum")
	CustID = Request.Form("custID")
	CustName = GetCustNameByCustNum(CustID)
	TruckNumber = Request.Form("truckNum")
	DriverUserNo = Trim(GetUserNumberByTruckNumber(TruckNumber))
	ReturnURL = Request.Form("returnPage")
	
	If DriverUserNo <> "*Not Found*" Then
		DriverCellNumber = getUserCellNumberModal(DriverUserNo)
	Else
		DriverCellNumber = ""
	End If
	
	AMDelivery = DeliveryIsAM(InvoiceNumber)
	CurrentlyPriority = DeliveryIsPriority(InvoiceNumber)
	
%>

	<script type="text/javascript">
	
		$(document).ready(function() {
	
			$('#btnTogglePriorityNoTexting').on('click', function(e) {

			    //get data-id attribute of the clicked service ticket
			    var InvoiceNumber = $("#txtInvoiceNumber").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val();
			    
			    //alert("ReturnURL : " & ReturnURL);
			    
				//close the service ticket options modal where we came from
				$('#deliveryBoardInvoiceOptionsModal').modal('hide');
				
				//turn off the automatic page refresh
				$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			    		    		
		    	$.ajax({
					type:"POST",
					//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
					url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
					data: "action=ToggleDeliveryAsPriorityFromDeliveryBoardModal&invoiceNum=" + encodeURIComponent(InvoiceNumber),
					success: function(response)
					 {
						 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
						   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
						 }
						 else {
						   location.reload();
						 }	
		             }
				});
	    	});	
	    	
	    	

			$('#btnRemovePriorityWithTexting').on('click', function(e) {
			    
			    //get data-id attribute of the clicked service ticket
			    var InvoiceNumber = $("#txtInvoiceNumber").val();
			    var ticketCustID = $("#txtCustID").val();
			    var DriverUserNo = $("#txtDriverUserNo").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val();
	    
				//close the service ticket options modal where we came from
				$('#deliveryBoardInvoiceOptionsModal').modal('hide');
								
				
				//open the modal window that asks if you would like to send a text message to the driver
				$("#deliveryBoardMarkAsNotPriorityWithTextingModal").modal('show');
				
				
				//if they want to send a text message to the driver,
				//call the appropriate function that will mark the ticket as urgent and send the text message
				$("#modal-btn-yes-send-text").on("click", function(){
					
					$("#deliveryBoardMarkAsNotPriorityWithTextingModal").modal('hide');
					
			    	$.ajax({
						type:"POST",
						//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
						url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
						data: "action=RemovePriorityFromDeliveryBoardModalAndSendText&returnURL=" + encodeURIComponent(ReturnURL) + "&invoiceNum=" + encodeURIComponent(InvoiceNumber) + "&custID=" + encodeURIComponent(ticketCustID) + "&userNo=" + encodeURIComponent(DriverUserNo),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	
			             }
					});
					
				});

				
				//if they do not want to send a text message to the driver,
				//call the appropriate function that will mark the ticket as urgent only
				$("#modal-btn-no-text").on("click", function(){
				
					$("#deliveryBoardMarkAsNotPriorityWithTextingModal").modal('hide');
					
			    	$.ajax({
						type:"POST",
						//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
						url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
						data: "action=RemovePriorityFromDeliveryBoardModal&invoiceNum=" + encodeURIComponent(InvoiceNumber),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	

			             }
					});
					
				});				
			    		    		
	    	});	
	    	
	


			$('#btnMarkAsPriorityWithTexting').on('click', function(e) {

			    //get data-id attribute of the clicked service ticket
			    var InvoiceNumber = $("#txtInvoiceNumber").val();
			    var ticketCustID = $("#txtCustID").val();
			    var DriverUserNo = $("#txtDriverUserNo").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val();
			 
			    
				//close the service ticket options modal where we came from
				$('#deliveryBoardInvoiceOptionsModal').modal('hide');
								
				//open the modal window that asks if you would like to send a text message to the driver
				$("#deliveryBoardMarkAsPriorityWithTextingModal").modal('show');
				
				
				//if they want to send a text message to the driver,
				//call the appropriate function that will mark the delivery as priority and send the text message
				$("#modal-btn-yes-send-priority-text").on("click", function(){
					
					$("#deliveryBoardMarkAsPriorityWithTextingModal").modal('hide');
					
			    	$.ajax({
						type:"POST",
						//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
						url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
						data: "action=MarkDeliveryAsPriorityFromDeliveryBoardModalAndSendText&returnURL=" + encodeURIComponent(ReturnURL) + "&invoiceNum=" + encodeURIComponent(InvoiceNumber) + "&custID=" + encodeURIComponent(ticketCustID) + "&userNo=" + encodeURIComponent(DriverUserNo),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	

			             }
					});
					
				});

				
				//if they do not want to send a text message to the driver,
				//call the appropriate function that will mark the delivery as priority only
				$("#modal-btn-no-priority-text").on("click", function(){
				
					$("#deliveryBoardMarkAsPriorityWithTextingModal").modal('hide');
					
			    	$.ajax({
						type:"POST",
						//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
						url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
						data: "action=MarkDeliveryAsPriorityFromDeliveryBoardModal&invoiceNum=" + encodeURIComponent(InvoiceNumber),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	

			             }
					});
					
				});				
			    		    		
	    	});	
	    	
	    	
	     	
			$('#deliveryBoardAddAlertModal').on('show.bs.modal', function(e) {
		
				//close the service ticket options modal where we came from
				//$('#deliveryBoardInvoiceOptionsModal').modal('hide');
				
				//turn off the automatic page refresh
				//$('#switchAutomaticRefresh').prop('checked', true).trigger("change");		
			
			    //get data-id attribute of the clicked service ticket
			    var InvoiceNumber = $("#txtInvoiceNumber").val();
			    var ticketCustID = $("#txtCustID").val();
			    var DriverUserNo = $("#txtDriverUserNo").val();
			    var CustName = $("#txtCustName").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val();
	
			    //populate the textbox with the id of the clicked prospect
			    $(e.currentTarget).find('input[name="txtInvoiceNumber"]').val(InvoiceNumber);
			    	    
			    var $modal = $(this);
		
	    		$modal.find('#myDeliveryBoardLabelAdd').html("Create Delivery Alert for " + CustName + " - Invoice #" + InvoiceNumber);
	    		
		    	$.ajax({
					type:"POST",
					//url: "../inc/InSightFuncs_AjaxForRoutingModals.asp",
					url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
					cache: false,
					data: "action=GetContentForAddDeliveryBoardAlertModal&myInvoiceNumber=" + encodeURIComponent(InvoiceNumber),
					success: function(response)
					 {
		               	 $modal.find('#deliveryBoardAddModalContent').html(response);
		             }
		    	});
			    
			});



	 
	    	
			$('#deliveryBoardEditAlertModal').on('show.bs.modal', function(e) {
			
				//close the service ticket options modal where we came from
				//$('#deliveryBoardInvoiceOptionsModal').modal('hide');
						
				//turn off the automatic page refresh
				//$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
							    
			    //get data-id attribute of the clicked service ticket
			    var InvoiceNumber = $("#txtInvoiceNumber").val();
			    var ticketCustID = $("#txtCustID").val();
			    var DriverUserNo = $("#txtDriverUserNo").val();
			    var CustName = $("#txtCustName").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val();
			    
			    $(e.currentTarget).find('input[name="txtInvoiceNumber"]').val(InvoiceNumber);
			    	    
			    var $modal = $(this);
		
	    		$modal.find('#myDeliveryBoardLabelEdit').html("Edit Delivery Alert for " + CustName + " - Invoice #" + InvoiceNumber);
	    		
		    	$.ajax({
					type:"POST",
					//url: "../inc/InSightFuncs_AjaxForRoutingModals.asp",
					url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
					cache: false,
					data: "action=GetContentForEditDeliveryBoardAlertModal&myInvoiceNumber=" + encodeURIComponent(InvoiceNumber),
					success: function(response)
					 {
		               	 $modal.find('#deliveryBoardEditModalContent').html(response);
		             }
		    	});
			    
			});
			
			
   			$("#btnDeleteDeliveryAlert").on("click", function(e){
			
				//close the service ticket options modal where we came from
				$('#deliveryBoardInvoiceOptionsModal').modal('hide');		

				//turn off the automatic page refresh
				//$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
							
			    //get data-id attribute of the clicked service ticket
			    var InvoiceNumber = $("#txtInvoiceNumber").val();
			    var ticketCustID = $("#txtCustID").val();
			    var DriverUserNo = $("#txtDriverUserNo").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val();
			    
			    $(e.currentTarget).find('input[name="txtInvoiceNumber"]').val(InvoiceNumber);
		    	    
		    	$.ajax({
					type:"POST",
					//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
					url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
					data: "action=DeleteAlertFromDeliveryBoardModal&invoiceNum=" + encodeURIComponent(InvoiceNumber),
					success: function(response)
					 {
						 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
						   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
						 }
						 else {
						   location.reload();
						 }	
		             }
				});
		    
			});
		    	
	    	
		    	
			
 	
			$('#btnMarkAsAMDeliveryWithTexting').on('click', function(e) {

			    //get data-id attribute of the clicked service ticket
			    var InvoiceNumber = $("#txtInvoiceNumber").val();
			    var ticketCustID = $("#txtCustID").val();
			    var DriverUserNo = $("#txtDriverUserNo").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val()
			    
				//close the service ticket options modal where we came from
				$('#deliveryBoardInvoiceOptionsModal').modal('hide');
				
				//open the modal window that asks if you would like to send a text message to the driver
				$("#deliveryBoardMarkAsAMDeliveryWithTextingModal").modal('show');
				
				
				//if they want to send a text message to the driver,
				//call the appropriate function that will mark the delivery as priority and send the text message
				$("#modal-btn-yes-send-am-text").on("click", function(){
					
					$("#deliveryBoardMarkAsAMDeliveryWithTextingModal").modal('hide');
					
			    	$.ajax({
						type:"POST",
						//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
						url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
						data: "action=MarkAsAMDeliveryFromDeliveryBoardModalAndSendText&returnURL=" + encodeURIComponent(ReturnURL) + "&invoiceNum=" + encodeURIComponent(InvoiceNumber) + "&custID=" + encodeURIComponent(ticketCustID) + "&userNo=" + encodeURIComponent(DriverUserNo),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	
			             }
					});
					
				});

				
				//if they do not want to send a text message to the driver,
				//call the appropriate function that will mark the delivery as priority only
				$("#modal-btn-no-am-text").on("click", function(){
				
					$("#deliveryBoardMarkAsPriorityWithTextingModal").modal('hide');
					
			    	$.ajax({
						type:"POST",
						//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
						url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
						data: "action=MarkAsAMDeliveryFromDeliveryBoardModal&invoiceNum=" + encodeURIComponent(InvoiceNumber),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	
			             }
					});
					
				});			    	
				
			});	
			
			
			
			
			

			$('#btnRemoveAMDeliveryWithTexting').on('click', function(e) {
			    
			    //get data-id attribute of the clicked service ticket
			    var InvoiceNumber = $("#txtInvoiceNumber").val();
			    var ticketCustID = $("#txtCustID").val();
			    var DriverUserNo = $("#txtDriverUserNo").val();
	    		var ReturnURL = $("#txtReturnURL").val();
	    		var BaseURL = $("#txtBaseURL").val()
	    		
				//close the service ticket options modal where we came from
				$('#deliveryBoardInvoiceOptionsModal').modal('hide');

				//open the modal window that asks if you would like to send a text message to the driver
				$("#deliveryBoardRemoveAMDeliveryWithTextingModal").modal('show');
				
	
			
				//if they want to send a text message to the driver,
				//call the appropriate function that will mark the ticket as urgent and send the text message
				$("#modal-btn-yes-remove-am-send-text").on("click", function(){
					
					$("#deliveryBoardRemoveAMDeliveryWithTextingModal").modal('hide');
					
			    	$.ajax({
						type:"POST",
						//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
						url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
						data: "action=RemoveAMDeliveryFromDeliveryBoardModalAndSendText&returnURL=" + encodeURIComponent(ReturnURL) + "&invoiceNum=" + encodeURIComponent(InvoiceNumber) + "&custID=" + encodeURIComponent(ticketCustID) + "&userNo=" + encodeURIComponent(DriverUserNo),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	
			             }
					});
					
				});

				
				//if they do not want to send a text message to the driver,
				//call the appropriate function that will mark the ticket as urgent only
				$("#modal-btn-no-remove-am-text").on("click", function(){
				
					$("#deliveryBoardRemoveAMDeliveryWithTextingModal").modal('hide');
					
			    	$.ajax({
						type:"POST",
						//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
						url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
						data: "action=RemoveAMDeliveryFromDeliveryBoardModal&invoiceNum=" + encodeURIComponent(InvoiceNumber),
						success: function(response)
						 {
							 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
							   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
							 }
							 else {
							   location.reload();
							 }	
			             }
					});
					
				});				
			    		    		
	    	});	
	    	
			

			$('#btnMarkAsAMDeliveryNoTexting').on('click', function(e) {

			    //get data-id attribute of the clicked service ticket
			    var InvoiceNumber = $("#txtInvoiceNumber").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val();
			    
				//close the service ticket options modal where we came from
				$('#deliveryBoardInvoiceOptionsModal').modal('hide');
				
				//turn off the automatic page refresh
				//$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			    		    		
		    	$.ajax({
					type:"POST",
					//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
					url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
					data: "action=MarkAsAMDeliveryFromDeliveryBoardModal&invoiceNum=" + encodeURIComponent(InvoiceNumber),
					success: function(response)
					 {
						 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
						   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
						 }
						 else {
						   location.reload();
						 }	
		             }
				});
	    	});	



	
			$('#btnRemoveAMDeliveryNoTexting').on('click', function(e) {

			    //get data-id attribute of the clicked service ticket
			    var InvoiceNumber = $("#txtInvoiceNumber").val();
			    var ReturnURL = $("#txtReturnURL").val();
			    var BaseURL = $("#txtBaseURL").val();
			    
				//close the service ticket options modal where we came from
				$('#deliveryBoardInvoiceOptionsModal').modal('hide');
				
				//turn off the automatic page refresh
				// $('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			    		    		
		    	$.ajax({
					type:"POST",
					//url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
					url: BaseURL + "inc/InSightFuncs_AjaxForRoutingModals.asp",
					data: "action=RemoveAMDeliveryFromDeliveryBoardModal&invoiceNum=" + encodeURIComponent(InvoiceNumber),
					success: function(response)
					 {
						 if (ReturnURL.indexOf("deliveryboardkiosknopaging") >=0) {
						   	window.location = BaseURL + "directLaunch/kiosks/routing/deliveryboardkiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
						 }
						 else {
						   location.reload();
						 }	
		             }
				});
	    	});	
	    	
	    	    		
		});
	</script>

	<input type="hidden" name="txtInvoiceNumber" id="txtInvoiceNumber" value="<%= InvoiceNumber %>">
	<input type="hidden" name="txtTruckNumber" id="txtTruckNumber" value="<%= TruckNumber %>">
	<input type="hidden" name="txtCustID" id="txtCustID" value="<%= CustID %>">
	<input type="hidden" name="txtCustName" id="txtCustName" value="<%= CustName %>">
	<input type="hidden" name="txtDriverUserNo" id="txtDriverUserNo" value="<%= DriverUserNo %>">
	<input type="hidden" name="txtReturnURL" id="txtReturnURL" value="<%= ReturnURL %>">
	<input type="hidden" name="txtBaseURL" id="txtBaseURL" value="<%= BaseURL %>">
	 
	<div class="row">
		<div class="center">
		
		<!--
		Driver User No: <%= DriverUserNo %><br>
		DriverCellNumber: <%= DriverCellNumber %><br>TruckNumber
		TruckNumber: <%= TruckNumber %><br>
		CustID: <%= CustID %><br>
		Return URL: <%= ReturnURL %><br>-->
		
			<% If DriverUserNo <> "*Not Found*" AND DriverCellNumber <> "" Then %>
					
				<div class="col-lg-11">
					<% If CurrentlyPriority = true Then %>
						<a class="btn btn-danger btn-lg btn-block btn-huge" id="btnRemovePriorityWithTexting" data-show="true" href="#" data-invoice-number="<%= InvoiceNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfDriver %>" data-target="#deliveryBoardMarkAsNotPriorityWithTextingModal" data-tooltip="true" data-title="Mark As Urgent" style="cursor:pointer;">Remove Priority Status</a>
					<% Else %>
						<a class="btn btn-danger btn-lg btn-block btn-huge" id="btnMarkAsPriorityWithTexting" data-show="true" href="#" data-invoice-number="<%= InvoiceNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfDriver %>" data-target="#deliveryBoardMarkAsPriorityWithTextingModal" data-tooltip="true" data-title="Mark As Urgent" style="cursor:pointer;">Mark Delivery As Priority</a>
					<% End If %>
				</div>
					
			<% Else	%>
			
				<div class="col-lg-11">
					<% If CurrentlyPriority = true Then %>
						<a href="#" class="btn btn-danger btn-lg btn-block btn-huge" id="btnTogglePriorityNoTexting">Remove Priority Status</a>
					<% Else %>
						<a href="#" class="btn btn-danger btn-lg btn-block btn-huge" id="btnTogglePriorityNoTexting">Mark Delivery As Priority</a>
					<% End If %>
				</div>
					
			<% End If %>
			
	
			<% If Session("UserNo") <> "" AND Session("UserNo") <> "0" Then %>
				<% If DeliveryAlertSet(InvoiceNumber,Session("UserNo")) Then %>	
			        <div class="col-lg-11">
			            <a class="btn btn-info btn-lg btn-block btn-huge" id="btnEditDeliveryAlert" data-show="true" data-toggle="modal" href="#deliveryBoardEditAlertModal" data-invoice-number="<%= InvoiceNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-target="#deliveryBoardEditAlertModal" data-tooltip="true" data-title="Edit Delivery Alert" style="cursor:pointer;">Edit Delivery Alert For This Ticket</a>
			        </div>  
				        
			        <div class="col-lg-11">
			            <a class="btn btn-warning btn-lg btn-block btn-huge" id="btnDeleteDeliveryAlert" data-show="true" href="#" data-invoice-number="<%= InvoiceNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-target="#deliveryBoardDeleteAlertModal" data-tooltip="true" data-title="Delete Delivery Alert" style="cursor:pointer;">Delete Delivery Alert For This Ticket</a>
			        </div>  
		        <% Else %>
			        <div class="col-lg-11">
			            <a class="btn btn-warning btn-lg btn-block btn-huge" id="btnAddDeliveryAlert" data-show="true" data-toggle="modal" href="#deliveryBoardAddAlertModal" data-invoice-number="<%= InvoiceNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-target="#deliveryBoardAddAlertModal" data-tooltip="true" data-title="Set Delivery Alert" style="cursor:pointer;">Create a Delivery Alert For This Ticket</a>
			        </div>
		        <% End If %>
	         <% End If %>
	        	        
	        
			<% If DriverUserNo <> "*Not Found*" AND DriverCellNumber <> "" Then %>
					
				<div class="col-lg-11">
					<% If AMDelivery = true Then %>
						<a class="btn btn-success btn-lg btn-block btn-huge" id="btnRemoveAMDeliveryWithTexting" data-show="true" href="#" data-invoice-number="<%= InvoiceNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfDriver %>" data-target="#deliveryBoardRemoveAMDeliverytWithTextingModal" data-tooltip="true" data-title="AM Delivery" style="cursor:pointer;">Remove AM Delivery Status</a>
					<% Else %>
						<a class="btn btn-success btn-lg btn-block btn-huge" id="btnMarkAsAMDeliveryWithTexting" data-show="true" href="#" data-invoice-number="<%= InvoiceNumber %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfDriver %>" data-target="#deliveryBoardMarkAsAMDeliverytWithTextingModal" data-tooltip="true" data-title="AM Delivery" style="cursor:pointer;">Mark Delivery As AM Delivery</a>
					<% End If %>
				</div>
					
			<% Else	%>
			
				<div class="col-lg-11">
					<% If AMDelivery = true Then %>
						<a href="#" class="btn btn-success btn-lg btn-block btn-huge" id="btnRemoveAMDeliveryNoTexting">Remove AM Delivery Status</a>
					<% Else %>
						<a href="#" class="btn btn-success btn-lg btn-block btn-huge" id="btnMarkAsAMDeliveryNoTexting">Mark Delivery As AM Delivery</a>
					<% End If %>
				</div>
					
			<% End If %>
	        
	        
	     </div>					        
	</div>	
	

<%
End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub ToggleDeliveryAsPriorityFromDeliveryBoardModal() 

	invoiceNum = Request.Form("invoiceNum")

	SQLTogglePriority = "SELECT * FROM RT_DeliveryBoard WHERE ivsNum = " & invoiceNum 
	
	Set cnnTogglePriority = Server.CreateObject("ADODB.Connection")
	cnnTogglePriority.open (Session("ClientCnnString"))
	
	Set rsTogglePriority = Server.CreateObject("ADODB.Recordset")
	rsTogglePriority.CursorLocation = 3 
	Set rsTogglePriority = cnnTogglePriority.Execute(SQLTogglePriority)
	
	If Not rsTogglePriority.Eof Then
		'All records should be the same so just check the 1st one
		If rsTogglePriority("Priority") <> 1 Then 
			Priority = 1
			CreateAuditLogEntry "Delivery Board Invoice Priority Changed","Delivery Board Invoice Priority Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - marked as PRIORITY"
		Else
			Priority = 0
			CreateAuditLogEntry "Delivery Board Invoice Priority Changed","Delivery Board Invoice Priority Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - no longer marked as PRIORITY"		
		End If	
		SQL = "UPDATE RT_DeliveryBoard Set Priority = " & Priority & " WHERE ivsNum = " & invoiceNum
		Set rsTogglePriority = cnnTogglePriority.Execute(SQL)
	
	Else
		'There were no details found which means:
		'1. Advanced dispatch is not on
		'2. The ticket hasn't been dispatched yet
		'So we mark the header instead
		
		SQL = "SELECT * FROM RT_DeliveryBoard WHERE invoiceNumber = " & invoiceNum 
		Set rsTogglePriority = cnnTogglePriority.Execute(SQL)
		If Not rsTogglePriority.Eof Then
			'All records should be the same so just check the 1st one
			If rsTogglePriority("Priority") <> 1 Then 
				Priority = 1
				CreateAuditLogEntry "Delivery Board Invoice Priority Changed","Delivery Board Invoice Priority Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - marked as PRIORITY"
			Else
				Priority = 0
				CreateAuditLogEntry "Delivery Board Invoice Priority Changed","Delivery Board Invoice Priority Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - no longer marked as PRIORITY"		
			End If	
			SQL = "UPDATE RT_DeliveryBoard Set Priority = " & Priority & " WHERE ivsNum = " & invoiceNum
			Set rsTogglePriority = cnnTogglePriority.Execute(SQL)
		End If
	End If
	
	set rsTogglePriority = Nothing
	cnnTogglePriority.Close
	Set cnnTogglePriority = Nothing
	

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************





'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub MarkDeliveryAsPriorityFromDeliveryBoardModalAndSendText() 

	invoiceNum = Request.Form("invoiceNum")
	CustID = Request.Form("custID")
	DriverUserNo = Request.Form("userNo")
	ReturnURL = Request.Form("returnURL")

	SQLMarkDeliveryAsPriority = "UPDATE RT_DeliveryBoard SET Priority = 1 WHERE ivsNum = " & invoiceNum 
	Set cnnMarkDeliveryAsPriority = Server.CreateObject("ADODB.Connection")
	cnnMarkDeliveryAsPriority.open (Session("ClientCnnString"))
	
	Set rsMarkDeliveryAsPriority = Server.CreateObject("ADODB.Recordset")
	rsMarkDeliveryAsPriority.CursorLocation = 3 
	Set rsMarkDeliveryAsPriority = cnnMarkDeliveryAsPriority.Execute(SQLMarkDeliveryAsPriority)

	CreateAuditLogEntry "Delivery Board Invoice Priority Changed","Delivery Board Invoice Priority Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - marked as PRIORITY"
		
	set rsMarkDeliveryAsPriority = Nothing
	cnnMarkDeliveryAsPriority.Close
	Set cnnMarkDeliveryAsPriority = Nothing
	

	'**********************
	'Send text 
	'**********************
	

	If getUserCellNumber(DriverUserNo) <> "" Then
	
		Send_To = getUserCellNumber(DriverUserNo)

		URL = BaseURL & "inc/sendtext.php"
		QString = "?n=" & Replace(getUserCellNumber(DriverUserNo),"-","")
			
		QString = QString & "&u1=" & EzTextingUserID()
		QString = QString & "&u2=" & EzTextingPassword()
		
		QString = QString & "&t=NOTIFICATION"
		
		If InStr(ReturnURL, "DeliveryBoardKioskNoPaging") Then
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & ReturnURL & "?pp=" & Session("PassPhrase") & "&cl=" & Session("ClientKey") & "&ri=" & Session("RefreshInterval"))
		Else
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & ReturnURL)
		End If	

		'Text message should alert them that the delivery is a priority
		If GetCustNameByCustNum(CustID) <> "" Then
			txtMSG = "The delivery for " & EZTexting_Filter1(Replace(GetCustNameByCustNum(CustID),"&"," ")) & " (Invoice " & invoiceNum & ") has been marked as a PRIORITY delivery. "
		Else
			txtMSG = "The delivery for " & GetTerm("Account") & ": " & CustID & " (Invoice " & invoiceNum & ") has been marked as a PRIORITY delivery. "
		End If
		
		QString = QString & "&m=" & txtMSG 

		QString = QString & "    Tap the link to see the details for this delivery    "

		QString = QString & Server.URLEncode(baseURL & "directlaunch/routing/moreinfo_statuschange_from_email_or_text.asp?i=" & invoiceNum & "&u=" & DriverUserNo & "&c=" & CustID & "&cl=" & MUV_READ("SERNO"))

		QString = QString &  "&cty=" & GetCompanyCountry()
		QString = Replace(Qstring," ", "%20")

		Response.Redirect (URL & Qstring)

		Description = "A priority delivery text message was sent to " & GetUserDisplayNameByUserNo(DriverUserNo) & " (" & getUserCellNumber(DriverUserNo) & ") at " & NOW()
		CreateAuditLogEntry "Delivery Board Routing System","Priority delivery text message sent","Minor",0,Description
		
	Else
	
		' Could not send dispatch test, no address on file
		emailBody = "Insight was unable to send an priority delivery text message to " & GetUserDisplayNameByUserNo(DriverUserNo) & ". No cell number on file"
		If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com"
		SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send priority delivery text message",emailBody,GetTerm("Service"),"Missing Cell Number"
		Description = "Insight was unable to send a priority delivery message to " & GetUserDisplayNameByUserNo(DriverUserNo) & ". No cell number on file"
		CreateAuditLogEntry "Delivery Board Routing System","Unable to send priority delivery text message","Major",0,Description

	End If

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub MarkDeliveryAsPriorityFromDeliveryBoardModal() 

	invoiceNum = Request.Form("invoiceNum")
	CustID = Request.Form("custID")
	DriverUserNo = Request.Form("userNo")

	SQLMarkDeliveryAsPriority = "UPDATE RT_DeliveryBoard SET Priority = 1 WHERE ivsNum = " & invoiceNum 
	Set cnnMarkDeliveryAsPriority = Server.CreateObject("ADODB.Connection")
	cnnMarkDeliveryAsPriority.open (Session("ClientCnnString"))
	
	Set rsMarkDeliveryAsPriority = Server.CreateObject("ADODB.Recordset")
	rsMarkDeliveryAsPriority.CursorLocation = 3 
	Set rsMarkDeliveryAsPriority = cnnMarkDeliveryAsPriority.Execute(SQLMarkDeliveryAsPriority)

	CreateAuditLogEntry "Delivery Board Invoice Priority Changed","Delivery Board Invoice Priority Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - marked as PRIORITY"
		
	set rsMarkDeliveryAsPriority = Nothing
	cnnMarkDeliveryAsPriority.Close
	Set cnnMarkDeliveryAsPriority = Nothing

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub RemovePriorityFromDeliveryBoardModalAndSendText() 

	invoiceNum = Request.Form("invoiceNum")
	CustID = Request.Form("custID")
	DriverUserNo = Request.Form("userNo")
	ReturnURL = Request.Form("returnURL")

	SQLRemoveDeliveryAsPriority = "UPDATE RT_DeliveryBoard SET Priority = 0 WHERE ivsNum = " & invoiceNum 
	Set cnnRemoveDeliveryAsPriority = Server.CreateObject("ADODB.Connection")
	cnnRemoveDeliveryAsPriority.open (Session("ClientCnnString"))
	
	Set rsRemoveDeliveryAsPriority = Server.CreateObject("ADODB.Recordset")
	rsRemoveDeliveryAsPriority.CursorLocation = 3 
	Set rsRemoveDeliveryAsPriority = cnnRemoveDeliveryAsPriority.Execute(SQLRemoveDeliveryAsPriority)

	CreateAuditLogEntry "Delivery Board Invoice Priority Changed","Delivery Board Invoice Priority Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - no longer marked as PRIORITY"
		
	set rsRemoveDeliveryAsPriority = Nothing
	cnnRemoveDeliveryAsPriority.Close
	Set cnnRemoveDeliveryAsPriority = Nothing
	

	'**********************
	'Send text 
	'**********************
	

	If getUserCellNumber(DriverUserNo) <> "" Then
	
		Send_To = getUserCellNumber(DriverUserNo)

		URL = BaseURL & "inc/sendtext.php"
		QString = "?n=" & Replace(getUserCellNumber(DriverUserNo),"-","")
			
		QString = QString & "&u1=" & EzTextingUserID()
		QString = QString & "&u2=" & EzTextingPassword()
		
		QString = QString & "&t=NOTIFICATION"
		
		If InStr(ReturnURL, "DeliveryBoardKioskNoPaging") Then
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & ReturnURL & "?pp=" & Session("PassPhrase") & "&cl=" & Session("ClientKey") & "&ri=" & Session("RefreshInterval"))
		Else
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & ReturnURL)
		End If	

		'Text message should alert them that delivery is no longer a priority
		If GetCustNameByCustNum(CustID) <> "" Then
			txtMSG = "The delivery for " & EZTexting_Filter1(Replace(GetCustNameByCustNum(CustID),"&"," ")) & " (Invoice " & invoiceNum & ") is NO LONGER a priority delivery. "
		Else
			txtMSG = "The delivery for " & GetTerm("Account") & ": " & CustID & " (Invoice " & invoiceNum & ") is NO LONGER a priority delivery. "
		End If
		
		QString = QString & "&m=" & txtMSG 

		QString = QString & "    Tap the link to see the details for this delivery    "

		QString = QString & Server.URLEncode(baseURL & "directlaunch/routing/moreinfo_statuschange_from_email_or_text.asp?i=" & invoiceNum & "&u=" & DriverUserNo & "&c=" & CustID & "&cl=" & MUV_READ("SERNO"))

		QString = QString &  "&cty=" & GetCompanyCountry()
		QString = Replace(Qstring," ", "%20")

		Response.Redirect (URL & Qstring)

		Description = "A no longer a priority delivery text message was sent to " & GetUserDisplayNameByUserNo(DriverUserNo) & " (" & getUserCellNumber(DriverUserNo) & ") at " & NOW()
		CreateAuditLogEntry "Delivery Board Routing System","No longer a priority delivery text message sent","Minor",0,Description
		
	Else
			
		' Could not send dispatch test, no address on file
		emailBody = "Insight was unable to send a no longer a priority delivery text message to " & GetUserDisplayNameByUserNo(DriverUserNo) & ". No cell number on file"
		If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com"
		SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send no longer a priority delivery text message",emailBody,GetTerm("Service"),"Missing Cell Number"
		Description = "Insight was unable to send no longer a priority delivery message to " & GetUserDisplayNameByUserNo(DriverUserNo) & ". No cell number on file"
		CreateAuditLogEntry "Delivery Board Routing System","Unable to send no longer a priority delivery text message","Major",0,Description
		

	End If

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub RemovePriorityFromDeliveryBoardModal() 

	invoiceNum = Request.Form("invoiceNum")
	CustID = Request.Form("custID")
	DriverUserNo = Request.Form("userNo")

	SQLRemoveDeliveryAsPriority = "UPDATE RT_DeliveryBoard SET Priority = 0 WHERE ivsNum = " & invoiceNum 
	Set cnnRemoveDeliveryAsPriority = Server.CreateObject("ADODB.Connection")
	cnnRemoveDeliveryAsPriority.open (Session("ClientCnnString"))
	
	Set rsRemoveDeliveryAsPriority = Server.CreateObject("ADODB.Recordset")
	rsRemoveDeliveryAsPriority.CursorLocation = 3 
	Set rsRemoveDeliveryAsPriority = cnnRemoveDeliveryAsPriority.Execute(SQLRemoveDeliveryAsPriority)

	CreateAuditLogEntry "Delivery Board Invoice Priority Changed","Delivery Board Invoice Priority Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - no longer marked as PRIORITY"
		
	set rsRemoveDeliveryAsPriority = Nothing
	cnnRemoveDeliveryAsPriority.Close
	Set cnnRemoveDeliveryAsPriority = Nothing

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub AddAlertFromDeliveryBoardModal() 

	AlertName = "Delivery Board - Invoice #: " & invoiceNum & " User: " & GetUserDisplayNameByUserNo(Session("UserNo"))
	Enabled = vbTrue
	
	invoiceNum = Request.Form("invoiceNum")
	Condition = Request.Form("condition")
	Emailto = Request.Form("emailto") 
	AdditionalEmails = Request.Form("addlemails")
	Textto = Request.Form("textto")
	AdditionalTexts = Request.Form("addltexts")
		
	ReferenceValue = invoiceNum
	
	If AdditionalEmails <> "" Then
		AdditionalEmails = Trim(AdditionalEmails)
		AdditionalEmails = Replace(AdditionalEmails,",",";") ' Common for the user to type , instead of ; So we fix it
		If Right(AdditionalEmails,1)=";" Then AdditionalEmails = Left(AdditionalEmails,Len(AdditionalEmails)-1)
	End If
	
	If AdditionalTexts <> "" Then
		AdditionalTexts = Trim(AdditionalTexts)
		AdditionalTexts = Replace(AdditionalTexts,",",";") ' Common for the user to type , instead of ; So we fix it
		If Right(AdditionalTexts,1)=";" Then AdditionalTexts = Left(AdditionalTexts,Len(AdditionalTexts)-1)
	End If
	
	SQLAddDeliveryAlert = "INSERT INTO SC_Alerts (AlertType,AlertName,Condition,EmailToUserNos, "
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & "AdditionalEmails,Enabled ,TextToUserNos,AdditionalText,ReferenceValue,CreatedByUserNo,ReferenceField)"
	SQLAddDeliveryAlert = SQLAddDeliveryAlert &  " VALUES (" 
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & "'DeliveryBoard'"
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & ",'" & AlertName & "'"
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & ",'" & Condition & "'"
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & ",'" & Emailto & "'"
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & ",'" & AdditionalEmails & "'"
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & ","  & Enabled 
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & ",'" & Textto & "'"	
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & ",'" & AdditionalTexts & "'"	
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & ",'" & ReferenceValue & "'"
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & "," & Session("UserNo") 
	SQLAddDeliveryAlert = SQLAddDeliveryAlert & ",'Invoice')"	
	
		
	Set cnnAddDeliveryAlert = Server.CreateObject("ADODB.Connection")
	cnnAddDeliveryAlert.open (Session("ClientCnnString"))
	
	Set rsAddDeliveryAlert = Server.CreateObject("ADODB.Recordset")
	rsAddDeliveryAlert.CursorLocation = 3 
	Set rsAddDeliveryAlert = cnnAddDeliveryAlert.Execute(SQLAddDeliveryAlert)
	set rsAddDeliveryAlert = Nothing
	
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the delivery board alert: " & AlertName
	CreateAuditLogEntry "Delivery Board Alert Added","Delivery Board Alert Added","Major",0,Description

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub EditAlertFromDeliveryBoardModal() 

	invoiceNum = Request.Form("invoiceNum")
	Condition = Request.Form("condition")
	Emailto = Request.Form("emailto") 
	AdditionalEmails = Request.Form("addlemails")
	Textto = Request.Form("textto")
	AdditionalTexts = Request.Form("addltexts")
		
	ReferenceValue = invoiceNum
	
	AlertName = "Delivery Board - Invoice #: " & invoiceNum & " User: " & GetUserDisplayNameByUserNo(Session("UserNo"))
	
	'*******************************************************************
	'Lookup the record as it exists now so we can fillin the audit trail
	
	SQL = "SELECT * FROM SC_Alerts WHERE AlertType = 'DeliveryBoard' AND "
	SQL = SQL & " CreatedByUserNo = " & Session("UserNo")
	SQL = SQL & " AND  ReferenceValue = '" & Request.Form("txtIvsNum") & "'"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		Orig_Condition = rs("Condition")
		Orig_Emailto = rs("EmailToUserNos") 
		Orig_AdditionalEmails = rs("AdditionalEmails")
		Orig_Textto = rs("TextToUserNos")
		Orig_AdditionalTexts = rs("AdditionalText")
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	'***********************************************************************
	'End Lookup the record as it exists now so we can fillin the audit trail
	'***********************************************************************
	
	SQLEditDeliveryAlert = "UPDATE SC_Alerts SET "
	SQLEditDeliveryAlert = SQLEditDeliveryAlert &  "Condition = '" & Condition & "',"
	SQLEditDeliveryAlert = SQLEditDeliveryAlert &  "EmailToUserNos = '" & Emailto & "',"
	SQLEditDeliveryAlert = SQLEditDeliveryAlert &  "AdditionalEmails = '" & AdditionalEmails & "',"
	SQLEditDeliveryAlert = SQLEditDeliveryAlert &  "TextToUserNos = '" & Textto & "',"
	SQLEditDeliveryAlert = SQLEditDeliveryAlert &  "AdditionalText = '" & AdditionalTexts & "'"
	SQLEditDeliveryAlert = SQLEditDeliveryAlert &  " WHERE AlertType = 'DeliveryBoard' AND "
	SQLEditDeliveryAlert = SQLEditDeliveryAlert & " CreatedByUserNo = " & Session("UserNo")
	SQLEditDeliveryAlert = SQLEditDeliveryAlert & " AND  ReferenceValue = '" & ReferenceValue  & "'"
		
	Set cnnEditDeliveryAlert = Server.CreateObject("ADODB.Connection")
	cnnEditDeliveryAlert.open (Session("ClientCnnString"))
	
	Set rsEditDeliveryAlert = Server.CreateObject("ADODB.Recordset")
	rsEditDeliveryAlert.CursorLocation = 3 
	Set rsEditDeliveryAlert = cnnEditDeliveryAlert.Execute(SQLEditDeliveryAlert)
	set rsEditDeliveryAlert = Nothing
	
	Description = ""
	If Orig_Condition <> Condition Then
		Description = Description & "  Delivery Board Alert trigger changed from " & Orig_Condition & " to " & Condition
	End If
	If Orig_Emailto <> Emailto Then
		Description = Description & "  Users to send email to changed from " & Orig_Emailto & " to " & Emailto
	End If
	If Orig_AdditionalEmails <> AdditionalEmails Then
		Description = Description & "  Additional emails changed from " & Orig_AdditionalEmails & " to " & AdditionalEmails
	End If
	If Orig_Textto <> Textto Then
		Description = Description & "  Users to send texts to changed from " & Orig_Orig_Textto& " to " & Textto
	End If
	If Orig_AdditionalTexts <> AdditionalTexts Then
		Description = Description & "  Additional text messages changed from " & Orig_AdditionalTexts & " to " & AdditionalTexts
	End If
	
	CreateAuditLogEntry "Delivery Board Alert Edited","Delivery Board Alert Edited","Major",0,Description

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub DeleteAlertFromDeliveryBoardModal() 

	invoiceNum = Request.Form("invoiceNum")
	
	AlertName = "Delivery Board - Invoice #: " & invoiceNum & " User: " & GetUserDisplayNameByUserNo(Session("UserNo"))

	SQLDeleteDeliveryAlert = "DELETE FROM SC_Alerts "
	SQLDeleteDeliveryAlert = SQLDeleteDeliveryAlert &  " WHERE AlertType = 'DeliveryBoard' AND "
	SQLDeleteDeliveryAlert = SQLDeleteDeliveryAlert & " CreatedByUserNo = " & Session("UserNo")
	SQLDeleteDeliveryAlert = SQLDeleteDeliveryAlert & " AND  ReferenceValue = '" & invoiceNum & "'"
	
	Set cnnDeleteDeliveryAlert = Server.CreateObject("ADODB.Connection")
	cnnDeleteDeliveryAlert.open (Session("ClientCnnString"))
	
	Set rsDeleteDeliveryAlert = Server.CreateObject("ADODB.Recordset")
	rsDeleteDeliveryAlert.CursorLocation = 3 
	Set rsDeleteDeliveryAlert = cnnDeleteDeliveryAlert.Execute(SQLDeleteDeliveryAlert)
	set rsDeleteDeliveryAlert = Nothing
	
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " deleted the delivery board alert: " & AlertName
	CreateAuditLogEntry "Alert Deleted","Alert Deleted","Major",0,Description

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************






'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub MarkAsAMDeliveryFromDeliveryBoardModalAndSendText() 

	invoiceNum = Request.Form("invoiceNum")
	CustID = Request.Form("custID")
	DriverUserNo = Request.Form("userNo")
	ReturnURL = Request.Form("returnURL")

	SQLMarkDeliveryAsAM = "UPDATE RT_DeliveryBoard SET AMorPM = 'AM' WHERE ivsNum = " & invoiceNum 
	Set cnnMarkDeliveryAsAM = Server.CreateObject("ADODB.Connection")
	cnnMarkDeliveryAsAM.open (Session("ClientCnnString"))
	
	Set rsMarkDeliveryAsAM = Server.CreateObject("ADODB.Recordset")
	rsMarkDeliveryAsAM.CursorLocation = 3 
	Set rsMarkDeliveryAsAM = cnnMarkDeliveryAsAM.Execute(SQLMarkDeliveryAsAM)

	CreateAuditLogEntry "Delivery Board Invoice AM/PM Changed","Delivery Board Invoice AM/PM Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - marked as an AM Delivery"
		
	set rsMarkDeliveryAsAM = Nothing
	cnnMarkDeliveryAsAM.Close
	Set cnnMarkDeliveryAsAM = Nothing
	

	'**********************
	'Send text 
	'**********************
	

	If getUserCellNumber(DriverUserNo) <> "" Then
	
		Send_To = getUserCellNumber(DriverUserNo)

		URL = BaseURL & "inc/sendtext.php"
		QString = "?n=" & Replace(getUserCellNumber(DriverUserNo),"-","")
			
		QString = QString & "&u1=" & EzTextingUserID()
		QString = QString & "&u2=" & EzTextingPassword()
		
		QString = QString & "&t=NOTIFICATION"
		
		If InStr(ReturnURL, "DeliveryBoardKioskNoPaging") Then
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & ReturnURL & "?pp=" & Session("PassPhrase") & "&cl=" & Session("ClientKey") & "&ri=" & Session("RefreshInterval"))
		Else
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & ReturnURL)
		End If	
		
		'Text message should alert them that the delivery is an AM delivery
		If GetCustNameByCustNum(CustID) <> "" Then
			txtMSG = "The delivery for " & EZTexting_Filter1(Replace(GetCustNameByCustNum(CustID),"&"," ")) & " (Invoice " & invoiceNum & ") has been marked as an AM delivery. "
		Else
			txtMSG = "The delivery for " & GetTerm("Account") & ": " & CustID & " (Invoice " & invoiceNum & ") has been marked as an AM delivery. "
		End If
		
		QString = QString & "&m=" & txtMSG 

		QString = QString & "    Tap the link to see the details for this delivery    "

		QString = QString & Server.URLEncode(baseURL & "directlaunch/routing/moreinfo_statuschange_from_email_or_text.asp?i=" & invoiceNum & "&u=" & DriverUserNo & "&c=" & CustID & "&cl=" & MUV_READ("SERNO"))

		QString = QString &  "&cty=" & GetCompanyCountry()
		QString = Replace(Qstring," ", "%20")

		Response.Redirect (URL & Qstring)

		Description = "An AM delivery text message was sent to " & GetUserDisplayNameByUserNo(DriverUserNo) & " (" & getUserCellNumber(DriverUserNo) & ") at " & NOW()
		CreateAuditLogEntry "Delivery Board Routing System","AM delivery text message sent","Minor",0,Description
		
	Else
	
		' Could not send dispatch test, no address on file
		emailBody = "Insight was unable to send an AM delivery text message to " & GetUserDisplayNameByUserNo(DriverUserNo) & ". No cell number on file"
		If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com"
		SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send AM delivery text message",emailBody,GetTerm("Service"),"Missing Cell Number"
		Description = "Insight was unable to send an AM delivery message to " & GetUserDisplayNameByUserNo(DriverUserNo) & ". No cell number on file"
		CreateAuditLogEntry "Delivery Board Routing System","Unable to send AM delivery text message","Major",0,Description

	End If

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub MarkAsAMDeliveryFromDeliveryBoardModal() 

	invoiceNum = Request.Form("invoiceNum")
	CustID = Request.Form("custID")
	DriverUserNo = Request.Form("userNo")

	SQLMarkDeliveryAsAM = "UPDATE RT_DeliveryBoard SET AMorPM = 'AM' WHERE ivsNum = " & invoiceNum 
	Set cnnMarkDeliveryAsAM = Server.CreateObject("ADODB.Connection")
	cnnMarkDeliveryAsAM.open (Session("ClientCnnString"))
	
	Set rsMarkDeliveryAsAM = Server.CreateObject("ADODB.Recordset")
	rsMarkDeliveryAsAM.CursorLocation = 3 
	Set rsMarkDeliveryAsAM = cnnMarkDeliveryAsAM.Execute(SQLMarkDeliveryAsAM)

	CreateAuditLogEntry "Delivery Board Invoice AM/PM Changed","Delivery Board Invoice AM/PM Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - marked as AM Delivery"
		
	set rsMarkDeliveryAsAM = Nothing
	cnnMarkDeliveryAsAM.Close
	Set cnnMarkDeliveryAsAM = Nothing

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub RemoveAMDeliveryFromDeliveryBoardModalAndSendText() 

	invoiceNum = Request.Form("invoiceNum")
	CustID = Request.Form("custID")
	DriverUserNo = Request.Form("userNo")
	ReturnURL = Request.Form("returnURL")

	SQLRemoveDeliveryAsAM = "UPDATE RT_DeliveryBoard SET AMorPM = '' WHERE ivsNum = " & invoiceNum 
	Set cnnRemoveDeliveryAsAM = Server.CreateObject("ADODB.Connection")
	cnnRemoveDeliveryAsAM.open (Session("ClientCnnString"))
	
	Set rsRemoveDeliveryAsAM = Server.CreateObject("ADODB.Recordset")
	rsRemoveDeliveryAsAM.CursorLocation = 3 
	Set rsRemoveDeliveryAsAM = cnnRemoveDeliveryAsAM.Execute(SQLRemoveDeliveryAsAM)

	CreateAuditLogEntry "Delivery Board Invoice AM/PM Changed","Delivery Board Invoice AM/PM Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - no longer marked as AM Delivery"
		
	set rsRemoveDeliveryAsAM = Nothing
	cnnRemoveDeliveryAsAM.Close
	Set cnnRemoveDeliveryAsAM = Nothing
	

	'**********************
	'Send text 
	'**********************
	

	If getUserCellNumber(DriverUserNo) <> "" Then
	
		Send_To = getUserCellNumber(DriverUserNo)

		URL = BaseURL & "inc/sendtext.php"
		
		QString = "?n=" & Replace(getUserCellNumber(DriverUserNo),"-","")
			
		QString = QString & "&u1=" & EzTextingUserID()
		QString = QString & "&u2=" & EzTextingPassword()
		
		QString = QString & "&t=NOTIFICATION"
		
		If InStr(ReturnURL, "DeliveryBoardKioskNoPaging") Then
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & ReturnURL & "?pp=" & Session("PassPhrase") & "&cl=" & Session("ClientKey") & "&ri=" & Session("RefreshInterval"))
		Else
			QString = QString & "&R=Location: " & Server.URLEncode(BaseURL & ReturnURL)
		End If	
		
		'Text message should alert them that delivery is no longer an AM Delivery
		If GetCustNameByCustNum(CustID) <> "" Then
			txtMSG = "The delivery for " & EZTexting_Filter1(Replace(GetCustNameByCustNum(CustID),"&"," ")) & " (Invoice " & invoiceNum & ") is NO LONGER an AM delivery. "
		Else
			txtMSG = "The delivery for " & GetTerm("Account") & ": " & CustID & " (Invoice " & invoiceNum & ") is NO LONGER an AM delivery. "
		End If
		
		
		QString = QString & "&m=" & txtMSG 

		QString = QString & "    Tap the link to see the details for this delivery    "

		QString = QString & Server.URLEncode(baseURL & "directlaunch/routing/moreinfo_statuschange_from_email_or_text.asp?i=" & invoiceNum & "&u=" & DriverUserNo & "&c=" & CustID & "&cl=" & MUV_READ("SERNO"))

		QString = QString &  "&cty=" & GetCompanyCountry()
		QString = Replace(Qstring," ", "%20")

		Response.Redirect (URL & Qstring)

		Description = "A no longer AM delivery text message was sent to " & GetUserDisplayNameByUserNo(DriverUserNo) & " (" & getUserCellNumber(DriverUserNo) & ") at " & NOW()
		CreateAuditLogEntry "Delivery Board Routing System","No longer AM delivery text message sent","Minor",0,Description
		
	Else
			
		' Could not send dispatch test, no address on file
		emailBody = "Insight was unable to send a no longer an AM delivery text message to " & GetUserDisplayNameByUserNo(DriverUserNo) & ". No cell number on file"
		If Instr(ucase(sURL),"DEV") <> 0 Then SEND_TO = "rich@ocsaccess.com"
		SendMail "mailsender@" & maildomain ,SEND_TO,"Unable to send no longer an AM delivery text message",emailBody,GetTerm("Service"),"Missing Cell Number"
		Description = "Insight was unable to send no longer an AM delivery message to " & GetUserDisplayNameByUserNo(DriverUserNo) & ". No cell number on file"
		CreateAuditLogEntry "Delivery Board Routing System","Unable to send no longer an AM delivery text message","Major",0,Description
		

	End If

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub RemoveAMDeliveryFromDeliveryBoardModal() 

	invoiceNum = Request.Form("invoiceNum")
	CustID = Request.Form("custID")
	DriverUserNo = Request.Form("userNo")

	SQLRemoveDeliveryAsAM = "UPDATE RT_DeliveryBoard SET AMorPM = '' WHERE ivsNum = " & invoiceNum 
	Set cnnRemoveDeliveryAsAM = Server.CreateObject("ADODB.Connection")
	cnnRemoveDeliveryAsAM.open (Session("ClientCnnString"))
	
	Set rsRemoveDeliveryAsAM = Server.CreateObject("ADODB.Recordset")
	rsRemoveDeliveryAsAM.CursorLocation = 3 
	Set rsRemoveDeliveryAsAM = cnnRemoveDeliveryAsAM.Execute(SQLRemoveDeliveryAsAM)

	CreateAuditLogEntry "Delivery Board Invoice AM/PM Changed","Delivery Board Invoice AM/PM Changed","Minor",0,"Delivery Invoice #: " & invoiceNum & " - no longer marked as AM Delivery"
		
	set rsRemoveDeliveryAsAM = Nothing
	cnnRemoveDeliveryAsAM.Close
	Set cnnRemoveDeliveryAsAM = Nothing

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