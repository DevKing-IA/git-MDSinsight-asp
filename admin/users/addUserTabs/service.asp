<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="service">

   <% If MUV_Read("serviceModuleOn") = "Enabled" Then %>
		
	   	<div class="col-md-3 enable-disable">
	   	
			<p style="margin-bottom:10px">									
				<input type="checkbox" id="chkUserCreateNewServiceTicket" name="chkUserCreateNewServiceTicket">		    
				User Can Create New <%= GetTerm("Service") %> Ticket
			</p>

			<p style="margin-bottom:10px">									
				<input type="checkbox" id="chkUserAccessServiceDispatchCenter" name="chkUserAccessServiceDispatchCenter">		    
				User Can Access <%= GetTerm("Service") %> Dipatch Center
			</p>
	   	
			<p style="margin-bottom:10px">									
				<input type="checkbox" id="chkUserAccessServiceActionsModalButton" name="chkUserAccessServiceActionsModalButton">		    
				User Can Access <%= GetTerm("Service") %> Actions Button
			</p>
 
			<p style="margin-bottom:10px">									
				<input type="checkbox" id="chkUserAccessServiceDispatchButton" name="chkUserAccessServiceDispatchButton">		    
				User Can Access <%= GetTerm("Service") %> Dispatch Button
			</p>

			<p style="margin-bottom:10px">									
				<input type="checkbox" id="chkUserAccessServiceCloseCancelButton" name="chkUserAccessServiceCloseCancelButton">		    
				User Can Access <%= GetTerm("Service") %> Close/Cancel Ticket Button
			</p>
           
		</div>
		
	   	<div class="col-md-3 enable-disable">
	   	
			<p style="margin-bottom:10px">									
				<input type="checkbox" id="chkUserCreateEquipmentSymptomCodesOnTheFly" name="chkUserCreateEquipmentSymptomCodesOnTheFly">		    
				User Allowed To Edit <%= GetTerm("Service") %> Symptoms On The Fly
			</p>
		
			<p style="margin-bottom:10px">									
				<input type="checkbox" id="chkUserCreateEquipmentProblemCodesOnTheFly" name="chkUserCreateEquipmentProblemCodesOnTheFly">		    
				User Allowed To Edit <%= GetTerm("Service") %> Problems On The Fly
			</p>
			
			<p style="margin-bottom:10px">									
				<input type="checkbox" id="chkUserCreateEquipmentResolutionCodesOnTheFly" name="chkUserCreateEquipmentResolutionCodesOnTheFly">		    
				User Allowed To Edit <%= GetTerm("Service") %> Resolutions On The Fly
			</p>
			
		</div>

	<% End If %>
</div>
<!-- Tab ends here -->