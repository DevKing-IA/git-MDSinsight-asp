<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="service">
   <% If MUV_Read("serviceModuleOn") = "Enabled" Then %>	
		
	   	<div class="col-md-3 enable-disable">
	   	
			<p style="margin-bottom:10px">									
				<% If userCreateNewServiceTicket = 1 Then %>
					<input type="checkbox" checked id="chkUserCreateNewServiceTicket" name="chkUserCreateNewServiceTicket">
				<% Else %>
					<input type="checkbox" unchecked id="chkUserCreateNewServiceTicket" name="chkUserCreateNewServiceTicket">		    
				<% End If %>
				
				User Can Create New <%= GetTerm("Service") %> Ticket
			</p>

			<p style="margin-bottom:10px">									
				<% If userAccessServiceDispatchCenter = 1 Then %>
					<input type="checkbox" checked id="chkUserAccessServiceDispatchCenter" name="chkUserAccessServiceDispatchCenter">
				<% Else %>
					<input type="checkbox" unchecked id="chkUserAccessServiceDispatchCenter" name="chkUserAccessServiceDispatchCenter">		    
				<% End If %>
				
				User Can Access <%= GetTerm("Service") %> Dipatch Center
			</p>
	   	
			<p style="margin-bottom:10px">									
				<% If userAccessServiceActionsModalButton = 1 Then %>
					<input type="checkbox" checked id="chkUserAccessServiceActionsModalButton" name="chkUserAccessServiceActionsModalButton">
				<% Else %>
					<input type="checkbox" unchecked id="chkUserAccessServiceActionsModalButton" name="chkUserAccessServiceActionsModalButton">		    
				<% End If %>
				
				User Can Access <%= GetTerm("Service") %> Actions Button
			</p>
 
			<p style="margin-bottom:10px">									
				<% If userAccessServiceDispatchButton = 1 Then %>
					<input type="checkbox" checked id="chkUserAccessServiceDispatchButton" name="chkUserAccessServiceDispatchButton">
				<% Else %>
					<input type="checkbox" unchecked id="chkUserAccessServiceDispatchButton" name="chkUserAccessServiceDispatchButton">		    
				<% End If %>
				
				User Can Access <%= GetTerm("Service") %> Dispatch Button
			</p>

			<p style="margin-bottom:10px">									
				<% If userAccessServiceCloseCancelButton = 1 Then %>
					<input type="checkbox" checked id="chkUserAccessServiceCloseCancelButton" name="chkUserAccessServiceCloseCancelButton">
				<% Else %>
					<input type="checkbox" unchecked id="chkUserAccessServiceCloseCancelButton" name="chkUserAccessServiceCloseCancelButton">		    
				<% End If %>
				
				User Can Access <%= GetTerm("Service") %> Close/Cancel Ticket Button
			</p>
           
		</div>
		
	   	<div class="col-md-3 enable-disable">
	   	
			<p style="margin-bottom:10px">									
				<% If userCreateEquipmentSymptomCodesOnTheFly = 1 Then %>
					<input type="checkbox" checked id="chkUserCreateEquipmentSymptomCodesOnTheFly" name="chkUserCreateEquipmentSymptomCodesOnTheFly">
				<% Else %>
					<input type="checkbox" unchecked id="chkUserCreateEquipmentSymptomCodesOnTheFly" name="chkUserCreateEquipmentSymptomCodesOnTheFly">		    
				<% End If %>
				
				User Allowed To Edit <%= GetTerm("Service") %> Symptoms On The Fly
			</p>
		
			<p style="margin-bottom:10px">									
				<% If userCreateEquipmentProblemCodesOnTheFly = 1 Then %>
					<input type="checkbox" checked id="chkUserCreateEquipmentProblemCodesOnTheFly" name="chkUserCreateEquipmentProblemCodesOnTheFly">
				<% Else %>
					<input type="checkbox" unchecked id="chkUserCreateEquipmentProblemCodesOnTheFly" name="chkUserCreateEquipmentProblemCodesOnTheFly">		    
				<% End If %>
				
				User Allowed To Edit <%= GetTerm("Service") %> Problems On The Fly
			</p>
			
			<p style="margin-bottom:10px">									
				<% If userCreateEquipmentResolutionCodesOnTheFly = 1 Then %>
					<input type="checkbox" checked id="chkUserCreateEquipmentResolutionCodesOnTheFly" name="chkUserCreateEquipmentResolutionCodesOnTheFly">
				<% Else %>
					<input type="checkbox" unchecked id="chkUserCreateEquipmentResolutionCodesOnTheFly" name="chkUserCreateEquipmentResolutionCodesOnTheFly">		    
				<% End If %>
				
				User Allowed To Edit <%= GetTerm("Service") %> Resolutions On The Fly
			</p>
			
		</div>

	<% End If %>
</div>
<!-- Tab ends here -->

