<!-- Tab starts here -->

<div role="tabpanel" class="tab-pane fade" id="leftnav">


		<!-- userAccessInvControl line !-->
	   	<div class="col-md-3 enable-disable">
 			
			<% If MUV_Read("OrderAPIModuleOn") = "Enabled" Then %>	
				<p style="margin-bottom:10px">								
					<% If userLeftNavAPIModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavAPIModule" name="chkUserLeftNavAPIModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavAPIModule" name="chkUserLeftNavAPIModule">		    
					<% End If %>
				
					User can see <%= GetTerm("API") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavAPIModule" name="chkUserLeftNavAPIModule" value="off">
			<% End If %>
            
 				
			<% If MUV_Read("biModuleOn") = "Enabled" Then %>
				<p style="margin-bottom:10px">										
					<% If userLeftNavBIModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavBIModule" name="chkUserLeftNavBIModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavBIModule" name="chkUserLeftNavBIModule">		    
					<% End If %>
					
					User can see <%= GetTerm("Business Intelligence") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavBIModule" name="chkUserLeftNavBIModule" value="off">
			<% End If %>


 			
			<% If MUV_Read("OrderAPIModuleOn") = "Enabled" Then %>
				<p style="margin-bottom:10px">										
					<% If userLeftNavProspectingModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavProspectingModule" name="chkUserLeftNavProspectingModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavProspectingModule" name="chkUserLeftNavProspectingModule">		    
					<% End If %>
					
					User can see <%= GetTerm("Prospecting") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavProspectingModule" name="chkUserLeftNavProspectingModule" value="off">
			<% End If %>


			<% If MUV_Read("custServiceOn") = 1 Then %>	
				<p style="margin-bottom:10px">								
					<% If userLeftNavCustomerServiceModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavCustomerServiceModule" name="chkUserLeftNavCustomerServiceModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavCustomerServiceModule" name="chkUserLeftNavCustomerServiceModule">		    
					<% End If %>
					
					User can see <%= GetTerm("Customer Service") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavCustomerServiceModule" name="chkUserLeftNavCustomerServiceModule" value="off">
			<% End If %>
			
	
			<% If MUV_Read("equipmentModuleOn") = "Enabled" Then %>	
				<p style="margin-bottom:10px">								
					<% If userLeftNavEquipmentModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavEquipmentModule" name="chkUserLeftNavEquipmentModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavEquipmentModule" name="chkUserLeftNavEquipmentModule">		    
					<% End If %>
					
					User can see <%= GetTerm("Equipment") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavEquipmentModule" name="chkUserLeftNavEquipmentModule" value="off">
			<% End If %>
			

			<% If MUV_Read("InventoryControlModuleOn") = "Enabled" Then %>	
				<p style="margin-bottom:10px">								
					<% If userLeftNavInventoryControlModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavInventoryControlModule" name="chkUserLeftNavInventoryControlModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavInventoryControlModule" name="chkUserLeftNavInventoryControlModule">		    
					<% End If %>
					
					User can see <%= GetTerm("Inventory Control") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavInventoryControlModule" name="chkUserLeftNavInventoryControlModule" value="off">
			<% End If %>
			

			<% If cint(MUV_Read("arModuleOn")) = 1 Then %>
				<p style="margin-bottom:10px">									
					<% If userLeftNavAccountsReceivableModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavAccountsReceivableModule" name="chkUserLeftNavAccountsReceivableModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavAccountsReceivableModule" name="chkUserLeftNavAccountsReceivableModule">		    
					<% End If %>
					
					User can see <%= GetTerm("Accounts Receivable") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavAccountsReceivableModule" name="chkUserLeftNavAccountsReceivableModule" value="off">
			<% End If %>

           
		</div>
		
	







	
		<!-- userAccessInvControl line !-->
	   	<div class="col-md-3 enable-disable">

	
			<% If MUV_Read("apModuleOn") = "Enabled" Then %>	
				<p style="margin-bottom:10px">							
					<% If userLeftNavAccountsPayableModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavAccountsPayableModule" name="chkUserLeftNavAccountsPayableModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavAccountsPayableModule" name="chkUserLeftNavAccountsPayableModule">		    
					<% End If %>
					
					User can see <%= GetTerm("Accounts Payable") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavAccountsPayableModule" name="chkUserLeftNavAccountsPayableModule" value="off">
			<% End If %>
			

			<% If MUV_Read("serviceModuleOn") = "Enabled" Then %>	
				<p style="margin-bottom:10px">								
					<% If userLeftNavServiceModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavServiceModule" name="chkUserLeftNavServiceModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavServiceModule" name="chkUserLeftNavServiceModule">		    
					<% End If %>
					
					User can see <%= GetTerm("Service") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavServiceModule" name="chkUserLeftNavServiceModule" value="off">
			<% End If %>

	
			<% If MUV_Read("routingModuleOn") = "Enabled" Then %>
				<p style="margin-bottom:10px">									
					<% If userLeftNavRoutingModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavRoutingModule" name="chkUserLeftNavRoutingModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavRoutingModule" name="chkUserLeftNavRoutingModule">		    
					<% End If %>
					
					User can see <%= GetTerm("Routing") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavRoutingModule" name="chkUserLeftNavRoutingModule" value="off">
			<% End If %>
			

			<% If MUV_Read("quickbooksModuleOn") = "Enabled" Then %>	
				<p style="margin-bottom:10px">								
					<% If userLeftNavQuickbooksModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavQuickbooksModule" name="chkUserLeftNavQuickbooksModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavQuickbooksModule" name="chkUserLeftNavQuickbooksModule">		    
					<% End If %>
					
					User can see <%= GetTerm("QuickBooks") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavQuickbooksModule" name="chkUserLeftNavQuickbooksModule" value="off">
			<% End If %>


			<% If MUV_Read("FilterTrax") = "1" Then %>	
				<p style="margin-bottom:10px">								
					<% If userLeftNavFiltertraxModule = 1 Then %>
						<input type="checkbox" checked id="chkUserLeftNavFiltertraxModule" name="chkUserLeftNavFiltertraxModule">
					<% Else %>
						<input type="checkbox" unchecked id="chkUserLeftNavFiltertraxModule" name="chkUserLeftNavFiltertraxModule">		    
					<% End If %>
					
					User can see <%= GetTerm("FilterTrax") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavFiltertraxModule" name="chkUserLeftNavFiltertraxModule" value="off">
			<% End If %>


			<p style="margin-bottom:10px">								
				<% If userLeftNavSystem = 1 Then %>
					<input type="checkbox" checked id="chkUserLeftNavSystem" name="chkUserLeftNavSystem">
				<% Else %>
					<input type="checkbox" unchecked id="chkUserLeftNavSystem" name="chkUserLeftNavSystem">		    
				<% End If %>
				
				User can see <%= GetTerm("System") %> menu link
			</p>

                                                                
		</div>

</div>
<!-- Tab ends here -->

		