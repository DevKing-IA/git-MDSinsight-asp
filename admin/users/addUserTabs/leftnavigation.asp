<!-- Tab starts here -->

<div role="tabpanel" class="tab-pane fade" id="leftnav">


		<!-- userAccessInvControl line !-->
	   	<div class="col-md-3 enable-disable">
 			
			<% If MUV_Read("OrderAPIModuleOn") = "Enabled" Then %>	
				<p style="margin-bottom:10px">								
					<input type="checkbox" id="chkUserLeftNavAPIModule" name="chkUserLeftNavAPIModule">		    
					User can see <%= GetTerm("API") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavAPIModule" name="chkUserLeftNavAPIModule" value="off">
			<% End If %>
            
 				
			<% If MUV_Read("biModuleOn") = "Enabled" Then %>
				<p style="margin-bottom:10px">										
					<input type="checkbox" id="chkUserLeftNavBIModule" name="chkUserLeftNavBIModule">		    
					User can see <%= GetTerm("Business Intelligence") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavBIModule" name="chkUserLeftNavBIModule" value="off">
			<% End If %>


 			
			<% If MUV_Read("OrderAPIModuleOn") = "Enabled" Then %>
				<p style="margin-bottom:10px">										
					<input type="checkbox" id="chkUserLeftNavProspectingModule" name="chkUserLeftNavProspectingModule">		    
					User can see <%= GetTerm("Prospecting") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavProspectingModule" name="chkUserLeftNavProspectingModule" value="off">
			<% End If %>


			<% If MUV_Read("custServiceOn") = 1 Then %>	
				<p style="margin-bottom:10px">								
					<input type="checkbox" id="chkUserLeftNavCustomerServiceModule" name="chkUserLeftNavCustomerServiceModule">		    
					User can see <%= GetTerm("Customer Service") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavCustomerServiceModule" name="chkUserLeftNavCustomerServiceModule" value="off">
			<% End If %>
			
	
			<% If MUV_Read("equipmentModuleOn") = "Enabled" Then %>	
				<p style="margin-bottom:10px">								
					<input type="checkbox" id="chkUserLeftNavEquipmentModule" name="chkUserLeftNavEquipmentModule">		    
					User can see <%= GetTerm("Equipment") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavEquipmentModule" name="chkUserLeftNavEquipmentModule" value="off">
			<% End If %>
			

			<% If MUV_Read("InventoryControlModuleOn") = "Enabled" Then %>	
				<p style="margin-bottom:10px">								
					<input type="checkbox" id="chkUserLeftNavInventoryControlModule" name="chkUserLeftNavInventoryControlModule">		    
					User can see <%= GetTerm("Inventory Control") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavInventoryControlModule" name="chkUserLeftNavInventoryControlModule" value="off">
			<% End If %>
			

			<% If cint(MUV_Read("arModuleOn")) = 1 Then %>
				<p style="margin-bottom:10px">									
					<input type="checkbox" id="chkUserLeftNavAccountsReceivableModule" name="chkUserLeftNavAccountsReceivableModule">		    
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
					<input type="checkbox" id="chkUserLeftNavAccountsPayableModule" name="chkUserLeftNavAccountsPayableModule">		    
					User can see <%= GetTerm("Accounts Payable") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavAccountsPayableModule" name="chkUserLeftNavAccountsPayableModule" value="off">
			<% End If %>
			

			<% If MUV_Read("serviceModuleOn") = "Enabled" Then %>	
				<p style="margin-bottom:10px">								
					<input type="checkbox" id="chkUserLeftNavServiceModule" name="chkUserLeftNavServiceModule">		    
					User can see <%= GetTerm("Service") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavServiceModule" name="chkUserLeftNavServiceModule" value="off">
			<% End If %>

	
			<% If MUV_Read("routingModuleOn") = "Enabled" Then %>
				<p style="margin-bottom:10px">									
					<input type="checkbox" id="chkUserLeftNavRoutingModule" name="chkUserLeftNavRoutingModule">		    
					User can see <%= GetTerm("Routing") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavRoutingModule" name="chkUserLeftNavRoutingModule" value="off">
			<% End If %>
			

			<% If MUV_Read("quickbooksModuleOn") = "Enabled" Then %>	
				<p style="margin-bottom:10px">								
					<input type="checkbox" id="chkUserLeftNavQuickbooksModule" name="chkUserLeftNavQuickbooksModule">		    
					User can see <%= GetTerm("QuickBooks") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavQuickbooksModule" name="chkUserLeftNavQuickbooksModule" value="off">
			<% End If %>


			<% If MUV_Read("FilterTrax") = "1" Then %>	
				<p style="margin-bottom:10px">								
					<input type="checkbox" id="chkUserLeftNavFiltertraxModule" name="chkUserLeftNavFiltertraxModule">		    
					User can see <%= GetTerm("FilterTrax") %> module menu link
				</p>
			<% Else %>
				<input type="hidden" id="chkUserLeftNavFiltertraxModule" name="chkUserLeftNavFiltertraxModule" value="off">
			<% End If %>


			<p style="margin-bottom:10px">								
				<input type="checkbox" id="chkUserLeftNavSystem" name="chkUserLeftNavSystem">		    
				User can see <%= GetTerm("System") %> menu link
			</p>

                                                                
		</div>

</div>
<!-- Tab ends here -->

		