<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="inventorycontrol">
   <% If MUV_Read("InventoryControlModuleOn") = "Enabled" Then %>
		<!-- userAccessInvControl line !-->
	   	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
			<strong><%=GetTerm("Inventory Control")%> Access</strong>
			<div class="radio">
				<label>
					<input type="radio" name="optInvControlAccessType" id="optInvControlAccessTypeNone" value="NONE" <% If userInventoryControlAccessType ="NONE" Then Response.Write("checked")%>>
					None
				</label>
			</div>
			<div class="radio">
				<label>
					<input type="radio" name="optInvControlAccessType" id="optInvControlAccessTypeReadOnly" value="READONLY" <% If userInventoryControlAccessType ="READONLY" Then Response.Write("checked")%>>
					Read Only
				</label>
			</div>
			<div class="radio">
				<label>
					<input type="radio" name="optInvControlAccessType" id="optInvControlAccessTypeReadWrite" value="READWRITE" <% If userInventoryControlAccessType ="READWRITE" Then Response.Write("checked")%>>
					Read / Write 
				</label>
			</div>
			
			<br>
			
			Allow Mobile Inventory Access
			<input type="checkbox" id="chkInvControlMobileAccess"  name="chkInvControlMobileAccess" <% If userMobileInventoryControlAccess = True Then Response.Write("checked")%>>
			
		</div>
	<% End If %>
</div>
<!-- Tab ends here -->