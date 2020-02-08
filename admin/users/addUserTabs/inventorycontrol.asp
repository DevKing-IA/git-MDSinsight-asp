<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="inventorycontrol">
   <% If MUV_Read("InventoryControlModuleOn") = "Enabled" Then %>
		<!-- userAccessInvControl line !-->
	   	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
			<strong><%=GetTerm("Inventory Control")%> Access</strong>
			<div class="radio">
				<label>
					<input type="radio" name="optInvControlAccessType" id="optInvControlAccessTypeNone" value="NONE">
					None
				</label>
			</div>
			<div class="radio">
				<label>
					<input type="radio" name="optInvControlAccessType" id="optInvControlAccessTypeReadOnly" value="READONLY">
					Read Only
				</label>
			</div>
			<div class="radio">
				<label>
					<input type="radio" name="optInvControlAccessType" id="optInvControlAccessTypeReadWrite" value="READWRITE">
					Read / Write 
				</label>
			</div>
			
			<br>
			
			Allow Mobile Inventory Access
			<input type="checkbox" id="chkInvControlMobileAccess"  name="chkInvControlMobileAccess">
			
		</div>
	<% End If %>
</div>
<!-- Tab ends here -->