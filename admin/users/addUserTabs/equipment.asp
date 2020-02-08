<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="equipment">
   <% If MUV_Read("equipmentModuleOn") = "Enabled" Then %>
		<!-- userAccessInvControl line !-->
	   	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
			<strong><%=GetTerm("Equipment")%> Access</strong>
			
			<br>
			
			User Can Edit Equipment Tables on the Fly
			<input type="checkbox" id="chkUserCanEditEqpTablesOnFly"  name="chkUserCanEditEqpTablesOnFly">
			
		</div>
	<% End If %>
</div>
<!-- Tab ends here -->