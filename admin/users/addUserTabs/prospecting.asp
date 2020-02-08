<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="prospecting">
   <% If  MUV_Read("prospectingModuleOn")  = "Enabled" Then %>
	   			<!-- userAccessCRM line !-->
			   	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
					<strong><%=GetTerm("Prospecting")%> Access</strong>
					<div class="radio">
						<label>
							<input type="radio" name="optCRMAccessType" id="optCRMAccessTypeNone" value="NONE">
							None
						</label>
					</div>
					<div class="radio">
						<label>
							<input type="radio" name="optCRMAccessType" id="optCRMAccessTypeReadOnly" value="READONLY">
							Read Only
						</label>
					</div>
					<div class="radio">
						<label>
							<input type="radio" name="optCRMAccessType" id="optCRMAccessTypeWriteOwned" value="WRITEOWNED">
							Read / Write owned prospects
						</label>
					</div>
					<div class="radio">
						<label>
							<input type="radio" name="optCRMAccessType" id="optCRMAccessTypeReadWrite" value="READWRITE">
							Read / Write all prospects
						</label>
					</div>
					
					<br>
					
					Allow Access To <%=GetTerm("Prospecting")%> Add/Edit Menu
					<input type="checkbox" checked id="chkCRMAddEditAccess"  name="chkCRMAddEditAccess">

					<br><br>
					
					Allow Access To Delete <%=GetTerm("Prospect")%> Button
					<input type="checkbox" unchecked id="chkCRMDeleteAccess"  name="chkCRMDeleteAccess">
					
					<br><br>
					
					Allowed To Edit <%=GetTerm("Prospect")%> On The Fly
					<input type="checkbox" unchecked id="chkUserEditCRMOnTheFly"  name="chkUserEditCRMOnTheFly">
					
					
				</div>
				
				<!-- download email line !-->
				<div class="row">
					<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
						<strong>Download this user's email</strong>
						<input type="checkbox" checked id="chkDownloadEmail"  name="chkDownloadEmail">
					</div>
				</div>
				<!-- eof download email line !-->

				<!-- make calendar entries line !-->
				<div class="row">
					<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
						<strong>Set appointments and meetings in user's calendar</strong>
						<input type="checkbox" checked id="chkUpdateCalendar"  name="chkUpdateCalendar">
					</div>
				</div>
				<!-- make calendar entries line !-->

				
			<%End If%>
			
			
        
</div>
<!-- Tab ends here -->