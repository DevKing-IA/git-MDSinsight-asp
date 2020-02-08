<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="prospecting">
   <%If MUV_Read("prospectingModuleOn")  = "Enabled" Then %>
		<!-- userAccessCRM line !-->
		<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
			<strong>Prospecting  Access</strong>
			
			<div class="radio">
				<label>
					<input type="radio" name="optCRMAccessType" id="optCRMAccessTypeNone" value="NONE" <% If userCRMAccessType ="NONE" Then Response.Write(" checked ")%> >
					None
					</label>
			</div>
			<div class="radio">
				<label>
				<input type="radio" name="optCRMAccessType" id="optCRMAccessTypeReadOnly" value="READONLY" <% If userCRMAccessType ="READONLY" Then Response.Write(" checked ")%> >
				Read Only
				</label>
			</div>
			<div class="radio">
				<label>
					<input type="radio" name="optCRMAccessType" id="optCRMAccessTypeWriteOwned" value="WRITEOWNED" <% If userCRMAccessType ="WRITEOWNED" Then Response.Write(" checked ")%> >
					Read / Write owned <%=GetTerm("prospects")%>
				</label>
			</div>
			<div class="radio">
				<label>
					<input type="radio" name="optCRMAccessType" id="optCRMAccessTypeReadWrite" value="READWRITE" <% If userCRMAccessType ="READWRITE" Then Response.Write(" checked ")%> >
					Read / Write all <%=GetTerm("prospects")%>
				</label>
			</div>
			
			<br>
			
			

		</div>
		
		<!-- download email line !-->
		<div class="row">
			<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
            
            <p>
				 Download this user's email 
				<% If userDownloadEmail = True Then %>
					<input type="checkbox" checked id="chkDownloadEmail"  name="chkDownloadEmail">
				<% Else %>
					<input type="checkbox" unchecked id="chkDownloadEmail"  name="chkDownloadEmail">		    
				<%End If%>
                </p>
                
                
                <p>
                Allow Access To <%=GetTerm("Prospecting")%> Add/Edit Menu
			<% If userProspectingAddEditAccess = True Then %>
				<input type="checkbox" checked id="chkCRMAddEditAccess"  name="chkCRMAddEditAccess">
			<% Else %>
				<input type="checkbox" unchecked id="chkCRMAddEditAccess"  name="chkCRMAddEditAccess">		    
			<%End If%>
            </p>

			<p>											
			Allow Access To Delete <%=GetTerm("Prospect")%> Button
			<% If userCRMDeleteAccess = True Then %>
				<input type="checkbox" checked id="chkCRMDeleteAccess"  name="chkCRMDeleteAccess">
			<% Else %>
				<input type="checkbox" unchecked id="chkCRMDeleteAccess"  name="chkCRMDeleteAccess">		    
			<%End If%>
            </p>
              
 			<p>											
			Allowed To Edit <%=GetTerm("Prospect")%> On The Fly
			<% If userEditCRMOnTheFly = True Then %>
				<input type="checkbox" checked id="chkUserEditCRMOnTheFly"  name="chkUserEditCRMOnTheFly">
			<% Else %>
				<input type="checkbox" unchecked id="chkUserEditCRMOnTheFly"  name="chkUserEditCRMOnTheFly">		    
			<%End If%>
            </p>
           
            <p>
            Set appointments and meetings in user's calendar 
				<% If userUpdateCalendar = True Then %>
					<input type="checkbox" checked id="chkUpdateCalendar"  name="chkUpdateCalendar">
				<% Else %>
					<input type="checkbox" unchecked id="chkUpdateCalendar"  name="chkUpdateCalendar">		    
				<%End If%>
            </p>
            
			</div>
		</div>
		<!-- eof download email line !-->

	 

	<% End If %>
                                    
</div>
<!-- Tab ends here -->