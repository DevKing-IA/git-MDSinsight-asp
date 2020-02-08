<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="api">

   <%If MUV_Read("orderAPIModuleOn")  = "Enabled" Then %>
		<!-- orderAPIModuleOn line !-->
		<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
			<strong>Order API  Access</strong>
			
			<div class="radio">
				<label>
					<input type="radio" name="optOrderAPIAccessType" id="optOrderAPIAccessTypeNone" value="NONE" <% If userOrderAPIAccessType ="NONE" Then Response.Write(" checked ")%> >
					None
					</label>
			</div>
			<div class="radio">
				<label>
				<input type="radio" name="optOrderAPIAccessType" id="optOrderAPIAccessTypeRead" value="READ" <% If userOrderAPIAccessType ="READY" Then Response.Write(" checked ")%> >
				Read
				</label>
			</div>
			<div class="radio">
				<label>
					<input type="radio" name="optOrderAPIAccessType" id="optOrderAPIAccessTypeReadResend" value="READ_RESEND" <% If userOrderAPIAccessType ="READ_RESEND" Then Response.Write(" checked ")%> >
					Read & Re-Send
				</label>
			</div>
		</div>
	<% End If %>
                                    
</div>
<!-- Tab ends here -->