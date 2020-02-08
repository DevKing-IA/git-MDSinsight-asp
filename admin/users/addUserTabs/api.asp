<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="api">

  <%If MUV_Read("orderAPIModuleOn")  = "Enabled" Then %>
				<!-- order api line !-->
				<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
					<strong>Order API  Access</strong>
					
					<div class="radio">
						<label>
							<input type="radio" name="optOrderAPIAccessType" id="optOrderAPIAccessTypeNone" value="NONE">
							None
							</label>
					</div>
					<div class="radio">
						<label>
						<input type="radio" name="optOrderAPIAccessType" id="optOrderAPIAccessTypeRead" value="READ">
						Read
						</label>
					</div>
					<div class="radio">
						<label>
							<input type="radio" name="optOrderAPIAccessType" id="optOrderAPIAccessTypeReadResend" value="READ_RESEND">
							Read & Re-Send
						</label>
					</div>
				</div>
		<% End If %>
        
</div>
<!-- Tab ends here -->