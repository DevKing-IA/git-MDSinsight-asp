
 <!--Start Show Session Expire Warning Popup here -->
    <div class="modal fade" id="session-expire-warning-modal" aria-hidden="true" data-keyboard="false" data-backdrop="static" tabindex="-1" role="dialog" aria-labelledby="myModalLabel">
        <div class="modal-dialog" role="document">
            <div class="modal-content">
                <div class="modal-header">                  
                    <h4 class="modal-title">Session Expiration Warning</h4>
                </div>
                <div class="modal-body">
                    Your Insight session will expire in <span id="seconds-timer"></span> seconds. Do you want to extend the session?
                </div>
                <div class="modal-footer">
                    <button id="btnOk" type="button" class="btn btn-default" style="padding: 6px 12px; margin-bottom: 0; font-size: 14px; font-weight: normal; border: 1px solid transparent; border-radius: 4px;  background-color: #428bca; color: #FFF;">Yes, Keep Me Logged In</button>
                    <!--<button id="btnSessionExpiredCancelled" type="button" class="btn btn-default" data-dismiss="modal" style="padding: 6px 12px; margin-bottom: 0; font-size: 14px; font-weight: normal; border: 1px solid transparent; border-radius: 4px; background-color: #428bca; color: #FFF;">Cancel</button>-->
                    <button id="btnLogoutNow" type="button" class="btn btn-default" style="padding: 6px 12px; margin-bottom: 0; font-size: 14px; font-weight: normal; border: 1px solid transparent; border-radius: 4px;  background-color: #428bca; color: #FFF;">Logout Now</button>
                </div>
            </div>
        </div>
    </div>
    <!--End Show Session Expire Warning Popup here -->
    <!--Start Show Session Expire Popup here -->
    <div class="modal fade" id="session-expired-modal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel">
        <div class="modal-dialog" role="document">
            <div class="modal-content">
                <div class="modal-header">
                    <h4 class="modal-title">Session Expired</h4>
                </div>
                <div class="modal-body">
                    Your session has expired.
                </div>
                <div class="modal-footer">
                    <button id="btnExpiredOk" onclick="sessionExpiredClicked()" type="button" class="btn btn-primary" data-dismiss="modal" style="padding: 6px 12px; margin-bottom: 0; font-size: 14px; font-weight: normal; border: 1px solid transparent; border-radius: 4px; background-color: #428bca; color: #FFF;">Ok</button>
                </div>
            </div>
        </div>
    </div>
	<script src="<%= BaseURL %>js/sessionTimeout.js"></script>
	 <!-- This will init session -->
	<script>
	 initSessionMonitor();
	</script>
	

 
  </body>
</html>