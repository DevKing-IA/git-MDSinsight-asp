<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="fieldservice">

	<!-- filter change row !-->
	<div class="col-lg-12" >
		<p><br>Show filter changes & PM calls within the following routes</p>
		<div class="form-control" style="height: auto; padding: 5px;">
		<% 'Get all Routes
			SQL9 = "SELECT DISTINCT RouteNum FROM AR_Customer Order By RouteNum"
			Set cnn9 = Server.CreateObject("ADODB.Connection")
			cnn9.open (Session("ClientCnnString"))
			Set rs9 = Server.CreateObject("ADODB.Recordset")
			rs9.CursorLocation = 3 
			Set rs9 = cnn9.Execute(SQL9)
			If not rs9.EOF Then
				Do
					checked = ""
					Response.Write( "<label class='btn btn-default btn-xs' style='width: 85px; text-align: left; font-size: 14px; margin: 0 3px 3px 0;'><input "&checked&" type='checkbox' name='txtRoutes' value='"&rs9("RouteNum")&"'> "&rs9("RouteNum")&"</label>")
					rs9.movenext
				Loop until rs9.eof
			End If
			set rs9 = Nothing
			cnn9.close
			set cnn9 = Nothing
		%>
		</div>
	</div>

                       
                                    
	<!-- No activity messages !-->
	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 box-border">
 			<div class="row">
               	<div class="col-lg-12">
					Turn on No Activity 'nag' messages
					<select class="form-control custom-select" name="seluserNoActivityNagMessageOverride_FS" id="seluserNoActivityNagMessageOverride_FS">
						<option value="Use Global">Use Global Setting</option>	
						<option value="Yes">Yes</option>	
						<option value="No">No</option>	
					</select>
				</div>
			</div>
					
			<!-- "nag" column -->
			<div class="row ">
				<div class="col-lg-12">

					<div class="row">
						<div class="col-lg-12">Start sending messages if there has been No Activity by 
							<select class="form-control custom-select" id="seluserNoActivityNagTimeOfDay_FS" name="seluserNoActivityNagTimeOfDay_FS">			
								<option value="00:00">-Midnight-</option>
								<option value="00:15">12:15 AM</option>
								<option value="00:00">12:30 AM</option>
								<option value="00:45">12:45 AM</option>
								<option value="1:00">1:00 AM</option>
								<option value="1:15">1:15 AM</option>
								<option value="1:30">1:30 AM</option>
								<option value="1:45">1:45 AM</option>
								<option value="2:00">2:00 AM</option>
								<option value="2:15">2:15 AM</option>
								<option value="2:30">2:30 AM</option>
								<option value="2:45">2:45 AM</option>
								<option value="3:00">3:00 AM</option>
								<option value="3:15">3:15 AM</option>
								<option value="3:30">3:30 AM</option>
								<option value="3:45">3:45 AM</option>
								<option value="4:00">4:00 AM</option>
								<option value="4:15">4:15 AM</option>
								<option value="4:30">4:30 AM</option>
								<option value="4:45">4:45 AM</option>
								<option value="5:00">5:00 AM</option>
								<option value="5:15">5:15 AM</option>
								<option value="5:30">5:30 AM</option>
								<option value="5:45">5:45 AM</option>
								<option value="6:00">6:00 AM</option>
								<option value="6:15">6:15 AM</option>
								<option value="6:30">6:30 AM</option>
								<option value="6:45">6:45 AM</option>
								<option value="7:00">7:00 AM</option>
								<option value="7:15">7:15 AM</option>
								<option value="7:30">7:30 AM</option>
								<option value="7:45">7:45 AM</option>
								<option value="8:00">8:00 AM</option>
								<option value="8:15">8:15 AM</option>
								<option value="8:30">8:30 AM</option>
								<option value="8:45">8:45 AM</option>
								<option value="9:00">9:00 AM</option>
								<option value="9:15">9:15 AM</option>
								<option value="930">9:30 AM</option>
								<option value="945">9:45 AM</option>
								<option value="10:00">10:00 AM</option>
								<option value="10:15">10:15 AM</option>
								<option value="10:30">10:30 AM</option>
								<option value="10:45">10:45 AM</option>
								<option value="11:00">11:00 AM</option>
								<option value="11:15">11:15 AM</option>
								<option value="11:30">11:30 AM</option>
								<option value="11:45">11:45 AM</option>
								<option value="12:00">-Noon-</option>
								<option value="12:15">12:15 PM</option>
								<option value="12:30">12:30 PM</option>
								<option value="12:45">12:45 PM</option>
								<option value="13:00">1:00 PM</option>
								<option value="13:15">1:15 PM</option>
								<option value="13:30">1:30 PM</option>
								<option value="13:45">1:45 PM</option>
								<option value="14:00">2:00 PM</option>
								<option value="14:15">2:15 PM</option>
								<option value="14:30">2:30 PM</option>
								<option value="14:45">2:45 PM</option>
								<option value="15:00">3:00 PM</option>
								<option value="15:15">3:15 PM</option>
								<option value="15:30">3:30 PM</option>
								<option value="15:45">3:45 PM</option>
								<option value="16:00">4:00 PM</option>
								<option value="16:15">4:15 PM</option>
								<option value="16:30">4:30 PM</option>
								<option value="16:45">4:45 PM</option>
								<option value="17:00">5:00 PM</option>
								<option value="17:15">5:15 PM</option>
								<option value="17:30">5:30 PM</option>
								<option value="17:45">5:45 PM</option>
								<option value="18:00">6:00 PM</option>
								<option value="18:15">6:15 PM</option>
								<option value="18:30">6:30 PM</option>
								<option value="18:45">6:45 PM</option>
								<option value="19:00">7:00 PM</option>
								<option value="19:15">7:15 PM</option>
								<option value="19:30">7:30 PM</option>
								<option value="19:45">7:45 PM</option>
								<option value="20:00">8:00 PM</option>
								<option value="20:15">8:15 PM</option>
								<option value="20:30">8:30 PM</option>
								<option value="20:45">8:45 PM</option>
								<option value="21:00">9:00 PM</option>
								<option value="21:15">9:15 PM</option>
								<option value="21:30">9:30 PM</option>
								<option value="21:45">9:45 PM</option>
								<option value="22:00">10:00 PM</option>
								<option value="22:15">10:15 PM</option>
								<option value="22:30">10:30 PM</option>
								<option value="22:45">10:45 PM</option>
								<option value="23:00">11:00 PM</option>
								<option value="23:15">11:15 PM</option>
								<option value="23:30">11:30 PM</option>
								<option value="23:45">11:45 PM</option>
		 					</select>
						</div>
					</div>
            	
					<div class="row">
                    	<div class="col-lg-12">Send when there has been No Activity for 
							<select class="form-control custom-select" id="seluserNoActivityNagMinutes_FS" name="seluserNoActivityNagMinutes_FS">
								<%
									For x = 15 to 180 Step 5 ' 3 hours
										If x mod 60 = 0 Then
											Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
										Else
											Response.Write("<option value='" & x & "'>" & x & "</option>")
										End If
									Next
								%>
							</select>&nbsp;minutes
						</div>
					</div>
                
					<div class="row">
						<div class="col-lg-12">Continue to send every
							<select class="form-control custom-select" id="seluserNoActivityNagIntervalMinutes_FS" name="seluserNoActivityNagIntervalMinutes_FS">
								<%
									For x = 10 to 120 Step 5 ' 2 hours
										If x mod 60 = 0 Then
											Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
										Else
											Response.Write("<option value='" & x & "'>" & x & "</option>")
										End If
									Next
								%>
							</select>&nbsp;minutes
						</div>
					</div>
                 
					<div class="row">
						<div class="col-lg-12">Send a maximum of 
							<select class="form-control custom-select" id="seluserNoActivityNagMessageMaxToSendPerStop_FS" name="seluserNoActivityNagMessageMaxToSendPerStop_FS">
								<%
									For x = 1 to 10
										Response.Write("<option value='" & x & "'>" & x & "</option>")
									Next
								%>
							</select>&nbsp;messages each time a 'No Activity' event occurs
						</div>
					</div>
                 
					<div class="row">
						<div class="col-lg-12">Send a maxium of 
							<select class="form-control custom-select"  id="seluserNoActivityNagMessageMaxToSendPerDriverPerDay_FS" name="seluserNoActivityNagMessageMaxToSendPerDriverPerDay_FS">
								<%
									For x = 1 to 25
										Response.Write("<option value='" & x & "'>" & x & "</option>")
									Next
								%>
							</select>&nbsp;messages to a technician on any given day
						</div>
					</div>
                 
					<div class="row">
						<div class="col-lg-12">Send method 
							<select class="form-control custom-select"   id="seluserNoActivityNagMessageSendMethod_FS" name="seluserNoActivityNagMessageSendMethod_FS">
								<option value="Text">Text Message Only</option>
								<!--<option value="Email">Email Only</option>
								<option value="TextThenEmail">Text - If no cell number, send email</option>
								<option value="EmailThenText">Email - If no valid email address, send text</option>
								<option value="Both">Both</option>-->
							</select>
						</div>
					</div>
            
				</div>
            </div>
            <!-- eof "nag" column -->
	</div>
	<!-- eof No activity messages !-->


   
</div>
<!-- Tab ends here -->