<style type="text/css">
.form-driver{
	width:40% !important;
}

.force-driver{
	margin-left:15px;
}

.custom-select{
	margin-bottom:20px;
}

.box-border{
	margin-left:0px;
	border:2px solid #fff;
}

.no-activity-box{
	margin-top:75px;
}
</style>

<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="routing">

	
 
 
			
	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 ">
    
    <!-- Route number !-->
	<div class="col-lg-12">
    <div class="row">
	Force Driver to Set Next Stop (mobile webapp)
	<select class="form-control form-driver" name="seluserForceNextStopSelectionOverride" id="seluserForceNextStopSelectionOverride">
		<option value="Use Global" <% If userForceNextStopSelectionOverride = "Use Global" Then Response.Write(" selected ")%>>Use Global Setting</option>	
		<option value="Yes" <% If userForceNextStopSelectionOverride = "Yes" Then Response.Write(" selected ")%>>Yes</option>	
		<option value="No" <% If userForceNextStopSelectionOverride = "No" Then Response.Write(" selected ")%>>No</option>	
	</select>
    </div>
	</div>
	<!-- eof Route number !-->

		<!-- "nag" column -->
		<div class="row ">
			<div class="col-lg-12 box-border">

				<!-- line -->
				<div class="row">
					<div class="col-lg-12">
						Turn on Next Stop 'nag' messages
						<select class="form-control custom-select" name="seluserNextStopNagMessageOverride" id="seluserNextStopNagMessageOverride">
							<option value="Use Global" <% If userNextStopNagMessageOverride= "Use Global" Then Response.Write(" selected ")%>>Use Global Setting</option>	
							<option value="Yes" <% If userNextStopNagMessageOverride= "Yes" Then Response.Write(" selected ")%>>Yes</option>	
							<option value="No" <% If userNextStopNagMessageOverride= "No" Then Response.Write(" selected ")%>>No</option>	
						</select>
					</div>
				</div>
				<!-- eof line -->
				
				<div class="row">
                	<div class="col-lg-12">Send when the Next Stop has not been set for 
						<select class="form-control custom-select" id="seluserNextStopNagMinutes" name="seluserNextStopNagMinutes">
							<%
								For x = 5 to 180 Step 5 ' 3 hours
									If x mod 60 = 0 Then
										If x = cint(userNextStopNagMinutes) Then 
											Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
										else
											Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
										End If
									Else
										If x = cint(userNextStopNagMinutes) Then 
											Response.Write("<option value='" & x & "' selected>" & x & "</option>")
										Else
											Response.Write("<option value='" & x & "'>" & x & "</option>")
										End If
									End If
								Next
							%>
						</select>&nbsp;minutes
					</div>
				</div>
                 
				<div class="row">
					<div class="col-lg-12">Continue to send every
						<select class="form-control custom-select" id="seluserNextStopNagIntervalMinutes" name="seluserNextStopNagIntervalMinutes">
							<%
								For x = 10 to 120 Step 5 ' 2 hours
									If x mod 60 = 0 Then
										If x = cint(userNextStopNagIntervalMinutes) Then 
											Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
										else
											Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
										End If
									Else
										If x = cint(userNextStopNagIntervalMinutes) Then 
											Response.Write("<option value='" & x & "' selected>" & x & "</option>")
										Else
											Response.Write("<option value='" & x & "'>" & x & "</option>")
										End If
									End If
								Next
							%>
						</select>&nbsp;minutes
					</div>
				</div>
                 
				<div class="row">
					<div class="col-lg-12">Send a maximum of 
						<select class="form-control custom-select" id="seluserNextStopNagMessageMaxToSendPerStop" name="seluserNextStopNagMessageMaxToSendPerStop">
							<%
								For x = 1 to 10
									If x = cint(userNextStopNagMessageMaxToSendPerStop) Then 
										Response.Write("<option value='" & x & "' selected>" & x & "</option>")
									Else
										Response.Write("<option value='" & x & "'>" & x & "</option>")
									End If
								Next
							%>
						</select>&nbsp;messages each time a 'No Next Stop' event occurs
					</div>
				</div>
                 
				<div class="row">
					<div class="col-lg-12">Send a maxium of 
						<select class="form-control custom-select"  id="seluserNextStopNagMessageMaxToSendThisDriverPerDay" name="seluserNextStopNagMessageMaxToSendThisDriverPerDay">
							<%
								For x = 1 to 25
									If x = cint(userNextStopNagMessageMaxToSendThisDriverPerDay) Then 
										Response.Write("<option value='" & x & "' selected>" & x & "</option>")
									Else
										Response.Write("<option value='" & x & "'>" & x & "</option>")
									End If
								Next
							%>
						</select>&nbsp;messages to this driver on any given day
					</div>
				</div>
                 
				<div class="row">
					<div class="col-lg-12">Send method 
						<select class="form-control custom-select"   id="seluserNextStopNagMessageSendMethod" name="seluserNextStopNagMessageSendMethod">
							<option value="Text"<%If userNextStopNagMessageSendMethod = "Text" Then Response.Write(" selected ")%>>Text Message Only</option>
							<!--<option value="Email"<%If userNextStopNagMessageSendMethod = "Email" Then Response.Write(" selected ")%>>Email Only</option>
							<option value="TextThenEmail"<%If userNextStopNagMessageSendMethod = "TextThenEmail" Then Response.Write(" selected ")%>>Text - If no cell number, send email</option>
							<option value="EmailThenText"<%If userNextStopNagMessageSendMethod = "EmailThenText" Then Response.Write(" selected ")%>>Email - If no valid email address, send text</option>
							<option value="Both"<%If userNextStopNagMessageSendMethod = "Both" Then Response.Write(" selected ")%>>Both</option>-->
						</select>
					</div>
				</div>
                            
            </div>
		</div>
        <!-- eof "nag" column -->
                                        
	</div>
    <!-- eof Additional driver setup info !-->
                                    
                                    
                                    
		<!-- No activity messages !-->
		<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 box-border no-activity-box" id="pnlDriverAdditional" >
 			<div class="row">
               	<div class="col-lg-12">
					Turn on No Activity 'nag' messages
					<select class="form-control custom-select" name="seluserNoActivityNagMessageOverride" id="seluserNoActivityNagMessageOverride">
						<option value="Use Global" <% If userNoActivityNagMessageOverride = "Use Global" Then Response.Write(" selected ")%>>Use Global Setting</option>	
						<option value="Yes" <% If userNoActivityNagMessageOverride = "Yes" Then Response.Write(" selected ")%>>Yes</option>	
						<option value="No" <% If userNoActivityNagMessageOverride = "No" Then Response.Write(" selected ")%>>No</option>	
					</select>
				</div>
			</div>
					
			<!-- "nag" column -->
			<div class="row ">
				<div class="col-lg-12">

					<div class="row">
						<div class="col-lg-12">Start sending messages if there has been No Activity by 
							<select class="form-control custom-select" id="seluserNoActivityNagTimeOfDay" name="seluserNoActivityNagTimeOfDay">			
								<option value="00:00"<%If userNoActivityNagTimeOfDay = "00:00" Then Response.Write(" selected ")%>>-Midnight-</option>
								<option value="00:15"<%If userNoActivityNagTimeOfDay = "00:15" Then Response.Write(" selected ")%>>12:15 AM</option>
								<option value="00:30"<%If userNoActivityNagTimeOfDay = "00:30" Then Response.Write(" selected ")%>>12:30 AM</option>
								<option value="00:45"<%If userNoActivityNagTimeOfDay = "00:45" Then Response.Write(" selected ")%>>12:45 AM</option>
								<option value="1:00"<%If userNoActivityNagTimeOfDay = "1:00" Then Response.Write(" selected ")%>>1:00 AM</option>
								<option value="1:15"<%If userNoActivityNagTimeOfDay = "1:15" Then Response.Write(" selected ")%>>1:15 AM</option>
								<option value="1:30"<%If userNoActivityNagTimeOfDay = "1:30" Then Response.Write(" selected ")%>>1:30 AM</option>
								<option value="1:45"<%If userNoActivityNagTimeOfDay = "1:45" Then Response.Write(" selected ")%>>1:45 AM</option>
								<option value="2:00"<%If userNoActivityNagTimeOfDay = "2:00" Then Response.Write(" selected ")%>>2:00 AM</option>
								<option value="2:15"<%If userNoActivityNagTimeOfDay = "2:15" Then Response.Write(" selected ")%>>2:15 AM</option>
								<option value="2:30"<%If userNoActivityNagTimeOfDay = "2:30" Then Response.Write(" selected ")%>>2:30 AM</option>
								<option value="2:45"<%If userNoActivityNagTimeOfDay = "2:45" Then Response.Write(" selected ")%>>2:45 AM</option>
								<option value="3:00"<%If userNoActivityNagTimeOfDay = "3:00" Then Response.Write(" selected ")%>>3:00 AM</option>
								<option value="3:15"<%If userNoActivityNagTimeOfDay = "3:15" Then Response.Write(" selected ")%>>3:15 AM</option>
								<option value="3:30"<%If userNoActivityNagTimeOfDay = "3:30" Then Response.Write(" selected ")%>>3:30 AM</option>
								<option value="3:45"<%If userNoActivityNagTimeOfDay = "3:45" Then Response.Write(" selected ")%>>3:45 AM</option>
								<option value="4:00"<%If userNoActivityNagTimeOfDay = "4:00" Then Response.Write(" selected ")%>>4:00 AM</option>
								<option value="4:15"<%If userNoActivityNagTimeOfDay = "4:15" Then Response.Write(" selected ")%>>4:15 AM</option>
								<option value="4:30"<%If userNoActivityNagTimeOfDay = "4:30" Then Response.Write(" selected ")%>>4:30 AM</option>
								<option value="4:45"<%If userNoActivityNagTimeOfDay = "4:45" Then Response.Write(" selected ")%>>4:45 AM</option>
								<option value="5:00"<%If userNoActivityNagTimeOfDay = "5:00" Then Response.Write(" selected ")%>>5:00 AM</option>
								<option value="5:15"<%If userNoActivityNagTimeOfDay = "5:15" Then Response.Write(" selected ")%>>5:15 AM</option>
								<option value="5:30"<%If userNoActivityNagTimeOfDay = "5:30" Then Response.Write(" selected ")%>>5:30 AM</option>
								<option value="5:45"<%If userNoActivityNagTimeOfDay = "5:45" Then Response.Write(" selected ")%>>5:45 AM</option>
								<option value="6:00"<%If userNoActivityNagTimeOfDay = "6:00" Then Response.Write(" selected ")%>>6:00 AM</option>
								<option value="6:15"<%If userNoActivityNagTimeOfDay = "6:15" Then Response.Write(" selected ")%>>6:15 AM</option>
								<option value="6:30"<%If userNoActivityNagTimeOfDay = "6:30" Then Response.Write(" selected ")%>>6:30 AM</option>
								<option value="6:45"<%If userNoActivityNagTimeOfDay = "6:45" Then Response.Write(" selected ")%>>6:45 AM</option>
								<option value="7:00"<%If userNoActivityNagTimeOfDay = "7:00" Then Response.Write(" selected ")%>>7:00 AM</option>
								<option value="7:15"<%If userNoActivityNagTimeOfDay = "7:15" Then Response.Write(" selected ")%>>7:15 AM</option>
								<option value="7:30"<%If userNoActivityNagTimeOfDay = "7:30" Then Response.Write(" selected ")%>>7:30 AM</option>
								<option value="7:45"<%If userNoActivityNagTimeOfDay = "7:45" Then Response.Write(" selected ")%>>7:45 AM</option>
								<option value="8:00"<%If userNoActivityNagTimeOfDay = "8:00" Then Response.Write(" selected ")%>>8:00 AM</option>
								<option value="8:15"<%If userNoActivityNagTimeOfDay = "8:15" Then Response.Write(" selected ")%>>8:15 AM</option>
								<option value="8:30"<%If userNoActivityNagTimeOfDay = "8:30" Then Response.Write(" selected ")%>>8:30 AM</option>
								<option value="8:45"<%If userNoActivityNagTimeOfDay = "8:45" Then Response.Write(" selected ")%>>8:45 AM</option>
								<option value="9:00"<%If userNoActivityNagTimeOfDay = "9:00" Then Response.Write(" selected ")%>>9:00 AM</option>
								<option value="9:15"<%If userNoActivityNagTimeOfDay = "9:15" Then Response.Write(" selected ")%>>9:15 AM</option>
								<option value="9:30"<%If userNoActivityNagTimeOfDay = "9:30" Then Response.Write(" selected ")%>>9:30 AM</option>
								<option value="9:45"<%If userNoActivityNagTimeOfDay = "9:45" Then Response.Write(" selected ")%>>9:45 AM</option>
								<option value="10:00"<%If userNoActivityNagTimeOfDay = "10:00" Then Response.Write(" selected ")%>>10:00 AM</option>
								<option value="10:15"<%If userNoActivityNagTimeOfDay = "10:15" Then Response.Write(" selected ")%>>10:15 AM</option>
								<option value="10:30"<%If userNoActivityNagTimeOfDay = "10:30" Then Response.Write(" selected ")%>>10:30 AM</option>
								<option value="10:45"<%If userNoActivityNagTimeOfDay = "10:45" Then Response.Write(" selected ")%>>10:45 AM</option>
								<option value="11:00"<%If userNoActivityNagTimeOfDay = "11:00" Then Response.Write(" selected ")%>>11:00 AM</option>
								<option value="11:15"<%If userNoActivityNagTimeOfDay = "11:15" Then Response.Write(" selected ")%>>11:15 AM</option>
								<option value="11:30"<%If userNoActivityNagTimeOfDay = "11:30" Then Response.Write(" selected ")%>>11:30 AM</option>
								<option value="11:45"<%If userNoActivityNagTimeOfDay = "11:45" Then Response.Write(" selected ")%>>11:45 AM</option>
								<option value="12:00"<%If userNoActivityNagTimeOfDay = "12:00" Then Response.Write(" selected ")%>>-Noon-</option>
								<option value="12:15"<%If userNoActivityNagTimeOfDay = "12:15" Then Response.Write(" selected ")%>>12:15 PM</option>
								<option value="12:30"<%If userNoActivityNagTimeOfDay = "12:30" Then Response.Write(" selected ")%>>12:30 PM</option>
								<option value="12:45"<%If userNoActivityNagTimeOfDay = "12:45" Then Response.Write(" selected ")%>>12:45 PM</option>
								<option value="13:00"<%If userNoActivityNagTimeOfDay = "13:00" Then Response.Write(" selected ")%>>1:00 PM</option>
								<option value="13:15"<%If userNoActivityNagTimeOfDay = "13:15" Then Response.Write(" selected ")%>>1:15 PM</option>
								<option value="13:30"<%If userNoActivityNagTimeOfDay = "13:30" Then Response.Write(" selected ")%>>1:30 PM</option>
								<option value="13:45"<%If userNoActivityNagTimeOfDay = "13:45" Then Response.Write(" selected ")%>>1:45 PM</option>
								<option value="14:00"<%If userNoActivityNagTimeOfDay = "14:00" Then Response.Write(" selected ")%>>2:00 PM</option>
								<option value="14:15"<%If userNoActivityNagTimeOfDay = "14:15" Then Response.Write(" selected ")%>>2:15 PM</option>
								<option value="14:30"<%If userNoActivityNagTimeOfDay = "14:30" Then Response.Write(" selected ")%>>2:30 PM</option>
								<option value="14:45"<%If userNoActivityNagTimeOfDay = "14:45" Then Response.Write(" selected ")%>>2:45 PM</option>
								<option value="15:00"<%If userNoActivityNagTimeOfDay = "15:00" Then Response.Write(" selected ")%>>3:00 PM</option>
								<option value="15:15"<%If userNoActivityNagTimeOfDay = "15:15" Then Response.Write(" selected ")%>>3:15 PM</option>
								<option value="15:30"<%If userNoActivityNagTimeOfDay = "15:30" Then Response.Write(" selected ")%>>3:30 PM</option>
								<option value="15:45"<%If userNoActivityNagTimeOfDay = "15:45" Then Response.Write(" selected ")%>>3:45 PM</option>
								<option value="16:00"<%If userNoActivityNagTimeOfDay = "16:00" Then Response.Write(" selected ")%>>4:00 PM</option>
								<option value="16:15"<%If userNoActivityNagTimeOfDay = "16:15" Then Response.Write(" selected ")%>>4:15 PM</option>
								<option value="16:30"<%If userNoActivityNagTimeOfDay = "16:30" Then Response.Write(" selected ")%>>4:30 PM</option>
								<option value="16:45"<%If userNoActivityNagTimeOfDay = "16:45" Then Response.Write(" selected ")%>>4:45 PM</option>
								<option value="17:00"<%If userNoActivityNagTimeOfDay = "17:00" Then Response.Write(" selected ")%>>5:00 PM</option>
								<option value="17:15"<%If userNoActivityNagTimeOfDay = "17:15" Then Response.Write(" selected ")%>>5:15 PM</option>
								<option value="17:30"<%If userNoActivityNagTimeOfDay = "17:30" Then Response.Write(" selected ")%>>5:30 PM</option>
								<option value="17:45"<%If userNoActivityNagTimeOfDay = "17:45" Then Response.Write(" selected ")%>>5:45 PM</option>
								<option value="18:00"<%If userNoActivityNagTimeOfDay = "18:00" Then Response.Write(" selected ")%>>6:00 PM</option>
								<option value="18:15"<%If userNoActivityNagTimeOfDay = "18:15" Then Response.Write(" selected ")%>>6:15 PM</option>
								<option value="18:30"<%If userNoActivityNagTimeOfDay = "18:30" Then Response.Write(" selected ")%>>6:30 PM</option>
								<option value="18:45"<%If userNoActivityNagTimeOfDay = "18:45" Then Response.Write(" selected ")%>>6:45 PM</option>
								<option value="19:00"<%If userNoActivityNagTimeOfDay = "19:00" Then Response.Write(" selected ")%>>7:00 PM</option>
								<option value="19:15"<%If userNoActivityNagTimeOfDay = "19:15" Then Response.Write(" selected ")%>>7:15 PM</option>
								<option value="19:30"<%If userNoActivityNagTimeOfDay = "19:30" Then Response.Write(" selected ")%>>7:30 PM</option>
								<option value="19:45"<%If userNoActivityNagTimeOfDay = "19:45" Then Response.Write(" selected ")%>>7:45 PM</option>
								<option value="20:00"<%If userNoActivityNagTimeOfDay = "20:00" Then Response.Write(" selected ")%>>8:00 PM</option>
								<option value="20:15"<%If userNoActivityNagTimeOfDay = "20:15" Then Response.Write(" selected ")%>>8:15 PM</option>
								<option value="20:30"<%If userNoActivityNagTimeOfDay = "20:30" Then Response.Write(" selected ")%>>8:30 PM</option>
								<option value="20:45"<%If userNoActivityNagTimeOfDay = "20:45" Then Response.Write(" selected ")%>>8:45 PM</option>
								<option value="21:00"<%If userNoActivityNagTimeOfDay = "21:00" Then Response.Write(" selected ")%>>9:00 PM</option>
								<option value="21:15"<%If userNoActivityNagTimeOfDay = "21:15" Then Response.Write(" selected ")%>>9:15 PM</option>
								<option value="21:30"<%If userNoActivityNagTimeOfDay = "21:30" Then Response.Write(" selected ")%>>9:30 PM</option>
								<option value="21:45"<%If userNoActivityNagTimeOfDay = "21:45" Then Response.Write(" selected ")%>>9:45 PM</option>
								<option value="22:00"<%If userNoActivityNagTimeOfDay = "22:00" Then Response.Write(" selected ")%>>10:00 PM</option>
								<option value="22:15"<%If userNoActivityNagTimeOfDay = "22:15" Then Response.Write(" selected ")%>>10:15 PM</option>
								<option value="22:30"<%If userNoActivityNagTimeOfDay = "22:30" Then Response.Write(" selected ")%>>10:30 PM</option>
								<option value="22:45"<%If userNoActivityNagTimeOfDay = "22:45" Then Response.Write(" selected ")%>>10:45 PM</option>
								<option value="23:00"<%If userNoActivityNagTimeOfDay = "23:00" Then Response.Write(" selected ")%>>11:00 PM</option>
								<option value="23:15"<%If userNoActivityNagTimeOfDay = "23:15" Then Response.Write(" selected ")%>>11:15 PM</option>
								<option value="23:30"<%If userNoActivityNagTimeOfDay = "23:30" Then Response.Write(" selected ")%>>11:30 PM</option>
								<option value="23:45"<%If userNoActivityNagTimeOfDay = "23:45" Then Response.Write(" selected ")%>>11:45 PM</option>	
		 					</select>
						</div>
					</div>
            	
					<div class="row">
                    	<div class="col-lg-12">Send when there has been No Activity for 
							<select class="form-control custom-select" id="seluserNoActivityNagMinutes" name="seluserNoActivityNagMinutes">
								<%
									For x = 15 to 180 Step 5 ' 3 hours
										If x mod 60 = 0 Then
											If x = cint(userNoActivityNagMinutes) Then 
												Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
											else
												Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
											End If
										Else
											If x = cint(userNoActivityNagMinutes) Then 
												Response.Write("<option value='" & x & "' selected>" & x & "</option>")
											Else
												Response.Write("<option value='" & x & "'>" & x & "</option>")
											End If
										End If
									Next
								%>
							</select>&nbsp;minutes
						</div>
					</div>
                
					<div class="row">
						<div class="col-lg-12">Continue to send every
							<select class="form-control custom-select" id="seluserNoActivityNagIntervalMinutes" name="seluserNoActivityNagIntervalMinutes">
								<%
									For x = 10 to 120 Step 5 ' 2 hours
										If x mod 60 = 0 Then
											If x = cint(userNoActivityNagIntervalMinutes) Then 
												Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
											else
												Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
											End If
										Else
											If x = cint(userNoActivityNagIntervalMinutes) Then 
												Response.Write("<option value='" & x & "' selected>" & x & "</option>")
											Else
												Response.Write("<option value='" & x & "'>" & x & "</option>")
											End If
										End If
									Next
								%>
							</select>&nbsp;minutes
						</div>
					</div>
                 
					<div class="row">
						<div class="col-lg-12">Send a maximum of 
							<select class="form-control custom-select" id="seluserNoActivityNagMessageMaxToSendPerStop" name="seluserNoActivityNagMessageMaxToSendPerStop">
								<%
									For x = 1 to 10
										If x = cint(userNoActivityNagMessageMaxToSendPerStop) Then 
											Response.Write("<option value='" & x & "' selected>" & x & "</option>")
										Else
											Response.Write("<option value='" & x & "'>" & x & "</option>")
										End If
									Next
								%>
							</select>&nbsp;messages each time a 'No Activity' event occurs
						</div>
					</div>
                 
					<div class="row">
						<div class="col-lg-12">Send a maxium of 
							<select class="form-control custom-select"  id="seluserNoActivityNagMessageMaxToSendPerDriverPerDay" name="seluserNoActivityNagMessageMaxToSendPerDriverPerDay">
								<%
									For x = 1 to 25
										If x = cint(userNoActivityNagMessageMaxToSendPerDriverPerDay) Then 
											Response.Write("<option value='" & x & "' selected>" & x & "</option>")
										Else
											Response.Write("<option value='" & x & "'>" & x & "</option>")
										End If
									Next
								%>
							</select>&nbsp;messages to a driver on any given day
						</div>
					</div>
                 
					<div class="row">
						<div class="col-lg-12">Send method 
							<select class="form-control custom-select"   id="seluserNoActivityNagMessageSendMethod" name="seluserNoActivityNagMessageSendMethod">
								<option value="Text"<%If userNoActivityNagMessageSendMethod = "Text" Then Response.Write(" selected ")%>>Text Message Only</option>
								<!--<option value="Email"<%If userNoActivityNagMessageSendMethod = "Email" Then Response.Write(" selected ")%>>Email Only</option>
								<option value="TextThenEmail"<%If userNoActivityNagMessageSendMethod = "TextThenEmail" Then Response.Write(" selected ")%>>Text - If no cell number, send email</option>
								<option value="EmailThenText"<%If userNoActivityNagMessageSendMethod = "EmailThenText" Then Response.Write(" selected ")%>>Email - If no valid email address, send text</option>
								<option value="Both"<%If userNoActivityNagMessageSendMethod = "Both" Then Response.Write(" selected ")%>>Both</option>-->
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

