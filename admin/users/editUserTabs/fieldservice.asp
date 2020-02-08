<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="fieldservice">

<%	
	'****************************************
	' This part show all the filter routes if
	'the filter change module is turned on   
	'****************************************	
	If filterChangeModuleOn() Then %>
	<!-- filter routes row !-->
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
					If InStr(","&userFilterRoutes&",",","&rs9("RouteNum")&",") > 0 Then
						checked = " checked='checked'"
					End If
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
<% End If%>	
   
   
<%	

If MUV_READ("SERNO")="1230d" Then                

	'**********************************************
	' If the service module advanced dispatch is on
	' and we are importing service tickets from the
	' backend, this part shows the service route(s)
	' the user should be associated with.
	' Primarily written for US Coffee
	'****************************************	
	If advancedDispatchIsOn() Then %>
	<!-- service routes row !-->
	<div class="col-lg-12" >
		<p><br>Assign imported service tickets with the route(s) below, to this user</p>
		<div class="form-control" style="height: auto; padding: 5px;">
			<% 'Get all Service Routes
			SQL9 = "SELECT DISTINCT OADRIV, TYNAME FROM VAI_VCPRPDL Where OAORDS = 'S'"' Order By Cast(OADRIV as int)"
			Set cnn9 = Server.CreateObject("ADODB.Connection")
			cnn9.open (Session("ClientCnnString"))
			Set rs9 = Server.CreateObject("ADODB.Recordset")
			rs9.CursorLocation = 3 
			Set rs9 = cnn9.Execute(SQL9)
			If not rs9.EOF Then
				Do
				checked = ""
					'If InStr(","&userFilterRoutes&",",","&rs9("RouteNum")&",") > 0 Then
					'	checked = " checked='checked'"
					'End If
					Response.Write("<label class='btn btn-default btn-xs' style='width: 250px; text-align: left; font-size: 14px; margin: 0 3px 3px 0;'>")
					Response.Write("<input "&checked&" type='checkbox' name='txtServiceRoutes' value='"&rs9("OADRIV")&"'> "&rs9("OADRIV")&"&nbsp;" & rs9("TYNAME")&"</label>")
					rs9.movenext
				Loop until rs9.eof
			End If
			set rs9 = Nothing
			cnn9.close
			set cnn9 = Nothing
			%>
		</div>
	</div>
	<% End If%> 
<% End If%>   
   
                           
                                    
	<!-- No activity messages !-->
	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 box-border">
		<div class="row">
           	<div class="col-lg-12">
				Turn on No Activity 'nag' messages
				<select class="form-control custom-select" name="seluserNoActivityNagMessageOverride_FS" id="seluserNoActivityNagMessageOverride_FS">
					<option value="Use Global" <% If userNoActivityNagMessageOverride_FS = "Use Global" Then Response.Write(" selected ")%>>Use Global Setting</option>	
					<option value="Yes" <% If userNoActivityNagMessageOverride_FS = "Yes" Then Response.Write(" selected ")%>>Yes</option>	
					<option value="No" <% If userNoActivityNagMessageOverride_FS = "No" Then Response.Write(" selected ")%>>No</option>	
				</select>
			</div>
		</div>
				
		<!-- "nag" column -->
		<div class="row ">
			<div class="col-lg-12">

				<div class="row">
					<div class="col-lg-12">Start sending messages if there has been No Activity by 
						<select class="form-control custom-select" id="seluserNoActivityNagTimeOfDay_FS" name="seluserNoActivityNagTimeOfDay_FS">			
							<option value="00:00"<%If userNoActivityNagTimeOfDay_FS = "00:00" Then Response.Write(" selected ")%>>-Midnight-</option>
							<option value="00:15"<%If userNoActivityNagTimeOfDay_FS = "00:15" Then Response.Write(" selected ")%>>12:15 AM</option>
							<option value="00:30"<%If userNoActivityNagTimeOfDay_FS = "00:30" Then Response.Write(" selected ")%>>12:30 AM</option>
							<option value="00:45"<%If userNoActivityNagTimeOfDay_FS = "00:45" Then Response.Write(" selected ")%>>12:45 AM</option>
							<option value="1:00"<%If userNoActivityNagTimeOfDay_FS = "1:00" Then Response.Write(" selected ")%>>1:00 AM</option>
							<option value="1:15"<%If userNoActivityNagTimeOfDay_FS = "1:15" Then Response.Write(" selected ")%>>1:15 AM</option>
							<option value="1:30"<%If userNoActivityNagTimeOfDay_FS = "1:30" Then Response.Write(" selected ")%>>1:30 AM</option>
							<option value="1:45"<%If userNoActivityNagTimeOfDay_FS = "1:45" Then Response.Write(" selected ")%>>1:45 AM</option>
							<option value="2:00"<%If userNoActivityNagTimeOfDay_FS = "2:00" Then Response.Write(" selected ")%>>2:00 AM</option>
							<option value="2:15"<%If userNoActivityNagTimeOfDay_FS = "2:15" Then Response.Write(" selected ")%>>2:15 AM</option>
							<option value="2:30"<%If userNoActivityNagTimeOfDay_FS = "2:30" Then Response.Write(" selected ")%>>2:30 AM</option>
							<option value="2:45"<%If userNoActivityNagTimeOfDay_FS = "2:45" Then Response.Write(" selected ")%>>2:45 AM</option>
							<option value="3:00"<%If userNoActivityNagTimeOfDay_FS = "3:00" Then Response.Write(" selected ")%>>3:00 AM</option>
							<option value="3:15"<%If userNoActivityNagTimeOfDay_FS = "3:15" Then Response.Write(" selected ")%>>3:15 AM</option>
							<option value="3:30"<%If userNoActivityNagTimeOfDay_FS = "3:30" Then Response.Write(" selected ")%>>3:30 AM</option>
							<option value="3:45"<%If userNoActivityNagTimeOfDay_FS = "3:45" Then Response.Write(" selected ")%>>3:45 AM</option>
							<option value="4:00"<%If userNoActivityNagTimeOfDay_FS = "4:00" Then Response.Write(" selected ")%>>4:00 AM</option>
							<option value="4:15"<%If userNoActivityNagTimeOfDay_FS = "4:15" Then Response.Write(" selected ")%>>4:15 AM</option>
							<option value="4:30"<%If userNoActivityNagTimeOfDay_FS = "4:30" Then Response.Write(" selected ")%>>4:30 AM</option>
							<option value="4:45"<%If userNoActivityNagTimeOfDay_FS = "4:45" Then Response.Write(" selected ")%>>4:45 AM</option>
							<option value="5:00"<%If userNoActivityNagTimeOfDay_FS = "5:00" Then Response.Write(" selected ")%>>5:00 AM</option>
							<option value="5:15"<%If userNoActivityNagTimeOfDay_FS = "5:15" Then Response.Write(" selected ")%>>5:15 AM</option>
							<option value="5:30"<%If userNoActivityNagTimeOfDay_FS = "5:30" Then Response.Write(" selected ")%>>5:30 AM</option>
							<option value="5:45"<%If userNoActivityNagTimeOfDay_FS = "5:45" Then Response.Write(" selected ")%>>5:45 AM</option>
							<option value="6:00"<%If userNoActivityNagTimeOfDay_FS = "6:00" Then Response.Write(" selected ")%>>6:00 AM</option>
							<option value="6:15"<%If userNoActivityNagTimeOfDay_FS = "6:15" Then Response.Write(" selected ")%>>6:15 AM</option>
							<option value="6:30"<%If userNoActivityNagTimeOfDay_FS = "6:30" Then Response.Write(" selected ")%>>6:30 AM</option>
							<option value="6:45"<%If userNoActivityNagTimeOfDay_FS = "6:45" Then Response.Write(" selected ")%>>6:45 AM</option>
							<option value="7:00"<%If userNoActivityNagTimeOfDay_FS = "7:00" Then Response.Write(" selected ")%>>7:00 AM</option>
							<option value="7:15"<%If userNoActivityNagTimeOfDay_FS = "7:15" Then Response.Write(" selected ")%>>7:15 AM</option>
							<option value="7:30"<%If userNoActivityNagTimeOfDay_FS = "7:30" Then Response.Write(" selected ")%>>7:30 AM</option>
							<option value="7:45"<%If userNoActivityNagTimeOfDay_FS = "7:45" Then Response.Write(" selected ")%>>7:45 AM</option>
							<option value="8:00"<%If userNoActivityNagTimeOfDay_FS = "8:00" Then Response.Write(" selected ")%>>8:00 AM</option>
							<option value="8:15"<%If userNoActivityNagTimeOfDay_FS = "8:15" Then Response.Write(" selected ")%>>8:15 AM</option>
							<option value="8:30"<%If userNoActivityNagTimeOfDay_FS = "8:30" Then Response.Write(" selected ")%>>8:30 AM</option>
							<option value="8:45"<%If userNoActivityNagTimeOfDay_FS = "8:45" Then Response.Write(" selected ")%>>8:45 AM</option>
							<option value="9:00"<%If userNoActivityNagTimeOfDay_FS = "9:00" Then Response.Write(" selected ")%>>9:00 AM</option>
							<option value="9:15"<%If userNoActivityNagTimeOfDay_FS = "9:15" Then Response.Write(" selected ")%>>9:15 AM</option>
							<option value="9:30"<%If userNoActivityNagTimeOfDay_FS = "9:30" Then Response.Write(" selected ")%>>9:30 AM</option>
							<option value="9:45"<%If userNoActivityNagTimeOfDay_FS = "9:45" Then Response.Write(" selected ")%>>9:45 AM</option>
							<option value="10:00"<%If userNoActivityNagTimeOfDay_FS = "10:00" Then Response.Write(" selected ")%>>10:00 AM</option>
							<option value="10:15"<%If userNoActivityNagTimeOfDay_FS = "10:15" Then Response.Write(" selected ")%>>10:15 AM</option>
							<option value="10:30"<%If userNoActivityNagTimeOfDay_FS = "10:30" Then Response.Write(" selected ")%>>10:30 AM</option>
							<option value="10:45"<%If userNoActivityNagTimeOfDay_FS = "10:45" Then Response.Write(" selected ")%>>10:45 AM</option>
							<option value="11:00"<%If userNoActivityNagTimeOfDay_FS = "11:00" Then Response.Write(" selected ")%>>11:00 AM</option>
							<option value="11:15"<%If userNoActivityNagTimeOfDay_FS = "11:15" Then Response.Write(" selected ")%>>11:15 AM</option>
							<option value="11:30"<%If userNoActivityNagTimeOfDay_FS = "11:30" Then Response.Write(" selected ")%>>11:30 AM</option>
							<option value="11:45"<%If userNoActivityNagTimeOfDay_FS = "11:45" Then Response.Write(" selected ")%>>11:45 AM</option>
							<option value="12:00"<%If userNoActivityNagTimeOfDay_FS = "12:00" Then Response.Write(" selected ")%>>-Noon-</option>
							<option value="12:15"<%If userNoActivityNagTimeOfDay_FS = "12:15" Then Response.Write(" selected ")%>>12:15 PM</option>
							<option value="12:30"<%If userNoActivityNagTimeOfDay_FS = "12:30" Then Response.Write(" selected ")%>>12:30 PM</option>
							<option value="12:45"<%If userNoActivityNagTimeOfDay_FS = "12:45" Then Response.Write(" selected ")%>>12:45 PM</option>
							<option value="13:00"<%If userNoActivityNagTimeOfDay_FS = "13:00" Then Response.Write(" selected ")%>>1:00 PM</option>
							<option value="13:15"<%If userNoActivityNagTimeOfDay_FS = "13:15" Then Response.Write(" selected ")%>>1:15 PM</option>
							<option value="13:30"<%If userNoActivityNagTimeOfDay_FS = "13:30" Then Response.Write(" selected ")%>>1:30 PM</option>
							<option value="13:45"<%If userNoActivityNagTimeOfDay_FS = "13:45" Then Response.Write(" selected ")%>>1:45 PM</option>
							<option value="14:00"<%If userNoActivityNagTimeOfDay_FS = "14:00" Then Response.Write(" selected ")%>>2:00 PM</option>
							<option value="14:15"<%If userNoActivityNagTimeOfDay_FS = "14:15" Then Response.Write(" selected ")%>>2:15 PM</option>
							<option value="14:30"<%If userNoActivityNagTimeOfDay_FS = "14:30" Then Response.Write(" selected ")%>>2:30 PM</option>
							<option value="14:45"<%If userNoActivityNagTimeOfDay_FS = "14:45" Then Response.Write(" selected ")%>>2:45 PM</option>
							<option value="15:00"<%If userNoActivityNagTimeOfDay_FS = "15:00" Then Response.Write(" selected ")%>>3:00 PM</option>
							<option value="15:15"<%If userNoActivityNagTimeOfDay_FS = "15:15" Then Response.Write(" selected ")%>>3:15 PM</option>
							<option value="15:30"<%If userNoActivityNagTimeOfDay_FS = "15:30" Then Response.Write(" selected ")%>>3:30 PM</option>
							<option value="15:45"<%If userNoActivityNagTimeOfDay_FS = "15:45" Then Response.Write(" selected ")%>>3:45 PM</option>
							<option value="16:00"<%If userNoActivityNagTimeOfDay_FS = "16:00" Then Response.Write(" selected ")%>>4:00 PM</option>
							<option value="16:15"<%If userNoActivityNagTimeOfDay_FS = "16:15" Then Response.Write(" selected ")%>>4:15 PM</option>
							<option value="16:30"<%If userNoActivityNagTimeOfDay_FS = "16:30" Then Response.Write(" selected ")%>>4:30 PM</option>
							<option value="16:45"<%If userNoActivityNagTimeOfDay_FS = "16:45" Then Response.Write(" selected ")%>>4:45 PM</option>
							<option value="17:00"<%If userNoActivityNagTimeOfDay_FS = "17:00" Then Response.Write(" selected ")%>>5:00 PM</option>
							<option value="17:15"<%If userNoActivityNagTimeOfDay_FS = "17:15" Then Response.Write(" selected ")%>>5:15 PM</option>
							<option value="17:30"<%If userNoActivityNagTimeOfDay_FS = "17:30" Then Response.Write(" selected ")%>>5:30 PM</option>
							<option value="17:45"<%If userNoActivityNagTimeOfDay_FS = "17:45" Then Response.Write(" selected ")%>>5:45 PM</option>
							<option value="18:00"<%If userNoActivityNagTimeOfDay_FS = "18:00" Then Response.Write(" selected ")%>>6:00 PM</option>
							<option value="18:15"<%If userNoActivityNagTimeOfDay_FS = "18:15" Then Response.Write(" selected ")%>>6:15 PM</option>
							<option value="18:30"<%If userNoActivityNagTimeOfDay_FS = "18:30" Then Response.Write(" selected ")%>>6:30 PM</option>
							<option value="18:45"<%If userNoActivityNagTimeOfDay_FS = "18:45" Then Response.Write(" selected ")%>>6:45 PM</option>
							<option value="19:00"<%If userNoActivityNagTimeOfDay_FS = "19:00" Then Response.Write(" selected ")%>>7:00 PM</option>
							<option value="19:15"<%If userNoActivityNagTimeOfDay_FS = "19:15" Then Response.Write(" selected ")%>>7:15 PM</option>
							<option value="19:30"<%If userNoActivityNagTimeOfDay_FS = "19:30" Then Response.Write(" selected ")%>>7:30 PM</option>
							<option value="19:45"<%If userNoActivityNagTimeOfDay_FS = "19:45" Then Response.Write(" selected ")%>>7:45 PM</option>
							<option value="20:00"<%If userNoActivityNagTimeOfDay_FS = "20:00" Then Response.Write(" selected ")%>>8:00 PM</option>
							<option value="20:15"<%If userNoActivityNagTimeOfDay_FS = "20:15" Then Response.Write(" selected ")%>>8:15 PM</option>
							<option value="20:30"<%If userNoActivityNagTimeOfDay_FS = "20:30" Then Response.Write(" selected ")%>>8:30 PM</option>
							<option value="20:45"<%If userNoActivityNagTimeOfDay_FS = "20:45" Then Response.Write(" selected ")%>>8:45 PM</option>
							<option value="21:00"<%If userNoActivityNagTimeOfDay_FS = "21:00" Then Response.Write(" selected ")%>>9:00 PM</option>
							<option value="21:15"<%If userNoActivityNagTimeOfDay_FS = "21:15" Then Response.Write(" selected ")%>>9:15 PM</option>
							<option value="21:30"<%If userNoActivityNagTimeOfDay_FS = "21:30" Then Response.Write(" selected ")%>>9:30 PM</option>
							<option value="21:45"<%If userNoActivityNagTimeOfDay_FS = "21:45" Then Response.Write(" selected ")%>>9:45 PM</option>
							<option value="22:00"<%If userNoActivityNagTimeOfDay_FS = "22:00" Then Response.Write(" selected ")%>>10:00 PM</option>
							<option value="22:15"<%If userNoActivityNagTimeOfDay_FS = "22:15" Then Response.Write(" selected ")%>>10:15 PM</option>
							<option value="22:30"<%If userNoActivityNagTimeOfDay_FS = "22:30" Then Response.Write(" selected ")%>>10:30 PM</option>
							<option value="22:45"<%If userNoActivityNagTimeOfDay_FS = "22:45" Then Response.Write(" selected ")%>>10:45 PM</option>
							<option value="23:00"<%If userNoActivityNagTimeOfDay_FS = "23:00" Then Response.Write(" selected ")%>>11:00 PM</option>
							<option value="23:15"<%If userNoActivityNagTimeOfDay_FS = "23:15" Then Response.Write(" selected ")%>>11:15 PM</option>
							<option value="23:30"<%If userNoActivityNagTimeOfDay_FS = "23:30" Then Response.Write(" selected ")%>>11:30 PM</option>
							<option value="23:45"<%If userNoActivityNagTimeOfDay_FS = "23:45" Then Response.Write(" selected ")%>>11:45 PM</option>	
	 					</select>
					</div>
				</div>
        	
				<div class="row">
                	<div class="col-lg-12">Send when there has been No Activity for 
						<select class="form-control custom-select" id="seluserNoActivityNagMinutes_FS" name="seluserNoActivityNagMinutes_FS">
							<%
								For x = 15 to 180 Step 5 ' 3 hours
									If x mod 60 = 0 Then
										If x = cint(userNoActivityNagMinutes_FS) Then 
											Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
										else
											Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
										End If
									Else
										If x = cint(userNoActivityNagMinutes_FS) Then 
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
						<select class="form-control custom-select" id="seluserNoActivityNagIntervalMinutes_FS" name="seluserNoActivityNagIntervalMinutes_FS">
							<%
								For x = 10 to 120 Step 5 ' 2 hours
									If x mod 60 = 0 Then
										If x = cint(userNoActivityNagIntervalMinutes_FS) Then 
											Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
										else
											Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
										End If
									Else
										If x = cint(userNoActivityNagIntervalMinutes_FS) Then 
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
						<select class="form-control custom-select" id="seluserNoActivityNagMessageMaxToSendPerStop_FS" name="seluserNoActivityNagMessageMaxToSendPerStop_FS">
							<%
								For x = 1 to 10
									If x = cint(userNoActivityNagMessageMaxToSendPerStop_FS) Then 
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
						<select class="form-control custom-select"  id="seluserNoActivityNagMessageMaxToSendPerDriverPerDay_FS" name="seluserNoActivityNagMessageMaxToSendPerDriverPerDay_FS">
							<%
								For x = 1 to 25
									If x = cint(userNoActivityNagMessageMaxToSendPerDriverPerDay_FS) Then 
										Response.Write("<option value='" & x & "' selected>" & x & "</option>")
									Else
										Response.Write("<option value='" & x & "'>" & x & "</option>")
									End If
								Next
							%>
						</select>&nbsp;messages to a technician on any given day
					</div>
				</div>
             
				<div class="row">
					<div class="col-lg-12">Send method 
						<select class="form-control custom-select"   id="seluserNoActivityNagMessageSendMethod_FS" name="seluserNoActivityNagMessageSendMethod_FS">
							<option value="Text"<%If userNoActivityNagMessageSendMethod_FS = "Text" Then Response.Write(" selected ")%>>Text Message Only</option>
							<!--<option value="Email"<%If userNoActivityNagMessageSendMethod_FS = "Email" Then Response.Write(" selected ")%>>Email Only</option>
							<option value="TextThenEmail"<%If userNoActivityNagMessageSendMethod_FS = "TextThenEmail" Then Response.Write(" selected ")%>>Text - If no cell number, send email</option>
							<option value="EmailThenText"<%If userNoActivityNagMessageSendMethod_FS = "EmailThenText" Then Response.Write(" selected ")%>>Email - If no valid email address, send text</option>
							<option value="Both"<%If userNoActivityNagMessageSendMethod_FS = "Both" Then Response.Write(" selected ")%>>Both</option>-->
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