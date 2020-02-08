<!-- Tab starts here -->
<div role="tabpanel" class="tab-pane fade" id="loginaccess">

		<div class="col-xs-7 col-sm-7 col-md-7 col-lg-7 enable-disable">
			<strong>Click on Box Change Color To Allow/Deny Login Access For That Day/Hour</strong>
			
			<table style="margin-top:5px; margin-bottom:20px;">
				<tr><td style="background-color:#3498DB; padding:10px; color: #fff; text-align:center">LOGIN ALLOWED</td>
				<td style="background-color:#E74C3C; padding:10px; color: #fff; text-align:center">LOGIN RESTRICTED</td></tr>
			</table>
			
			<div id="day-schedule"></div>

		</div>
 
		<div class="col-xs-5 col-sm-5 col-md-5 col-lg-5 enable-disable">
			<strong>Do not allow login during schedule holidays/times</strong><br>(as defined in the Company Calendar)
			<% If userLoginDisableAccessHolidays = vbTrue Then %>
				<input type="checkbox" checked="checked" id="chkDisableHolidayLogins" name="chkDisableHolidayLogins">
			<% Else %>
				<input type="checkbox" id="chkDisableHolidayLogins" name="chkDisableHolidayLogins">		    
			<%End If%>

			<br><br>
			<button type="button" class="btn btn-primary" id="clearDayScheduler"><i class="fa fa-eraser"></i> Clear All Selections</button>
			<br><br>
			<div id="day-schedule-selected"></div>
			
		</div>
   
</div>

<%
	'********************************************************************
	'Get the current restricted login days/times from SQL
	'********************************************************************
		
	Set cnnGetCurrentRestrictions = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentRestrictions.open (Session("ClientCnnString"))
	Set rsGetCurrentRestrictions = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentRestrictions.CursorLocation = 3 

	Set cnnLoopThroughRestrictions = Server.CreateObject("ADODB.Connection")
	cnnLoopThroughRestrictions.open (Session("ClientCnnString"))
	Set rsLoopThroughRestrictions = Server.CreateObject("ADODB.Recordset")
	rsLoopThroughRestrictions.CursorLocation = 3 
	
	jsonDefaultSelectedDays = ""

	For LoopDayNo = 0 to 6
	
		SQLGetCurrentRestrictions = "SELECT InternalRecordIdentifier, UserNo, DayNo, StartRestrictedTime, EndRestrictedTime, "
		SQLGetCurrentRestrictions = SQLGetCurrentRestrictions & "(SELECT COUNT(*) AS Expr1 "
		SQLGetCurrentRestrictions = SQLGetCurrentRestrictions & " FROM SC_UserRestrictedLoginSchedule "
		SQLGetCurrentRestrictions = SQLGetCurrentRestrictions & " WHERE (DayNo = " & LoopDayNo & ")) AS DayNoCount "
		SQLGetCurrentRestrictions = SQLGetCurrentRestrictions & " FROM  SC_UserRestrictedLoginSchedule AS SC_UserRestrictedLoginSchedule_1 "
		SQLGetCurrentRestrictions = SQLGetCurrentRestrictions & " WHERE (userNo = " & UserNo & " AND DayNo = " & LoopDayNo & ") ORDER BY InternalRecordIdentifier"

		'Response.write(SQLGetCurrentRestrictions)
		
		Set rsGetCurrentRestrictions = cnnGetCurrentRestrictions.Execute(SQLGetCurrentRestrictions)
		
		If NOT rsGetCurrentRestrictions.EOF Then
		
			'Do While NOT rsGetCurrentRestrictions.EOF
			
				NumberOfTimeRangesThisDay = rsGetCurrentRestrictions("DayNoCount")
				
				
				If NumberOfTimeRangesThisDay > 1 Then
				
					'This day has multiple restriction ranges
					
					jsonDefaultSelectedDaysTimeSlot = ""

					''''''''''0': [['09:30', '11:00'], ['13:00', '16:30']],  FOR EXAMPLE	
					
	
					SQLLoopThroughRestrictions = "SELECT * FROM SC_UserRestrictedLoginSchedule "
					SQLLoopThroughRestrictions = SQLLoopThroughRestrictions & " WHERE (userNo = " & UserNo & " AND DayNo = " & LoopDayNo & ") "
					SQLLoopThroughRestrictions = SQLLoopThroughRestrictions & " ORDER BY InternalRecordIdentifier"
		
					'Response.write("<strong>SQLLoopThroughRestrictions</strong>: " & SQLLoopThroughRestrictions & "<br><br>")
					
					Set rsLoopThroughRestrictions = cnnLoopThroughRestrictions.Execute(SQLLoopThroughRestrictions)
				
					If NOT rsLoopThroughRestrictions.EOF Then
			
						Do While NOT rsLoopThroughRestrictions.EOF
						
							StartRestrictedTime = rsLoopThroughRestrictions("StartRestrictedTime")
							EndRestrictedTime = rsLoopThroughRestrictions("EndRestrictedTime")
						
							jsonDefaultSelectedDaysTimeSlot = jsonDefaultSelectedDaysTimeSlot & "['" & StartRestrictedTime & "', '" & EndRestrictedTime & "']" & ","
							
							rsLoopThroughRestrictions.MoveNext
						Loop
					
					End If
					
					jsonDefaultSelectedDaysTimeSlot = Left(jsonDefaultSelectedDaysTimeSlot,len(jsonDefaultSelectedDaysTimeSlot)-1)
					jsonDefaultSelectedDays = jsonDefaultSelectedDays & "'" & LoopDayNo & "' : [" & jsonDefaultSelectedDaysTimeSlot  & "],"
					
			
				Else
				
					'This day has one restriction range
					'''''''''''6': [['00:00', '24:00']] FOR EXAMPLE					
					StartRestrictedTime = rsGetCurrentRestrictions("StartRestrictedTime")
					EndRestrictedTime = rsGetCurrentRestrictions("EndRestrictedTime")
					jsonDefaultSelectedDays = jsonDefaultSelectedDays & "'" & LoopDayNo & "' : [['" & StartRestrictedTime & "', '" & EndRestrictedTime & "']]" & ","
					
				End If
				
				
			'rsGetCurrentRestrictions.MoveNext
			'Loop
		End If
		
	Next
	
	'remove last comma from the string
	If jsonDefaultSelectedDays <> "" Then
		jsonDefaultSelectedDays = Left(jsonDefaultSelectedDays,len(jsonDefaultSelectedDays)-1)
	End If

	cnnGetCurrentRestrictions.close
	cnnLoopThroughRestrictions.close

%>
  <script src="../../../js/DayScheduleSelector.js"></script>
  <script>
    (function ($) {
 
		  $( "#clearDayScheduler" ).click(function() {
		  
		  	$("#day-schedule").data('artsy.dayScheduleSelector').clear();
		  	
		  	var userNo = $("#txtUserNo").val();
		  	
			$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForAdminSettings.asp",
				data: "action=ClearLoginAccessForExistingUser&userNo="+encodeURIComponent(userNo),
				async: true,
				success: function(msg){
		            if(msg == "success"){
		            } 
				}
			});
		  	
		  });   
		  
		  
	      $("#day-schedule").dayScheduleSelector({
	        /* NOTE: These are currently being set directly in the JS file
	        days: [1, 2, 3, 5, 6],
	        interval: 15,
	        startTime: '09:50',
	        endTime: '21:06'
	        */
	      });
	      
	      //These are the default days that would show login access denied
	      //This string is built from the SC_UserRestrictedLoginSchedule table entries, if any
	      
	      //IMPORTANT NOTE: THIS FUNCTION IS NOT USED IN ADD NEW USER BECAUSE THERE ARE NO
	      //DFEAULT RESTRICTED TIMES SET UP IN THE DATABASE YET
	      $("#day-schedule").data('artsy.dayScheduleSelector').deserialize({
	        //'0': [['09:30', '11:00'], ['13:00', '16:30']],
	        //'0': [['00:00', '24:00']],
	        //'6': [['00:00', '24:00']]
	        <%= jsonDefaultSelectedDays %>
	      });
	      
		  $("#day-schedule").on('selected.artsy.dayScheduleSelector', function (e, selected) {
				  /* selected is an array of time slots selected this time. */
				  
				 //$('#day-schedule-selected').html("");
				 
		        hours = $("#day-schedule").data('artsy.dayScheduleSelector').serialize();
		
		        json = '[';
		
		        for (i = 0; i <= 6; i++) {           
		            for (y = 0; y < hours[i].length; y++) {
		                json += '{"day":'+i+',"start":"'+hours[i][y][0]+'","end":"'+hours[i][y][1]+'"},';
		            }
		        }
		        json = json.slice(0, -1) + ']';
		        
		        //$('#day-schedule-selected').append(json + "<br>");
		        
		        var userNo = $("#txtUserNo").val();
		        
				$.ajax({
					type:"POST",
					url: "../../inc/InSightFuncs_AjaxForAdminSettings.asp",
					data: "action=UpdateLoginAccessForExistingUser&userNo="+encodeURIComponent(userNo)+"&jsonString="+encodeURIComponent(json),
					async: true,
					success: function(msg){
					}
				});
			        
		
				var obj = $.parseJSON(json);
				$(obj).each(function(i,val){
				    $.each(val,function(key,value){
				        console.log(key+" : "+ value);
				        //$('#day-schedule-selected').append(key+" : "+ value + "<br>");
				});
			
				
			});
			
		
  
	  });
	       
    })($);
  </script>

<!-- Tab ends here -->