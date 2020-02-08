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
			
		</div>
   
</div>

<%

%>
  <script src="../../../js/DayScheduleSelector.js"></script>
  <script>
    (function ($) {
 
		  $( "#clearDayScheduler" ).click(function() {
		  
		  	$("#day-schedule").data('artsy.dayScheduleSelector').clear();
		  	
			$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForAdminSettings.asp",
				data: "action=ClearLoginAccessNewForUser",
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
	      	      
		  $("#day-schedule").on('selected.artsy.dayScheduleSelector', function (e, selected) {
				  /* selected is an array of time slots selected this time. */
				 
				 
		        hours = $("#day-schedule").data('artsy.dayScheduleSelector').serialize();
		
		        json = '[';
		
		        for (i = 0; i <= 6; i++) {           
		            for (y = 0; y < hours[i].length; y++) {
		                json += '{"day":'+i+',"start":"'+hours[i][y][0]+'","end":"'+hours[i][y][1]+'"},';
		            }
		        }
		        json = json.slice(0, -1) + ']';
		        
		        
				$.ajax({
					type:"POST",
					url: "../../inc/InSightFuncs_AjaxForAdminSettings.asp",
					data: "action=UpdateLoginAccessForNewUser&jsonString="+encodeURIComponent(json),
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