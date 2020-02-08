<!--#include file="../inc/header-prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->


<% 

txtInternalRecordIdentifier = Request.QueryString("i")
ProspectName = GetProspectNameByNumber(txtInternalRecordIdentifier)

'Read edit prospect tab color settings
SQL = "SELECT * FROM Settings_Global"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	CRMTileActivityColor = rs("CRMTileActivityColor")
	CRMTileStageColor = rs("CRMTileStageColor")
	CRMTileOwnerColor = rs("CRMTileOwnerColor")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing


If CRMTileActivityColor = "" Then CRMTileActivityColor = "#f1c40f"
If IsNull(CRMTileActivityColor) Then CRMTileActivityColor = "#f1c40f"

If CRMTileStageColor = "" Then CRMTileStageColor = "#e67e22"
If IsNull(CRMTileStageColor) Then CRMTileStageColor = "#e67e22"

If CRMTileOwnerColor = "" Then CRMTileOwnerColor = "#95a5a6"
If IsNull(CRMTileOwnerColor) Then CRMTileOwnerColor = "#95a5a6"


MaxActivityDaysWarningInit = GetCRMMaxActivityDaysWarning()
MaxActivityDaysPermittedInit = GetCRMMaxActivityDaysPermitted()

%>


<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.min.js"></script>
<link href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet">

<script>


	$(document).ready(function() {
	
		$('input, select, textarea').each(
		    function(index){  
		        var input = $(this);
		        //console.log(input.attr('name'));
		    }
		); 

       var maxDaysPermitted = '<%= MaxActivityDaysPermittedInit %>';
       maxDaysPermitted = parseInt(maxDaysPermitted);

		//In Global Settings, Zero Means User Can Pick Any Date In The Future
		if (maxDaysPermitted != 0) {
	        $('#datetimepickerNextActivity').datetimepicker({
	        	useCurrent: false,
	        	minDate:moment(),
	        	maxDate:moment().add('<%= MaxActivityDaysPermittedInit %>', 'days'),
                format: 'MM/DD/YYYY hh:mm A',
                ignoreReadonly: true,
                sideBySide: true,
	
			});    
		}
		else {
	        $('#datetimepickerNextActivity').datetimepicker({
	        	useCurrent: false,
	            minDate:moment(),
                format: 'MM/DD/YYYY hh:mm A',
                ignoreReadonly: true,
                sideBySide: true, 
	        	
			});      

		}  


		
		//Initially, hide both divs that show either the appointment or meeting fields
		$("#showEmailNewOwnerCheckbox").hide();
		$("#activityDateWarning").hide();
	    $("#showActivityAppointmentDuration").hide();
	    $("#showActivityMeetingDuration").hide();
	    $("#showActivityMeetingLocation").hide();


	    //When a user changes the new owner, determine whether to show checkbox to not send
	    //accept/reject prospect ownership email
	    
		$("#selNewProspectOwner").change(function() {
			
			myProspectID = $("#txtInternalRecordIdentifier").val(); 
			newOwnerUserNo = $("#selNewProspectOwner").val();
			
       		$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=CheckIfSelectedOwnerIsNotCurrentUser&myProspectID=" + encodeURIComponent(myProspectID)+ "&newOwnerUserNo=" + encodeURIComponent(newOwnerUserNo),
				success: function(response)
				 {
				 	if (response == "1") {
				 		$("#showEmailNewOwnerCheckbox").show();
				 	}
				 	else {
				 		$("#showEmailNewOwnerCheckbox").hide();
				 	}
	             }
			});	//end ajax post to data: "action=CheckIfSelectedOwnerIsNotCurrentUser"
			
		}); 
	
	
   			    
	    //When a user changes the next activity, there are several ajax posts that have to be made to determine
	    //whether or not to show a meeting or appointment, based on user type and activity type
	    
		$("#selProspectNextActivity").change(function() {

		    
		    //Hide both divs that show either the appointment or meeting fields whenever the customer changes
		    //a next activity, until we know what to display, if anything
		    
		    $("#showActivityAppointmentDuration").hide();
		    $("#showActivityMeetingDuration").hide();	
		    $("#showActivityMeetingLocation").hide();			

			//First, make an ajax post to determine whether or not this user's Outlook Calendar gets updated
			//when an activity change is made
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetAllowActivityUpdatesToUsersCalendarForModal",
				success: function(response)
				 {
				 	//if allowUpdatesToUsersCalendar is true, then we have to determine if we show a meeting or
				 	//appointment information in the modal. This is based on the next activity selected.
	               	
	               	if (response == 'True') {
	               	 
	               	 	//get the ID of the next activity that the user selected
	               	 	
	               	 	newActivityRecID = $("#selProspectNextActivity").val();
	               	 	
	               	 	//Now make a second ajax post here to check to show meeting or appointment div, or no div at all, 
	               	 	//based on the ID of the selected next activity
	               	 	
	               	 	$.ajax({
							type:"POST",
							url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
							cache: false,
							data: "action=GetActivityCalendarApptOrMeetingForModal&myActivityRecID=" + encodeURIComponent(newActivityRecID),
							success: function(response2)
							 {
							 	activityCalendarShowApptOrMeeting = response2;
							 	
							 	//If the returned value for the activity is 'Appointment', display the appointment div input fields
							 	
				               	if (activityCalendarShowApptOrMeeting == 'Appointment') 
				               	{
				               		$("#showActivityAppointmentDuration").show();
				               	}
				               	
				               	//If the returned value for the activity is 'Meeting', display the meeeting div input fields
				               	
				               	else if (activityCalendarShowApptOrMeeting == 'Meeting')  
				               	{
				               		myProspectID = $("#txtInternalRecordIdentifier").val();
				               		
				               		
				               		//If the activity is a 'Meeting' then we need to make a third ajax post to determine
				               		//the default location for this meeting. This comes from PR_Prospects
				               		
				               		$.ajax({
										type:"POST",
										url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
										cache: false,
										data: "action=GetMeetingLocationForModal&myProspectID=" + encodeURIComponent(myProspectID),
										success: function(response3)
										 {
										 	//Show the meeting div input fields and set the default value of the location textbox with the 
										 	//address information returned from the ajax post
										 	
										 	$("#showActivityMeetingDuration").show();
										 	$("#showActivityMeetingLocation").show();
							               	$("#txtMeetingLocation").val(response3);               	 
							             },
							            failure: function(response3)
										 {
										  	//If no address infomation was returned, just show the meeting div input fields, and do not
										  	//set the default value of the meeting location
										   	$("#showActivityMeetingDuration").show();
										   	$("#showActivityMeetingLocation").show();
							             }

									});	//end ajax post to data: "action=GetMeetingLocationForModal"
				               		
				               	}
				               	else {
				               			//Else, the activity is not a 'Meeting' or an 'Appointment' so make sure the divs are hidden
										$("#showActivityAppointmentDuration").hide();
										$("#showActivityMeetingDuration").hide();
										$("#showActivityMeetingLocation").hide();
											
				               	}// end if statement for activityCalendarShowApptOrMeeting 

           	 
				             }  //end success function for ajax post {show meeting or appointment for this activity}

						}); //end ajax post to data: "action=GetActivityCalendarApptOrMeetingForModal" {show meeting or appointment for this activity}
						
	               	 }	//end if for if (response == 'True') {user calendar gets updated with an activity change}
	               	 
					else{
						$("#showActivityAppointmentDuration").hide();
						$("#showActivityMeetingDuration").hide();	
						$("#showActivityMeetingLocation").hide();					

					}	      
	               	          	 
	             } //end success function for ajax post {user calendar gets updated with an activity change}
	             
			});//end ajax post to data: "action=GetAllowActivityUpdatesToUsersCalendarForModal"
			
		});	// end $("#selProspectNextActivity").change(function()	
				
				
				
				
        $("#datetimepickerNextActivity").on("dp.change", function(e) {

           var maxDaysWarning = '<%= MaxActivityDaysWarningInit %>';
           maxDaysWarning = parseInt(maxDaysWarning);
           
           
           //In Global Settings, Zero Means to Not Show Warning
           if (maxDaysWarning != 0) {
				   var selectedDateFromPicker = moment($("#datetimepickerNextActivity").find("input").val());
		
					var now = moment(new Date()); //todays date
					var duration = moment.duration(now.diff(selectedDateFromPicker));
					var activityDaysDifference = duration.asDays();
		
		           if (Math.abs(activityDaysDifference) > maxDaysWarning){
		           		$("#activityDateWarning").show();
		           		
		           }
		           else {
		           		$("#activityDateWarning").hide();
		           }
		     }
           
        });
        
			  
     
	});
</script>

<script language="JavaScript">
<!--

   function validateRecycleProspectForm()
    {

       if (document.frmRecycleProspect.selProspectNextActivity.value == "") {
            swal("Next activity cannot be blank.");
            return false;
       }
       
       if (document.frmRecycleProspect.txtNextActivityDueDate.value == "") {
            swal("Next activity due date cannot be blank.");
            return false;
       }				 
 
	   var radio = document.getElementsByName('radStage'); // get all radio buttons
	   var isChecked = 0; // default is 0 
	   for(var i=0; i<radio.length;i++) { // go over all the radio buttons with name 'radStage'
			if(radio[i].checked) isChecked = 1; // if one of them is checked - tell me
		}

 		if(isChecked == 0) { // if the default value stayed the same, check the first radio button
   			swal("Please select a stage.");
   			return false;
   		}
      
       return true;

    }
// -->
</script>   

<style type="text/css">

/*Colored Content Boxes
------------------------------------*/

	.container{
		width: 100%;
	}
	
	.quick-info-block {
	  padding: 3px 20px;
	  text-align: center;
	  margin-bottom: 20px;
	  border-radius: 7px;
	}
	
	.quick-info-block p{
	  color: #555;
	  font-size:14px;
	  text-align:left;
	}
	.quick-info-block h2 {
	  color: #fff;
	  font-size:20px;
	  margin-bottom:25px;
	}

	.quick-info-block h2.black {
	  color: #000;
	  font-size:20px;
	  margin-bottom:25px;
	}
	
	.quick-info-block h2 a:hover{
	  text-decoration: none;
	}
	
	.quick-info-block-light,
	.quick-info-block-default {
	  background: #fafafa;
	  border: solid 1px #eee; 
	}
	
	.quick-info-block-default:hover {
	  box-shadow: 0 0 8px #eee;
	}
	
	.quick-info-block-light p,
	.quick-info-block-light h2,
	.quick-info-block-default p,
	.quick-info-block-default h2 {
	  color: #555;
	}

	.quick-info-block-u {
	  background: #72c02c;
	}
	.quick-info-block-blue {
	  background: #3498db;
	}
	.quick-info-block-red {
	  background: #e74c3c;
	}
	.quick-info-block-sea {
	  background: #1abc9c;
	}
	.quick-info-block-grey {
	  background: #f8f8f8;
	}
	.quick-info-block-yellow {
	  background: #f1c40f;
	}
	.quick-info-block-orange {
	  background: #e67e22;
	}
	.quick-info-block-green {
	  background: #2ecc71;
	}
	.quick-info-block-purple {
	  background: #9b6bcc;
	}
	.quick-info-block-aqua {
	  background: #27d7e7;
	}
	.quick-info-block-brown {
	  background: #9c8061;
	}
	.quick-info-block-dark-blue {
	  background: #4765a0;
	}
	.quick-info-block-light-green {
	  background: #79d5b3;
	}
	.quick-info-block-dark {
	  background: #555;
	}
	.quick-info-block-light {
	  background: #ecf0f1;
	}
	
	textarea.form-control {
    	height: 100px; !important;
    	width:525px !important;
    	border-radius:3px !important;
	}
		
	hr.tile {
	    border: 0;
	    height: 3px;
	    background-image: linear-gradient(to right, rgba(0, 0, 0, 0), rgba(255, 255, 255, 0.95), rgba(0, 0, 0, 0));
	}

	.CRMTileActivityColor {
		<% Response.Write("background:" & CRMTileActivityColor & " !important;") %>
	}
	.CRMTileStageColor {
		<% Response.Write("background:" & CRMTileStageColor & " !important;") %>
	}
	.CRMTileOwnerColor {
		<% Response.Write("background:" & CRMTileOwnerColor & " !important;") %>
	}

	.red-line{
		border-left:3px solid red;
	}   

	 .radio {
	    position: relative;
	    display: inline;
	    margin-top: 10px;
	    margin-bottom: 20px;
	    color: #000;
		}
	  .radio .btn,
	  .radio-inline .btn {
	    padding-left: 2em;
	    min-width: 7em;
	    margin-top: 10px;
	    margin-left: 5px;
	  }
	 
	  .radio label,
	  .radio-inline label {
	    text-align: left;
	    padding-left: 0.5em;
	    color: #000;
	  }
	  
	input[type=checkbox] {
	  transform:scale(1.5, 1.5);
	}	  

</style>
<!-- eof css !-->
<h1 class="page-header"><i class="fa fa-fw fa-recycle"></i> <%= GetTerm("Recycle") %> Inactive Prospect - <%= ProspectName %>
	<!-- customize !-->
	<div class="col pull-right">
	</div>
	<!-- eof customize !-->
</h1>

		
<form autocomplete="off" action="<%= BaseURL %>prospecting/recycleInactiveProspect_submit.asp" method="POST" name="frmRecycleProspect" id="frmRecycleProspect" onsubmit="return validateRecycleProspectForm();" class="form-horizontal track-event-form bv-form">

<input autocomplete="false" name="hidden" type="text" style="display:none;">
<input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= txtInternalRecordIdentifier %>">

<div class="container pull-left">
	<div class="row">
	
      <div class="col-md-4">
 		<div class="quick-info-block CRMTileOwnerColor">
		<h2 class="heading-md"><i class="icon-2x color-light fa fa-user"></i>&nbsp;<%= GetTerm("Owner") %></h2>
		
		<hr class="tile">

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                <%
				      	CurrentOwnerUserNo = GetProspectOwnerNoByNumber(txtInternalRecordIdentifier)
				      	CurrentOwnerName = GetUserFirstAndLastNameByUserNo(CurrentOwnerUserNo)
	                %>
	                <p><strong>Previous Owner</strong>: <%= CurrentOwnerName %></p>
	                <input type="hidden" name="txtOrigOwnerUserNo" id="txtOrigOwnerUserNo" value="<%= CurrentOwnerUserNo %>">
	                
	                <hr class="tile">
	                
	                  <p>Choose New Owner Below (<strong>You Are Selected By Default</strong>):</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
                    		<select data-placeholder="Choose Owner" class="C_Country_Modal form-control" id="selNewProspectOwner" name="selNewProspectOwner"> 
							<option value="<%=Session("UserNo")%>"><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option>
					      	<%'Owner dropdown
				      	  	SQL = "SELECT UserNo, userFirstName, userLastName, userType FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
				      	  	SQL = SQL & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo") & " AND userEnabled = 1 AND "
							SQL = SQL & "(userType='Outside Sales' OR userType='Outside Sales Manager' OR userType='Admin' "
							SQL = SQL & "OR userType='Inside Sales' OR userType='Inside Sales Manager' OR userType='CSR' "
							SQL = SQL & "OR userType='CSR Manager') "
				      	  	SQL = SQL & "ORDER BY userFirstName, userLastName"
			
							Set cnn8 = Server.CreateObject("ADODB.Connection")
							cnn8.open (Session("ClientCnnString"))
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.CursorLocation = 3 
							Set rs = cnn8.Execute(SQL)
						
							If not rs.EOF Then
								Do
									FullName = rs("userFirstName") & " " & rs("userLastName")
									Response.Write("<option value='" & rs("UserNo") & "'>" & FullName & "</option>")
									rs.movenext
								Loop until rs.eof
							End If
							set rs = Nothing
							cnn8.close
							set cnn8 = Nothing
		      				%>
						</select>

	                   </div>
	                </div> 
	                	 
               </div>
               
               <div class="form-group" id="showEmailNewOwnerCheckbox" style="display:none;">
	                <div class="col-sm-12">
	                  <p>Do Not Send Accept/Reject Email To New Owner:&nbsp;&nbsp;<input type="checkbox" name="chkDoNotEmailNewOwner" id="chkDoNotEmailNewOwner"></p>
	                </div> 	 
               </div>
               

   			</div>

   			
   						
		<div class="quick-info-block CRMTileActivityColor">
		<h2 class="heading-md"><i class="icon-2x color-light fa fa-arrow-right"></i>&nbsp;<%= GetTerm("Next Activity") %></h2>
	
		<hr class="tile">
	
			<%	
			MaxActivityDaysWarning = GetCRMMaxActivityDaysWarning()
			MaxActivityDaysPermitted = GetCRMMaxActivityDaysPermitted()
			%>
			
			<input type="hidden" name="txtCRMMaxActivityDaysWarning" id="txtCRMMaxActivityDaysWarning" value="<%= MaxActivityDaysWarning %>">
			<input type="hidden" name="txtCRMMaxActivityDaysPermitted" id="txtCRMMaxActivityDaysPermitted" value="<%= MaxActivityDaysPermitted %>">
	
              <div class="form-group">        	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-cog"></i></div>
                    		<select data-placeholder="Choose Next Activity" class="C_Country_Modal form-control red-line" id="selProspectNextActivity" name="selProspectNextActivity"> 
							<option value="">Select Next Activity</option>
					      	<% 
					      	  	SQLNextActivity = "SELECT * FROM PR_Activities ORDER BY Activity"
	
								Set cnnNextActivity = Server.CreateObject("ADODB.Connection")
								cnnNextActivity.open (Session("ClientCnnString"))
								Set rsNextActivity = Server.CreateObject("ADODB.Recordset")
								rsNextActivity.CursorLocation = 3 
								Set rsNextActivity = cnnNextActivity.Execute(SQLNextActivity)
								If not rsNextActivity.EOF Then
									Do
										Response.Write("<option value='" & rsNextActivity("InternalRecordIdentifier") & "'>" & rsNextActivity("Activity")& "</option>")
										rsNextActivity.movenext
									Loop until rsNextActivity.eof
								End If
								set rsNextActivity = Nothing
								cnnNextActivity.close
								set cnnNextActivity = Nothing
							%>
							</select>

	                   </div>
	                </div> 	 
               </div>


               <div class="form-group" id="showActivityAppointmentDuration">
	                <div class="col-sm-12">
	                  <p style="text-align:left">Appointment Duration (for Outlook Calendar):</p>
						<!-- Get Default Appointment Duration from tblGlobalSettings -->
						<%
							EWSDefaultAppointmentDuration = GetPOSTParams("EWSDEFAULTAPPTDURATION")
						%>
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-clock-o" aria-hidden="true"></i></div>
						  	<select data-placeholder="Choose Appointment Duration" class="C_Country_Modal form-control" name="selAppointmentDuration" id="selAppointmentDuration">
								<%For x = 15 to 180 Step 5
									If x mod 60 = 0 Then
										If x = cint(EWSDefaultAppointmentDuration) Then 
											Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
										else
											Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
										End If
									Else
										If x = cint(EWSDefaultAppointmentDuration) Then 
											Response.Write("<option value='" & x & "' selected>" & x & " minutes</option>")
										Else
											Response.Write("<option value='" & x & "'>" & x & " minutes</option>")
										End If
									End If
								Next %>											
							</select>
	                   </div>
	                </div> 	 
               </div>
						

               <div class="form-group" id="showActivityMeetingDuration">
	                <div class="col-sm-12">
	                  <p style="text-align:left">Meeting Duration (for Outlook Calendar):</p>
					  <!-- Get Default Meeting Duration from tblGlobalSettings -->
					  <%
						EWSDefaultMeetingDuration = GetPOSTParams("EWSDEFAULTMEETINGDURATION")
					  %>
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-clock-o" aria-hidden="true"></i></div>
						  	<select data-placeholder="Choose Meeting Duration" class="C_Country_Modal form-control" name="selMeetingDuration" id="selMeetingDuration">
								<%For x = 15 to 300 Step 15
									If x mod 60 = 0 Then
										If x = cint(EWSDefaultMeetingDuration) Then 
											Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
										else
											Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
										End If
									Else
										If x = cint(EWSDefaultMeetingDuration) Then 
											Response.Write("<option value='" & x & "' selected>" & x & " minutes</option>")
										Else
											Response.Write("<option value='" & x & "'>" & x & " minutes</option>")
										End If
									End If
								Next %>												
							</select>
	                   </div>
	                </div> 	 
               </div>
              
				
				<div class="form-group" id="showActivityMeetingLocation">
	                <div class="col-sm-12">
	                  <p style="text-align:left">Meeting Location (for Outlook Calendar):</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card"></i></div>
	                    	<input type="text" class="form-control" id="txtMeetingLocation" placeholder="Meeting Location" name="txtMeetingLocation">
	                   </div>
	                </div> 
	           </div>    
	           					
				<div class="form-group">
					<div class="col-lg-12" style="margin-top:15px;" id="activityDateWarning" style="display:none">
						<div class="alert alert-danger">
						  <strong>Warning!</strong> This activity has been schedule beyond the recommended limit.
						</div>	
					</div>	
				</div>	
				
				<div class="form-group">
					<div class="col-lg-12">
						<div class="col-sm-4"><p style="margin-left:-14px; text-align:left">Next Activity Due Date</p></div>
						<div class="col-sm-8">
					        <div class="input-group date" id="datetimepickerNextActivity">
					            <input type="text" class="form-control red-line" id="txtNextActivityDueDate" name="txtNextActivityDueDate" placeholder="Click Calendar" readonly="readonly" />
					            <span class="input-group-addon">
					                <span class="glyphicon glyphicon-calendar"></span>
					            </span>
					        </div>
					    </div>
			    	</div>
				</div>   
			
	          <div class="form-group">        	                
	                <div class="col-lg-12">
	                  <div class="input-group">
	                  		<p>Notes For This Activity</p>
							<textarea class="form-control" id="txtNextActivityNotes" name="txtNextActivityNotes"></textarea>
	                   </div>
	                </div> 	 
	           </div>

		</div>
   			
        <!-- END QUICK INFO BOX -->
      </div><!-- end col-md-4 -->
	
   	<div class="col-md-4">
   						
            <div class="quick-info-block CRMTileStageColor">
            <h2 class="heading-md"><i class="icon-2x color-light fa fa-tasks"></i>&nbsp;<%= GetTerm("Stage") %></h2>

			<hr class="tile">
			
			<%	
			CurrentStageRecID = GetProspectCurrentStageIntRecIDByProspectNumber(txtInternalRecordIdentifier)
			
		  	SQLCurrentStageInfo = "SELECT TOP 1 * FROM PR_ProspectStages Where ProspectRecID = " & txtInternalRecordIdentifier & " AND " & " InternalRecordIdentifier = " & CurrentStageRecID & " ORDER BY RecordCreationDateTime DESC"
			
			Set cnnCurrentStageInfo = Server.CreateObject("ADODB.Connection")
			cnnCurrentStageInfo.open (Session("ClientCnnString"))
			Set rsCurrentStageInfo = Server.CreateObject("ADODB.Recordset")
			rsCurrentStageInfo.CursorLocation = 3 
			Set rsCurrentStageInfo = cnnCurrentStageInfo.Execute(SQLCurrentStageInfo)
			If not rsCurrentStageInfo.EOF Then
				ProspectCurrentStageNotes = rsCurrentStageInfo("Notes")
			End If
			set rsCurrentStageInfo = Nothing
			cnnCurrentStageInfo.close
			set cnnCurrentStageInfo = Nothing
				
			%>
			<div class="form-group">
				<div class="col-lg-12">
					<p><strong>Stage Prior to Inactive Status:</strong> <%= GetStageByNum(GetProspectCurrentStageByProspectNumber(txtInternalRecordIdentifier)) %></p>
					<p><strong>Prior Reason:</strong> <%= GetStageReasonByStageIntRecID(GetProspectCurrentStageIntRecIDByProspectNumber(txtInternalRecordIdentifier)) %></p>
					<p><strong>Prior Stage Notes:</strong> <%= ProspectCurrentStageNotes %></p>
					<input type="hidden" name="txtCurrentStageNo" value="<%= CurrentStageRecID %>">
				</div>
			</div>
			
            <h2 class="heading-md"><i class="icon-2x color-light fa fa-tasks"></i>&nbsp;Update <%= GetTerm("Stage") %></h2>

			<hr class="tile">
			
			
			<div class="form-group">
				<div class="col-sm-12" style="width:550px; margin-left:0px; margin-right:0px; text-align:center;">
					
					<div class="row">
						<div class="col-md-12 col-lg-12">
							<!--<h4 class="modal-title">1. Primary Stage</h4>-->
							<div class="form-group">
						      	<% 
						      		'Get all stages
						      	  	SQLStages = "SELECT * FROM PR_Stages WHERE StageType = 'Primary' ORDER BY SortOrder"
				
									Set cnnStages = Server.CreateObject("ADODB.Connection")
									cnnStages.open (Session("ClientCnnString"))
									Set rsStages = Server.CreateObject("ADODB.Recordset")
									rsStages.CursorLocation = 3 
									Set rsStages = cnnStages.Execute(SQLStages)
										
									If not rsStages.EOF Then
										Do
											%>
											<div class="radio">
												<label class="btn btn-default">
													<input name="radStage" id="radStage<%= rsStages("InternalRecordIdentifier") %>" value="<%= rsStages("InternalRecordIdentifier") %>" type="radio"><%= rsStages("Stage") %>							    
												</label>
											</div>
											<%													
											rsStages.movenext
										Loop until rsStages.eof
									End If
									set rsStages = Nothing
									cnnStages.close
									set cnnStages = Nothing
									
								%>										
							</div>
						</div>
					</div>

					<hr class="tile">
					

					<div class="row">
						<div class="col-md-12 col-lg-12">
							<!--<h4 class="modal-title">2. Secondary Stage</h4>-->
							<div class="form-group">
						      	<% 
						      		'Get all stages
						      	  	SQLStages = "SELECT * FROM PR_Stages WHERE StageType = 'Secondary' AND InternalRecordIdentifier > 1 ORDER BY SortOrder"
				
									Set cnnStages = Server.CreateObject("ADODB.Connection")
									cnnStages.open (Session("ClientCnnString"))
									Set rsStages = Server.CreateObject("ADODB.Recordset")
									rsStages.CursorLocation = 3 
									Set rsStages = cnnStages.Execute(SQLStages)
										
									If not rsStages.EOF Then
										Do
											%>
											<div class="radio">
												<label class="btn btn-default">
													<input name="radStage" id="radStage<%= rsStages("InternalRecordIdentifier") %>" value="<%= rsStages("InternalRecordIdentifier") %>" type="radio"><%= rsStages("Stage") %>							    
												</label>
											</div>
											<%													
											rsStages.movenext
										Loop until rsStages.eof
									End If
									set rsStages = Nothing
									cnnStages.close
									set cnnStages = Nothing
									
								%>										
							</div>
						</div>
					</div>
						

					<hr class="tile">
					
					
					<div class="form-group">
						<div class="col-md-12 col-lg-12">
							<div class="radio">
								<label class="btn btn-default success">
									<input name="radStage" id="radStageWon" value="radStageWon" type="radio">Won							    
								</label>
							</div>
						</div>
					</div>
						
						 
				 <hr class="tile">
		
		          <div class="form-group">        	                
		                <div class="col-lg-12">
		                  <div class="input-group">
		                  		<p>Notes For This Stage Change</p>
								<textarea class="form-control" id="txtStageNotes" name="txtStageNotes"></textarea>
		                   </div>
		                </div> 	 
		           </div>
			</div><!-- end quick info block -->
      	</div><!-- end col-md-4 -->
     
     	
 </div> <!-- end row --> 
</div> <!-- end container -->

        <div class="col-md-4">
		   	<div class="form-group pull-left">
				<div class="col-lg-12">
					<button class="btn btn-primary btn-lg btn-block" href="main.asp" role="button" type="submit"><%= GetTerm("RECYCLE") %> THIS PROSPECT <i class="fa fa-recycle" aria-hidden="true"></i></button>
				</div>
			</div>
      	</div> 

</div> <!-- end container -->
</div> <!-- end container -->


</form>



<!--#include file="../inc/footer-main.asp"-->