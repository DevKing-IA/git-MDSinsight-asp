
 
<!-- ******************************************************************************************************************************** -->
<!-- MODAL WINDOW DESIGN AND DEFINITIONS -->
<!-- ******************************************************************************************************************************** -->

	<!-- modal placeholder for edit prospect BusinessCard begins here !-->
	 <!-- Modal -->
	 
		<div class="modal fade" id="myProspectingModalEditBusinessCard" tabindex="-1" role="dialog" aria-labelledby="myProspectingModalEditBusinessCardLabel">		
		  <div class="modal-dialog" role="document" style="width:800px;">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalBusinessCardInfoAddLabel">Update <%= GetTerm("Business Card") %></h4>
		      </div>
		      <form name="frmEditProspectBusinessCardFromModal" id="frmEditProspectBusinessCardFromModal" action="editProspectBusinessCardFromModal_viewprospect.asp" method="POST" onsubmit="return validateEditProspectBusinessCard();">
		      		
		      	  <input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
		      
			      <div class="modal-body">  
					<div class="row">					
						<div class="col-lg-12" style="margin-bottom:15px;" id="prospectBusinessCardInfo"></div>							
			  			</div>
			       </div>
				   <div class="clearfix"></div>
						  

			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-primary">Update <%= GetTerm("Business Card") %></button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- edit prospect BusinessCard modal ends here !-->
	
	
	
	
	<!-- modal placeholder for edit prospect owner begins here !-->
	 <!-- Modal -->
	 
		<div class="modal fade" id="myProspectingModalEditOwner" tabindex="-1" role="dialog" aria-labelledby="myProspectingModalEditOwnerLabel">
		
		<script>
		
			$(document).ready(function() {
				
				$("#showEmailNewOwnerCheckbox").hide();
		    
			    //When a user changes the new owner, determine whether to show checkbox to not send
			    //accept/reject prospect ownership email
			    
			    $(document).on('change','[name="selProspectEditOwner"]',function(){

					myProspectID = $("#txtInternalRecordIdentifier").val(); 
					newOwnerUserNo = $("#selProspectEditOwner").val();
					
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
  
			}); //end document.ready() function
		
		</script>
		
		
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalOwnerAddLabel">Update <%= GetTerm("Owner") %></h4>
		      </div>
		      <form name="frmEditProspectOwnerFromModal" id="frmEditProspectOwnerFromModal" action="editProspectOwnerFromModal_viewprospect.asp" method="POST">
		      		
		      	  <input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
		      	  <input type="hidden" name="txtOrigOwnerUserNo" id="txtOrigOwnerUserNo" value="">
		      
			      <div class="modal-body">     
						<div class="col-lg-12" style="margin-bottom:15px;" id="prospectOwnerDropdown">
						</div>	
						
		               <div class="form-group" id="showEmailNewOwnerCheckbox" style="display:none;">
			                <div class="col-sm-12">
			                  <p>Do Not Send Accept/Reject Email To New Owner:&nbsp;&nbsp;<input type="checkbox" name="chkDoNotEmailNewOwner" id="chkDoNotEmailNewOwner"></p>
			                </div> 	 
		               </div>
												
			       </div>
				   <div class="clearfix"></div>
						  

			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-primary">Update <%= GetTerm("Owner") %></button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- edit prospect owner modal ends here !-->


	
	
	<!-- modal placeholder for edit prospect comments begins here !-->
	 <!-- Modal -->
	 
		<div class="modal fade" id="myProspectingModalEditComments" tabindex="-1" role="dialog" aria-labelledby="myProspectingModalEditCommentsLabel">		
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalCommentsAddLabel">Update Prospect <%= GetTerm("Comments") %></h4>
		      </div>
		      <form name="frmEditProspectCommentsFromModal" id="frmEditProspectCommentsFromModal" action="editProspectCommentsFromModal_viewprospect.asp" method="POST">
		      		
		      	  <input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
		      
			      <div class="modal-body">     
						<div class="col-lg-12" style="margin-bottom:15px;" id="prospectComments">
						</div>							
			       </div>
				   <div class="clearfix"></div>
						  

			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-primary">Update <%= GetTerm("Comments") %></button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- edit prospect comments modal ends here !-->



	<!-- modal placeholder for edit prospect Opportunity begins here !-->
	 <!-- Modal -->
	 
		<div class="modal fade" id="myProspectingModalEditOpportunity" tabindex="-1" role="dialog" aria-labelledby="myProspectingModalEditOpportunityLabel">		
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalOpportunityInfoAddLabel">Update <%= GetTerm("Opportunity") %></h4>
		      </div>
		      <form name="frmEditProspectOpportunityFromModal" id="frmEditProspectOpportunityFromModal" action="editProspectOpportunityFromModal_viewprospect.asp" method="POST">
		      		
		      	  <input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
		      
			      <div class="modal-body">  
					<div class="row">					
						<div class="col-lg-12" style="margin-bottom:15px;" id="prospectOpportunityInfo"></div>							
			  			</div>
			       </div>
				   <div class="clearfix"></div>
						  

			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-primary">Update <%= GetTerm("Opportunity") %></button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- edit prospect Opportunity modal ends here !-->

	


	<!-- modal placeholder for edit prospect Current Supplier Info begins here !-->
	 <!-- Modal -->
	 
		<div class="modal fade" id="myProspectingModalEditCurrentSupplierInfo" tabindex="-1" role="dialog" aria-labelledby="myProspectingModalEditCurrentSupplierInfoLabel">		
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalCurrentSupplierInfoAddLabel">Update <%= GetTerm("Current Supplier Info") %></h4>
		      </div>
		      <form name="frmEditProspectCurrentSupplierInfoFromModal" id="frmEditProspectCurrentSupplierInfoFromModal" action="editProspectCurrentSupplierInfoFromModal_viewprospect.asp" method="POST">
		      		
		      	  <input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
				   
			      <div class="modal-body">     
						<div class="col-lg-12" style="margin-bottom:15px;" id="prospectCurrentSupplierInfo">
						</div>							
			       </div>
				   <div class="clearfix"></div>	  

			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-primary">Update <%= GetTerm("Current Supplier Info") %></button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- edit prospect Current Supplier Info modal ends here !-->



	<!-- modal placeholder for edit prospect Competitor Source begins here !-->
	 <!-- Modal -->
	 
		<div class="modal fade" id="myProspectingModalEditCompetitorSource" tabindex="-1" role="dialog" aria-labelledby="myProspectingModalEditCompetitorSourceLabel">		
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">

			   <style type="text/css">
					
					fieldset.group  { 
					  margin: 0; 
					  padding: 0; 
					  margin-bottom: 1.25em; 
					  padding: .125em; 
					} 
					
					fieldset.group legend { 
					  margin: 0; 
					  padding: 0; 
					  font-weight: bold; 
					  margin-left: 20px; 
					  font-size: 100%; 
					  color: black; 
					} 
					
					
					ul.checkbox  { 
					  margin: 0; 
					  padding: 0; 
					  margin-left: 20px !important; 
					  list-style: none; 
					  text-align: left !important;
					} 
					
					ul.checkbox li input { 
					  margin-right: .25em; 
					} 
					
					ul.checkbox li { 
					  border: 1px transparent solid; 
					  display:inline-block;
					  width:12em;
					} 
					
					ul.checkbox li label { 
					  margin-left: ; 
					} 
					
					.checkbox label{
						color:#000 !important;
						margin-top: 0px;
					}
					
					.checkbox label, .radio label {
					    min-height: 20px;
					    padding-left: 20px;
					    margin-bottom: 0;
					    font-weight: 400;
					    cursor: pointer;
					    color: #000;
					}					
			  </style>
			  
			  
				<script language="JavaScript">
				<!--
				
				   function validateEditProspectCompetitorSourceForm()
				    {
								    
				 		var chkd = document.frmEditProspectCompetitorSourceFromModal.chkBottledWater.checked || +
				 			document.frmEditProspectCompetitorSourceFromModal.chkFilteredWater.checked|| +
				 			document.frmEditProspectCompetitorSourceFromModal.chkOCS.checked|| +
				 			document.frmEditProspectCompetitorSourceFromModal.chkOCS_Supply.checked|| +
				 			document.frmEditProspectCompetitorSourceFromModal.chkOfficeSupplies.checked|| +
				 			document.frmEditProspectCompetitorSourceFromModal.chkVending.checked|| +
				 			document.frmEditProspectCompetitorSourceFromModal.chkMicroMarket.checked|| +
				 			document.frmEditProspectCompetitorSourceFromModal.chkPantry.checked;
							
							if (chkd == true)
							{
						       if (document.frmEditProspectCompetitorSourceFromModal.txtPrimaryCompetitor.value == "") {
						            swal("Primary competitor must be selected if offerings are selected.");
						            return false;
						       }
							}
							else
							{
						       if (document.frmEditProspectCompetitorSourceFromModal.txtPrimaryCompetitor.value !== "") {
						            swal("You must select at least one offering for the primary competitor.");
						            return false;
						       }			
						    }   
										
				
				       return true;
				
				    }
				// -->
				</script>  
		    
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalCompetitorSourceAddLabel">Update <%= GetTerm("Primary Competitor") %></h4>
		      </div>
		      <form name="frmEditProspectCompetitorSourceFromModal" id="frmEditProspectCompetitorSourceFromModal" action="editProspectCompetitorSourceFromModal_viewprospect.asp" method="POST" onsubmit="return validateEditProspectCompetitorSourceForm();">
		      		
		      	  <input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
		      				  				   
			      <div class="modal-body">     
						<div class="col-lg-12" style="margin-bottom:15px;" id="prospectCompetitorSource">
						</div>							
			       </div>
				   <div class="clearfix"></div>
						  

			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-primary">Update <%= GetTerm("Primary Competitor") %></button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- edit prospect Competitor Source modal ends here !-->







	<!-- modal palceholder for edit prospect next activity begins here !-->
	 <!-- Modal -->
		<div class="modal fade" id="myProspectingModalEditActivity" tabindex="-1" role="dialog" aria-labelledby="myProspectingModalEditActivityLabel">
		
		<script>
		
			$(document).ready(function() {
			
				//Initially, hide both divs that show either the appointment or meeting fields
				
			    $("#showActivityAppointmentDuration").hide();
			    $("#showActivityMeetingDuration").hide();
			    			    
			    //When a user changes the next activity, there are several ajax posts that have to be made to determine
			    //whether or not to show a meeting or appointment, based on user type and activity type
			    
				$("#selProspectNextActivity").change(function() {
	
				    
				    //Hide both divs that show either the appointment or meeting fields whenever the customer changes
				    //a next activity, until we know what to display, if anything
				    
				    $("#showActivityAppointmentDuration").hide();
				    $("#showActivityMeetingDuration").hide();				

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
									               	$("#txtMeetingLocation").val(response3);               	 
									             },
									            failure: function(response3)
												 {
												  	//If no address infomation was returned, just show the meeting div input fields, and do not
												  	//set the default value of the meeting location
												   	$("#showActivityMeetingDuration").show();
									             }

											});	//end ajax post to data: "action=GetMeetingLocationForModal"
						               		
						               	}
						               	else {
						               			//Else, the activity is not a 'Meeting' or an 'Appointment' so make sure the divs are hidden
												$("#showActivityAppointmentDuration").hide();
												$("#showActivityMeetingDuration").hide();
													
						               	}// end if statement for activityCalendarShowApptOrMeeting 

	               	 
						             }  //end success function for ajax post {show meeting or appointment for this activity}

								}); //end ajax post to data: "action=GetActivityCalendarApptOrMeetingForModal" {show meeting or appointment for this activity}
								
			               	 }	//end if for if (response == 'True') {user calendar gets updated with an activity change}
			               	 
							else{
								$("#showActivityAppointmentDuration").hide();
								$("#showActivityMeetingDuration").hide();						

							}	      
			               	          	 
			             } //end success function for ajax post {user calendar gets updated with an activity change}
			             
					});//end ajax post to data: "action=GetAllowActivityUpdatesToUsersCalendarForModal"
					
				});	// end $("#selProspectNextActivity").change(function()		    
  
			}); //end document.ready() function
		
		</script>
		
		<script language="JavaScript">
		<!--
		
		   function validateNextActivitySubmit()
		    {
				
			   var selProspectNextActivity = $("#selProspectNextActivity option:selected").val();
			   
		       if (selProspectNextActivity == "" || selProspectNextActivity == "-1") {
		            swal("Required: Please select a next activity.");
		            return false;
		       }
			   
			   		    				       
			   var nextActivityDueDate = $("#txtProspectEditNextActivityDate").val();
			   		    
		       if (nextActivityDueDate == "") {
		            swal("Required: Please select a due date for the next actvity status.");
		            return false;
		       }
		       
		       return true;
		
		    }
		// -->
		</script>  
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalAddLabel">Set Prospect Activity</h4>
		      </div>
		      <form name="frmEditProspectNextActivityFromModal" id="frmEditProspectNextActivityFromModal" action="editProspectNextActivityFromModal_viewprospect.asp" method="POST" onsubmit="return validateNextActivitySubmit()">
		      		
		      	  <input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
		      	  <input type="hidden" name="txtActivityRecID" id="txtActivityRecID" value="">

			      <div class="modal-body">     
						            					  
					  	<div class="col-lg-12" id="prospectCurrentActivitySummary">
					  	<!-- Content for the current activity in this modal will be generated and written here -->
						<!-- Content generated by Sub GetProspectActivityInformationForModal() in InsightFuncs_AjaxForProspectingModals.asp -->

					  	</div>
					
						<div class="col-lg-12" style="margin-bottom:15px;">
							<div class="form-group">
								<div class="col-lg-4" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Mark This Activity As:</label>
								</div>
								<div class="col-lg-8">			
								  	<select class="form-control-modal" name="selProspectCurrentActivityStatus" id="selProspectCurrentActivityStatus">
									      <option value="Completed">COMPLETED</option>
									      <option value="Cancelled">CANCELLED</option>
									      <option value="Rescheduled">RESCHEDULED</option>
									</select>
								</div>
							</div>
						</div>	
							
						<div class="col-lg-12">	
							<div class="form-group">
							  <label for="prospectEditNextActivityNotes">Notes:</label>
							  <textarea class="form-control" rows="5" id="txtProspectEditNextActivityNotes" name="txtProspectEditNextActivityNotes"></textarea>
							</div>
						</div>
	
					
						<div class="col-lg-12">
							<div class="form-group">
								<div class="col-lg-5" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Select a New Next Activity:</label>
								</div>
								<div class="col-lg-7">			
								  	<select class="form-control-modal" name="selProspectNextActivity" id="selProspectNextActivity">
							      	<% 
							      	  	
										
									%>									
									</select>
								</div>
							</div>
						</div>	
						
						<div class="col-lg-12" style="margin-top:15px;" id="showActivityAppointmentDuration">
							<div class="form-group">
								<div class="col-lg-5" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Appointment Duration (for Outlook Calendar):</label>
								</div>
								<!-- Get Default Appointment Duration from tblGlobalSettings -->
								<%
									EWSDefaultAppointmentDuration = GetPOSTParams("EWSDEFAULTAPPTDURATION")
								%>
								<div class="col-lg-7">		
								  	<select class="form-control" name="selAppointmentDuration" id="selAppointmentDuration">
										<%For x = 15 to 180 Step 5
											If x mod 60 = 0 Then
												If x = cint(EWSDefaultAppointmentDuration) Then 
													Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
												else
													Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
												End If
											Else
												If x = cint(EWSDefaultAppointmentDuration) Then 
													Response.Write("<option value='" & x & "' selected>" & x & "</option>")
												Else
													Response.Write("<option value='" & x & "'>" & x & "</option>")
												End If
											End If
										Next %>
										
									</select>
								</div>
							</div>
						</div>
						
						
						
						<div class="col-lg-12" style="margin-top:15px;" id="showActivityMeetingDuration">
							<div class="form-group">
								<div class="col-lg-5" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Meeting Duration (for Outlook Calendar):</label>
								</div>
								<!-- Get Default Meeting Duration from tblGlobalSettings -->
								<%
									EWSDefaultMeetingDuration = GetPOSTParams("EWSDEFAULTMEETINGDURATION")
								%>
								<div class="col-lg-7">												
								  	<select class="form-control" name="selMeetingDuration" id="selMeetingDuration">
										<%For x = 15 to 300 Step 15
											If x mod 60 = 0 Then
												If x = cint(EWSDefaultMeetingDuration) Then 
													Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
												else
													Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
												End If
											Else
												If x = cint(EWSDefaultMeetingDuration) Then 
													Response.Write("<option value='" & x & "' selected>" & x & "</option>")
												Else
													Response.Write("<option value='" & x & "'>" & x & "</option>")
												End If
											End If
										Next %>												
									</select>
								</div>
							</div>
							<div class="form-group">
								<div class="col-lg-12" style="padding-left:0px; margin-top:15px;">
									<div class="form-group">
									  <label for="txtMeetingLocation">Meeting Location (for Outlook Calendar):</label>
									  <input class="form-control" type="text" id="txtMeetingLocation" name="txtMeetingLocation">
									</div>
								
								</div>
							</div>
						</div>
												
						<div class="col-lg-12" style="margin-top:15px;" id="activityDateWarning" style="display:none">
							<div class="alert alert-danger">
							  <strong>Warning!</strong> This activity has been schedule beyond the recommended limit.
							</div>	
						</div>	
													
						<div class="col-lg-12" style="margin-top:15px;">	
							<div class="form-group">

								<div class="col-lg-5" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Due Date:</label>
								</div>
								<div class="col-lg-7">								  	
					                <div class="input-group date" id="datetimepickerNextActivity">
					                    <input type="text" class="form-control" name="txtProspectEditNextActivityDate" id="txtProspectEditNextActivityDate">
					                    <span class="input-group-addon">
					                        <span class="glyphicon glyphicon-calendar"></span>
					                    </span>
					                </div>
					             </div>
							</div>
						</div>
	
						
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-primary">Update Prospect Activity</button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- edit prospect activity modal ends here !-->
	
	

	
	


	<!-- modal placeholder for edit prospect stage begins here !-->
	 <!-- Modal -->
	 
		<div class="modal fade" id="myProspectingModalEditStage" tabindex="-1" role="dialog" aria-labelledby="myProspectingModalEditStageLabel">
		
		<style>
		
		 .radio {
		    position: relative !important;
		    display: inline !important;
		    margin-top: 10px !important;
		    margin-bottom: 20px !important;
			}
		  .radio .btn,
		  .radio-inline .btn {
		    padding-left: 2em !important;
		    min-width: 7em !important;
		    margin-top: 10px !important;
		    margin-left: 5px !important;
		  }
		 
		  .radio label,
		  .radio-inline label {
		    text-align: left !important;
		    padding-left: 0.5em !important;
		    color:#000 !important;
		  }
		</style>
		
		<script>
			$(document).ready(function() {
			
			    $("#showUnqualifyingReasons").hide();
			    $("#showLostReasons").hide();
			    
				$('input[type=radio][name=radStage]').change(function() {
				        if (this.id !== 'radStage0') {
				            $("#showUnqualifyingReasons").hide();
				        }
				        if (this.id !== 'radStageLost') {
				            $("#showLostReasons").hide();
				        }
				    });			    
			    
			    $("input[id$='radStage0']").click(function() {
			        $("#showUnqualifyingReasons").show();
			    });
			    
			    $("input[id$='radStageLost']").click(function() {
			        $("#showLostReasons").show();
			    });
	
				    
			});
		
		</script>
		
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalAddLabel">Set Prospect Stage</h4>
		      </div>
		      <form name="frmEditProspectStageFromModal" id="frmEditProspectStageFromModal" action="editProspectStageFromModal_viewprospect.asp" method="POST">
		      		
		      	  <input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
		      	  <input type="hidden" name="txtStageRecID" id="txtStageRecID" value="">
		      
			      <div class="modal-body">     
						            					  
					  	<div class="col-lg-12" id="prospectCurrentStageSummary">
					  	<!-- Content for the current activity in this modal will be generated and written here -->
						<!-- Content generated by Sub GetProspectStageInformationForModal() in InsightFuncs_AjaxForProspectingModals.asp -->

					  	</div>
						
	
						<div class="row"> <!-- You can also position the row if need be. -->
							<div class="col-md-12 col-lg-12"><!-- set width of column I wanted mine to stretch most of the screen-->
								<hr style="min-width:85%; background-color:#eee !important; height:1px;">
							</div>
						 </div>
					
						<div class="row">
							<div class="col-md-12 col-lg-12">
								<!--<h4 class="modal-title">1. Primary Stage</h4>-->
								<div class="form-group stageprimarygroup">
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

						<div class="row"> <!-- You can also position the row if need be. -->
							<div class="col-md-12 col-lg-12"><!-- set width of column I wanted mine to stretch most of the screen-->
								<hr style="min-width:85%; background-color:#eee !important; height:1px;">
							</div>
						 </div>
						 
						<div class="row">
							<div class="col-md-12 col-lg-12">
								<!--<h4 class="modal-title">2. Secondary Stage</h4>-->
								<div class="form-group stagesecondarygroup">
							      	<% 
							      		'Get all stages
							      	  	SQLStages = "SELECT * FROM PR_Stages WHERE StageType = 'Secondary' ORDER BY SortOrder"
					
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
							

						<div class="col-lg-12" style="margin-top:15px;" id="showUnqualifyingReasons">
							<div class="form-group">
								<div class="col-lg-4" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Select a reason why this prospect is unqualified:</label>
								</div>
								<div class="col-lg-8">			
								  	<select class="form-control-modal" name="selUnqualifyingReasons" id="selUnqualifyingReasons">
							      	<% 
							      		'Get all reasons
							      	  	SQL9 = "SELECT * FROM PR_Reasons WHERE ReasonType='Unqualifying' OR ReasonType='Unqualifying and Lost' ORDER BY InternalRecordIdentifier"
					
										Set cnn9 = Server.CreateObject("ADODB.Connection")
										cnn9.open (Session("ClientCnnString"))
										Set rs9 = Server.CreateObject("ADODB.Recordset")
										rs9.CursorLocation = 3 
										Set rs9 = cnn9.Execute(SQL9)
											
										If not rs9.EOF Then
											Do
												Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Reason") & "</option>")
												rs9.movenext
											Loop until rs9.eof
										End If
										set rs9 = Nothing
										cnn9.close
										set cnn9 = Nothing
										
									%>									
								</select>
							</div>
						</div>
						</div>
						
						<div class="row"> <!-- You can also position the row if need be. -->
							<div class="col-md-12 col-lg-12"><!-- set width of column I wanted mine to stretch most of the screen-->
								<hr style="min-width:85%; background-color:#eee !important; height:1px;">
							</div>
						 </div>
	

						<div class="row">
							<div class="col-md-12 col-lg-12">
								<!--<h4 class="modal-title">3. Final Stage</h4>-->
								<div class="form-group">
									<div class="radio">
										<label class="btn btn-default success">
											<input name="radStage" id="radStageWon" value="radStageWon" type="radio">Won							    
										</label>
									</div>
									<div class="radio">
										<label class="btn btn-default warning">
											<input name="radStage" id="radStageLost" value="radStageLost" type="radio">Lost							    
										</label>
									</div>
								</div>
							</div>
						</div>
											
						<div class="col-lg-12" style="margin-top:15px;" id="showLostReasons">
							<div class="form-group">
								<div class="col-lg-4" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Select a reason why this prospect is lost:</label>
								</div>
								<div class="col-lg-8">			
								  	<select class="form-control-modal" name="selLostReasons" id="selLostReasons">
							      	<% 
							      		'Get all reasons
							      	  	SQL9 = "SELECT * FROM PR_Reasons WHERE ReasonType='Lost' OR ReasonType='Unqualifying and Lost' ORDER BY InternalRecordIdentifier"
					
										Set cnn9 = Server.CreateObject("ADODB.Connection")
										cnn9.open (Session("ClientCnnString"))
										Set rs9 = Server.CreateObject("ADODB.Recordset")
										rs9.CursorLocation = 3 
										Set rs9 = cnn9.Execute(SQL9)
											
										If not rs9.EOF Then
											Do
												Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Reason") & "</option>")
												rs9.movenext
											Loop until rs9.eof
										End If
										set rs9 = Nothing
										cnn9.close
										set cnn9 = Nothing
										
									%>										
									</select>
								</div>
							</div>
						</div>
						
						<div class="row"> <!-- You can also position the row if need be. -->
							<div class="col-md-12 col-lg-12"><!-- set width of column I wanted mine to stretch most of the screen-->
								<hr style="min-width:85%; background-color:#eee !important; height:1px;">
							</div>
						 </div>
	
						<div class="row">					
							<div class="col-lg-12">	
								<div class="form-group">
								  <label for="prospectEditStageNotes">Notes For This Change:</label>
								  <textarea class="form-control" rows="5" id="txtProspectEditStageNotes" name="txtProspectEditStageNotes"></textarea>
								</div>
							</div>
						</div>
							
			       </div>
			      <div class="modal-footer">
                  <div class="pull-left">
                  <%If userCanEditCRMOnTheFly(Session("UserNO")) = True AND GetCRMAddEditMenuPermissionLevel(Session("UserNO")) = vbTrue Then%>
                  	<button type="button" class="btn btn-success"  data-toggle="modal" data-show="true" href="#" data-prospect-id="3" data-target="#ONTHEFLYmodalStage" data-tooltip="true">Add A Stage</button>
                  <%End If%>
                  </div>
                  <div class="pull-right">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-primary">Update Prospect Stage</button>
                  </div>  
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- add prospects to group modal ends here !-->
	

<!-- ******************************************************************************************************************************** -->
<!-- END MODAL WINDOW DESIGN AND DEFINITIONS -->
<!-- ******************************************************************************************************************************** -->


