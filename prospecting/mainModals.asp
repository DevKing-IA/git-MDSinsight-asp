<!-- ******************************************************************************************************************************** -->
<!-- MODAL WINDOW DESIGN AND DEFINITIONS -->
<!-- ******************************************************************************************************************************** -->

<!------------------------------------------------------------------------------>	
<!-- modal for filtering data in the prospecting table/grid view !-->
<!------------------------------------------------------------------------------>

<!--#include file="mainModalCustomizeDataFilterValues.asp"-->

<!------------------------------------------------------------------------------>	
<!-- END modal for filtering data in the prospecting table/grid view !-->
<!------------------------------------------------------------------------------>


<!------------------------------------------------------------------------------>	
<!-- modal for showing and hiding columns in the prospecting table/grid view !-->
<!------------------------------------------------------------------------------>
<div class="modal fade bs-modal-show-hide-columns" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel" aria-hidden="true">
<div class="modal-dialog modal-lg modal-height">
<div class="modal-content">

	<style type="text/css">
	.ativa-scroll{
		max-height: 300px
	}
	</style>
	
	<!-- modal scroll !-->
	<script type="text/javascript">
		$(document).ready(ajustamodal);
		$(window).resize(ajustamodal);
		function ajustamodal() {
		//var altura = $(window).height() - 155; //value corresponding to the modal heading + footer
		var altura = $(window).height() - 205; //value corresponding to the modal heading + footer
		$(".ativa-scroll").css({"height":altura,"overflow-y":"auto"});
	}
	</script>
	<!-- eof modal scroll !-->

  <div class="modal-header">
    <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
    <h4 class="modal-title" id="myModalLabel" align="center">Show/Hide <%= GetTerm("Prospecting") %> Columns</h4>
  </div>

	<form method="post" action="mainCustomizeSaveShowHideColumnValues.asp" name="frmProspectingCustomizeColumns">

      <!-- insert content in here !-->
      <div class="modal-body ativa-scroll">
 	      	
  	      	<!-- filtering !-->
	      	<div class="container-fluid">
		      	<div class="row">
 		      	
 		      	<!-- left column !-->
 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column">
	 		      	<h4><br>Column Names</h4>
 		      	</div>
 		      	<!-- eof left column !-->
 		      	
 		      	<!-- right column !-->
 		      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
	 		      	
		      	<!-- row !-->
		      	<div class="row">
     	
		      	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
			      	<strong>Click To Show/Hide</strong>
		      	</div>
				<%
					'************************
					'Read Settings_Reports
					'************************
					SQL = "SELECT * from Settings_Reports where ReportNumber = 1400 AND PoolForProspecting = 'Live' AND UserNo = " & Session("userNo")
					SQL = "SELECT * from Settings_Reports where ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = '" & customFilterReportNameForSQL & "'"
					Set cnn8 = Server.CreateObject("ADODB.Connection")
					cnn8.open (Session("ClientCnnString"))
					Set rs = Server.CreateObject("ADODB.Recordset")
					Set rs= cnn8.Execute(SQL)
					UseSettings_Reports = False
					If NOT rs.EOF Then
						UseSettings_Reports = True
						showHideColumns = rs("ReportSpecificData1")
					End If
					'****************************
					'End Read Settings_Reports
					'****************************
				%>
		      	<div class="col-lg-9 col-md-9 col-sm-12 col-xs-12">
					
					<div class="ck-button">
					<label><input type="checkbox" value="col_address" name="chkCol_address" <% If InStr(showHideColumns,"col_address") Then Response.Write("checked='checked'") %>><span>Street Address</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_city" name="chkCol_city" <% If InStr(showHideColumns,"col_city") Then Response.Write("checked='checked'") %>><span>City</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_state" name="chkCol_state" <% If InStr(showHideColumns,"col_state") Then Response.Write("checked='checked'") %>><span>State</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_zip" name="chkCol_zip" <% If InStr(showHideColumns,"col_zip") Then Response.Write("checked='checked'") %>><span>Zip Code</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_leadsource" name="chkCol_leadsource" <% If InStr(showHideColumns,"col_leadsource") Then Response.Write("checked='checked'") %>><span>Lead Source</span></label>
					</div>
					<!--<div class="ck-button">
					<label><input type="checkbox" value="col_leadsource" name="chkCol_stage" <% If InStr(showHideColumns,"col_stage") OR showHideColumns = "" Then Response.Write("checked='checked'") %>><span>Stage</span></label>
					</div>-->
					<div class="ck-button">
					<label><input type="checkbox" value="col_industry" name="chkCol_industry" <% If InStr(showHideColumns,"col_industry") Then Response.Write("checked='checked'") %>><span>Industry</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_numemployees" name="chkCol_numemployees" <% If InStr(showHideColumns,"col_numemployees") Then Response.Write("checked='checked'") %>><span>Number of Employees</span></label>
					</div>
					<!--<div class="ck-button">
					<label><input type="checkbox" value="col_owner" name="chkCol_owner" <% If InStr(showHideColumns,"col_owner") Then Response.Write("checked='checked'") %>><span>Prospect Owner</span></label>
					</div>-->
					<div class="ck-button">
					<label><input type="checkbox" value="col_createddate" name="chkCol_createddate" <% If InStr(showHideColumns,"col_createddate") Then Response.Write("checked='checked'") %>><span>Prospect Created Date</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_createdby" name="chkCol_createdby" <% If InStr(showHideColumns,"col_createdby") Then Response.Write("checked='checked'") %>><span>Prospect Created By</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_telemarketer" name="chkCol_telemarketer" <% If InStr(showHideColumns,"col_telemarketer") Then Response.Write("checked='checked'") %>><span>Telemarketer</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_numpantries" name="chkCol_numpantries" <% If InStr(showHideColumns,"col_numpantries") Then Response.Write("checked='checked'") %>><span>Number of Pantries</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_prospectid" name="chkCol_prospectid" <% If InStr(showHideColumns,"col_prospectid") Then Response.Write("checked='checked'") %>><span>Prospect ID</span></label>
					</div>

		      	</div>
 		      	  		      			      	
		      	</div>
		      	<!-- eof row !-->
		      	
		      	
 		      	</div>
 		      	<!-- eof right column !-->
 		      	
		      	</div>
	      	</div>
       
       </div>
      <!-- eof content insertion !-->
      
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
         <a href="#" onClick="document.frmProspectingCustomizeColumns.submit()"><button type="button" class="btn btn-primary">Save Show/Hide Columns</button></a>     
      </div>
      </form>
    </div>
  </div>
</div>
<!------------------------------------------------------------------------------>
<!-- modal for showing and hiding columns in the prospecting table/grid view !-->
<!------------------------------------------------------------------------------>	
	

	<!-- delete prospect modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="myProspectingDeleteModal" tabindex="-1" role="dialog" aria-labelledby="myProspectingDeleteLabel">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingDeleteModalLabel">Delete Prospect(s)</h4>
		      </div>
		      <form name="frmDeleteProspects" id="frmDeleteProspects" method="post" action="deleteProspectsFromModal.asp">
			      <div class="modal-body">
	
					<div class="col-lg-12" id="deleteProspectInfo">
					</div>
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-danger" data-dismiss="modal" onclick="frmDeleteProspects.submit()">Delete Prospect(s)</button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- deleted prospect modal ends here !-->
	
	<!-- add notes to prospect modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="myProspectingAddMultipleNotesModal" tabindex="-1" role="dialog" aria-labelledby="myProspectingAddMultipleNotesModalLabel">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingAddMultipleNotesModalLabel">Add Note To Prospect(s)</h4>
		      </div>
		      <form name="frmAddNotesToProspects" id="frmAddNotesToProspects" method="post" action="#">
			      <div class="modal-body">
	
					<div class="col-lg-12" id="addnotesProspectInfo">
					</div>
                    
			<div class="form-group col-lg-12">
				
    			
    				<textarea class="form-control textarea required" rows="4" id="txtProspectingNote" name="txtProspectingNote"></textarea>
    			
 			</div>                    
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-danger">Add Note To Prospect(s)</button>
			      </div>
                  <input type="hidden" id="addnotesmultipleids" name="addnotesmultipleids" value="">
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- add notes prospect modal ends here !-->    

	<!-- export prospect modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="myProspectingExportModal" tabindex="-1" role="dialog" aria-labelledby="myProspectingExportLabel">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingExportModalLabel">Export Prospect(s)</h4>
		      </div>
		      <form name="frmExportProspects" id="frmExportProspects" method="post" action="exportProspectsFromModal.asp">
			      <div class="modal-body">
	
					<div class="col-lg-12" id="exportProspectInfo">
					</div>
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-info" data-dismiss="modal" onclick="frmExportProspects.submit()">Export Prospect(s)</button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- export prospect modal ends here !-->

	<!-- new prospect group modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="myProspectingWatchModal" tabindex="-1" role="dialog" aria-labelledby="myProspectingWatchLabel">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalLabel">Watch This Prospect</h4>
		      </div>
		      <form name="frmCreateNewProspectWatch" id="frmCreateNewProspectWatch">
			      <div class="modal-body">
	
					<div class="col-lg-12">
						<div class="form-group">
							<label class="col-sm-12 control-label">Some Fields Here</label>
							<input type="text" class="form-control required" id="txtProspectingWatch" name="txtProspectingWatch">
						</div>
					</div>
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="button" class="btn btn-primary" data-dismiss="modal">Watch Prospect</button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- new prospect group modal ends here !-->
	

	<!-- new prospect group modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="saveAsNewProspectFilterView" tabindex="-1" role="dialog" aria-labelledby="saveAsNewProspectFilterView">
		  <div class="modal-dialog" role="document">
		  
		 			  
				<script language="JavaScript">
				<!--
				
				   function validateViewName()
				    {
					
					   var viewNameInputField = $("#txtNewFilterReportViewName").val();
					   var viewNameSelectBox = $("#selExistingFilterViewNames option:selected").val();
					   		    
				       if (viewNameInputField == "" && viewNameSelectBox == "") {
				            swal("Please select a name or enter a new name to save this view.");
				            return false;
				       }
				       
				       if (viewNameInputField == "Default" || viewNameInputField == "DEFAULT" || viewNameInputField == "default" || viewNameInputField == "Current" || viewNameInputField == "current" || viewNameInputField == "CURRENT"|| viewNameInputField == "All Prospects" || viewNameInputField == "all prospects" || viewNameInputField == "ALL PROSPECTS") {
				            swal("A view cannot be named DEFAULT, CURRENT or ALL PROSPECTS, they are reserved names.");
				            return false;
				       }
				
				       return true;
				
				    }
				// -->
				</script>  

		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalLabel">Save This View</h4>
		      </div>
		      <form name="frmCreateNewFilterReportView" id="frmCreateNewFilterReportView">
			      <div class="modal-body">
	
					<div class="col-lg-12">
						<div class="form-group">
							<label class="col-sm-12 control-label" style="margin-left:-15px;">Save This View As:</label>
							
					      	<%'Report View Name Dropdown
					      	
					      	CurrentViewName = MUV_READ("CRMVIEWSTATE")
					      	CurrentViewNameForSQL = Replace(MUV_READ("CRMVIEWSTATE"),"'","''")
					      	 
					  	  	SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' "
					  	  	SQL = SQL & " AND UserReportName <> 'Current'  AND UserReportName <> 'Default' AND UserReportName <> 'All Prospects' "
					  	  	SQL = SQL & " ORDER BY UserReportName "
					
							Set cnn8 = Server.CreateObject("ADODB.Connection")
							cnn8.open (Session("ClientCnnString"))
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.CursorLocation = 3 
							Set rs = cnn8.Execute(SQL)
						
							If NOT rs.EOF Then
							%>
								<!-- Display Report View Names -->
								<select class="form-control when-line" style="width:100%;height:50px;display:inline;margin-left:0px;" name="selExistingFilterViewNames" id="selExistingFilterViewNames">
								<option value=""> -- Choose An Existing View Name To Overwrite -- </option>
								<%
									Do
										selReportName = Replace(rs("UserReportName"),"''","'")
										If MUV_READ("CRMVIEWSTATE") = selReportName Then
											%><option value="<%= selReportName %>" selected="selected"><%= selReportName %></option><%
										Else
											%><option value="<%= selReportName %>"><%= selReportName %></option><%
										End If
										rs.movenext
									Loop until rs.eof
								%>		
								</select>
								<!-- End Display Report View Names -->
							<%
							End If
							set rs = Nothing
							cnn8.close
							set cnn8 = Nothing
					      	%>
							
							<br><br>
							<label class="col-sm-12 control-label" style="margin-left:-15px;">Enter a New View Name:</label>
							<input type="text" class="form-control required"  style="width:100%;height:50px;display:inline;margin-left:0px;" id="txtNewFilterReportViewName" name="txtNewFilterReportViewName">
						</div>
					</div>
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="button" class="btn btn-primary" onclick="if (validateViewName()) saveAsNewProspectFilterView();" id="saveFilterReportViewNameButton">Save Filter View</button>
			      </div/>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- new prospect group modal ends here !-->
	
	
	
	

	<!-- new prospect group modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="editFilterViewName" tabindex="-1" role="dialog" aria-labelledby="editFilterViewName">
		  <div class="modal-dialog" role="document">
		  
		 			  
				<script language="JavaScript">
				<!--
				
				   function validateEditViewName()
				    {
								    				       
					   var viewName = $("#txtUpdatedFilterReportViewName").val();
					   		    
				       if (viewName == "") {
				            swal("Please enter a name to save this view.");
				            return false;
				       }				       
				       	
				       if (viewName == "Default" || viewName == "DEFAULT" || viewName == "default" || viewName == "Current" || viewName == "current" || viewName == "CURRENT"|| viewName == "All Prospects" || viewName == "all prospects" || viewName == "ALL PROSPECTS") {
				            swal("A view cannot be named DEFAULT, CURRENT or ALL PROSPECTS, they are reserved names.");
				            return false;
				       }
				
				       return true;
				
				    }
				// -->
				</script>  

		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalLabel">Save This View</h4>
		      </div>
		      <form name="frmEditFilterViewName" id="frmEditFilterViewName">
			      <div class="modal-body">
	
					<div class="col-lg-12">
						<div class="form-group">
							<label class="col-sm-12 control-label" style="margin-left:-15px;">Save This View As:</label>
							<% CurrentViewName = MUV_READ("CRMVIEWSTATE") %>

							<input type="text" class="form-control required" id="txtUpdatedFilterReportViewName" name="txtUpdatedFilterReportViewName">

							<input type="hidden" name="originalViewName" id="originalViewName" value="<%= CurrentViewName %>">
						</div>
					</div>
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="button" class="btn btn-primary" onclick="if (validateEditViewName()) renameProspectFilterView();" id="updateFilterReportViewNameButton">Update Filter Name</button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- new prospect group modal ends here !-->
	
	
				  
	<!-- delete prospect view modal starts here !-->
		<div class="modal fade" id="deleteProspectView" tabindex="-1" role="dialog" aria-labelledby="myProspectingRecycleLabel">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">		
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><span class="glyphicon glyphicon-remove" aria-hidden="true"></span></button>
					<h4 class="modal-title custom_align" id="Heading">Delete this View</h4>
				</div>
				
				<form name="frmDeleteProspectView" id="frmDeleteProspectView" method="post" action="deleteProspectFilterViewFromModal.asp">

					<div class="modal-body">
						<h3><%= CurrentViewName %></h3>
						<div class="alert alert-danger"><span class="glyphicon glyphicon-warning-sign"></span> Are you sure you want to delete this view?</div>
						<input type="hidden" name="viewNameToDelete" id="viewNameToDelete" value="<%= CurrentViewName %>">
					</div>
					
					<div class="modal-footer ">
						<button type="button" class="btn btn-default" data-dismiss="modal"><span class="glyphicon glyphicon-remove"></span> No</button>
						<button type="button" class="btn btn-success" onclick="frmDeleteProspectView.submit()"><span class="glyphicon glyphicon-ok-sign"></span> Yes, Delete This View</button>
					</div>
				
				</form>
			</div>
			<!-- /.modal-content --> 
		</div>
		<!-- /.modal-dialog --> 
	</div>
 	<!-- delete prospect view modal ends here !-->   
	
	

	<!-- modal palceholder for edit prospect next activity begins here !-->
	 <!-- Modal -->
	 
		<div class="modal fade" id="myProspectingModalEditActivity" tabindex="-1" role="dialog" aria-labelledby="myProspectingModalEditActivityLabel">
				
		<script>
		
//common function to populate selectboxes
function PopulateSelecBoxes(selectid,selectednumber){
    $.ajax({
        type: "POST",
        url: 'onthefly_selectboxes.asp',
        data: ({ section : selectid, action:'edit',selectedvalue:selectednumber }),
        dataType: "html",
        success: function(data) {
            $("#"+selectid).html(data);
        },
        error: function() {
            alert('Error occured');
        }
    });	
}		
		
			$(document).ready(function() {
				
// below code added by nurba
// 04/26/2019		
	var myActivityRecID = '<%=ActivityRecID%>';			 
					 PopulateSelecBoxes('selProspectNextActivity',myActivityRecID);	  
					 
 //-------------------------------------------------------------------------------
	// Industry select box change
    $( "#selProspectNextActivity").change(function() {
		var val = $( "#selProspectNextActivity option:selected").val();
		if (val== -1){
			//deselect add new row
			$('#selProspectNextActivity option[selected="selected"]').each(
				function() {
					$(this).removeAttr('selected');
				}
			);

			// mark the first option as selected
			$("#selProspectNextActivity option:first").attr('selected','selected');
			
			//show modal
			$('#ONTHEFLYmodalNextActivity').modal('show');
			
		}
	});
	
	
	//Next activity modal window submit
	$('#frmAddNextActivity').submit(function(e) {
		
		if ($('#frmAddNextActivity #txtActivity').val()==''){
			 swal("Activity name can not be blank.");
			return false;
		}
		
		$("#ONTHEFLYmodalNextActivity .btn-primary").html("Saving...");
        $.ajax({
            type: "POST",
            url: "onthefly_nextactivity_submit.asp",
            data: $('#frmAddNextActivity').serialize(),
            success: function(response) {
				PopulateSelecBoxes('selProspectNextActivity','<%=myActivityRecID%>');
				$("#ONTHEFLYmodalNextActivity .modal-body").html('Next Activity added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
				
            },
            error: function() {
				$("#ONTHEFLYmodalNextActivity .btn-primary").html("Save");
            }
        });
        return false;
    });
//-------------------------------------------------------------------------------

//end nurba						
			
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
		        <h4 class="modal-title" id="myProspectingModalAddLabel">Edit Prospect Activity</h4>
		      </div>
		      <form name="frmEditProspectNextActivityFromModal" id="frmEditProspectNextActivityFromModal" action="editProspectNextActivityFromModal.asp" method="POST" onsubmit="return validateNextActivitySubmit()">
		      		
		      	  <input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="">
		      	  <input type="hidden" name="txtActivityRecID" id="txtActivityRecID" value="">

			      <div class="modal-body">     
						            					  
					  	<div class="col-lg-12" id="prospectCurrentActivitySummary">
					  	<!-- Content for the current activity in this modal will be generated and written here -->
						<!-- Content generated by Sub GetProspectActivityInformationForModal() in InsightFuncs_AjaxForProspectingModals.asp -->

					  	</div>
					
						<div class="col-lg-12">
							<div class="form-group">
								<div class="col-lg-4" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Mark This Activity As:</label>
								</div>
								<div class="col-lg-8">			
								  	<select class="form-control" name="selProspectCurrentActivityStatus" id="selProspectCurrentActivityStatus">
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
	
					
						<div class="col-lg-12" style="margin-top:15px;">
							<div class="form-group">
								<div class="col-lg-4" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Select a New Next Activity:</label>
								</div>
								<div class="col-lg-8">			
								  	<select class="form-control" name="selProspectNextActivity" id="selProspectNextActivity">
							      	<% 
							      	  
										
									%>									
									</select>
								</div>
							</div>
						</div>	
						
						<div class="col-lg-12" style="margin-top:15px;" id="showActivityAppointmentDuration">
							<div class="form-group">
								<div class="col-lg-4" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Appointment Duration (for Outlook Calendar):</label>
								</div>
								<!-- Get Default Appointment Duration from tblGlobalSettings -->
								<%
									EWSDefaultAppointmentDuration = GetPOSTParams("EWSDEFAULTAPPTDURATION")
								%>
								<div class="col-lg-8">	
	
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
								<div class="col-lg-4" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Meeting Duration (for Outlook Calendar):</label>
								</div>
								<!-- Get Default Meeting Duration from tblGlobalSettings -->
								<%
									EWSDefaultMeetingDuration = GetPOSTParams("EWSDEFAULTMEETINGDURATION")
								%>
								<div class="col-lg-8">	
											
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
							  	<label for="prospectEditNextActivityDate">Due Date:</label>
				                <div class="input-group date" id="datetimepicker1">
				                    <input type="text" class="form-control" name="txtProspectEditNextActivityDate" id="txtProspectEditNextActivityDate">
				                    <span class="input-group-addon">
				                        <span class="glyphicon glyphicon-calendar"></span>
				                    </span>
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
	<!-- add prospects to group modal ends here !-->
	

<!-- modal window next activity -->
<!--#include file="onthefly_nextactivity.asp"--> 
<!-- end modal window nextactivity -->   	


	<!-- modal palceholder for edit prospect stage begins here !-->
	 <!-- Modal -->
	 
		<div class="modal fade" id="myProspectingModalEditStage" tabindex="-1" role="dialog" aria-labelledby="myProspectingModalEditStageLabel">
		
		<style>
		
		 .radio {
		    position: relative;
		    display: inline;
		    margin-top: 10px;
		    margin-bottom: 20px;
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
	
 //-------------------------------------------------------------------------------
	
	
	//Next activity modal window submit
	$('#frmAddStageOnthefly').submit(function(e) {
		
		if ($('#frmAddStageOnthefly #txtstagedescription').val()==''){
			 swal("Stage description can not be blank.");
			return false;
		}
		
		$("#ONTHEFLYmodalStage .btn-primary").html("Saving...");
        $.ajax({
            type: "POST",
            url: "onthefly_addstage_submit.asp",
            data: $('#frmAddStageOnthefly').serialize(),
            success: function(response) {
				
				if (response!=0){
					var stagedesc = $('#frmAddStageOnthefly #txtstagedescription').val();
					var stagetype = $('#frmAddStageOnthefly #selStageType').val();

					var str = '<div class="radio">';
						str += '<label class="btn btn-default">'
						str += '<input name="radStage" id="radStage'+response+'" value="'+response+'" type="radio">'+stagedesc+'</label>';
						str += '</div>'
						
						if (stagetype=='Primary'){
							$('.stageprimarygroup').append(str);
						} else {
							$('.stagesecondarygroup').append(str);
						}
				}
				
				$("#ONTHEFLYmodalStage .modal-body").html('New stage added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
				
            },
            error: function() {
				$("#ONTHEFLYmodalStage .btn-primary").html("Save");
            }
        });
        return false;
    });
//-------------------------------------------------------------------------------	
				    
			});
		
		</script>
		
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalAddLabel">Edit Prospect Stage</h4>
		      </div>
		      <form name="frmEditProspectStageFromModal" id="frmEditProspectStageFromModal" action="editProspectStageFromModal.asp" method="POST">
		      		
		      	  <input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="">
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
								  	<select class="form-control" name="selUnqualifyingReasons" id="selUnqualifyingReasons">
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
								  	<select class="form-control" name="selLostReasons" id="selLostReasons">
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
	
	

	<!-- recycle prospect modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="myProspectingRecycleModal" tabindex="-1" role="dialog" aria-labelledby="myProspectingRecycleLabel">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingRecycleModalLabel">Recycle This Prospect</h4>
		      </div>
		      <form name="frmProspectRecycle" id="frmProspectRecycle">
			      <div class="modal-body">
	
					<div class="col-lg-12">
						<div class="form-group">
							<label class="col-sm-12 control-label">Some Fields Here</label>
							<input type="text" class="form-control required" id="txtProspectRecycle" name="txtProspectRecycle">
						</div>
					</div>
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="button" class="btn btn-primary" data-dismiss="modal">Recycle Prospect</button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- recycle prospect modal ends here !-->
	
<!-- modal window stage -->
<!--#include file="onthefly_addstage.asp"--> 
<!-- end modal window stage -->   	


<!-- ******************************************************************************************************************************** -->
<!-- END MODAL WINDOW DESIGN AND DEFINITIONS -->
<!-- ******************************************************************************************************************************** -->
	
