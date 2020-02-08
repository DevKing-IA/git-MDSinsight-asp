<!--#include file="../inc/header-prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->


<%


 'Read edit prospect tab color settings
SQL = "SELECT * FROM Settings_Global"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	CRMTabLogColor = rs("CRMTabLogColor")
	CRMTabProductsColor = rs("CRMTabProductsColor")
	CRMTabEquipmentColor = rs("CRMTabEquipmentColor")
	CRMTabDocumentsColor = rs("CRMTabDocumentsColor")
	CRMTabLocationColor = rs("CRMTabLocationColor")
	CRMTabContactsColor = rs("CRMTabContactsColor")
	CRMTabCompetitorsColor = rs("CRMTabCompetitorsColor")
	CRMTabOpportunityColor = rs("CRMTabOpportunityColor")
	CRMTabAuditTrailColor	 = rs("CRMTabAuditTrailColor")
	CRMTileOfferingColor = rs("CRMTileOfferingColor")
	CRMTileCompetitorColor = rs("CRMTileCompetitorColor")
	CRMTileDollarsColor = rs("CRMTileDollarsColor")
	CRMTileActivityColor = rs("CRMTileActivityColor")
	CRMTileStageColor = rs("CRMTileStageColor")
	CRMTileOwnerColor = rs("CRMTileOwnerColor")
	CRMTileCommentsColor = rs("CRMTileCommentsColor")
	CRMHideLocationTab = rs("CRMHideLocationTab")
	CRMHideProductsTab = rs("CRMHideProductsTab")
	CRMHideEquipmentTab = rs("CRMHideEquipmentTab")	
End If

SQL = "SELECT * FROM Settings_Prospecting"
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	TabSocialMediaColor = rs("TabSocialMediaColor")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing

If CRMTabLogColor = "" Then CRMTabLogColor = "#D8F9D1"
If IsNull(CRMTabLogColor) Then CRMTabLogColor = "#D8F9D1"

If CRMTabProductsColor = "" Then CRMTabProductsColor = "#D8F9D1"
If IsNull(CRMTabProductsColor) Then CRMTabProductsColor = "#D8F9D1"

If CRMTabEquipmentColor = "" Then CRMTabEquipmentColor = "#FFA500"
If IsNull(CRMTabEquipmentColor) Then CRMTabEquipmentColor = "#FFA500"

If CRMTabDocumentsColor = "" Then CRMTabDocumentsColor = "#F6F6F6"
If IsNull(CRMTabDocumentsColor) Then CRMTabDocumentsColor = "#F6F6F6"

If CRMTabLocationColor = "" Then CRMTabLocationColor = "#D8F9D1"
If IsNull(CRMTabLocationColor) Then CRMTabLocationColor = "#D8F9D1"

If CRMTabContactsColor = "" Then CRMTabContactsColor = "#FCB3B3"
If IsNull(CRMTabContactsColor) Then CRMTabContactsColor = "#FCB3B3"

If CRMTabCompetitorsColor = "" Then CRMTabCompetitorsColor = "#FCB3B3"
If IsNull(CRMTabCompetitorsColor) Then CRMTabCompetitorsColor = "#FCB3B3"

If CRMTabOpportunityColor = "" Then CRMTabOpportunityColor = "#FFA500"
If IsNull(CRMTabOpportunityColor) Then CRMTabOpportunityColor = "#FFA500"

If CRMTabAuditTrailColor = "" Then CRMTabAuditTrailColor = "#FFA500"
If IsNull(CRMTabAuditTrailColor) Then CRMTabAuditTrailColor = "#FFA500"

If CRMTileOfferingColor = "" Then CRMTileOfferingColor = "#3498db"
If IsNull(CRMTileOfferingColor) Then CRMTileOfferingColor = "#3498db"

If CRMTileCompetitorColor = "" Then CRMTileCompetitorColor = "#9b6bcc"
If IsNull(CRMTileCompetitorColor) Then CRMTileCompetitorColor = "#9b6bcc"

If CRMTileDollarsColor = "" Then CRMTileDollarsColor = "#2ecc71"
If IsNull(CRMTileDollarsColor) Then CRMTileDollarsColor = "#2ecc71"

If CRMTileActivityColor = "" Then CRMTileActivityColor = "#f1c40f"
If IsNull(CRMTileActivityColor) Then CRMTileActivityColor = "#f1c40f"

If CRMTileStageColor = "" Then CRMTileStageColor = "#e67e22"
If IsNull(CRMTileStageColor) Then CRMTileStageColor = "#e67e22"

If CRMTileOwnerColor = "" Then CRMTileOwnerColor = "#95a5a6"
If IsNull(CRMTileOwnerColor) Then CRMTileOwnerColor = "#95a5a6"

If CRMTileCommentsColor = "" Then CRMTileCommentsColor = "#d43f3a"
If IsNull(CRMTileCommentsColor) Then CRMTileCommentsColor = "#d43f3a"


MaxActivityDaysWarningInit = GetCRMMaxActivityDaysWarning()
MaxActivityDaysPermittedInit = GetCRMMaxActivityDaysPermitted()

%>

<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.min.js"></script>
<link href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet">

<script>

	function highlightBlankFields(){
	  $(".input-group .form-control").each(function() {
	     var val = $(this).val();
	     if(val == "" || val == 0) {
	       $(this).css({ backgroundColor:'#ffff99' });
	     }
	     else {
	     	$(this).css({ backgroundColor:'#fff' });
	     }
	  });
	}

	$(window).load(function()
	{
	   var phones = [{ "mask": "(###) ###-####"}];
	    $('#txtPhoneNumber').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtCellPhoneNumber').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtFaxNumber').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	        
	});

	$(document).ready(function() {
	

		highlightBlankFields();
		
		$(".input-group .form-control").blur(function() {
		  highlightBlankFields()
		});	
		
		
		$('input, select, textarea').each(
		    function(index){  
		        var input = $(this);
		        //console.log(input.attr('name'));
		    }
		); 
		
	  $("#txtNumEmployees").change(function() {
	  
	  		if ($("#txtProjectedGPSpend").val() == '')
	  		{
				intRecID = $("#txtNumEmployees").val();
				projGPSpend = $("#" + intRecID).val();
				$("#txtProjectedGPSpend").val(projGPSpend);
			}
	  });
	  

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
	    
		$("#selProspectOwner").change(function() {
			
			myProspectID = $("#txtInternalRecordIdentifier").val(); 
			newOwnerUserNo = $("#selProspectOwner").val();
			
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
								 	$("#showActivityMeetingDuration").show();
								 	$("#showActivityMeetingLocation").show();
             	 
				               		
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
	function isValidPhone(p) {
	  //var phoneRe = /^[2-9]\d{2}[2-9]\d{2}\d{4}$/;
	  //var phoneRe = /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/;
	  var phoneRe = /^(1\s|1|)?((\(\d{3}\))|\d{3})(\-|\s)?(\d{3})(\-|\s)?(\d{4})$/;
	  var digits = p.replace(/\D/g, "");
	  return phoneRe.test(digits);
	}
	
	function isValidEmail(email) 
	{
	    var re = /\S+@\S+\.\S+/;
	    return re.test(email);
	}	

   function validateAddProspectForm()
    {
    
       if (document.frmAddProspect.txtCompanyName.value == "") {
            swal("Company name cannot be blank.");
            return false;
       }
       if ((document.frmAddProspect.txtEmailAddress.value !== "") && (isValidEmail(document.frmAddProspect.txtEmailAddress.value) == false)) {
           swal("The email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
       if ((document.frmAddProspect.txtPhoneNumber.value !== "") && (isValidPhone(document.frmAddProspect.txtPhoneNumber.value) == false)) {
           swal("The phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
       if ((document.frmAddProspect.txtPhoneNumberExt.value !== "") && (document.frmAddProspect.txtPhoneNumber.value == "")) {
           swal("A phone extension was added with no phone number. Please enter a phone number or clear the extension.");
           return false;
       }       
       if ((document.frmAddProspect.txtCellPhoneNumber.value !== "") && (isValidPhone(document.frmAddProspect.txtCellPhoneNumber.value) == false)) {
           swal("The cell phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
       if ((document.frmAddProspect.txtFaxNumber.value !== "") && (isValid(document.frmAddProspect.txtFaxNumber.value) == false)) {
           swal("The fax number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }

	   
 		var chkd = document.frmAddProspect.chkBottledWater.checked || +
 			document.frmAddProspect.chkFilteredWater.checked|| +
 			document.frmAddProspect.chkOCS.checked|| +
 			document.frmAddProspect.chkOCS_Supply.checked|| +
 			document.frmAddProspect.chkOfficeSupplies.checked|| +
 			document.frmAddProspect.chkVending.checked|| +
 			document.frmAddProspect.chkMicroMarket.checked|| +
 			document.frmAddProspect.chkPantry.checked;
			
			if (chkd == true)
			{
		       if (document.frmAddProspect.txtPrimaryCompetitor.value == "") {
		            swal("Primary competitor must be selected if offerings are selected.");
		            return false;
		       }
			}
			else
			{
		       if (document.frmAddProspect.txtPrimaryCompetitor.value !== "") {
		            swal("You must select at least one offering for the primary competitor.");
		            return false;
		       }			
		    }   
			

       if (document.frmAddProspect.selProspectNextActivity.value == "99999999") {
            swal("Next activity cannot be blank.");
            return false;
       }
       
       if (document.frmAddProspect.txtNextActivityDueDate.value == "") {
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
	  color: #fff;
	  font-size:16px;
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
    	width:385px !important;
    	border-radius:3px !important;
	}
		
	hr.tile {
	    border: 0;
	    height: 3px;
	    background-image: linear-gradient(to right, rgba(0, 0, 0, 0), rgba(255, 255, 255, 0.95), rgba(0, 0, 0, 0));
	}
	
	.CRMTabLogColor{
		<% Response.Write("background:" & CRMTabLogColor & " !important;") %>
	}
	.CRMTabProductsColor{
		<% Response.Write("background:" & CRMTabProductsColor & " !important;") %>
	}
	.CRMTabEquipmentColor{
		<% Response.Write("background:" & CRMTabEquipmentColor & " !important;") %>
	}
	.CRMTabDocumentsColor{
		<% Response.Write("background:" & CRMTabDocumentsColor & " !important;") %>
	}
	.CRMTabLocationColor{
		<% Response.Write("background:" & CRMTabLocationColor & " !important;") %>
	}
	.CRMTabContactsColor{
		<% Response.Write("background:" & CRMTabContactsColor & " !important;") %>
	}
	.CRMTabCompetitorsColor{
		<% Response.Write("background:" & CRMTabCompetitorsColor & " !important;") %>
	}
	.CRMTabOpportunityColor{
		<% Response.Write("background:" & CRMTabOpportunityColor & " !important;") %>
	}
	.CRMTabAuditTrailColor{
		<% Response.Write("background:" & CRMTabAuditTrailColor & " !important;") %>
	}
	
	
	.CRMTileOfferingColor {
		<% Response.Write("background:" & CRMTileOfferingColor & " !important;") %>
	}
	.CRMTileCompetitorColor {
		<% Response.Write("background:" & CRMTileCompetitorColor & " !important;") %>
	}
	.CRMTileDollarsColor {
		<% Response.Write("background:" & CRMTileDollarsColor & " !important;") %>
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
	.CRMTileCommentsColor {
		<% Response.Write("background:" & CRMTileCommentsColor & " !important;") %>
	}


	.red-line{
		border-left:3px solid red;
	}   
	
	input[type=checkbox] {
	  transform:scale(1.5, 1.5);
	}	  
	

</style>
<!-- eof css !-->
<%
	txtFirstName = Request.form("txtFirstName")
	txtLastName = Request.form("txtLastName")
	txtCompanyName = Request.form("txtCompanyName")
	txtPhoneNumber = Request.form("txtPhoneNumber")
	txtAddress = Request.form("txtAddress")

%>
<h1 class="page-header"><i class="fa fa-fw fa-asterisk"></i> Add New Prospect
	<!-- customize !-->
	<div class="col pull-right">
	</div>
	<!-- eof customize !-->
</h1>

		
<form autocomplete="off" action="<%= BaseURL %>prospecting/addProspect_submit.asp" method="POST" name="frmAddProspect" id="frmAddProspect" onsubmit="return validateAddProspectForm();" class="form-horizontal track-event-form bv-form">
<input autocomplete="false" name="hidden" type="text" style="display:none;">
<div class="container pull-left">
<div class="row">
      <div class="col-md-3">

		<div class="quick-info-block quick-info-block-grey">
		<h2 class="heading-md black"><i class="icon-2x color-light fa fa-user-circle"></i>&nbsp;Business Card - <%= txtCompanyName %></h2>

              <div class="form-group">
 
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
                    		<select data-placeholder="Choose Suffix, Mr., Mrs., etc." class="C_Country_Modal form-control" id="txtSuffix" name="txtSuffix">  
                    			<option value="">Salutation, Mr., Mrs., etc.</option>  
                    			<option value="Mr.">Mr.</option>
								<option value="Mrs.">Mrs.</option>
								<option value="Miss">Miss</option>
								<option value="Dr.">Dr.</option>
								<option value="Ms.">Ms.</option>                     
							</select>
	                    	
	                   </div>
	                </div> 
              
               </div>

              <div class="form-group">
	                            
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtFirstName" placeholder="First Name" name="txtFirstName" value="<%= txtFirstName %>">
	                   </div>
	                </div> 
	                
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtLastName" placeholder="Last Name" name="txtLastName" value="<%= txtLastName %>">
	                   </div>
	                </div> 
	 
               </div>
               

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-briefcase"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtCompanyName" placeholder="Company Name" name="txtCompanyName" value="<%= txtCompanyName %>">
	                   </div>
	                </div> 
	 
               </div>
               

              <div class="form-group">
	                            
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-id-card-o"></i></div>
                    		<select data-placeholder="Choose Job Title" class="C_Country_Modal form-control" id="txtTitle" name="txtTitle">  
                    			
							</select>
   	                   </div>
	                </div> 
	 
               </div>

              <div class="form-group">
	                            
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine1" placeholder="Street Address" name="txtAddressLine1" value="<%= txtAddress %>">
	                   </div>
	                </div> 
	           </div>     
	                
	                
	          <div class="form-group">
	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card-o"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine2" placeholder="Suite, Floor #, etc." name="txtAddressLine2">
	                   </div>
	                </div> 
	 
               </div>



              <div class="form-group">
	                            
	                <div class="col-sm-9">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-book"></i></div>
	                    	<input type="text" class="form-control" id="txtCity" placeholder="City" name="txtCity">
	                   </div>
	                </div> 
	                
	          </div>     
	          <div class="form-group">

	                <div class="col-sm-7">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-book"></i></div>
                    		<select data-placeholder="Choose State" class="C_Country_Modal form-control" id="txtState" name="txtState"> 
                    			<option value="">State</option>
								<!--#include file="statelist.asp"-->
							</select>				
		
	                   </div>
	                </div> 
	                <div class="col-sm-5">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-book"></i></div>
	                    	<input type="text" class="form-control" id="txtZipCode" placeholder="Zip" name="txtZipCode">
	                   </div>
	                </div> 
	 
               </div>
               
                
              <div class="form-group">
	                            	                
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-globe"></i></div>
                    		<select data-placeholder="Choose State" class="C_Country_Modal form-control" id="txtCountry" name="txtCountry"> 
								<!--#include file="countrylist.asp"-->
							</select>

	                   </div>
	                </div> 
	                	 
               </div>
              
              <div class="form-group">
	                            
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtEmailAddress" placeholder="Email Address" name="txtEmailAddress">
	                   </div>
	                </div> 
	          </div>    
	          <div class="form-group">

	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-globe"></i></div>
	                    	<input type="text" class="form-control" id="txtWebsiteURL" placeholder="Company Website URL" name="txtWebsiteURL">
	                   </div>
	                </div> 
	                
               </div>
               

              <div class="form-group">

	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-phone"></i></div>
	                    	<input type="text" class="form-control" id="txtPhoneNumber" placeholder="Phone Number" name="txtPhoneNumber" value="<%= txtPhoneNumber %>">
	                   </div>
	                </div> 
	 
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-phone"></i></div>
	                    	<input type="text" class="form-control" id="txtPhoneNumberExt" placeholder="Extension" name="txtPhoneNumberExt" value="<%= txtPhoneNumberExt %>">
	                   </div>
	                </div> 

	                	 
               </div>

                      
               <div class="form-group">
               
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-mobile"></i></div>
	                    	<input type="text" class="form-control" id="txtCellPhoneNumber" placeholder="Cell Phone Number" name="txtCellPhoneNumber">
	                   </div>
	                </div> 
	                            
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-fax"></i></div>
	                    	<input type="text" class="form-control" id="txtFaxNumber" placeholder="Fax Number" name="txtFaxNumber">
	                   </div>
	                </div> 
	                	 
               </div>
               <div class="form-group">
	                            
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-building"></i></div>
                    		<select data-placeholder="Choose Industry" class="C_Country_Modal form-control" id="txtIndustry" name="txtIndustry"> 
                    		
						</select>


	                   </div>
	                </div> 
	                	 
               </div>
              
                    
			</div>
        <!-- END QUICK INFO BOX -->
        
        
      </div><!-- end col-md-6 -->



      <div class="col-md-3">

		<div class="quick-info-block CRMTileCommentsColor">
		<h2 class="heading-md"><i class="icon-2x color-light fa fa-comment"></i>&nbsp;<%= GetTerm("Comments") %></h2>
	          <div class="form-group">        	                
	                <div class="col-sm-12">
	                  <div class="input-group">
							<textarea class="form-control" id="txtComments" name="txtComments"></textarea>
	                   </div>
	                </div> 	 
	           </div>
		</div>
	  

		<div class="quick-info-block CRMTileDollarsColor">
		<h2 class="heading-md"><i class="icon-2x color-light fa fa-usd"></i>&nbsp;<%= GetTerm("Opportunity") %></h2>

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-credit-card-alt"></i></div>
	                    	<input type="number" class="form-control" id="txtProjectedGPSpend" placeholder="Projected GP Spend (numbers only)" name="txtProjectedGPSpend">
	                   </div>
	                </div> 
	                	 
               </div>

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-users"></i></div>
                    		<select data-placeholder="Choose State" class="C_Country_Modal form-control" id="txtNumEmployees" name="txtNumEmployees"> 
                    			
							</select>
				  	  			<%
				  	  			'Get GP Spend From Employee Range
									SQL9 = "SELECT *, Cast(LEFT(Range,CHARINDEX('-',Range)-1) as int) as Expr1 FROM PR_EmployeeRangeTable "
									SQL9 = SQL9 & "order by Expr1"

									Set cnn9 = Server.CreateObject("ADODB.Connection")
									cnn9.open (Session("ClientCnnString"))
									Set rs9 = Server.CreateObject("ADODB.Recordset")
									rs9.CursorLocation = 3 
									Set rs9 = cnn9.Execute(SQL9)
										
									If not rs9.EOF Then
										Do
											%><input type="hidden" value="<%= rs9("ProjectedGPSpend") %>" id="<%= rs9("InternalRecordIdentifier") %>"><%
											rs9.movenext
										Loop until rs9.eof
									End If
									set rs9 = Nothing
									cnn9.close
									set cnn9 = Nothing
								%>
							

	                   </div>
	                </div> 
	                	 
               </div>

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-apple"></i></div>
                    		<select data-placeholder="Choose State" class="C_Country_Modal form-control" id="txtNumPantries" name="txtNumPantries"> 
                    			<option value="">Select # Pantries</option>
								<% For i = 0 To 50 %>
								  <option value="<%= i %>"><%= i %></option>
								<% Next %>							
							</select>

	                   </div>
	                </div> 
	                	 
               </div>
               
               <div class="form-group">
               		<div class="col-sm-6"><p>Bldg. Lease Expires Date</p></div>
					<div class="col-sm-6">
				        <div class="input-group date" id="datetimepickerLeaseExpiresDate">
				            <input type="text" class="form-control" id="txtLeaseExpirationDate" name="txtLeaseExpirationDate" placeholder="Click Calendar" readonly />
				            <span class="input-group-addon">
				                <span class="glyphicon glyphicon-calendar"></span>
				            </span>
				        </div>
				    </div>
				</div>   
				<script type="text/javascript">
		            $(function () {
		            	
		                $('#datetimepickerLeaseExpiresDate').datetimepicker({
		                   minDate: moment(),
		                   format: 'MM/DD/YYYY',
		                   ignoreReadonly: true
		                
		                });  
		            });
		        </script>    
                
                
               <div class="form-group">
               		<div class="col-sm-6"><p>Contract Expiration Date</p></div>
					<div class="col-sm-6">
				        <div class="input-group date" id="datetimepickerContractExpireDate">
				            <input type="text" class="form-control" id="txtContractExpirationDate" name="txtContractExpirationDate" placeholder="Click Calendar" readonly />
				            <span class="input-group-addon">
				                <span class="glyphicon glyphicon-calendar"></span>
				            </span>
				        </div>
				    </div>
				</div>   
				<script type="text/javascript">
		            $(function () {
		            	
		                $('#datetimepickerContractExpireDate').datetimepicker({
		                   minDate: moment(),
		                   format: 'MM/DD/YYYY',
		                   ignoreReadonly: true
		                
		                });  
		            });
		        </script>                     
			</div>
        <!-- END QUICK INFO BOX -->
        
 		<div class="quick-info-block CRMTileOwnerColor">
		<h2 class="heading-md"><i class="icon-2x color-light fa fa-user"></i>&nbsp;<%= GetTerm("Owner") %></h2>

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
                    		<select data-placeholder="Choose Owner" class="C_Country_Modal form-control" id="selProspectOwner" name="selProspectOwner"> 
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

   			
        <!-- END QUICK INFO BOX -->
      </div><!-- end col-md-6 -->
	
	<div class="col-md-3">
        
   			
    			
		<div class="quick-info-block CRMTileOfferingColor">
		<h2 class="heading-md"><i class="icon-2x color-light fa fa-clock-o"></i>&nbsp;<%= GetTerm("Current Supplier Info") %></h2>
              <div class="form-group">        	                
	                <div class="col-sm-12">
	                  <div class="input-group">
							<textarea class="form-control" id="txtCurrentOffering" name="txtCurrentOffering"></textarea>
	                   </div>
	                </div> 	 
               </div>
   		</div>
   			
  
       

            <div class="quick-info-block CRMTileCompetitorColor">
            <h2 class="heading-md"><i class="icon-2x color-light fa fa-user-circle-o"></i>&nbsp;<%= GetTerm("Primary Competitor") %></h2>
            
              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
                    		<select data-placeholder="Choose Telemarketer" class="C_Country_Modal form-control" id="txtTelemarketerUserNo" name="txtTelemarketerUserNo"> 
							<option value="">Select a Telemarketer</option>
                            
					      	<%'Telemarketer dropdown

				      	  	SQL = "SELECT UserNo, userFirstName, userLastName, userType FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
				      	  	SQL = SQL & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo") & " AND userEnabled = 1"
				      	  	SQL = SQL & " AND userType = 'Telemarketing' "
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
            



              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <div class="input-group">
                    	<div class="input-group-addon"><i class="fa fa-external-link"></i></div>
                		<select data-placeholder="Choose Lead Source" class="C_Country_Modal form-control" id="txtLeadSource" name="txtLeadSource"> 
                    		
						</select>
	                   </div>
	                </div> 
	                	 
               </div>

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <div class="input-group">
                    	<div class="input-group-addon"><i class="fa fa-coffee"></i></div>
                		<select data-placeholder="Choose Primary Competitor" class="C_Country_Modal form-control" id="txtPrimaryCompetitor" name="txtPrimaryCompetitor"> 
                    		
						</select>

	                   </div>
	                </div> 
	                	                
	                	 
               </div>

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
						/*color:#fff !important;*/
					}
					
					.checkbox label, .radio label {
					    min-height: 20px;
					    padding-left: 20px;
					    margin-bottom: 0;
					    font-weight: 400;
					    cursor: pointer;
					    color: #fff;
					}					
			  </style>


              <div class="form-group">
              
              		<h2 class="heading-md" style="margin-bottom:5px;"><i class="icon-2x color-light fa fa-industry"></i>&nbsp;Offerings</h2>
					
					<div class="col-sm-12">
					
						<fieldset class="group"> 
							<ul class="checkbox"> 
							  <li><input type="checkbox" id="chkBottledWater" name="chkBottledWater" ><label for="chkBottledWater">Bottled Water</label></li> 
							  <li><input type="checkbox" id="chkFilteredWater" name="chkFilteredWater"><label for="chkFilteredWater">Filtered Water</label></li> 
							  <li><input type="checkbox" id="chkOCS" name="chkOCS"><label for="chkOCS">OCS</label></li> 
							  <li><input type="checkbox" id="chkOCS_Supply" name="chkOCS_Supply"><label for="chkOCS_Supply">OCS Supply</label></li> 
							  <li><input type="checkbox" id="chkOfficeSupplies" name="chkOfficeSupplies"><label for="chkOfficeSupplies">Office Supplies</label></li> 
							  <li><input type="checkbox" id="chkVending" name="chkVending"><label for="chkVending">Vending</label></li> 
							  <li><input type="checkbox" id="chkMicroMarket" name="chkMicroMarket"><label for="chkMicroMarket">Micromarket</label></li>
							  <li><input type="checkbox" id="chkPantry" name="chkPantry"><label for="chkPantry">Pantry</label></li>
							</ul> 
						</fieldset> 
					</div>	 
					
               </div>
               
 				<div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-id-card-o"></i></div>
	                    	<input type="text" class="form-control" id="txtFormerCustomerNumber" placeholder="Former Customer #" name="txtFormerCustomerNumber">
	                   </div>
	                </div> 
   	 
               </div>              
               
               <div class="form-group">
               		<div class="col-sm-6"><p>Former Customer Cancel Date</p></div>
					<div class="col-sm-6">
				        <div class="input-group date" id="datetimepickerCancelDate">
				            <input type="text" class="form-control" id="txtFormerCustomerCancelDate" name="txtFormerCustomerCancelDate" placeholder="Click Calendar" readonly />
				            <span class="input-group-addon">
				                <span class="glyphicon glyphicon-calendar"></span>
				            </span>
				        </div>
				    </div>
				</div>   
 
               

   			</div>
   			
   			
   		</div>
   		
   	<div class="col-md-3">
   						
		<div class="quick-info-block CRMTileActivityColor">
		<h2 class="heading-md"><i class="icon-2x color-light fa fa-arrow-right"></i>&nbsp;<%= GetTerm("Next Activity") %></h2>

			<input type="hidden" name="txtCRMMaxActivityDaysWarning" id="txtCRMMaxActivityDaysWarning" value="<%= MaxActivityDaysWarning %>">
			<input type="hidden" name="txtCRMMaxActivityDaysPermitted" id="txtCRMMaxActivityDaysPermitted" value="<%= MaxActivityDaysPermitted %>">

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-cog"></i></div>
                    		<select data-placeholder="Choose Next Activity" class="C_Country_Modal form-control red-line" id="selProspectNextActivity" name="selProspectNextActivity"> 
							<option value="99999999">Select Next Activity</option>
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
			
				<div class="col-sm-6"><p>Next Activity Due Date</p></div>
				<div class="col-sm-6">
			        <div class="input-group date" id="datetimepickerNextActivity">
			            <input type="text" class="form-control red-line" id="txtNextActivityDueDate" name="txtNextActivityDueDate" placeholder="Click Calendar" readonly />
			            <span class="input-group-addon">
			                <span class="glyphicon glyphicon-calendar"></span>
			            </span>
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
		
			
			<script type="text/javascript">
	            $(function () {
	                $('#datetimepickerNextActivity').datetimepicker({
	                   minDate: moment(),
	                   format: 'MM/DD/YYYY',
	                   ignoreReadonly: true		                
	                });
	                
	            });
	        </script>            
           

		</div>
   			
   			
   			
   			
            <div class="quick-info-block CRMTileStageColor">
            <h2 class="heading-md"><i class="icon-2x color-light fa fa-tasks"></i>&nbsp;<%= GetTerm("Stage") %></h2>
            

			<style>
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
			</style>

			<div class="form-group">
				<div class="col-sm-12" style="width:400px; margin-left:0px; margin-right:0px; text-align: center;">
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
                	 
			<hr class="tile">

			<div class="form-group">
				<div class="col-sm-12" style="width:400px; margin-left:0px; margin-right:0px; text-align: center;">

					      	<% 
					      		'Get all stages
					      	  	SQLStages = "SELECT * FROM PR_Stages WHERE StageType = 'Secondary' AND InternalRecordIdentifier <> 0 ORDER BY SortOrder"
			
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
				
		
	          <div class="form-group">        	                
	                <div class="col-lg-12">
	                  <div class="input-group">
	                  		<p>Notes For This Stge</p>
							<textarea class="form-control" id="txtStageNotes" name="txtStageNotes"></textarea>
	                   </div>
	                </div> 	 
	           </div>
		
	
      </div><!-- end col-md-6 -->
      

 </div> <!-- end row -->
        
        
<div class="form-group pull-right">
	<div class="col-lg-12" style="margin-top:90px;">
		<button class="btn btn-primary btn-lg btn-block" href="main.asp" role="button" type="submit">SAVE THIS PROSPECT <i class="fa fa-floppy-o" aria-hidden="true"></i></button>
	</div>
</div>
        
</div> <!-- end container -->

</form>
 
 <!-- tabs js !-->
 <script type="text/javascript">
 $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
  e.target // newly activated tab
  e.relatedTarget // previous active tab
})
 </script>
 <!-- eof tabs js !-->
 
 <!-- below codes added by nurba
 #date 03/13/2019
 on the fly modify selectbox options
 -->
<%
'check if user has  perssion to add records on the fly

If userCanEditCRMOnTheFly(Session("UserNO")) = True Then
%> 
<script>
$( document ).ready(function() {
	
	//populate lead source
	PopulateSelecBoxes('txtLeadSource',-1);
	
	//populate industry
	PopulateSelecBoxes('txtIndustry',-1);
	
	//populate employee range
	PopulateSelecBoxes('txtNumEmployees',-1);
	
	//populate primary competitor
	PopulateSelecBoxes('txtPrimaryCompetitor',-1);
	
	//populate contact title
	PopulateSelecBoxes('txtTitle',-1);
	
	
	
//-------------------------------------------------------------------------------	
	// lead source select box change
    $( "#txtLeadSource").change(function() {
		var val = $( "#txtLeadSource option:selected").val();
		if (val== -1){
			//deselect add new row
			$('#txtLeadSource option[selected="selected"]').each(
				function() {
					$(this).removeAttr('selected');
				}
			);

			// mark the first option as selected
			$("#txtLeadSource option:first").attr('selected','selected');
			
			//show modal
			$('#ONTHEFLYmodalLeadSource').modal('show');
			
		}
	});
	
	//lead source modal window submit
	$('#frmAddLeadSource').submit(function(e) {
		
		if ($('#frmAddLeadSource #txtLeadSource').val()==''){
			 swal("Lead source can not be blank.");
			return false;
		}
		
		$("#ONTHEFLYmodalLeadSource .btn-primary").html("Saving...");
        $.ajax({
            type: "POST",
            url: "onthefly_LeadSource_submit.asp",
            data: $('#frmAddLeadSource').serialize(),
            success: function(response) {
				PopulateSelecBoxes('txtLeadSource',-1);
				$("#ONTHEFLYmodalLeadSource .modal-body").html('Lead Source added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
                //alert(response['response']);
            },
            error: function() {
				$("#ONTHEFLYmodalLeadSource .btn-primary").html("Save");
                alert('Error add lead source');
            }
        });
        return false;
    });
//-------------------------------------------------------------------------------


//-------------------------------------------------------------------------------
	// Industry select box change
    $( "#txtIndustry").change(function() {
		var val = $( "#txtIndustry option:selected").val();
		if (val== -1){
			//deselect add new row
			$('#txtIndustry option[selected="selected"]').each(
				function() {
					$(this).removeAttr('selected');
				}
			);

			// mark the first option as selected
			$("#txtIndustry option:first").attr('selected','selected');
			
			//show modal
			$('#ONTHEFLYmodalIndustry').modal('show');
			
		}
	});
	
	//Industry modal window submit
	$('#frmAddIndustry').submit(function(e) {
		
		if ($('#frmAddIndustry #txtIndustry').val()==''){
			 swal("Industry can not be blank.");
			return false;
		}
		
		$("#ONTHEFLYmodalIndustry .btn-primary").html("Saving...");
        $.ajax({
            type: "POST",
            url: "onthefly_industry_submit.asp",
            data: $('#frmAddIndustry').serialize(),
            success: function(response) {
				PopulateSelecBoxes('txtIndustry',-1);
				$("#ONTHEFLYmodalIndustry .modal-body").html('Industry added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
                //alert(response['response']);
            },
            error: function() {
				$("#ONTHEFLYmodalIndustry .btn-primary").html("Save");
                alert('Error add industry');
            }
        });
        return false;
    });
//-------------------------------------------------------------------------------



//-------------------------------------------------------------------------------
	// Employee range select box change
    $( "#txtNumEmployees").change(function() {
		var val = $( "#txtNumEmployees option:selected").val();
		if (val== -1){
			//deselect add new row
			$('#txtNumEmployees option[selected="selected"]').each(
				function() {
					$(this).removeAttr('selected');
				}
			);

			// mark the first option as selected
			$("#txtNumEmployees option:first").attr('selected','selected');
			
			//show modal
			$('#ONTHEFLYmodalEmployeeRange').modal('show');
			
		}
	});
	
	//Employee range modal window submit
	$('#frmAddEmployeeRange').submit(function(e) {
		
		
        if ($('#frmAddEmployeeRange #txtEmployeeRange1').val() == "") {
            swal("Beginning Employee Range can not be blank.");
            return false;
        }
        
        if (!isInt($('#frmAddEmployeeRange #txtEmployeeRange1').val())) {
             swal("Beginning Employee Range must be a whole number.");
            return false;
        }

         if ($('#frmAddEmployeeRange #txtEmployeeRange2').val() == "") {
            swal("Ending Employee Range can not be blank.");
            return false;
        }
        
        if (!isInt($('#frmAddEmployeeRange #txtEmployeeRange2').val())) {
             swal("Ending Employee Range must be a whole number.");
            return false;
        }
        
		if(parseInt($('#frmAddEmployeeRange #txtEmployeeRange1').val()) > parseInt($('#frmAddEmployeeRange #txtEmployeeRange2').val()))
		{
            swal("Ending employee range must be greater than beginning employee range.");
            return false;
		}
		 
		if(parseInt($('#frmAddEmployeeRange #txtEmployeeRange1').val())==parseInt($('#frmAddEmployeeRange #txtEmployeeRange2').val()))
		{
	        swal("Beginning and ending employee ranges cannot be equal");
	        return false;
		
		}  
				
		
		$("#ONTHEFLYmodalEmployeeRange .btn-primary").html("Saving...");
        $.ajax({
            type: "POST",
            url: "onthefly_EmployeeRange_submit.asp",
            data: $('#frmAddEmployeeRange').serialize(),
            success: function(response) {
				PopulateSelecBoxes('txtNumEmployees',-1);
				$("#ONTHEFLYmodalEmployeeRange .modal-body").html('Employee range added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
                //alert(response['response']);
            },
            error: function() {
				$("#ONTHEFLYmodalEmployeeRange .btn-primary").html("Save");
                alert('Error add industry');
            }
        });
        return false;
    });
//-------------------------------------------------------------------------------	

//-------------------------------------------------------------------------------
	// primary competitor select box change
    $( "#txtPrimaryCompetitor").change(function() {
		var val = $( "#txtPrimaryCompetitor option:selected").val();
		if (val== -1){
			//deselect add new row
			$('#txtPrimaryCompetitor option[selected="selected"]').each(
				function() {
					$(this).removeAttr('selected');
				}
			);

			// mark the first option as selected
			$("#txtPrimaryCompetitor option:first").attr('selected','selected');
			
			//show modal
			$('#ONTHEFLYmodalCompetitor').modal('show');
			
		}
	});
	
	//competitor modal window submit
	$('#frmAddCompetitor').submit(function(e) {
		
		
        if ($('#frmAddCompetitor #txtCompetitorName').val() == "") {
            swal("Competitor name cannot be blank.");
            return false;
        }

        if ($('#frmAddCompetitor #txtCompetitorAddressInfo').val() == "") {
            swal("Please enter address information for competitor.");
            return false;
        }
		/*
		if ($('#frmAddCompetitor #txtCompetitorWebsite').val() == "") {
            swal("Competitor web site cannot be blank.");
            return false;
        }

        if ($('#frmAddCompetitor #txtCompetitorAdditionalNotes').val() == "") {
            swal("Please enter additional notes for competitor.");
            return false;
        }
		*/		
		
		$("#ONTHEFLYmodalCompetitor .btn-primary").html("Saving...");
        $.ajax({
            type: "POST",
            url: "onthefly_competitor_submit.asp",
            data: $('#frmAddCompetitor').serialize(),
            success: function(response) {
				PopulateSelecBoxes('txtPrimaryCompetitor',-1);
				$("#ONTHEFLYmodalCompetitor .modal-body").html('Competitor added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
                //alert(response['response']);
            },
            error: function() {
				$("#ONTHEFLYmodalCompetitor .btn-primary").html("Save");
                alert('Error add competitor');
            }
        });
        return false;
    });
//-------------------------------------------------------------------------------

//-------------------------------------------------------------------------------
	// contact title select box change
    $( "#txtTitle").change(function() {
		var val = $( "#txtTitle option:selected").val();
		if (val== -1){
			//deselect add new row
			$('#txtTitle option[selected="selected"]').each(
				function() {
					$(this).removeAttr('selected');
				}
			);

			// mark the first option as selected
			$("#txtTitle option:first").attr('selected','selected');
			
			//show modal
			$('#ONTHEFLYmodalContactTitle').modal('show');
			
		}
	});
	
	//contact title modal window submit
	$('#frmAddContactTitle').submit(function(e) {
		
		if ($('#frmAddContactTitle #txtContactTitle').val()==''){
			 swal("Job Title can not be blank.");
			return false;
		}
		
		$("#ONTHEFLYmodalContactTitle .btn-primary").html("Saving...");
        $.ajax({
            type: "POST",
            url: "onthefly_contacttitle_submit.asp",
            data: $('#frmAddContactTitle').serialize(),
            success: function(response) {
				PopulateSelecBoxes('txtTitle',-1);
				$("#ONTHEFLYmodalContactTitle .modal-body").html('Job Title added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
                //alert(response['response']);
            },
            error: function() {
				$("#ONTHEFLYmodalContactTitle .btn-primary").html("Save");
                alert('Error add lead source');
            }
        });
        return false;
    });
//-------------------------------------------------------------------------------	
	
});

//common function to populate selectboxes
function PopulateSelecBoxes(selectid){
    $.ajax({
        type: "POST",
        url: 'onthefly_selectboxes.asp',
        data: ({ section : selectid, action:'add' }),
        dataType: "html",
        success: function(data) {
            $("#"+selectid).html(data);
        },
        error: function() {
            alert('Error occured');
        }
    });	
}
</script> 

<!-- modal window add lead source -->
<!--#include file="onthefly_leadsource.asp"--> 
<!-- end modal window add lead source -->

<!-- modal window add industries -->
<!--#include file="onthefly_industry.asp"--> 
<!-- end modal window add industries -->

<!-- modal window employee range -->
<!--#include file="onthefly_employeerange.asp"--> 
<!-- end modal window employee range -->

<!-- modal window competitor -->
<!--#include file="onthefly_competitor.asp"--> 
<!-- end modal window competitor -->
 
<!-- modal window contact title -->
<!--#include file="onthefly_contacttitle.asp"--> 
<!-- end modal window contact title -->

<%
End If 'userCanEditCRMOnTheFly
%><!-- nurba end--><!--#include file="../inc/footer-main.asp"-->