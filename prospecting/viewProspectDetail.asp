<!--#include file="../inc/header-prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="viewProspectDetailStylesheet.asp"-->

<% 

InternalRecordIdentifier = Request.QueryString("i") 
If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")

OpenTabNum = 1 'Dfault open tab #
OpenTabNum = Request.QueryString("t") 

MaxActivityDaysWarningInit = GetCRMMaxActivityDaysWarning()
MaxActivityDaysPermittedInit = GetCRMMaxActivityDaysPermitted()


Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " viewed the prospect, <strong>" & GetProspectNameByNumber(InternalRecordIdentifier) & "</strong>."
CreateAuditLogEntry GetTerm("Prospecting") & ", prospect detail viewed",GetTerm("Prospecting") & ", prospect detail viewed","Minor",0,Description

Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " viewed this prospect."
Record_PR_Activity InternalRecordIdentifier,Description,Session("UserNo")

%>

<%

SQL = "SELECT * FROM PR_Prospects where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)

If not rs.EOF Then
		InternalRecordIdentifier = rs("InternalRecordIdentifier")
		Company = rs("Company")
		Street= rs("Street")
		City= rs("City")
		State= rs("State")
		PostalCode = rs("PostalCode")
		Country= rs("Country")
		Suite= rs("Floor_Suite_Room__c")						
		Website= rs("Website")								
		LeadSourceNumber = rs("LeadSourceNumber")
		LeadSource = GetLeadSourceByNum(LeadSourceNumber)				
		StageNumber = GetProspectCurrentStageByProspectNumber(InternalRecordIdentifier)
		IndustryNumber = rs("IndustryNumber")	
		Industry = GetIndustryByNum(IndustryNumber)											
		OwnerUserNo = rs("OwnerUserNo")				
		CreatedDate= rs("CreatedDate")
		CreatedByUserNo= rs("CreatedByUserNo")				
		TelemarketerUserNo = rs("TelemarketerUserNo")
		Telemarketer = GetUserDisplayNameByUserNo(TelemarketerUserNo)
		ProjectedGPSpend= rs("ProjectedGPSpend")
		NumberOfPantries = rs("NumberOfPantries")
		EmployeeRangeNumber = rs("EmployeeRangeNumber")
		NumEmployees = GetEmployeeRangeByNum(EmployeeRangeNumber)
		CreatedDate = rs("CreatedDate")
		FormerCustNum = rs("FormerCustNum")
		CancelDate = rs("CancelDate")
		LeaseExpirationDate = rs("LeaseExpirationDate")	
		ContractExpirationDate = rs("ContractExpirationDate")
		Comments = rs("Comments")
		CurrentOffering = rs("CurrentOffering")
		LastVerifiedDate = rs("LastVerifiedDate")			
		
		full_address = Street & " " & Suit & ", " & City & ", " & State & ", " & PostalCode

End If

PrimaryCompetitorID = GetPrimaryCompetitorIDByProspectNumber(InternalRecordIdentifier)


%>

<!-- datetime picker !-->
<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.min.js"></script>
<link href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet">
<!-- end datetime picker !-->

<!-- date picker !-->
<link rel="stylesheet" href="<%= baseURL %>css/datepicker/BeatPicker.min.css"/>
<script src="<%= baseURL %>js/datepicker/BeatPicker.min.js"></script>
<!-- eof date picker !-->

<!-- function that gets the value of the tab when it is clicked and then
updates the value of a hidden form field so when the page posts, it returns
back to the tab that was previously opened -->

<script type="text/javascript">

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


	$(function () {
		$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
		var target = $(e.target).attr("href");
		$('input[name="txtTab"]').val(target);
		});
	})
	
	$(document).ready(function() {
	  $('#filter-contacts').keyup(function() {
	    //alert('Handler for .keyup() called.');
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

	        $('#datetimepickerVerifyDate').datetimepicker({
	        	useCurrent: true,
				minDate:moment(),
                format: 'MM/DD/YYYY',
                ignoreReadonly: true,
                sideBySide: true 
	        	
			}); 
			$('#txtProspectEditVerifyDate').datetimepicker({
	        	useCurrent: true,
				minDate:moment(),
                format: 'MM/DD/YYYY',
                ignoreReadonly: true,
                sideBySide: true 
	        	
			}); 

		$('body').on('hidden.bs.modal', function () {
			if($('.modal.in').length > 0)
			{
				$('body').addClass('modal-open');
			}
		});
								 
		       
		$('#myProspectingModalEditBusinessCard').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myProspectID = $(e.relatedTarget).data('prospect-id');	
	
	    	var $modal = $(this);
	    	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetProspectBusinessCardInformationForModal&myProspectID=" + encodeURIComponent(myProspectID),
				success: function(response)
				 {
	               	 $modal.find('#prospectBusinessCardInfo').html(response);
					 
// below code added by nurba
// 03/15/2019					 
					 PopulateSelecBoxes('txtIndustry','<%=IndustryNumber%>');	  
					 
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
			//$('#myProspectingModalEditBusinessCard').modal('hide');
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
				PopulateSelecBoxes('txtIndustry','<%=IndustryNumber%>');
				$("#ONTHEFLYmodalIndustry .modal-body").html('Industry added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
				
            },
            error: function() {
				$("#ONTHEFLYmodalIndustry .btn-primary").html("Save");
                //alert('Error add industry');
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
			//$('#myProspectingModalEditBusinessCard').modal('hide');
			$('#ONTHEFLYmodalContactTitle').modal('show');
			
		}
	});
	
	//contact title modal window submit
	$('#frmAddContactTitle').submit(function(e) {
		
		if ($('#frmAddContactTitle #txtTitle').val()==''){
			 swal("Contact title can not be blank.");
			return false;
		}
		
		$("#ONTHEFLYmodalContactTitle .btn-primary").html("Saving...");
        $.ajax({
            type: "POST",
            url: "onthefly_contacttitle_submit.asp",
            data: $('#frmAddContactTitle').serialize(),
            success: function(response) {
				PopulateSelecBoxes('txtTitle','');
				$("#ONTHEFLYmodalContactTitle .modal-body").html('Contact title added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
				
            },
            error: function() {
				$("#ONTHEFLYmodalContactTitle .btn-primary").html("Save");
                //alert('Error add industry');
            }
        });
        return false;
    });	
//end nurba
					              	 
	             },
	             failure: function(response)
				 {
				   $modal.find('#prospectBusinessCardInfo').html("Failed");
	             }
			});
	    	    		    
		});
 
 





		$('#myProspectingModalEditOwner').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myProspectID = $(e.relatedTarget).data('prospect-id');	
		    var myOwnerUserNo = $(e.relatedTarget).data('owner-no');
		    //populate the textbox with the user number of the current owner
		    $(e.currentTarget).find('input[name="txtOrigOwnerUserNo"]').val(myOwnerUserNo);
	
	    	var $modal = $(this);
	    	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetProspectOwnerInformationForModal&myProspectID=" + encodeURIComponent(myProspectID) + "&myOwnerUserNo=" + encodeURIComponent(myOwnerUserNo),
				success: function(response)
				 {
	               	 $modal.find('#prospectOwnerDropdown').html(response);	               	 
	             },
	             failure: function(response)
				 {
				   $modal.find('#prospectOwnerDropdown').html("Failed");
	             }
			});
	    	    		    
		});

		
		$('#myProspectingModalEditComments').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myProspectID = $(e.relatedTarget).data('prospect-id');	
	
	    	var $modal = $(this);
	    	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetProspectCommentsInformationForModal&myProspectID=" + encodeURIComponent(myProspectID),
				success: function(response)
				 {
	               	 $modal.find('#prospectComments').html(response);	               	 
	             },
	             failure: function(response)
				 {
				   $modal.find('#prospectComments').html("Failed");
	             }
			});
	    	    		    
		});
		
		
		$('#myProspectingModalEditOpportunity').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myProspectID = $(e.relatedTarget).data('prospect-id');	
	
	    	var $modal = $(this);
	    	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetProspectOpportunityInformationForModal&myProspectID=" + encodeURIComponent(myProspectID),
				success: function(response)
				 {
	               	 $modal.find('#prospectOpportunityInfo').html(response);
					 
// below code added by nurba
// 03/15/2019					 
					 PopulateSelecBoxes('txtNumEmployees','<%=EmployeeRangeNumber%>');	  
					 
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
				PopulateSelecBoxes('txtNumEmployees','<%=EmployeeRangeNumber%>');
				$("#ONTHEFLYmodalEmployeeRange .modal-body").html('Employee range added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
            },
            error: function() {
				$("#ONTHEFLYmodalEmployeeRange .btn-primary").html("Save");
                //alert('Error add industry');
            }
        });
        return false;
    });
//-------------------------------------------------------------------------------	

//end nurba					 
					 	               	 
	             },
	             failure: function(response)
				 {
				   $modal.find('#prospectOpportunityInfo').html("Failed");
	             }
			});
	    	    		    
		});
		
		
 
		$('#myProspectingModalEditCurrentSupplierInfo').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myProspectID = $(e.relatedTarget).data('prospect-id');	
	
	    	var $modal = $(this);
	    	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetProspectCurrentSupplierInformationForModal&myProspectID=" + encodeURIComponent(myProspectID),
				success: function(response)
				 {
	               	 $modal.find('#prospectCurrentSupplierInfo').html(response);	               	 
	             },
	             failure: function(response)
				 {
				   $modal.find('#prospectCurrentSupplierInfo').html("Failed");
	             }
			});
	    	    		    
		});
		
		$('#myProspectingModalEditCompetitorSource').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myProspectID = $(e.relatedTarget).data('prospect-id');	
	
	    	var $modal = $(this);
	    	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetProspectCompetitorSourceInformationForModal&myProspectID=" + encodeURIComponent(myProspectID),
				success: function(response)
				 {
	               	 $modal.find('#prospectCompetitorSource').html(response);	
					 
// below code added by nurba
// 03/15/2019					 
					 PopulateSelecBoxes('txtPrimaryCompetitor','<%=PrimaryCompetitorID%>');
					 PopulateSelecBoxes('txtLeadSource','<%=LeadSourceNumber%>');	  
					 
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
				PopulateSelecBoxes('txtLeadSource','<%=LeadSourceNumber%>');
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
				PopulateSelecBoxes('txtPrimaryCompetitor','<%=PrimaryCompetitorID%>');
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
//end nurba
					                	 
	             },
	             failure: function(response)
				 {
				   $modal.find('#prospectCompetitorSource').html("Failed");
	             }
			});
	    	    		    
		});
		
		

	//verify date modal window submit
	$('#frmUpdateVerifyDate').submit(function(e) {
		
		
        if ($('#frmUpdateVerifyDate #txtProspectEditVerifyDate').val() == "") {
            swal("Last verify date name cannot be blank.");
            return false;
        }


		$("#frmUpdateVerifyDate .btn-primary").html("Saving...");
        $.ajax({
            type: "POST",
            url: "onthefly_lastverifieddate_submit.asp",
            data: $('#frmUpdateVerifyDate').serialize(),
            success: function(response) {
				$('#txtlastverfiydate').html($('#frmUpdateVerifyDate #txtProspectEditVerifyDate').val());
				var datediff = days_between($('#frmUpdateVerifyDate #txtProspectEditVerifyDate').val());
				if (datediff==0) {
					$('#txtlastverfiydaterange').html("today");
				} else if (datediff==1){
					$('#txtlastverfiydaterange').html("1 day ago");
				} else {
					$('#txtlastverfiydaterange').html(datediff+" days ago");
						
				}
				$("#myProspectingModalEditVerifyDate .modal-body").html('Last verify date updated successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
            },
            error: function() {
				$("#frmUpdateVerifyDate .btn-primary").html("Save");
            }
        });
        return false;
    });
			
			
	//verify date modal window submit
	$('.btnUpdateVerifyDateAuto').click(function(e) {
		
		//$("#frmUpdateVerifyDate .btn-primary").html("Saving...");
		$(".btnUpdateVerifyDateAutoStatus").addClass("fa-spin");
		
        $.ajax({
            type: "POST",
            url: "onthefly_lastverifieddate_auto.asp",
			data: { "dateInternalRecordIdentifier": "<%=InternalRecordIdentifier%>" },
            success: function(response) {
				$('#txtlastverfiydate').html(response);
				$('#txtlastverfiydaterange').html("today");
				$(".btnUpdateVerifyDateAutoStatus").removeClass("fa-spin");				
				
            },
            error: function() {
				//$("#frmUpdateVerifyDate .btn-primary").html("Save");
				$(".btnUpdateVerifyDateAutoStatus").removeClass("fa-spin");
				 
            }
        });
        return false;
    });
			
			
						
function days_between(date1) {

    // The number of milliseconds in one day
    var ONE_DAY = 1000 * 60 * 60 * 24;
	
	var arrdate = date1.split("/");
	date1 = new Date(arrdate[2]+"-"+arrdate[0]+"-"+arrdate[1]);

    // Convert both dates to milliseconds
    var date1_ms = date1.getTime();
	var today = new Date();
    var date2_ms = today.getTime();

    // Calculate the difference in milliseconds
    var difference_ms = date2_ms-date1_ms;

    // Convert back to days and return
    return Math.round(difference_ms/ONE_DAY);

}				
		
		$('#myProspectingModalEditActivity').on('show.bs.modal', function(e) {

		
		    //get data-id attribute of the clicked prospect
		    var myProspectID = $(e.relatedTarget).data('prospect-id');	
		    var myActivityRecID = $(e.relatedTarget).data('activity-id');
		    //populate the textbox with the id of the clicked prospect
			

		    $(e.currentTarget).find('input[name="txtActivityRecID"]').val(myActivityRecID);
		    	    
		    var $modal = $(this);
		
			$("#activityDateWarning").hide();
			
			
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetInitialActivityAppmtOrMeeting&myActivityRecID="+myActivityRecID,
				success: function(response)
				 {
					 
					 
					 
// below code added by nurba
// 04/26/2019					 
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
					 
					 
					 
					 
						if (response == "Appointment")
						{					
						    $("#showActivityAppointmentDuration").show();
						    $("#showActivityMeetingDuration").hide();									
						}
						else if (response == "Meeting")
						{					
							$("#showActivityAppointmentDuration").hide();     		
						               		
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
						    $("#showActivityAppointmentDuration").hide();
						    $("#showActivityMeetingDuration").hide();
		           		}
	               	 
	             }
			});
		
	    
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetProspectActivityInformationForModal&myProspectID=" + encodeURIComponent(myProspectID) + "&myActivityRecID=" + encodeURIComponent(myActivityRecID),
				success: function(response)
				 {
	               	 $modal.find('#prospectCurrentActivitySummary').html(response);	               	 
	             },
	             failure: function(response)
				 {
				   $modal.find('#prospectCurrentActivitySummary').html("Failed");
	             }
			});
			
		    
		});
	
	    
		$('#myProspectingModalEditStage').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myProspectID = $(e.relatedTarget).data('prospect-id');	 
		    var myStageRecID = $(e.relatedTarget).data('stage-id');
		    //populate the textbox with the id of the clicked prospect

		    $(e.currentTarget).find('input[name="txtStageRecID"]').val(myStageRecID);
		    	    
		    var $modal = $(this);
		
	    
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetProspectStageInformationForModal&myProspectID=" + encodeURIComponent(myProspectID) + "&myStageRecID=" + encodeURIComponent(myStageRecID),
				success: function(response)
				 {
	               	 $modal.find('#prospectCurrentStageSummary').html(response);
	             },
	             failure: function(response)
				 {
				   $modal.find('#prospectCurrentStageSummary').html("Failed");
	             }
			});
			
// below code added by nurba
// 04/28/2019					 
  
					 
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

//end nurba			
		    
		});
	  
	});	
	
	
	function ajaxRowMode(type, id, mode) {
	
		$('#ajaxRow'+type+'-'+id).attr("class", "ajaxRow"+mode);
		if(id==0){
			$('#ajaxRow'+type+'-' + 0 + '').remove();
		}	
	
		 $(".ajaxRowEdit").find('input[disabled="true"]').each(function () {
		     $(this).removeAttr("disabled");
		 });
		 
		
	   var phonesTab = [{ "mask": "(###) ###-####"}];
	    $('#txtPhoneTab' + id).inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtCellTab' + id).inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtFaxTab' + id).inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	
			if (mode=='Edit'){

			curcontactid = $('#ajaxRowContacts-' + id + ' [data-type="ContactTitleNumber"]').val();

			$('#ajaxRowContacts-' + curcontactid + ' [data-type="ContactTitleNumber"]').empty();
						
						var titlerow='';
												
						
						$.each(ContactTitles, function (key, ContactTitle) {
							//console.log(ContactTitle.id);
							if (ContactTitle.id== -1){
								titlerow ='<option value="'+ContactTitle.id+'"  style="font-weight:bold">'+ContactTitle.title+'</option>';
							} else if (ContactTitle.id== id) {
								titlerow ='<option value="'+ContactTitle.id+'" selected>'+ContactTitle.title+'</option>';
							} else {
								titlerow ='<option value="'+ContactTitle.id+'">'+ContactTitle.title+'</option>';
							}
							$('#ajaxRowContacts-' + curcontactid + ' [data-type="ContactTitleNumber"]').append(titlerow);
						});
			}
									
	}
	
</script>




<!-- title / lead owner !-->
<div class="row">
	<div class="page-header">

		<div class="col-lg-3">
			<%
		
			SQLContacts1 = "SELECT * FROM PR_ProspectContacts WHERE ProspectIntRecID = " & InternalRecordIdentifier & " AND PrimaryContact = 1"
			
			Set cnnContacts1 = Server.CreateObject("ADODB.Connection")
			cnnContacts1.open (Session("ClientCnnString"))
			Set rsContacts1 = Server.CreateObject("ADODB.Recordset")
			rsContacts1.CursorLocation = 3 
			Set rsContacts1 = cnnContacts1.Execute(SQLContacts1)
			
			If not rsContacts1.EOF Then
			
			  	primarySuffix = rsContacts1("Suffix")
			  	primaryFirstName = rsContacts1("FirstName")
				primaryLastName = rsContacts1("LastName")	
				primaryTitleNumber = rsContacts1("ContactTitleNumber")
				primaryTitle = GetContactTitleByNum(primaryTitleNumber)
				primaryEmail = rsContacts1("Email") 
				primaryPhone = rsContacts1("Phone")
				primaryPhoneExt = rsContacts1("PhoneExt")
				primaryCell = rsContacts1("Cell")
				primaryFax = rsContacts1("Fax")

								
			End If
			Set rsContacts1 = Nothing
			cnnContacts1.Close
			Set cnnContacts1 = Nothing
				
			%>
            <div class="business-card">
				<% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
					<!-- User Has READONLY Access -->
				<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
					<a style="position:absolute; right:20px" class="pull-right" data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditBusinessCard" data-tooltip="true" data-title="Edit Business Card"><button class="btn btn-success" role="button" type="button"><i class="fas fa-pen-square fa-lg" aria-hidden="true"></i></button></a>					
				<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
					<!-- User Has WRITEOWNED Access and Is Not The Owner -->				
				<% ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then %>
					<a style="position:absolute; right:20px" class="pull-right" data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditBusinessCard" data-tooltip="true" data-title="Edit Business Card"><button class="btn btn-success" role="button" type="button"><i class="fas fa-pen-square fa-lg" aria-hidden="true"></i></button></a>					
				<% End If %>
								   
                <div class="media">
                    <div class="media-left">
                        <img class="media-object img-circle profile-img" src="http://s3.amazonaws.com/37assets/svn/765-default-avatar.png">
                        <small style="margin-left:5px;">(<%=InternalRecordIdentifier%>)</small><br>
                    </div>
                    <div class="media-body">
                    	<h2 class="company"><%= Company %></h2>
                        <h2 class="name"><%= primarySuffix %>&nbsp;<%= primaryFirstName %>&nbsp;<%= primaryLastName %></h2>
                        
                        <% If primaryTitle <> "0" Then %>
                        	<div class="job"><%= primaryTitle %></div>
                        <% End If %>
                        
                        <div class="address"><nobr><a href="https://maps.google.com/?q=<%=full_address%>" target="_blank"><%= Street %> &nbsp; <%= Suite %></a></nobr></div>
                        
                        
                        <% If State <> "" AND City <> "" AND PostalCode <> "" Then %>
                        	<div class="address"><a href="https://maps.google.com/?q=<%=full_address%>" target="_blank"><%= City %>, <%= State %>&nbsp; <%= PostalCode %></a></div>
                        <% End If %>
                        
                        <% If primaryPhone <> "" Then %>
                        	<div class="phone"><i class="fa fa-phone" aria-hidden="true"></i>&nbsp;&nbsp;<%= primaryPhone %>
                        	<% If primaryPhoneExt <> "" Then %>
                        		&nbsp;&nbsp;Ext. <%= primaryPhoneExt %>
                        	<% End If %>
                        	</div>
                        <% End If %>
                         
                        <% If primaryCell <> "" Then %>
                        	<div class="cell"><i class="fa fa-mobile fa-lg" aria-hidden="true"></i>&nbsp;&nbsp;&nbsp;<%= primaryCell %> (cell)</div>
                        <% End If %>
                        
                        <% If primaryFax <> "" Then %>
                        	<div class="fax"><i class="fa fa-fax" aria-hidden="true"></i>&nbsp;&nbsp;<%= primaryFax %> (fax)</div>
                        <% End If %>                        
                        
                        <% If primaryEmail <> "" Then %>
                        	<div class="mail"><i class="fa fa-envelope" aria-hidden="true"></i>&nbsp;&nbsp;<a href="mailto:<%= primaryEmail %>"><%= primaryEmail %></a></div>
                        <% End If %>
                     
                        <% If Industry <> "" Then %>
                        	<div class="address"><%= Industry %></div>
                        <% End If %>
                        
                        
                        <% If Website <> "" Then %>
                        	<div class="website"><i class="fa fa-globe" aria-hidden="true"></i>&nbsp;&nbsp;<a href="http://<%= Website %>" target="_blank"><%= Website %></a></div>
                        <% End If %>

                    </div>
                    <div class="media-footer">
                    <strong>Last verified: <span id="txtlastverfiydate">
                    <%
					If IsDate(LastVerifiedDate) Then
						Response.Write(DateValue(LastVerifiedDate))
					Else
						Response.Write("never")
					End If
					
					%>
                    </span> <%'=primaryLastVerified%></strong>
                    <% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
					<!-- User Has READONLY Access -->
				<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
					<a  data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditVerifyDate" data-tooltip="true" data-title="Edit Last Verified Date"><button class="btn btn-success" role="button" type="button" style="padding:3px; line-height:10px"><i class="fas fa-pen-square fa-sm" aria-hidden="true"></i></button></a>					
				<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
					<!-- User Has WRITEOWNED Access and Is Not The Owner -->				
				<% ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then %>
					<a  data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditVerifyDate" data-tooltip="true" data-title="Edit Last Verified Date"><button class="btn btn-success" role="button" type="button" style="padding:3px; line-height:10px"><i class="fas fa-pen-square fa-sm" aria-hidden="true"></i></button></a>	<button class="btn btn-success  btnUpdateVerifyDateAuto" role="button" type="button" style="padding:3px; line-height:10px"><i class="fas fa-sync-alt  btnUpdateVerifyDateAutoStatus"></i></button>				
				<% End If %>
                <%
				If IsDate(LastVerifiedDate) Then
						
						Response.Write("<br><strong  id=""txtlastverfiydaterange"">")
						diffdays =  DateDiff("d",LastVerifiedDate,Date)
						If diffdays=0 Then
							Response.Write("today")
						ElseIf 	diffdays = 1 Then
							Response.Write(diffdays & " day ago")
						Else
							Response.Write(diffdays & " days ago")
						End If
						Response.Write("</strong>")
					Else
						'Response.Write("never")
					End If
				%>
                    </div>
                </div>
            </div>
             <div class="quick-info-block CRMTileOwnerColor">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-user"></i>&nbsp;<%= GetTerm("Owner") %>
 					<% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
						<!-- User Has READONLY Access -->
					<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
						<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-owner-no="<%= OwnerUserNo %>" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditOwner" data-tooltip="true" data-title="Edit Prospect Owner"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					
					<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
						<!-- User Has WRITEOWNED Access and Is Not The Owner -->				
					<% ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then %>
						<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-owner-no="<%= OwnerUserNo %>" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditOwner" data-tooltip="true" data-title="Edit Prospect Owner"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					
					<% End If %>                	
                </h2>
                <hr class="tile">
                <p><% If OwnerUserNo <> 0 Then Response.Write(GetUserDisplayNameByUserNo(OwnerUserNo)) %></p>                        
            </div>

			<a class="btn btn-primary btn-lg btn-block" href="main.asp" role="button" style="margin-top:15px;"><i class="fa fa-arrow-left"></i> &nbsp;Back To <%= GetTerm("Prospect") %> List</a>
			
		</div>

		<div class="col-lg-3">
 
            
             <div class="quick-info-block CRMTileCommentsColor">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-comment"></i>&nbsp;<%= GetTerm("Comments") %>
 					<% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
						<!-- User Has READONLY Access -->
					<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
						<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditComments" data-tooltip="true" data-title="Edit Prospect Comments"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					
					<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
						<!-- User Has WRITEOWNED Access and Is Not The Owner -->				
					<% ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then %>
						<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditComments" data-tooltip="true" data-title="Edit Prospect Comments"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					
					<% End If  %>                    
                </h2>
                <hr class="tile"> 
                
                <% If Comments = "" Then %>
                	<p>(Not Entered)</p> 
                <% Else %>
                	<p><%= Comments %></p>
				<% End If %>  
				
                    
            </div>

            <div class="quick-info-block CRMTileDollarsColor">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-usd"></i>&nbsp;<%= GetTerm("Opportunity") %>
 					<% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
						<!-- User Has READONLY Access -->
					<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
						<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditOpportunity" data-tooltip="true" data-title="Edit Prospect <%= GetTerm("Opportunity") %>"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					
					<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
						<!-- User Has WRITEOWNED Access and Is Not The Owner -->				
					<% ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then %>
						<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditOpportunity" data-tooltip="true" data-title="Edit Prospect <%= GetTerm("Opportunity") %>"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					
					<% End If  %>                  	
                </h2>
                <hr class="tile">
                <% If ProjectedGPSpend = 0 Then %>
                	<p>Projected GP Spend (Not Entered)</p> 
                <% Else %>
                	<p>Projected GP Spend <%= FormatCurrency(ProjectedGPSpend,2) %></p>
				<% End If %>  
				
                <% If NumEmployees = "0" OR NumEmployees = "" Then %>
                	<p># Employees (Not Entered)</p> 
                <% Else %>
                	<p># Employees <%= NumEmployees %></p>
				<% End If %>    
                
                <p># Pantries <%= NumberOfPantries %></p>    
                 
                <% If IsNull(LeaseExpirationDate) OR LeaseExpirationDate="1/1/1900" Then %>
                	<p>Lease Expiration Date (Not Entered)</p> 
                <% Else %>
                	<p>Lease Expiration Date <%= LeaseExpirationDate %></p>
				<% End If %>    
 
                <% If IsNull(ContractExpirationDate) OR ContractExpirationDate="1/1/1900" Then %>
                	<p>Contract Expiration Date (Not Entered)</p> 
                <% Else %>
                	<p>Contract Expiration Date <%= ContractExpirationDate %></p>
				<% End If %>     
				        
            </div>

		</div>
            

		<div class="col-lg-3">

            <div class="quick-info-block CRMTileOfferingColor">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-clock-o"></i>&nbsp;<%= GetTerm("Current Supplier Info") %>
					<% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
						<!-- User Has READONLY Access -->
					<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
						<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditCurrentSupplierInfo" data-tooltip="true" data-title="Edit Prospect <%= GetTerm("Current Supplier Info") %>"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					
					<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
						<!-- User Has WRITEOWNED Access and Is Not The Owner -->				
					<% ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then %>
						<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditCurrentSupplierInfo" data-tooltip="true" data-title="Edit Prospect <%= GetTerm("Current Supplier Info") %>"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					
					<% End If  %>                
                </h2>

                <hr class="tile">

                <% If CurrentOffering = "" Then %>
                	<p>(Not Entered)</p> 
                <% Else %>
                	<p><%= CurrentOffering %></p>
				<% End If %>  
                                        
            </div>

				<%
					PrimaryCompetitorID = GetPrimaryCompetitorIDByProspectNumber(InternalRecordIdentifier)
					
					If PrimaryCompetitorID <> "" Then
						PrimaryCompetitorName = GetCompetitorByNum(PrimaryCompetitorID)
					
						SQLCompetitors1 = "SELECT * FROM PR_ProspectCompetitors WHERE CompetitorRecID = " & PrimaryCompetitorID & " AND ProspectRecID = " &  InternalRecordIdentifier
						
						Set cnnCompetitors1 = Server.CreateObject("ADODB.Connection")
						cnnCompetitors1.open (Session("ClientCnnString"))
						Set rsCompetitors1 = Server.CreateObject("ADODB.Recordset")
						rsCompetitors1.CursorLocation = 3 
						Set rsCompetitors1 = cnnCompetitors1.Execute(SQLCompetitors1)
						
						If not rsCompetitors1.EOF Then
						
							BottledWater = rsCompetitors1 ("BottledWater")
							FilteredWater = rsCompetitors1 ("FilteredWater")
							OCS = rsCompetitors1 ("OCS")
							OCS_Supply = rsCompetitors1 ("OCS_Supply")
							OfficeSupplies = rsCompetitors1 ("OfficeSupplies")
							Vending = rsCompetitors1 ("Vending")
							Micromarket = rsCompetitors1 ("Micromarket")
							Pantry = rsCompetitors1 ("Pantry")
											
						End If
						Set rsCompetitors1 = Nothing
						cnnCompetitors1.Close
						Set cnnCompetitors1 = Nothing
						
						
						If BottledWater = vbTrue Then BottledWater = "Bottled Water" Else BottledWater = ""
						If FilteredWater = vbTrue Then FilteredWater = "Filtered Water" Else FilteredWater = ""
						If OCS = vbTrue Then OCS = "OCS" Else OCS = ""
						If OCS_Supply = vbTrue Then OCS_Supply = "OCS Supply" Else OCS_Supply = ""
						If OfficeSupplies = vbTrue Then OfficeSupplies = "Office Supplies " Else OfficeSupplies = ""
						If Vending = vbTrue Then Vending = "Vending" Else Vending = ""
						If Micromarket = vbTrue Then Micromarket = "Micromarkets" Else Micromarket = ""
						If Pantry = vbTrue Then Pantry = "Pantry" Else Pantry = ""
					Else

						PrimaryCompetitorName = ""
						BottledWater = ""
						FilteredWater = ""
						OCS = ""
						OCS_Supply = ""
						OfficeSupplies = ""
						Vending = ""
						Micromarket = ""
						Pantry = ""
					
					End If
						
					%>
			
		            <div class="quick-info-block CRMTileCompetitorColor">
		                <h2 class="heading-md"><i class="icon-2x color-light fa fa-user-circle-o"></i>&nbsp;<%= GetTerm("Primary Competitor") %>
							<% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
								<!-- User Has READONLY Access -->
							<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
								<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditCompetitorSource" data-tooltip="true" data-title="Edit Prospect <%= GetTerm("Primary Competitor") %>"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
								<!-- User Has WRITEOWNED Access and Is Not The Owner -->				
							<% ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then %>
								<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditCompetitorSource" data-tooltip="true" data-title="Edit Prospect <%= GetTerm("Primary Competitor") %>"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					
							<% End If  %>    
                		</h2>
		                <hr class="tile">
		                

		                <% If Telemarketer = "0" OR TelemarketerUserNo = 0 Then %>
		                	<p>Telemarketer (Not Entered)</p> 
		                <% Else %>
		                	<p>Telemarketer: (<%= Telemarketer %>)</p>
						<% End If %>  
		                
		                <% If LeadSource = "0" OR LeadSourceNumber = 0 Then %>
		                	<p>Lead Source (Not Entered)</p> 
		                <% Else %>
		                	<p>Lead Source (<%= LeadSource %>)</p>
						<% End If %>  
		                
		                <% If PrimaryCompetitorName <> "" Then %>
		                	<p>Primary Competitor: <%= PrimaryCompetitorName %></p>
		                <% Else %>
		                	<p>Primary Competitor (Not Entered)</p>
		                <% End If %>
		                
		                <% If BottledWater <> "" AND FilteredWater <> "" AND OCS <> "" AND OCS_Supply <> "" AND OfficeSupplies <> "" AND Vending <> "" AND Micromarket <> "" AND Pantry <> "" Then %>
		                	<p>Primary Competitor Offerings: 
		                	<% If BottledWater <> "" Then Response.Write(BottledWater) %>
		                	
		                	<% If BottledWater <> "" Then %>
		                		<% If FilteredWater <> "" Then Response.Write(", " & FilteredWater) %>
		                	<% Else %>
		                		<% If FilteredWater <> "" Then Response.Write(FilteredWater) %>
							<% End If %>
							
							<% If FilteredWater <> "" Then %>
		                		<% If OCS <> "" Then Response.Write(", " & OCS) %>
		                	<% Else %>
		                		<% If OCS <> "" Then Response.Write(OCS) %>
		                	<% End If %>
		                	
		                	<% If OCS <> "" Then %>
		                		<% If OCS_Supply <> "" Then Response.Write(", " & OCS_Supply) %>
		                	<% Else %>	
		                		<% If OCS_Supply <> "" Then Response.Write(OCS_Supply) %>
		                	<% End If %>	
		                	
		                	<% If OCS_Supply <> "" Then %>
		                		<% If OfficeSupplies <> "" Then Response.Write(", " & OfficeSupplies) %>
		                	<% Else %>
		                		<% If OfficeSupplies <> "" Then Response.Write(OfficeSupplies) %>
		                	<% End If %>
		                	
		                	<% If OfficeSupplies <> "" Then %>
		                		<% If Vending <> "" Then Response.Write(", " & Vending) %>
		                	<% Else %>
		                		<% If Vending <> "" Then Response.Write(Vending) %>
		                	<%  End If %>
		                	
		                	<% If Vending <> "" Then %>
		                		<% If Micromarket <> "" Then Response.Write(", " & Micromarket) %>
		                	<% Else %>	
		                		<% If Micromarket <> "" Then Response.Write(Micromarket) %>
							<%  End If %>
							
							<% If Micromarket <> "" Then %>
		                		<% If Pantry <> "" Then Response.Write(", " & Pantry) %>
		                	<% Else %>
		                		<% If Pantry <> "" Then Response.Write(Pantry) %>
		                	<% End If %>
		                </p>
		                <% Else %>
		                	<p>Primary Competitor Offerings (Not Entered)</p>
		                <% End If %>
								                
		                  
		                <hr class="tile">
		                
		                <% If FormerCustNum = "0" OR FormerCustNum = "" Then %>
		                	<p>Former Customer # (Not Entered)</p> 
		                <% Else %>
		                	<p>Former Customer #: (<%= FormerCustNum %>)</p>
						<% End If %>  
		                
		                <% If IsNull(CancelDate) OR CancelDate ="1/1/1900" Then %>
		                	<p>Cancel Date (Not Entered)</p> 
		                <% Else %>
		                	<p>Cancel Date: <%= CancelDate %></p>
						<% End If %>     
		                   
		            </div>
		</div>

		<div class="col-lg-3">
	
			<%
		
			SQLContacts1 = "SELECT * FROM PR_ProspectActivities where ProspectRecID = " & InternalRecordIdentifier & " AND Status IS NULL"
			
			Set cnnContacts1 = Server.CreateObject("ADODB.Connection")
			cnnContacts1.open (Session("ClientCnnString"))
			Set rsContacts1 = Server.CreateObject("ADODB.Recordset")
			rsContacts1.CursorLocation = 3 
			Set rsContacts1 = cnnContacts1.Execute(SQLContacts1)
			
			If not rsContacts1.EOF Then
				ActivityRecID = rsContacts1("ActivityRecID")
			  	nextActivity = GetActivityByNum(rsContacts1("ActivityRecID"))
				nextActivityDueDate = FormatDateTime(rsContacts1("ActivityDueDate"),2) & " " & FormatDateTime(rsContacts1("ActivityDueDate"),3)
				daysOld = DateDiff("d",rsContacts1("RecordCreationDateTime"),Now())
				daysOverdue = DateDiff("d",rsContacts1("ActivityDueDate"),Now())	
							
			End If
			Set rsContacts1 = Nothing
			cnnContacts1.Close
			Set cnnContacts1 = Nothing
				
			%>
			
			<% If ActivityRecID <> "" Then %>
			
		            <div class="quick-info-block CRMTileActivityColor">
		                <h2 class="heading-md"><i class="icon-2x color-light fa fa-arrow-right"></i>&nbsp;<%= GetTerm("Next Activity") %> 
							<% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
								<!-- User Has READONLY Access -->
							<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
								<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-activity-id="<%= ActivityRecID %>" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditActivity" data-tooltip="true" data-title="Edit Prospect Activity"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>
							<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
								<!-- User Has WRITEOWNED Access and Is Not The Owner -->				
							<% ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then %>
								<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-activity-id="<%= ActivityRecID %>" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditActivity" data-tooltip="true" data-title="Edit Prospect Activity"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					
							<% End If  %>
		                </h2> 
		                <hr class="tile">
		                <p><%= nextActivity %></p>
		                <hr class="tile">
		                <p><i class="fa fa-calendar-check-o" aria-hidden="true"></i>&nbsp;<strong>Due Date:</strong>&nbsp;&nbsp;<span class="nextActivityDate"><%= formatDateTime(nextActivityDueDate,2) %></span>&nbsp;&nbsp;<span class="nextActivityTime"><%= formatDateTime(nextActivityDueDate,3) %></span></p> 
		                <p><i class="fa fa-calendar-times-o" aria-hidden="true"></i>&nbsp;<strong>Days Since Created:</strong>&nbsp;&nbsp;<%= daysOld %></p>
		                <% If daysOverdue > 0 Then %>
		                	<p class="overdue"><i class="fa fa-times-circle-o" aria-hidden="true"></i>&nbsp;<strong>Days Overdue:</strong>&nbsp;&nbsp;<%= daysOverdue %></p> 
		                <% End If %>    
		            </div>
			<% Else %>
		
				    <div class="quick-info-block CRMTileActivityColor">
		                <h2 class="heading-md"><i class="icon-2x color-light fa fa-arrow-right"></i>&nbsp;<%= GetTerm("Next Activity") %></h2> 
		                <hr class="tile">
		                <p>No Next Activity</p>
		                <hr class="tile">
		                <p><i class="fa fa-calendar-check-o" aria-hidden="true"></i>&nbsp;<strong>Due Date:</strong>&nbsp;&nbsp;<span class="nextActivityDate">NA</span></p> 
		                <p><i class="fa fa-calendar-times-o" aria-hidden="true"></i>&nbsp;<strong>Days Since Created:</strong>&nbsp;&nbsp;NA</p>
		                <% If daysOverdue > 0 Then %>
		                	<p class="overdue"><i class="fa fa-times-circle-o" aria-hidden="true"></i>&nbsp;<strong>Days Overdue:</strong>&nbsp;&nbsp;<%= daysOverdue %></p> 
		                <% End If %>    
		            </div>
			<% End If %>
			
            <div class="quick-info-block CRMTileStageColor">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-tasks"></i>&nbsp;<%= GetTerm("Stage") %> 
					<% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
						<!-- User Has READONLY Access -->
					<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
						<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-stage-id="<%= GetProspectCurrentStageByProspectNumber(InternalRecordIdentifier) %>" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditStage" data-tooltip="true" data-title="Edit Prospect Stage"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>						<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
						<!-- User Has WRITEOWNED Access and Is Not The Owner -->				
					<% ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then %>
						<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-stage-id="<%= GetProspectCurrentStageByProspectNumber(InternalRecordIdentifier) %>" data-prospect-id="<%= InternalRecordIdentifier %>" data-target="#myProspectingModalEditStage" data-tooltip="true" data-title="Edit Prospect Stage"><button class="btn btn-success" role="button" type="button"><i class="fa fa-pencil-square fa-lg" aria-hidden="true"></i></button></a>					
					<% End If  %>             
                </h2> 
                <hr class="tile">
                <p><%= GetStageByNum(StageNumber) %></p>
                <hr class="tile">
                <p><i class="fa fa-pencil-square-o" aria-hidden="true"></i>&nbsp;Last Change Date:&nbsp;&nbsp;<%= GetProspectLastStageChangeDateByProspectNumber(InternalRecordIdentifier) %></p>
                <p><i class="fa fa-calendar-o" aria-hidden="true"></i>&nbsp;Days Since Qualified:&nbsp;&nbsp;XYZ</p>
                <div class="progressbarsone" progress="<%= GetPercentForStage(StageNumber)%>%"></div>       
            </div>
            
		
			
			<!--<a class="btn btn-danger btn-lg btn-block" href="#" role="button">Create Note</a>-->
            

		</div>
		
		
	</div>
</div>
<!-- eof title / lead owner !-->

		 
<!-- tabs start here !-->
<div class="bottom-table">
	<div class="row">
		<div class="col-lg-12">
			<div class="bottom-tabs-section">

				<!-- tab navigation !-->
				<ul class="nav nav-tabs" role="tablist">
					<li role='presentation' class="active"><a href='#log' class='CRMTabLogColor' aria-controls='notes' role='tab' data-toggle='tab'><%= GetTerm("Journal") %> (<%=NumberOfLogItemsByProspectNumber(InternalRecordIdentifier)%>)</a></li>
					<% If CRMHideProductsTab = 0 Then %>
						<li role='presentation'><a href='#products' class='CRMTabProductsColor' aria-controls='documents' role='tab' data-toggle='tab'><%= GetTerm("Products") %></a></li>
					<% End If %>
					<% If CRMHideEquipmentTab = 0 Then %>
						<li role='presentation'><a href='#equipment' class='CRMTabEquipmentColor' aria-controls='documents' role='tab' data-toggle='tab'><%= GetTerm("Equipment") %></a></li>
					<% End If %>
					<li role='presentation'><a href='#documents' class='CRMTabDocumentsColor' aria-controls='documents' role='tab' data-toggle='tab'><%= GetTerm("Documents") %> (<%=NumberOfDocumentsByProspectNumber(InternalRecordIdentifier)%>)</a></li>  
					<li role='presentation'><a href='#contacts' class='CRMTabContactsColor' aria-controls='contacts' role='tab' data-toggle='tab'><%= GetTerm("Contacts") %> (<%=NumberOfContactsByProspectNumber(InternalRecordIdentifier)%>)</a></li>
					<li role='presentation'><a href='#competitors' class='CRMTabCompetitorsColor' aria-controls='general' role='tab' data-toggle='tab'><%= GetTerm("Competitors") %> (<%=NumberOfCompetitorsByProspectNumber(InternalRecordIdentifier)%>)</a></li>
					<% If CRMHideLocationTab = 0 Then %>
						<li role='presentation'><a href='#location' class='CRMTabLocationColor' aria-controls='general' role='tab' data-toggle='tab'><%= GetTerm("Location") %></a></li>
					<% End If %>
                    <li role='presentation'><a href='#socialmedia' class='CRMTabSocialMediaColor' aria-controls='audit' role='tab' data-toggle='tab'><%= GetTerm("Social Media") %> (<%=NumberOfSocialMediaByProspectNumber(InternalRecordIdentifier)%>)</a></li>
					<li role='presentation'><a href='#audit' class='CRMTabAuditTrailColor' aria-controls='audit' role='tab' data-toggle='tab'><%= GetTerm("Audit Trail") %></a></li>
					<!--<li role='presentation'><a href="main.asp" class="btn btn-secondary active" role="button" aria-pressed="true"><span class="return">Back To <%= GetTerm("Prospecting") %> Main</span></a></li>-->
				</ul>
				<!-- eof tab navigation -->
			
			
				<% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
					<!-- User Has READONLY Access -->
					<div class="tab-content">
						<!--#include file="viewProspectReadOnly_log_tab.asp"-->
						<% If CRMHideProductsTab = 0 Then %>
							<!--#include file="viewProspectReadOnly_products_tab.asp"-->
						<% End If %>
						<% If CRMHideEquipmentTab = 0 Then %>
							<!--#include file="viewProspectReadOnly_equipment_tab.asp"-->
						<% End If %>
						<!--#include file="viewProspectReadOnly_documents_tab.asp"-->
						<!--#include file="viewProspectReadOnly_contacts_tab.asp"-->
						<!--#include file="viewProspectReadOnly_competitors_tab.asp"-->
						<% If CRMHideLocationTab = 0 Then %>
							<!--#include file="viewProspectReadOnly_location_tab.asp"-->
						<% End If %>
                        <!--#include file="viewProspectReadOnly_socialmedia_tab.asp"-->
						<!--#include file="viewProspectReadOnly_audit_tab.asp"-->
					</div>							
				<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
					<div class="tab-content">
						<!--#include file="viewProspect_log_tab.asp"-->
						<% If CRMHideProductsTab = 0 Then %>
							<!--#include file="viewProspect_products_tab.asp"-->
						<% End If %>
						<% If CRMHideEquipmentTab = 0 Then %>
							<!--#include file="viewProspect_equipment_tab.asp"-->
						<% End If %>
						<!--#include file="viewProspect_documents_tab.asp"-->
						<!--#include file="viewProspect_contacts_tab.asp"-->
						<!--#include file="viewProspect_competitors_tab.asp"-->
						<% If CRMHideLocationTab = 0 Then %>
							<!--#include file="viewProspect_location_tab.asp"-->
						<% End If %>
                        
                        <!--#include file="viewProspect_socialmedia_tab.asp"-->
						<!--#include file="viewProspect_audit_tab.asp"-->
					</div>	
				<% ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
					<!-- User Has WRITEOWNED Access and Is Not The Owner -->
					<div class="tab-content">
						<!--#include file="viewProspectReadOnly_log_tab.asp"-->
						<% If CRMHideProductsTab = 0 Then %>
							<!--#include file="viewProspectReadOnly_products_tab.asp"-->
						<% End If %>
						<% If CRMHideEquipmentTab = 0 Then %>
							<!--#include file="viewProspectReadOnly_equipment_tab.asp"-->
						<% End If %>
						<!--#include file="viewProspectReadOnly_documents_tab.asp"-->
						<!--#include file="viewProspectReadOnly_contacts_tab.asp"-->
						<!--#include file="viewProspectReadOnly_competitors_tab.asp"-->
						<% If CRMHideLocationTab = 0 Then %>
							<!--#include file="viewProspectReadOnly_location_tab.asp"-->
						<% End If %>
                        <!--#include file="viewProspectReadOnly_socialmedia_tab.asp"-->
						<!--#include file="viewProspectReadOnly_audit_tab.asp"-->
					</div>											
				<% ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then %>
					<div class="tab-content">
						<!--#include file="viewProspect_log_tab.asp"-->
						<% If CRMHideProductsTab = 0 Then %>
							<!--#include file="viewProspect_products_tab.asp"-->
						<% End If %>
						<% If CRMHideEquipmentTab = 0 Then %>
							<!--#include file="viewProspect_equipment_tab.asp"-->
						<% End If %>
						<!--#include file="viewProspect_documents_tab.asp"-->
						<!--#include file="viewProspect_contacts_tab.asp"-->
						<!--#include file="viewProspect_competitors_tab.asp"-->
						<% If CRMHideLocationTab = 0 Then %>
							<!--#include file="viewProspect_location_tab.asp"-->
						<% End If %>
                        <!--#include file="viewProspect_socialmedia_tab.asp"-->
						<!--#include file="viewProspect_audit_tab.asp"-->
					</div>											
				<% End If  %>

				
			</div>
		</div>
	</div>
</div>

<%
set rs = Nothing
cnn8.close
set cnn8 = Nothing
%>

 <!-- tabs js  !-->
 <script type="text/javascript">
 $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
  e.target // newly activated tab
  e.relatedTarget // previous active tab
})
   </script>

   <script>
$(document).ready(function(){
  $("#demo").on("hide.bs.collapse", function(){
    $(".btn-custom").html('<span class="glyphicon glyphicon-collapse-down"></span> Click to Expand');
  });
  $("#demo").on("show.bs.collapse", function(){
    $(".btn-custom").html('<span class="glyphicon glyphicon-collapse-up"></span> Click to Collapse');
  });
});
</script>
 <!-- eof tabs js !-->

 <!-- custom table search !-->

<script>

$(document).ready(function () {

    (function ($) {

        $('#filter-audit').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-audit tr').hide();
            $('.searchable-audit tr').filter(function () {
                return rex.test($(this).text());
            }).show();
        })
		

        
        $('#filter-notes').keyup(function () {
			var val = $("input[name='logtyperadiofilter']:checked"). val();
			
			var regstr = '';
			if (val==0){
				regstr = '';
			} else if (val==1){
				regstr = 'Stage Change';
			} else if (val==2){
				regstr = 'Email';
			} else if (val==3){
				regstr = 'Note';
			} else if (val==4){
				regstr = 'Activity';
			}	
			
			
            var rex2 = new RegExp(regstr, '');
			/*
            $('.searchable-notes tr').hide();
            $('.searchable-notes tr').filter(function () {
                return rex2.test($(this).text());
            }).show();	
			*/		

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-notes tr').hide();
            $('.searchable-notes tr').filter(function () {
                return rex.test($(this).text()) && rex2.test($(this).text());
            }).show();
        })
        
		$("input[name='logtyperadiofilter']").click(function () {
			var val = $("input[name='logtyperadiofilter']:checked"). val();
			
			var regstr = '';
			if (val==0){
				regstr = '';
			} else if (val==1){
				regstr = 'Stage Change';
			} else if (val==2){
				regstr = 'Email';
			} else if (val==3){
				regstr = 'Note';
			} else if (val==4){
				regstr = 'Activity';
			}	
			
			
            var rex2 = new RegExp(regstr, '');
			/*
            $('.searchable-notes tr').hide();
            $('.searchable-notes tr').filter(function () {
                return rex.test($(this).text());
            }).show();
			*/
			
			var rex = new RegExp($('#filter-notes').val(), 'i');
            $('.searchable-notes tr').hide();
            $('.searchable-notes tr').filter(function () {
                return rex.test($(this).text()) && rex2.test($(this).text());
            }).show();
			
        })
		
        $('#filter-documents').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-documents tr').hide();
            $('.searchable-documents tr').filter(function () {
                return rex.test($(this).text());
            }).show();
        })

        $('#filter-contacts').keyup(function () {

           var rex = new RegExp($(this).val(), 'i');
           $('.searchable-contacts tr').hide();		   
           $('.searchable-contacts tr').filter(function () {
               return rex.test($(this).text());
            }).show();
			
        })
 
        $('#filter-competitors').keyup(function () {

           var rex = new RegExp($(this).val(), 'i');
           $('.searchable-competitors tr').hide();
           $('.searchable-competitors tr').filter(function () {
               return rex.test($(this).text());
            }).show();
        })
        
       
        $('#filter-opportunity').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-opportunity tr').hide();
            $('.searchable-opportunity tr').filter(function () {
                return rex.test($(this).text());
            }).show();
        })
        

    }(jQuery));

});
</script>
<!-- eof custom table search !-->

<!-- progress bar !-->
<link rel="stylesheet" href="<%= BaseURL %>js/jprogress/jprogress.css">
<script src="<%= BaseURL %>js/jprogress/jprogress.js" type="text/javascript"></script>

<script>
    // activate jprogress
    $(".progressbars").jprogress();
    $(".progressbarsone").jprogress({
        background: "url(../js/jprogress/progress_bar_tiles.png)"
     });
</script>
<!-- eof progress bar !-->


<!-- checkboxes JS !-->
<script type="text/javascript">
    function changeState(el) {
        if (el.readOnly) el.checked=el.readOnly=false;
        else if (!el.checked) el.readOnly=el.indeterminate=true;
    }
</script>
<!-- eof checkboxes JS !-->

 
 
 
<!-- ******************************************************************************************************************************** -->
<!-- MODAL WINDOW DESIGN AND DEFINITIONS -->
<!-- ******************************************************************************************************************************** -->

	<!--#include file="editProspectModals.asp"-->
    
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

<!-- modal window next activity -->
<!--#include file="onthefly_nextactivity.asp"--> 
<!-- end modal window nextactivity -->    

<!-- modal window update last verify date -->
<!--#include file="onthefly_lastverifieddate.asp"--> 
<!-- end modal window update last verify date -->   

<!-- modal window stage -->
<!--#include file="onthefly_addstage.asp"--> 
<!-- end modal window stage -->   

<!-- modal window contacts tab -->
<!--#include file="onthefly_contacttitle_forcontactstab.asp"--> 
<!-- end modal contacts tab -->   

<!-- ******************************************************************************************************************************** -->
<!-- END MODAL WINDOW DESIGN AND DEFINITIONS -->
<!-- ******************************************************************************************************************************** -->



<!--#include file="../inc/footer-main.asp"-->