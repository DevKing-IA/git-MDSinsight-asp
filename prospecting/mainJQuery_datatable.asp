
<script type="text/javascript">
		
	$(document).ready(function() {
	


    	$("[data-toggle=tooltip]").tooltip();	

       var maxDaysPermitted = '<%= MaxActivityDaysPermittedInit %>';
       maxDaysPermitted = parseInt(maxDaysPermitted);

		//In Global Settings, Zero Means User Can Pick Any Date In The Future
		if (maxDaysPermitted != 0) {
	        $('#datetimepicker1').datetimepicker({
	        	useCurrent: false,
	        	minDate:moment(),
	        	maxDate:moment().add('<%= MaxActivityDaysPermittedInit %>', 'days')
			});    
		}
		else {
	        $('#datetimepicker1').datetimepicker({
	        	useCurrent: false,
	        	minDate:moment()
			});      

		}  

        $("#datetimepicker1").on("dp.change", function(e) {

           var maxDaysWarning = '<%= MaxActivityDaysWarningInit %>';
           maxDaysWarning = parseInt(maxDaysWarning);
           
           
           //In Global Settings, Zero Means to Not Show Warning
           if (maxDaysWarning != 0) {
				   var selectedDateFromPicker = moment($("#datetimepicker1").find("input").val());
		
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
		
		$('#myProspectingModalEditActivity').on('show.bs.modal', function(e) {
		
		
		    //get data-id attribute of the clicked prospect
		    var myProspectID = $(e.relatedTarget).data('prospect-id');	
		    var myActivityRecID = $(e.relatedTarget).data('activity-id');
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtInternalRecordIdentifier"]').val(myProspectID);
		    $(e.currentTarget).find('input[name="txtActivityRecID"]').val(myActivityRecID);
		    	    
		    var $modal = $(this);
		
			$("#activityDateWarning").hide();
			
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetInitialActivityAppmtOrMeeting",
				success: function(response)
				 {
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
		    $(e.currentTarget).find('input[name="txtInternalRecordIdentifier"]').val(myProspectID);
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
		    
		});
		

	
		
		$('#myProspectingDeleteModal').on('show.bs.modal', function(e) {

	    	var $modal = $(this);
			var chkBoxArray = [];
			$(".checksingle:checked").each(function() {
			    chkBoxArray.push(this.id);
			    //alert(this.id);
			});			
	    	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetProspectDeleteInformationForModal&prospectsArray="+encodeURIComponent(chkBoxArray),
				success: function(response)
				 {
	               	 $modal.find('#deleteProspectInfo').html(response);	               	 
	             },
	             failure: function(response)
				 {
				   $modal.find('#deleteProspectInfo').html("Failed");
	             }
			});	    
		});
		
		$('#myProspectingAddMultipleNotesModal').on('show.bs.modal', function(e) {

	    	var $modal = $(this);
			var chkBoxArray = [];
			$(".checksingle:checked").each(function() {
			    chkBoxArray.push(this.id);
			});			
	    	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetProspectAddNotesInformationForModal&prospectsArray="+encodeURIComponent(chkBoxArray),
				success: function(response)
				 {
	               	 $modal.find('#addnotesProspectInfo').html(response);	               	 
	             },
	             failure: function(response)
				 {
				   $modal.find('#addnotesProspectInfo').html("Failed");
	             }
			});	    
						
		});

		$('#myProspectingExportModal').on('show.bs.modal', function(e) {

	    	var $modal = $(this);
			var chkBoxArray = [];
			$(".checksingle:checked").each(function() {
			    chkBoxArray.push(this.id);
			    //alert(this.id);
			});			
	    	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
				cache: false,
				data: "action=GetProspectExportInformationForModal&prospectsArray="+encodeURIComponent(chkBoxArray),
				success: function(response)
				 {
	               	 $modal.find('#exportProspectInfo').html(response);	               	 
	             },
	             failure: function(response)
				 {
				   $modal.find('#exportProspectInfo').html("Failed");
	             }
			});	    
		});

 //-------------------------------------------------------------------------------
	
	//Add Notes modal window submit
	$('#frmAddNotesToProspects').submit(function(e) {
		

			var chkBoxArray = [];
			$(".checksingle:checked").each(function() {
			    chkBoxArray.push(this.id);
			});				
		
		if ($('#frmAddNotesToProspects #txtProspectingNote').val()==''){
			 swal("Note can not be blank.");
			return false;
		}
		
		$('#frmAddNotesToProspects #addnotesmultipleids').val(chkBoxArray);						
		
		$("#myProspectingAddMultipleNotesModal .btn-danger").html("Saving...");
		
        $.ajax({
            type: "POST",
            url: "onthefly_addnotes_submit.asp",
            data: $('#frmAddNotesToProspects').serialize(),
            success: function(response) {		
				$("#myProspectingAddMultipleNotesModal .modal-footer").hide();			
				$("#myProspectingAddMultipleNotesModal .modal-body").html('Note added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');												
            },
            error: function() {
				$("#myProspectingAddMultipleNotesModal .btn-danger").html("Add Note To Prospect(s)");
            }
        });
        return false;
		
    });
//-------------------------------------------------------------------------------	

			
	    $("#PleaseWaitPanel").hide();
		$("#deletedSelectedProspects").hide();
		$("#exportProspects").hide();
		
		var selectedFilterVal = $("#selectFilteredView option:selected").val();
		
		if (selectedFilterVal == "Default" || selectedFilterVal == "All") {
			
			$("#btnRenameProspectFilterView").hide();
			$("#btnDeleteProspectFilterView").hide();			
			$("#btnRenameProspectFilterViewRecyclePool").hide();
			$("#btnDeleteProspectFilterViewRecyclePool").hide();
		}
		else if (selectedFilterVal == "") {
			
			$("#btnRenameProspectFilterView").hide();
			$("#btnDeleteProspectFilterView").hide();			
			$("#btnRenameProspectFilterViewRecyclePool").hide();
			$("#btnDeleteProspectFilterViewRecyclePool").hide();
		}


		$(".search").keyup(function () {
		
			var searchTerm = $(".search").val();
			var listItem = $('.results tbody').children('tr');
			var searchSplit = searchTerm.replace(/ /g, "'):containsi('")
			
			$.extend($.expr[':'], {'containsi': function(elem, i, match, array){
			    return (elem.textContent || elem.innerText || '').toLowerCase().indexOf((match[3] || "").toLowerCase()) >= 0;
				}
			});
			
			$(".results tbody tr").not(":containsi('" + searchSplit + "')").each(function(e){
				$(this).attr('visible','false');
			});
			
			$(".results tbody tr:containsi('" + searchSplit + "')").each(function(e){
				$(this).attr('visible','true');
			});
			
			var jobCount = $('.results tbody tr[visible="true"]').length;
			$('.counter').html('<strong>'+ jobCount + '</strong> prospects found containing <strong>' + searchTerm + '</strong>.');
			
			if(jobCount == '0') {
				$('.no-result').show();
			}
			else {
				$('.no-result').hide();
			}
			
			});			   
		
		});
	
    
	
	function saveAsNewProspectFilterView() {
	
	   var viewNameInputField = $("#txtNewFilterReportViewName").val();
	   var viewNameSelectBox = $("#selExistingFilterViewNames option:selected").val();
		
		$.ajax({		
			type:"POST",
			url: "createProspectFilterViewFromModalSaveAs.asp?viewNameInputField="+encodeURIComponent(viewNameInputField)+"&viewNameSelectBox="+encodeURIComponent(viewNameSelectBox),
			complete: function (data) {
				window.location.href = "main.asp";
			}
		})	
		
	}	
	
	function saveAsNewProspectFilterViewRecyclePool() {
	
	   var viewNameInputField = $("#txtNewFilterReportViewName").val();
	   var viewNameSelectBox = $("#selExistingFilterViewNames option:selected").val();
		
		$.ajax({		
			type:"POST",
			url: "createRecyclePoolProspectFilterViewFromModalSaveAs.asp?viewNameInputField="+encodeURIComponent(viewNameInputField)+"&viewNameSelectBox="+encodeURIComponent(viewNameSelectBox),
			complete: function (data) {
				window.location.href = "mainRecyclePool.asp";
			}
		})	
		
	}	

		
	function saveAsNewProspectFilterViewWonPool() {
	
	   var viewNameInputField = $("#txtNewFilterReportViewName").val();
	   var viewNameSelectBox = $("#selExistingFilterViewNames option:selected").val();
		
		$.ajax({		
			type:"POST",
			url: "createWonPoolProspectFilterViewFromModalSaveAs.asp?viewNameInputField="+encodeURIComponent(viewNameInputField)+"&viewNameSelectBox="+encodeURIComponent(viewNameSelectBox),
			complete: function (data) {
				window.location.href = "mainWonPool.asp";
			}
		})	
		
	}	


	function renameProspectFilterView() {
	
	   var newViewName = $("#txtUpdatedFilterReportViewName").val();
	   var originalViewName = $("#originalViewName").val();

		$.ajax({
			type:"POST",
			url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
			cache: false,
			data: "action=CheckIfViewNameExists&newViewName=" + encodeURIComponent(newViewName),
			success: function(response)
			{
				if (response == "False")
				{					
		       		$.ajax({
						type:"POST",
						url: "renameProspectFilterViewFromModal.asp?newViewName="+encodeURIComponent(newViewName)+"&originalViewName="+encodeURIComponent(originalViewName),
						complete: function (data)
						{
							window.location.href = "main.asp";
			            }
					})
							
				}
				else {
				    swal("The view name you have entered already exists. Please enter a different name.");
           			return false;
           		}

	        }
		})
	}	
	



	
	function renameProspectFilterViewRecyclePool() {
	
	   var newViewName = $("#txtUpdatedFilterReportViewNameRecPool").val();
	   var originalViewName = $("#originalViewNameRecPool").val();
	   
		$.ajax({
			type:"POST",
			url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
			cache: false,
			data: "action=CheckIfViewNameExistsRecyclePool&newViewName=" + encodeURIComponent(newViewName),
			success: function(response)
			{				
				if (response == "False")
				{					
		       		$.ajax({
						type:"POST",
						url: "renameRecyclePoolProspectFilterViewFromModal.asp?newViewName="+encodeURIComponent(newViewName)+"&originalViewName="+encodeURIComponent(originalViewName),
						complete: function (data) 
						{
							//alert(data);
							window.location.href = "mainRecyclePool.asp";
			            }
					})
							
				}
				else {
				    swal("The view name you have entered already exists. Please enter a different name.");
           			return false;
           		}

	        }
		})
	}	
	
	



	
	function renameProspectFilterViewWonPool() {
	
	   var newViewName = $("#txtUpdatedFilterReportViewNameWonPool").val();
	   var originalViewName = $("#originalViewNameWonPool").val();
	   
		$.ajax({
			type:"POST",
			url: "../inc/InSightFuncs_AjaxForProspectingModals.asp",
			cache: false,
			data: "action=CheckIfViewNameExistsWonPool&newViewName=" + encodeURIComponent(newViewName),
			success: function(response)
			{				
				if (response == "False")
				{					
		       		$.ajax({
						type:"POST",
						url: "renameWonPoolProspectFilterViewFromModal.asp?newViewName="+encodeURIComponent(newViewName)+"&originalViewName="+encodeURIComponent(originalViewName),
						complete: function (data) 
						{
							//alert(data);
							window.location.href = "mainWonPool.asp";
			            }
					})
							
				}
				else {
				    swal("The view name you have entered already exists. Please enter a different name.");
           			return false;
           		}

	        }
		})
	}	
	
	
		    
	
	function updateStageForProspect(prospectIntRecID) {
		var selectedStage = $("#selStageChange " + prospectIntRecID).val();	
		alert(selectedStage);
		$.ajax({		
			type:"POST",
			url: "updateProspectStageFromGridView.asp?pid="+prospectIntRecID+"&stage="+selectedStage,
			complete: function (data) {
				//$("#saveProspectGroup").html("success!");
				//window.location.reload(true); 
			}
		})	
		
	}	
	
		   	
</script>

