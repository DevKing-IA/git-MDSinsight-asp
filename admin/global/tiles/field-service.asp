<!--#include file="../../../inc/header.asp"-->
<!--#include file="../../../inc/InsightFuncs.asp"-->

<link rel="stylesheet" href="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.css" type="text/css">
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.js"></script>

<%
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))

	' Check to see if the Salesman file exists
	SalesmanFileExists = True
						
	'Only if the backend has a salesman table
	On Error Goto 0
	Set rs = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rs = cnn8.Execute("SELECT TOP 1 * FROM Salesman")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			SalesmanFileExists = False
		End If
	End IF

	

	SQL = "SELECT * FROM Settings_Global"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		ServiceColorsOn = rs("ServiceColorsOn")
		ServiceNormalAlertColor = rs("ServiceNormalAlertColor")
		ServicePriorityColor = rs("ServicePriorityColor")
		ServicePriorityAlertColor = rs("ServicePriorityAlertColor")			
		NoActivityNagMessageONOFF_FS = rs("NoActivityNagMessageONOFF_FS")
		NoActivityNagMinutes_FS = rs("NoActivityNagMinutes_FS")
		NoActivityNagIntervalMinutes_FS = rs("NoActivityNagIntervalMinutes_FS")
		NoActivityNagMessageMaxToSendPerStop_FS = rs("NoActivityNagMessageMaxToSendPerStop_FS")
		NoActivityNagMessageMaxToSendPerDriverPerDay_FS = rs("NoActivityNagMessageMaxToSendPerDriverPerDay_FS")
		NoActivityNagMessageSendMethod_FS = rs("NoActivityNagMessageSendMethod_FS")
		NoActivityNagTimeOfDay_FS = rs("NoActivityNagTimeOfDay_FS")  		
		FS_SignatureOptional = rs("FS_SignatureOptional") 
		FS_TechCanDecline = rs("FS_TechCanDecline") 
		FSDefaultNotificationMethod  = rs("FSDefaultNotificationMethod") 
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	'**********************************************************************
	'Get DLink Settings From Settings_EmailService
	'**********************************************************************

	SQL = "SELECT * FROM Settings_EmailService "
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		DLinkInEmail = rs("IncludeACKInDispatchEmail")
		DLinkInText = rs("IncludeACKinDispatchText")
	End If
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
 
 
	'************************************************
	'The newest settings are in Settings_FieldService
	'************************************************
	SQL = "SELECT * FROM Settings_FieldService "
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		ServiceTicketCarryoverReportOnOff = rs("ServiceTicketCarryoverReportOnOff")
		ServiceTicketCarryoverReportToPrimarySalesman = rs("ServiceTicketCarryoverReportToPrimarySalesman")
		ServiceTicketCarryoverReportToSecondarySalesman = rs("ServiceTicketCarryoverReportToSecondarySalesman")
		ServiceTicketCarryoverReportEmailSubject = rs("ServiceTicketCarryoverReportEmailSubject")
		ServiceTicketCarryoverReportUserNos = rs("ServiceTicketCarryoverReportUserNos")		
		ServiceTicketCarryoverReportAdditionalEmails = rs("ServiceTicketCarryoverReportAdditionalEmails")	
		ServiceTicketCarryoverReportTextSummaryOnOff = rs("ServiceTicketCarryoverReportTextSummaryOnOff")
		ServiceTicketCarryoverReportTextSummaryUserNos = rs("ServiceTicketCarryoverReportTextSummaryUserNos")
		ServiceTicketCarryoverReportTeamIntRecIDs = rs("ServiceTicketCarryoverReportTeamIntRecIDs")
		ServiceTicketCarryoverReportIncludeRegions = rs("ServiceTicketCarryoverReportIncludeRegions")
		CarryoverReportInclCustType = rs("CarryoverReportInclCustType")		
		CarryoverReportInclTicketNum = rs("CarryoverReportInclTicketNum")		
		CarryoverReportShowRedoBreakdown = rs("CarryoverReportShowRedoBreakdown")	
		
		ServiceDayStartTime = rs("ServiceDayStartTime")
		ServiceDayEndTime = rs("ServiceDayEndTime")
		ServiceDayElapsedTimeCalculationMethod = rs("ServiceDayElapsedTimeCalculationMethod")
		
		FieldServiceNotesReportOnOff = rs("FieldServiceNotesReportOnOff")
		FieldServiceNotesReportUserNos = rs("FieldServiceNotesReportUserNos")
		FieldServiceNotesReportAdditionalEmails = rs("FieldServiceNotesReportAdditionalEmails")
		FieldServiceNotesReportEmailSubject = rs("FieldServiceNotesReportEmailSubject")	
			
		AutoDispatchUsersOnOff = rs("AutoDispatchUsersOnOff")	
		AutoDispatchUserNos = rs("AutoDispatchUserNos")	
	
		FS_ShowPartsButton = rs("ShowPartsButton")
		ServiceTicketScreenShowHoldTab = rs("ServiceTicketScreenShowHoldTab")
		
		ServiceTicketThresholdReportONOFF = rs("ServiceTicketThresholdReportONOFF")
		ServiceTicketThresholdReportOnlyUndispatched = rs("ServiceTicketThresholdReportOnlyUndispatched")
		ServiceTicketThresholdReportOnlySkipFilterChanges = rs("ServiceTicketThresholdReportOnlySkipFilterChanges")
		ServiceTicketThresholdReportThresholdHours = rs("ServiceTicketThresholdReportThresholdHours")
		ServiceTicketThresholdReportUserNos = rs("ServiceTicketThresholdReportUserNos")
		ServiceTicketThresholdReportAdditionalEmails = rs("ServiceTicketThresholdReportAdditionalEmails")	
		
		FSBoardKioskGlobalUseRegions = rs("FSBoardKioskGlobalUseRegions")
		FSBoardKioskGlobalTitleText = rs("FSBoardKioskGlobalTitleText")
		FSBoardKioskGlobalTitleTextFontColor = rs("FSBoardKioskGlobalTitleTextFontColor")
		FSBoardKioskGlobalTitleGradientColor = rs("FSBoardKioskGlobalTitleGradientColor")
		FSBoardKioskGlobalColorPieTimer = rs("FSBoardKioskGlobalColorPieTimer")
		FSBoardKioskGlobalColorAwaitingDispatch = rs("FSBoardKioskGlobalColorAwaitingDispatch")
		FSBoardKioskGlobalColorAwaitingAcknowledgement = rs("FSBoardKioskGlobalColorAwaitingAcknowledgement")
		FSBoardKioskGlobalColorDispatchAcknowledged = rs("FSBoardKioskGlobalColorDispatchAcknowledged")
		FSBoardKioskGlobalColorDispatchDeclined = rs("FSBoardKioskGlobalColorDispatchDeclined")
		FSBoardKioskGlobalColorEnRoute = rs("FSBoardKioskGlobalColorEnRoute")
		FSBoardKioskGlobalColorOnSite = rs("FSBoardKioskGlobalColorOnSite")
		FSBoardKioskGlobalColorRedoSwap = rs("FSBoardKioskGlobalColorRedoSwap")
		FSBoardKioskGlobalColorRedoWaitForParts = rs("FSBoardKioskGlobalColorRedoWaitForParts")
		FSBoardKioskGlobalColorRedoFollowUp = rs("FSBoardKioskGlobalColorRedoFollowUp")
		FSBoardKioskGlobalColorRedoUnableToWork = rs("FSBoardKioskGlobalColorRedoUnableToWork")
		FSBoardKioskGlobalColorClosed = rs("FSBoardKioskGlobalColorClosed")
		FSBoardKioskGlobalColorUrgent = rs("FSBoardKioskGlobalColorUrgent")	
	End If
		
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

%>

<!-- bootstrap timepicker !-->
<script type="text/javascript" src="http://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>	
<link href="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.css" rel="stylesheet" type="text/css">
<script src="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.min.js" type="text/javascript"></script>
<!-- eof bootstrap timepicker !-->

<!-- spectrum color picker !-->
<script src="<%= BaseURL %>/js/spectrum-color-picker/spectrum.js"></script>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>/js/spectrum-color-picker/spectrum.css">
<!-- eof spectrum color picker !-->


<style  type="text/css">

	.content-element{
	  margin:50px 0 0 50px;
	}
	.circles-list ol {
	  list-style-type: none;
	  margin-left: 1.25em;
	  padding-left: 2.5em;
	  counter-reset: li-counter;
	  border-left: 1px solid #3c763d;
	  position: relative; }
	
	.circles-list ol > li {
	  position: relative;
	  margin-bottom: 3.125em;
	  clear: both; }
	
	.circles-list ol > li:before {
	  position: absolute;
	  top: -0.5em;
	  font-family: "Open Sans", sans-serif;
	  font-weight: 600;
	  font-size: 1em;
	  left: -3.75em;
	  width: 2.25em;
	  height: 2.25em;
	  line-height: 2.25em;
	  text-align: center;
	  z-index: 9;
	  color: #3c763d;
	  border: 2px solid #3c763d;
	  border-radius: 50%;
	  content: counter(li-counter);
	  background-color: #DFF0D8;
	  counter-increment: li-counter; }
	  	
	.row .panel-row{
	    margin-top:40px;
	    padding: 0 10px;
	}
	
	.clickable{
	    cursor: pointer;   
	}
	
	.panel-heading span {
		margin-top: -20px;
		font-size: 15px;
	}

	.container {
		margin-bottom: 20px;
		margin-top: 20px;
		margin-left:0px;
		width: 100%;
	}

	.container .row {
		margin-bottom: 20px;
		/*margin-top: 20px;*/
	}

	.full-spectrum .sp-palette {
		max-width: 200px;
	}
	
	.tab-colors-box{
		padding:15px;
		border:2px solid #000;
		margin:0px 0px 15px 0px;
		width:100%;
		display:block;
		float:left;
	}
	
	.tab-colors-title strong{
		width:100%;
		text-align:center;
		display:block;
	}
	
	.tab-colors-title .row{
		margin-bottom:0px;
	}
	
	.line-full{
	 	margin-bottom:20px;
	}
	
	.multi-select{
		min-height:200px;
		min-width:180px;
	}
	
	.custom-select{
		width: auto !important;
		display:inline-block;
	}

	.multi-select-dispatch{
	  min-height: 160px;
	  margin-top: 20px;
	 }

	
	.select-large{
		min-width:40% !important;
	}
	
	.ui-timepicker-table td a{
		padding: 3px;
		width:auto;
		text-align: left;
		font-size: 11px;
	}	
	
	.ui-timepicker-table .ui-timepicker-title{
		font-size: 13px;
	}
	
	.ui-timepicker-table th.periods{
		font-size: 13px;
	}
	
	.ui-widget-header{
		background: #193048;
		border: 1px solid #193048;
	}
	
	#PleaseWaitPanel{
		position: fixed;
		left: 470px;
		top: 275px;
		width: 975px;
		height: 300px;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	}   

	.btn-huge{
	    padding: 18px 28px;
	    font-size: 22px;	    
	}	
		
</style>
<!-- eof spectrum color picker !-->


<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;<%= GetTerm("Field Service") %> Settings 
	<button id="toggle" class="btn btn-small btn-success"><i class="fas fa-arrows-v"></i>&nbsp;EXPAND/COLLAPSE ALL SETTINGS</button>
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>


<form method="post" action="field-service-submit.asp" name="frmFieldService" id="frmFieldService">

<div class="container">
	
	<script>
	
		function showSavingChangesDiv() {
		  document.getElementById('PleaseWaitPanel').style.display = "block";
		  setTimeout(function() {
		    document.getElementById('PleaseWaitPanel').style.display = "none";
		  },1500);
		   
		}
	
		$(document).ready(function() {
	
			$('.panel .panel-body').css('display','none');
			$('.panel-heading span.clickable').addClass('panel-collapsed');
			$('.panel-heading span.clickable').find('i').removeClass('glyphicon-chevron-up').addClass('glyphicon-chevron-down');
	
			$(document).on('click', '.panel-heading span.clickable', function(e){
			    var $this = $(this);
				if(!$this.hasClass('panel-collapsed')) {
					$this.parents('.panel').find('.panel-body').slideUp();
					$this.addClass('panel-collapsed');
					$this.find('i').removeClass('glyphicon-chevron-up').addClass('glyphicon-chevron-down');
				} else {
					$this.parents('.panel').find('.panel-body').slideDown();
					$this.removeClass('panel-collapsed');
					$this.find('i').removeClass('glyphicon-chevron-down').addClass('glyphicon-chevron-up');
				}
			});
			
			
	 		$("#toggle").click(function(){
	 		
	            if(!$('.panel-heading span.clickable').hasClass('panel-collapsed')) {
					$('.panel .panel-body').css('display','none');
					$('.panel-heading span.clickable').addClass('panel-collapsed');
					$('.panel-heading span.clickable').find('i').removeClass('glyphicon-chevron-up').addClass('glyphicon-chevron-down');
	            }
	            else {
					$('.panel .panel-body').css('display','block');
					$('.panel-heading span.clickable').removeClass('panel-collapsed');
					$('.panel-heading span.clickable').find('i').removeClass('glyphicon-chevron-down').addClass('glyphicon-chevron-up');
	            }
	        });	
	        
	
			$('#modalFieldServiceNotesReportScheduler').on('show.bs.modal', function(e) {
			    	    
			    var $modal = $(this);
		
		    	$.ajax({
					type:"POST",
					url: "../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
					cache: false,
					data: "action=GetContentForFieldServiceNotesReportScheduler",
					success: function(response)
					 {
		               	 $modal.find('#modalFieldServiceNotesReportSchedulerContent').html(response);               	 
		             },
		             failure: function(response)
					 {
					  	$modal.find('#modalFieldServiceNotesReportSchedulerContent').html("Failed");
			            //var height = $(window).height() - 600;
			            //$(this).find(".modal-body").css("max-height", height);
		             }
				});
				
			});
			        
	
	        $('#modalNotesReportAdditionalEmails').on('show.bs.modal', function (e) {
	            var $modal = $(this);
	        });
	        
			
			$('#modalServiceTicketCarryoverReportScheduler').on('show.bs.modal', function(e) {
			    	    
			    var $modal = $(this);
		
		    	$.ajax({
					type:"POST",
					url: "../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
					cache: false,
					data: "action=GetContentForServiceTicketCarryoverReportScheduler",
					success: function(response)
					 {
		               	 $modal.find('#modalServiceTicketCarryoverReportSchedulerContent').html(response);               	 
		             },
		             failure: function(response)
					 {
					  	$modal.find('#modalServiceTicketCarryoverReportSchedulerContent').html("Failed");
			            //var height = $(window).height() - 600;
			            //$(this).find(".modal-body").css("max-height", height);
		             }
				});
	
	        });
	        
	
	        $('#modalCarryoverReportAdditionalEmails').on('show.bs.modal', function (e) {
	            var $modal = $(this);
	        });
	        
	
	        $('#modalCarryoverReportTeamIntRecIDs').on('show.bs.modal', function (e) {
	
	            var $modal = $(this);
	
	            $.ajax({
	                type: "POST",
	                url: "../../../inc/InSightFuncs_AjaxForAdminSelectUsers.asp",
	                cache: false,
	                data: "action=GetContentForCarryoverReportTeamIntRecIDs",
	                success: function (response) {
	                    $modal.find('#modalCarryoverReportTeamIntRecIDsContent').html(response);
	                },
	                failure: function (response) {
	                    $modal.find('#modalCarryoverReportTeamIntRecIDsContent').html("Failed");
	                    //var height = $(window).height() - 600;
	                    //$(this).find(".modal-body").css("max-height", height);
	                }
	            });
	
	        });
	
			$("#chkServiceTicketCarryoverReportTextSummaryOnOff").change(function() {
			    if ( $(this).is(':checked') ) {
			        $("#btnTextSummaryUsersButton").show();
			        $("#btnTextSummaryUsersList").show();
			    } else {
			        $("#btnTextSummaryUsersButton").hide();
			        $("#btnTextSummaryUsersList").hide();
			    }
			});
			
	        
			
			$('#modalServiceTicketThresholdReportScheduler').on('show.bs.modal', function(e) {
			    	    
			    var $modal = $(this);
		
		    	$.ajax({
					type:"POST",
					url: "../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
					cache: false,
					data: "action=GetContentForServiceTicketThresholdReportScheduler",
					success: function(response)
					 {
		               	 $modal.find('#modalServiceTicketThresholdReportSchedulerContent').html(response);               	 
		             },
		             failure: function(response)
					 {
					  	$modal.find('#modalServiceTicketThresholdReportSchedulerContent').html("Failed");
			            //var height = $(window).height() - 600;
			            //$(this).find(".modal-body").css("max-height", height);
		             }
				});
	
	        });
	
	
	        $('#modalThresholdReportAdditionalEmails').on('show.bs.modal', function (e) {
	            var $modal = $(this);
	        });
				
	
			$('#lstExistingFieldServiceNotesReportUserIDs').multiselect({
			   buttonTitle: function(options, select) {
				    var selected = '';
				    options.each(function () {
				      selected += $(this).text() + ', ';
				    });
				    return selected.substr(0, selected.length - 2);
				  },
				buttonClass: 'btn btn-primary',
				buttonWidth: '425px',
				maxHeight: 400,
				dropRight:true,
				enableFiltering:true,
				filterPlaceholder:'Search',
				enableCaseInsensitiveFiltering:true,
				// possible options: 'text', 'value', 'both'
				filterBehavior:'text',
				includeFilterClearBtn:true,
				nonSelectedText:'No Users Selected For Notes Report',
				numberDisplayed: 20,
			    onChange: function() {
			        var selected = this.$select.val();
			        $("#lstSelectedFieldServiceNotesReportUserIDs").val(selected);
			        console.log(selected);
			        // ...
			    }
	    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current field service notes report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedFieldServiceNotesReportUserIDs").val();
			
			if (data) {
				//Make an array
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingFieldServiceNotesReportUserIDs").val(dataarray);
				// Then refresh
				$("#lstExistingFieldServiceNotesReportUserIDs").multiselect("refresh");
			}
			//*************************************************************************************************
			
			
			
			$('#lstExistingServiceTicketCarryoverReportUserIDs').multiselect({
			   buttonTitle: function(options, select) {
				    var selected = '';
				    options.each(function () {
				      selected += $(this).text() + ', ';
				    });
				    return selected.substr(0, selected.length - 2);
				  },
				buttonClass: 'btn btn-primary',
				buttonWidth: '425px',
				maxHeight: 400,
				dropRight:true,
				enableFiltering:true,
				filterPlaceholder:'Search',
				enableCaseInsensitiveFiltering:true,
				// possible options: 'text', 'value', 'both'
				filterBehavior:'text',
				includeFilterClearBtn:true,
				nonSelectedText:'No Users Selected For Carryover Report',
				numberDisplayed: 20,
			    onChange: function() {
			        var selected = this.$select.val();
			        $("#lstSelectedServiceTicketCarryoverReportUserIDs").val(selected);
			        console.log(selected);
			        // ...
			    }
	    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current carryover report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedServiceTicketCarryoverReportUserIDs").val();
			
			if (data) {
				//Make an array
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingServiceTicketCarryoverReportUserIDs").val(dataarray);
				// Then refresh
				$("#lstExistingServiceTicketCarryoverReportUserIDs").multiselect("refresh");
			}
			//*************************************************************************************************
	        
		
	
			
			
			
			$('#lstExistingServiceTicketCarryoverReportTeamIntRecIDs').multiselect({
			   buttonTitle: function(options, select) {
				    var selected = '';
				    options.each(function () {
				      selected += $(this).text() + ', ';
				    });
				    return selected.substr(0, selected.length - 2);
				  },
				buttonClass: 'btn btn-primary',
				buttonWidth: '425px',
				maxHeight: 400,
				dropRight:true,
				enableFiltering:true,
				filterPlaceholder:'Search',
				enableCaseInsensitiveFiltering:true,
				// possible options: 'text', 'value', 'both'
				filterBehavior:'text',
				includeFilterClearBtn:true,
				nonSelectedText:'No Teams Selected For Carryover Report',
				numberDisplayed: 20,
			    onChange: function() {
			        var selected = this.$select.val();
			        $("#lstSelectedServiceTicketCarryoverReportTeamIntRecIDs").val(selected);
			        console.log(selected);
			        // ...
			    }
	    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current carryover report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedServiceTicketCarryoverReportTeamIntRecIDs").val();
			
			if (data) {
				//Make an array
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingServiceTicketCarryoverReportTeamIntRecIDs").val(dataarray);
				// Then refresh
				$("#lstExistingServiceTicketCarryoverReportTeamIntRecIDs").multiselect("refresh");
			}
			//*************************************************************************************************
	        
	

						
			
			$('#lstExistingServiceTicketCarryoverReportTextSummmaryUserIDs').multiselect({
			   buttonTitle: function(options, select) {
				    var selected = '';
				    options.each(function () {
				      selected += $(this).text() + ', ';
				    });
				    return selected.substr(0, selected.length - 2);
				  },
				buttonClass: 'btn btn-primary',
				buttonWidth: '425px',
				maxHeight: 400,
				dropRight:true,
				enableFiltering:true,
				filterPlaceholder:'Search',
				enableCaseInsensitiveFiltering:true,
				// possible options: 'text', 'value', 'both'
				filterBehavior:'text',
				includeFilterClearBtn:true,
				nonSelectedText:'No Users Selected To Receive Text Summaries',
				numberDisplayed: 20,
			    onChange: function() {
			        var selected = this.$select.val();
			        $("#lstSelectedServiceTicketCarryoverReportTextSummmaryUserIDs").val(selected);
			        console.log(selected);
			        // ...
			    }
	    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current threshhold report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedServiceTicketCarryoverReportTextSummmaryUserIDs").val();
			//Make an array
			
			if (data) {
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingServiceTicketCarryoverReportTextSummmaryUserIDs").val(dataarray);
				// Then refresh
				$("#lstExistingServiceTicketCarryoverReportTextSummmaryUserIDs").multiselect("refresh");
			}
			//*************************************************************************************************

	
			
						
			
			$('#lstExistingServiceTicketThresholdReportUserNos').multiselect({
			   buttonTitle: function(options, select) {
				    var selected = '';
				    options.each(function () {
				      selected += $(this).text() + ', ';
				    });
				    return selected.substr(0, selected.length - 2);
				  },
				buttonClass: 'btn btn-primary',
				buttonWidth: '425px',
				maxHeight: 400,
				dropRight:true,
				enableFiltering:true,
				filterPlaceholder:'Search',
				enableCaseInsensitiveFiltering:true,
				// possible options: 'text', 'value', 'both'
				filterBehavior:'text',
				includeFilterClearBtn:true,
				nonSelectedText:'No Users Selected For Threshold Report',
				numberDisplayed: 20,
			    onChange: function() {
			        var selected = this.$select.val();
			        $("#lstSelectedServiceTicketThresholdReportUserNos").val(selected);
			        console.log(selected);
			        // ...
			    }
	    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current threshhold report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedServiceTicketThresholdReportUserNos").val();
			//Make an array
			
			if (data) {
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingServiceTicketThresholdReportUserNos").val(dataarray);
				// Then refresh
				$("#lstExistingServiceTicketThresholdReportUserNos").multiselect("refresh");
			}
			//*************************************************************************************************
					
	        
		});
	</script>
	

<% If MUV_READ("serviceModuleOn") = "Disabled" Then %>
	<div class="col-lg-6">
		Please contact support if you would like to activate the <%=GetTerm("Field Service")%> module.
	</div>
<% Else %>

	<div class="container">
	
		<%
			Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
			Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
			Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
			Response.Write("</div>")
			Response.Flush()
		%>
	
	
		<div class="row">
			<h3><i class="fad fa-sliders-h"></i>&nbsp;<%= GetTerm("Field Service") %> Master Settings</h3>
		</div>
	
	    <div class="row">
	    
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title"><%= GetTerm("Field Service") %> General Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
			         	<div class="row">
			         	   	<div class="col-lg-8">Service Day Start Time</div>
							<div class="col-lg-4">		        
								<div class="input-group bootstrap-timepicker timepicker">
								  	<input id="txtServiceDayStartTime" type="text" name="txtServiceDayStartTime" value="<%= ServiceDayStartTime %>" class="form-control">
								 	<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
								</div>		        	    
							</div>
						</div>
						
			         	<div class="row">
			         	   	<div class="col-lg-8">Service Day End Time</div>
							<div class="col-lg-4">		        
								<div class="input-group bootstrap-timepicker timepicker">
								  	<input id="txtServiceDayEndTime" type="text" name="txtServiceDayEndTime" value="<%= ServiceDayEndTime %>" class="form-control">
								 	<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
								</div>		        	    
							</div>
						</div>
		
			         	<div class="row">
			         	   	<div class="col-lg-9">Calculate service ticket elapsed time using <strong>defined business hours above</strong></div>
				         	<div class="col-lg-3">
								<% If ServiceDayElapsedTimeCalculationMethod = "Business" Then %>
									<input type="radio" name="optCalcElapsedTime" id="optCalcElapsedTime" value="Business" checked>
								<% Else %>
									<input type="radio" name="optCalcElapsedTime" id="optCalcElapsedTime" value="Business">				
								<% End If %>
							</div>
						</div>
		
			         	<div class="row">
			         	   	<div class="col-lg-9">Calculate service ticket elapsed time using <strong>actual elapsed minutes</strong></div>
				         	<div class="col-lg-3">
								<% If ServiceDayElapsedTimeCalculationMethod = "Actual" Then %>
									<input type="radio" name="optCalcElapsedTime" id="optCalcElapsedTime" value="Actual" checked>
								<% Else %>
									<input type="radio" name="optCalcElapsedTime" id="optCalcElapsedTime" value="Actual">
								<% End If %>
							</div>
						</div>
					
					</div>
				</div>
			</div>
			
			<div class="col-md-4">
				&nbsp;
			</div>
	
			<div class="col-md-4">
				&nbsp;
			</div>
			
		</div>
	
		<div class="row">
			<h3><i class="fad fa-palette"></i>&nbsp;<%= GetTerm("Field Service") %> Color &amp; Display Settings</h3>
		</div>
		
		<div class="row">
		
			<div class="col-md-4">
				<div class="panel panel-info">
					<div class="panel-heading">
						<h3 class="panel-title"><%= GetTerm("Field Service") %> Screen Highlight Colors &amp; Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
				         	<div class="row">
					         	<div class="col-lg-5">Hightlight Colors On</div>
									<%
									If cInt(ServiceColorsOn) = 0 Then
										Response.Write("<input type='checkbox' id='chkServiceColorsOn' name='chkServiceColorsOn' ")
									Else
										Response.Write("<input type='checkbox' id='chkServiceColorsOn' name='chkServiceColorsOn' checked ")
									End If
									Response.Write(">")%>
							</div>
														
					    
				         	<div class="row">
				         	   	<div class="col-lg-9">Normal <%=GetTerm("Customer")%> - Alert Sent</div>
					         	<div class="col-lg-3">
									<input type='text' id="txtNormalAlert" name="txtNormalAlert" value="<%= ServiceNormalAlertColor %>">
								</div>
							</div>
						    
				         	<div class="row">
				         	   	<div class="col-lg-9">Priority <%=GetTerm("Customer")%></div>
					         	<div class="col-lg-3">
									<input type='text' id="txtPriorityAccount" name="txtPriorityAccount" value="<%= ServicePriorityColor %>">
								</div>
							</div>
						    
				         	<div class="row">
				         	   	<div class="col-lg-9">Priority <%=GetTerm("Customer")%> - Alert Sent</div>
					         	<div class="col-lg-3">
									<input type='text' id="txtPriorityAccountAlert" name="txtPriorityAccountAlert" value="<%= ServicePriorityAlertColor %>">
								</div>
							</div>
					
					</div>
				</div>
			</div>
			
			
			<div class="col-md-4">
				<div class="panel panel-info">
					<div class="panel-heading">
						<h3 class="panel-title"><%= GetTerm("Field Service") %> Global Color Settings (Kiosk and Board)</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					
					<div class="panel-body">
					
			         	<div class="row">
			         	   	<div class="col-lg-4">Pie Timer Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorPieTimer" name="txtFSBoardKioskGlobalColorPieTimer" value="<%= FSBoardKioskGlobalColorPieTimer %>">
							</div>
						</div>

			         	<div class="row">
			         	   	<div class="col-lg-4">Urgent Ticket Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorUrgent" name="txtFSBoardKioskGlobalColorUrgent" value="<%= FSBoardKioskGlobalColorUrgent %>">
							</div>
						</div>

			         	<div class="row">
			         		<div class="col-lg-12">
								<div class="text-element circles-list">
									<ol>
										<li>
											<input type='text' id="txtFSBoardKioskGlobalColorAwaitingDispatch" name="txtFSBoardKioskGlobalColorAwaitingDispatch" value="<%= FSBoardKioskGlobalColorAwaitingDispatch %>">
											Awaiting Dispatch Color
										</li>
										<li>
											<input type='text' id="txtFSBoardKioskGlobalColorAwaitingAcknowledgement" name="txtFSBoardKioskGlobalColorAwaitingAcknowledgement" value="<%= FSBoardKioskGlobalColorAwaitingAcknowledgement %>">
											Awaiting Acknowledgement Color
										</li>
										<li>
											<input type='text' id="txtFSBoardKioskGlobalColorDispatchAcknowledged" name="txtFSBoardKioskGlobalColorDispatchAcknowledged" value="<%= FSBoardKioskGlobalColorDispatchAcknowledged %>">
											Dispatch Acknowledged Color
										</li>	
										<li>
											<input type='text' id="txtFSBoardKioskGlobalColorDispatchDeclined" name="txtFSBoardKioskGlobalColorDispatchDeclined" value="<%= FSBoardKioskGlobalColorDispatchDeclined %>">
											Dispatched Declined Color
										</li>
										<li>
											<input type='text' id="txtFSBoardKioskGlobalColorEnRoute" name="txtFSBoardKioskGlobalColorEnRoute" value="<%= FSBoardKioskGlobalColorEnRoute %>">
											En Route Color
										</li>
										<li>
											<input type='text' id="txtFSBoardKioskGlobalColorOnSite" name="txtFSBoardKioskGlobalColorOnSite" value="<%= FSBoardKioskGlobalColorOnSite %>">
											On Site Color
										</li>
										<li>
											<input type='text' id="txtFSBoardKioskGlobalColorRedoSwap" name="txtFSBoardKioskGlobalColorRedoSwap" value="<%= FSBoardKioskGlobalColorRedoSwap %>">
											Swap Color (Redo)<br><br>									
											<input type='text' id="txtFSBoardKioskGlobalColorRedoWaitForParts" name="txtFSBoardKioskGlobalColorRedoWaitForParts" value="<%= FSBoardKioskGlobalColorRedoWaitForParts %>">
											Wait For Parts Color (Redo)<br><br>
											<input type='text' id="txtFSBoardKioskGlobalColorRedoFollowUp" name="txtFSBoardKioskGlobalColorRedoFollowUp" value="<%= FSBoardKioskGlobalColorRedoFollowUp %>">
											Follow Up Color (Redo)<br><br>
											<input type='text' id="txtFSBoardKioskGlobalColorRedoUnableToWork" name="txtFSBoardKioskGlobalColorRedoUnableToWork" value="<%= FSBoardKioskGlobalColorRedoUnableToWork %>">
											Unable To Work Color (Redo)
										</li>
										<li>
											<input type='text' id="txtFSBoardKioskGlobalColorClosed" name="txtFSBoardKioskGlobalColorClosed" value="<%= FSBoardKioskGlobalColorClosed %>">
											Closed Ticket Color
										</li>										
									</ol>
								</div>
							</div>
						</div>

					    <!--
						
						
			         	<div class="row">
			         	   	<div class="col-lg-9">Urgent Ticket Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorUrgent" name="txtFSBoardKioskGlobalColorUrgent" value="<%= FSBoardKioskGlobalColorUrgent %>">
							</div>
						</div>
	
			         	<div class="row">
			         	   	<div class="col-lg-9">Awaiting Dispatch Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorAwaitingDispatch" name="txtFSBoardKioskGlobalColorAwaitingDispatch" value="<%= FSBoardKioskGlobalColorAwaitingDispatch %>">
							</div>
						</div>
					    
			         	<div class="row">
			         	   	<div class="col-lg-9">Awaiting Acknowledgement Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorAwaitingAcknowledgement" name="txtFSBoardKioskGlobalColorAwaitingAcknowledgement" value="<%= FSBoardKioskGlobalColorAwaitingAcknowledgement %>">
							</div>
						</div>
						
			         	<div class="row">
			         	   	<div class="col-lg-9">Dispatch Acknowledged Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorDispatchAcknowledged" name="txtFSBoardKioskGlobalColorDispatchAcknowledged" value="<%= FSBoardKioskGlobalColorDispatchAcknowledged %>">
							</div>
						</div>
		
			         	<div class="row">
			         	   	<div class="col-lg-9">Dispatched Declined Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorDispatchDeclined" name="txtFSBoardKioskGlobalColorDispatchDeclined" value="<%= FSBoardKioskGlobalColorDispatchDeclined %>">
							</div>
						</div>

			         	<div class="row">
			         	   	<div class="col-lg-9">En Route Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorEnRoute" name="txtFSBoardKioskGlobalColorEnRoute" value="<%= FSBoardKioskGlobalColorEnRoute %>">
							</div>
						</div>
						
			         	<div class="row">
			         	   	<div class="col-lg-9">On Site Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorOnSite" name="txtFSBoardKioskGlobalColorOnSite" value="<%= FSBoardKioskGlobalColorOnSite %>">
							</div>
						</div>
						
			         	<div class="row">
			         	   	<div class="col-lg-9">Swap Color (Redo)</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorRedoSwap" name="txtFSBoardKioskGlobalColorRedoSwap" value="<%= FSBoardKioskGlobalColorRedoSwap %>">
							</div>
						</div>
						
			         	<div class="row">
			         	   	<div class="col-lg-9">Wait For Parts Color (Redo)</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorRedoWaitForParts" name="txtFSBoardKioskGlobalColorRedoWaitForParts" value="<%= FSBoardKioskGlobalColorRedoWaitForParts %>">
							</div>
						</div>
		
			         	<div class="row">
			         	   	<div class="col-lg-9">Follow Up Color (Redo)</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorRedoFollowUp" name="txtFSBoardKioskGlobalColorRedoFollowUp" value="<%= FSBoardKioskGlobalColorRedoFollowUp %>">
							</div>
						</div>
					
			         	<div class="row">
			         	   	<div class="col-lg-9">Unable To Work Color (Redo)</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorRedoUnableToWork" name="txtFSBoardKioskGlobalColorRedoUnableToWork" value="<%= FSBoardKioskGlobalColorRedoUnableToWork %>">
							</div>
						</div>
						
			         	<div class="row">
			         	   	<div class="col-lg-9">Closed Ticket Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalColorClosed" name="txtFSBoardKioskGlobalColorClosed" value="<%= FSBoardKioskGlobalColorClosed %>">
							</div>
						</div>-->
						
					</div>
				</div>
			</div>
			
			<div class="col-md-4">
				<div class="panel panel-info">
					<div class="panel-heading">
						<h3 class="panel-title"><%= GetTerm("Field Service") %> Kiosk Specific Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
			         	<div class="row">
			         	   	<div class="col-lg-12">Title Text (in kiosk mode)
			         	   		<br>&nbsp;&nbsp;&nbsp;<small>*To include the current date, use <b><i>~today~</i></b>.
			         	   		<br>&nbsp;&nbsp;&nbsp;*To include the day of week name, use <b><i>~dow~</i></b>.
			         	   		<br>&nbsp;&nbsp;&nbsp;Example: Service Tickets for ~dow~ ~today~ will show<b>&nbsp;&nbsp;Service Tickets for <%=WeekDayName(Datepart("w",Now()))%>&nbsp;<%=FormatDateTime(Now(),2)%></b></small>
							</div>
						</div>
		
			         	<div class="row">
					       	<div class="col-lg-12">
								<input type='text' id="txtFSBoardKioskGlobalTitleText" name="txtFSBoardKioskGlobalTitleText" value="<%= FSBoardKioskGlobalTitleText %>" class="form-control">
								
							</div>
						</div>
			         					    
			         	<div class="row">
			         	   	<div class="col-lg-6">Title Bar and Border Gradient Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalTitleGradientColor" name="txtFSBoardKioskGlobalTitleGradientColor" value="<%= FSBoardKioskGlobalTitleGradientColor %>">
							</div>
						</div>
		
			         					    
			         	<div class="row">
			         	   	<div class="col-lg-6">Title Text Font Color</div>
				         	<div class="col-lg-3">
								<input type='text' id="txtFSBoardKioskGlobalTitleTextFontColor" name="txtFSBoardKioskGlobalTitleTextFontColor" value="<%= FSBoardKioskGlobalTitleTextFontColor %>">
							</div>
						</div>
					
	

			         	<div class="row">
			         	   	<div class="col-lg-6">Use Regions</div>
				         	<div class="col-lg-3">
				         		<%
				         			If Not IsNumeric(FSBoardKioskGlobalUseRegions) Then FSBoardKioskGlobalUseRegions = 0
									If cInt(FSBoardKioskGlobalUseRegions) = 0 Then
										Response.Write("<input type='checkbox' id='chkFSBoardKioskGlobalUseRegions' name='chkFSBoardKioskGlobalUseRegions' ")
									Else
										Response.Write("<input type='checkbox' id='chkFSBoardKioskGlobalUseRegions' name='chkFSBoardKioskGlobalUseRegions' checked ")
									End If
									Response.Write(">")
								%>
							</div>
						</div>

			         	<div class="row">
			         	   	<div class="col-lg-6">Show HOLD tab</div>
				         	<div class="col-lg-3">
				         		<%
				         			If Not IsNumeric(ServiceTicketScreenShowHoldTab) Then ServiceTicketScreenShowHoldTab = 0
									If cInt(ServiceTicketScreenShowHoldTab) = 0 Then
										Response.Write("<input type='checkbox' id='chkServiceTicketScreenShowHoldTab' name='chkServiceTicketScreenShowHoldTab' ")
									Else
										Response.Write("<input type='checkbox' id='chkServiceTicketScreenShowHoldTab' name='chkServiceTicketScreenShowHoldTab' checked ")
									End If
									Response.Write(">")
								%>
							</div>
						</div>
				
					</div>
				</div>
			</div>
			
			
		</div>
		
		
		<div class="row">
			<h3><i class="fad fa-comment-exclamation"></i>&nbsp;<%= GetTerm("Field Service") %> Alert Settings</h3>
		</div>
		
		
		<div class="row">
		
			<div class="col-md-4">
				<div class="panel panel-warning">
					<div class="panel-heading">
						<h3 class="panel-title"><%= GetTerm("Field Service") %> Dispatch Email/Text Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
			         	<div class="row">
				         	<div class="col-lg-8">Include Acknowledge / Decline Link in Email</div>
								<%
								If DLinkInEmail = 0 Then
									Response.Write("<input type='checkbox' id='chkDLinkInEmail' name='chkDLinkInEmail' ")
								Else
									Response.Write("<input type='checkbox' id='chkDLinkInEmail' name='chkDLinkInEmail' checked ")
								End If
								Response.Write(">")%>
						</div>
				    
			         	<div class="row">
				         	<div class="col-lg-8">Include Acknowledge / Decline Link in Text Message</div>
								<%
								If DLinkInText = 0 Then
									Response.Write("<input type='checkbox' id='chkDLinkInText' name='chkDLinkInText' ")
								Else
									Response.Write("<input type='checkbox' id='chkDLinkInText' name='chkDLinkInText' checked ")
								End If
								Response.Write(">")%>
						</div>
	
			         	<div class="row">
				         	<div class="col-lg-8">Signature is optional</div>
								<%
								If FS_SignatureOptional = 0 Then
									Response.Write("<input type='checkbox' id='chkFSSignatureOptional' name='chkFSSignatureOptional' ")
								Else
									Response.Write("<input type='checkbox' id='chkFSSignatureOptional' name='chkFSSignatureOptional' checked ")
								End If
								Response.Write(">")%>
						</div>
					
			         	<div class="row">
				         	<div class="col-lg-8">Techs can decline dispatches</div>
								<%
								If FS_TechCanDecline = 0 Then
									Response.Write("<input type='checkbox' id='chkFSTechCanDecline' name='chkFSTechCanDecline' ")
								Else
									Response.Write("<input type='checkbox' id='chkFSTechCanDecline' name='chkFSTechCanDecline' checked ")
								End If
								Response.Write(">")%>
						</div>
	
			         	<div class="row">
				         	<div class="col-lg-8">Show parts button on mobile menu</div>
								<%
								If FS_ShowPartsButton = 0 Then
									Response.Write("<input type='checkbox' id='chkFSShowPartsButton' name='chkFSShowPartsButton' ")
								Else
									Response.Write("<input type='checkbox' id='chkFSShowPartsButton' name='chkFSShowPartsButton' checked ")
								End If
								Response.Write(">")%>
						</div>
	
			         	<div class="row">
							<div class="col-lg-8">Default notification method</div>
								<select class="form-control custom-select" id="selFSDefaultNotificationMethod" name="selFSDefaultNotificationMethod">			
									<option value="Text"<%If FSDefaultNotificationMethod = "Test" Then Response.Write(" selected ")%>>Text</option>
									<option value="Email"<%If FSDefaultNotificationMethod = "Email" Then Response.Write(" selected ")%>>Email</option>
									<option value="Text & Email"<%If FSDefaultNotificationMethod = "Text & Email" Then Response.Write(" selected ")%>>Text & Email</option>
									<option value="None"<%If FSDefaultNotificationMethod   = "None" Then Response.Write(" selected ")%>>None</option>
								</select>
						</div>
	
			         	<div class="row">
				         	<div class="col-lg-8">Auto dispatch service tickets</div>
								<%
								If AutoDispatchUsersOnOff = 0 Then
									Response.Write("<input type='checkbox' id='chkAutoDispatchUsersOnOff' name='chkAutoDispatchUsersOnOff' ")
								Else
									Response.Write("<input type='checkbox' id='chkAutoDispatchUsersOnOff' name='chkAutoDispatchUsersOnOff' checked ")
								End If
								Response.Write(">")%>
						</div>
	
			         	<div class="row">
							<div class="col-lg-8">User(s) to auto dispatch&nbsp;
								<select class="form-control multi-select-dispatch" id="selAutoDispatchUserNos" name="selAutoDispatchUserNos" multiple>			
									<option value="0"<%If AutoDispatchUserNos = "" Then Response.Write(" selected ")%>>--- none from here ---</option>
									<option value="<%=Session("UserNo")%>"<%If UserInList(Session("UserNo"),AutoDispatchUserNos) = True Then Response.Write(" selected ")%>><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option>
							      	<%'Users dropdown
							      	 
						      	  	SQL = "SELECT UserNo, userFirstName, userLastName, userDisplayName FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
						      	  	SQL = SQL & "WHERE userEnabled = 1 AND userArchived <> 1 and UserNo <> " & Session("UserNo")
						      	  	SQL = SQL & " order by  userFirstName, userLastName"
					
									Set cnn8 = Server.CreateObject("ADODB.Connection")
									cnn8.open (Session("ClientCnnString"))
									Set rs = Server.CreateObject("ADODB.Recordset")
									rs.CursorLocation = 3 
									Set rs = cnn8.Execute(SQL)
								
									If not rs.EOF Then
										Do
											FullName = rs("userFirstName") & " " & rs("userLastName") & " (" & rs("userDisplayName") & ")"
											If UserInList(rs("UserNo"),AutoDispatchUserNos) = True Then
												Response.Write("<option value='" & rs("UserNo") & "' selected>" & FullName & "</option>")
											Else
												Response.Write("<option value='" & rs("UserNo") & "'>" & FullName & "</option>")
											End If
											rs.movenext
										Loop until rs.eof
									End If
									set rs = Nothing
									cnn8.close
									set cnn8 = Nothing
							      	%>
								</select>
								<br><strong>Use CTRL and SHIFT to make multiple selections</strong></div>
						</div>
					
					
					</div>
				</div>
			</div>
			
			<div class="col-md-4">
				<div class="panel panel-warning">
					<div class="panel-heading">
						<h3 class="panel-title"><%= GetTerm("Field Service") %> Technician Nag Alert Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
		         	<div class="row"style="padding-left:15px;padding-right:15px;">
	
			                <!-- no activity box -->
							<div class="nag-box2">
			                    <font color="Red"><strong><i>*Insight will begin sending these nag messages after the technician has changed the 
			                    status of any of their service tickets or the time of day has reached the time set below.</i></strong></font>
			
								<div class="row">
								 	<div class="col-lg-12">Send 'nag' messages when there has been a period of <strong>No Activity</strong>
										<%
										If NoActivityNagMessageONOFF_FS = 0 Then
											Response.Write("<input type='checkbox' id='chkNoActivityNagMessageONOFF_FS' name='chkNoActivityNagMessageONOFF_FS' ")
										Else
											Response.Write("<input type='checkbox' id='chkNoActivityNagMessageONOFF_FS' name='chkNoActivityNagMessageONOFF_FS' checked ")
										End If
										Response.Write(">")%>
									</div>
								</div>
								
								<div class="row">
									<div class="col-lg-12">Start sending messages if there has been No Activity by 
										<select class="form-control custom-select" id="selNoActivityNagTimeOfDay_FS" name="selNoActivityNagTimeOfDay_FS">			
											<option value="12:00"<%If NoActivityNagTimeOfDay_FS = "12:00" Then Response.Write(" selected ")%>>-Midnight-</option>
											<option value="12:15"<%If NoActivityNagTimeOfDay_FS = "12:15" Then Response.Write(" selected ")%>>12:15 AM</option>
											<option value="12:30"<%If NoActivityNagTimeOfDay_FS = "12:30" Then Response.Write(" selected ")%>>12:30 AM</option>
											<option value="12:45"<%If NoActivityNagTimeOfDay_FS = "12:45" Then Response.Write(" selected ")%>>12:45 AM</option>
											<option value="1:00"<%If NoActivityNagTimeOfDay_FS = "1:00" Then Response.Write(" selected ")%>>1:00 AM</option>
											<option value="1:15"<%If NoActivityNagTimeOfDay_FS = "1:15" Then Response.Write(" selected ")%>>1:15 AM</option>
											<option value="1:30"<%If NoActivityNagTimeOfDay_FS = "1:30" Then Response.Write(" selected ")%>>1:30 AM</option>
											<option value="1:45"<%If NoActivityNagTimeOfDay_FS = "1:45" Then Response.Write(" selected ")%>>1:45 AM</option>
											<option value="2:00"<%If NoActivityNagTimeOfDay_FS = "2:00" Then Response.Write(" selected ")%>>2:00 AM</option>
											<option value="2:15"<%If NoActivityNagTimeOfDay_FS = "2:15" Then Response.Write(" selected ")%>>2:15 AM</option>
											<option value="2:30"<%If NoActivityNagTimeOfDay_FS = "2:30" Then Response.Write(" selected ")%>>2:30 AM</option>
											<option value="2:45"<%If NoActivityNagTimeOfDay_FS = "2:45" Then Response.Write(" selected ")%>>2:45 AM</option>
											<option value="3:00"<%If NoActivityNagTimeOfDay_FS = "3:00" Then Response.Write(" selected ")%>>3:00 AM</option>
											<option value="3:15"<%If NoActivityNagTimeOfDay_FS = "3:15" Then Response.Write(" selected ")%>>3:15 AM</option>
											<option value="3:30"<%If NoActivityNagTimeOfDay_FS = "3:30" Then Response.Write(" selected ")%>>3:30 AM</option>
											<option value="3:45"<%If NoActivityNagTimeOfDay_FS = "3:45" Then Response.Write(" selected ")%>>3:45 AM</option>
											<option value="4:00"<%If NoActivityNagTimeOfDay_FS = "4:00" Then Response.Write(" selected ")%>>4:00 AM</option>
											<option value="4:15"<%If NoActivityNagTimeOfDay_FS = "4:15" Then Response.Write(" selected ")%>>4:15 AM</option>
											<option value="4:30"<%If NoActivityNagTimeOfDay_FS = "4:30" Then Response.Write(" selected ")%>>4:30 AM</option>
											<option value="4:45"<%If NoActivityNagTimeOfDay_FS = "4:45" Then Response.Write(" selected ")%>>4:45 AM</option>
											<option value="5:00"<%If NoActivityNagTimeOfDay_FS = "5:00" Then Response.Write(" selected ")%>>5:00 AM</option>
											<option value="5:15"<%If NoActivityNagTimeOfDay_FS = "5:15" Then Response.Write(" selected ")%>>5:15 AM</option>
											<option value="5:30"<%If NoActivityNagTimeOfDay_FS = "5:30" Then Response.Write(" selected ")%>>5:30 AM</option>
											<option value="5:45"<%If NoActivityNagTimeOfDay_FS = "5:45" Then Response.Write(" selected ")%>>5:45 AM</option>
											<option value="6:00"<%If NoActivityNagTimeOfDay_FS = "6:00" Then Response.Write(" selected ")%>>6:00 AM</option>
											<option value="6:15"<%If NoActivityNagTimeOfDay_FS = "6:15" Then Response.Write(" selected ")%>>6:15 AM</option>
											<option value="6:30"<%If NoActivityNagTimeOfDay_FS = "6:30" Then Response.Write(" selected ")%>>6:30 AM</option>
											<option value="6:45"<%If NoActivityNagTimeOfDay_FS = "6:45" Then Response.Write(" selected ")%>>6:45 AM</option>
											<option value="7:00"<%If NoActivityNagTimeOfDay_FS = "7:00" Then Response.Write(" selected ")%>>7:00 AM</option>
											<option value="7:15"<%If NoActivityNagTimeOfDay_FS = "7:15" Then Response.Write(" selected ")%>>7:15 AM</option>
											<option value="7:30"<%If NoActivityNagTimeOfDay_FS = "7:30" Then Response.Write(" selected ")%>>7:30 AM</option>
											<option value="7:45"<%If NoActivityNagTimeOfDay_FS = "7:45" Then Response.Write(" selected ")%>>7:45 AM</option>
											<option value="8:00"<%If NoActivityNagTimeOfDay_FS = "8:00" Then Response.Write(" selected ")%>>8:00 AM</option>
											<option value="8:15"<%If NoActivityNagTimeOfDay_FS = "8:15" Then Response.Write(" selected ")%>>8:15 AM</option>
											<option value="8:30"<%If NoActivityNagTimeOfDay_FS = "8:30" Then Response.Write(" selected ")%>>8:30 AM</option>
											<option value="8:45"<%If NoActivityNagTimeOfDay_FS = "8:45" Then Response.Write(" selected ")%>>8:45 AM</option>
											<option value="9:00"<%If NoActivityNagTimeOfDay_FS = "9:00" Then Response.Write(" selected ")%>>9:00 AM</option>
											<option value="9:15"<%If NoActivityNagTimeOfDay_FS = "9:15" Then Response.Write(" selected ")%>>9:15 AM</option>
											<option value="9:30"<%If NoActivityNagTimeOfDay_FS = "9:30" Then Response.Write(" selected ")%>>9:30 AM</option>
											<option value="9:45"<%If NoActivityNagTimeOfDay_FS = "9:45" Then Response.Write(" selected ")%>>9:45 AM</option>
											<option value="10:00"<%If NoActivityNagTimeOfDay_FS = "10:00" Then Response.Write(" selected ")%>>10:00 AM</option>
											<option value="10:15"<%If NoActivityNagTimeOfDay_FS = "10:15" Then Response.Write(" selected ")%>>10:15 AM</option>
											<option value="10:30"<%If NoActivityNagTimeOfDay_FS = "10:30" Then Response.Write(" selected ")%>>10:30 AM</option>
											<option value="10:45"<%If NoActivityNagTimeOfDay_FS = "10:45" Then Response.Write(" selected ")%>>10:45 AM</option>
											<option value="11:00"<%If NoActivityNagTimeOfDay_FS = "11:00" Then Response.Write(" selected ")%>>11:00 AM</option>
											<option value="11:15"<%If NoActivityNagTimeOfDay_FS = "11:15" Then Response.Write(" selected ")%>>11:15 AM</option>
											<option value="11:30"<%If NoActivityNagTimeOfDay_FS = "11:30" Then Response.Write(" selected ")%>>11:30 AM</option>
											<option value="11:45"<%If NoActivityNagTimeOfDay_FS = "11:45" Then Response.Write(" selected ")%>>11:45 AM</option>
											<option value="12:00"<%If NoActivityNagTimeOfDay_FS = "12:00" Then Response.Write(" selected ")%>>-Noon-</option>
											<option value="12:15"<%If NoActivityNagTimeOfDay_FS = "12:15" Then Response.Write(" selected ")%>>12:15 PM</option>
											<option value="12:30"<%If NoActivityNagTimeOfDay_FS = "12:30" Then Response.Write(" selected ")%>>12:30 PM</option>
											<option value="12:45"<%If NoActivityNagTimeOfDay_FS = "12:45" Then Response.Write(" selected ")%>>12:45 PM</option>
											<option value="13:00"<%If NoActivityNagTimeOfDay_FS = "13:00" Then Response.Write(" selected ")%>>1:00 PM</option>
											<option value="13:15"<%If NoActivityNagTimeOfDay_FS = "13:15" Then Response.Write(" selected ")%>>1:15 PM</option>
											<option value="13:30"<%If NoActivityNagTimeOfDay_FS = "13:30" Then Response.Write(" selected ")%>>1:30 PM</option>
											<option value="13:45"<%If NoActivityNagTimeOfDay_FS = "13:45" Then Response.Write(" selected ")%>>1:45 PM</option>
											<option value="14:00"<%If NoActivityNagTimeOfDay_FS = "14:00" Then Response.Write(" selected ")%>>2:00 PM</option>
											<option value="14:15"<%If NoActivityNagTimeOfDay_FS = "14:15" Then Response.Write(" selected ")%>>2:15 PM</option>
											<option value="14:30"<%If NoActivityNagTimeOfDay_FS = "14:30" Then Response.Write(" selected ")%>>2:30 PM</option>
											<option value="14:45"<%If NoActivityNagTimeOfDay_FS = "14:45" Then Response.Write(" selected ")%>>2:45 PM</option>
											<option value="15:00"<%If NoActivityNagTimeOfDay_FS = "15:00" Then Response.Write(" selected ")%>>3:00 PM</option>
											<option value="15:15"<%If NoActivityNagTimeOfDay_FS = "15:15" Then Response.Write(" selected ")%>>3:15 PM</option>
											<option value="15:30"<%If NoActivityNagTimeOfDay_FS = "15:30" Then Response.Write(" selected ")%>>3:30 PM</option>
											<option value="15:45"<%If NoActivityNagTimeOfDay_FS = "15:45" Then Response.Write(" selected ")%>>3:45 PM</option>
											<option value="16:00"<%If NoActivityNagTimeOfDay_FS = "16:00" Then Response.Write(" selected ")%>>4:00 PM</option>
											<option value="16:15"<%If NoActivityNagTimeOfDay_FS = "16:15" Then Response.Write(" selected ")%>>4:15 PM</option>
											<option value="16:30"<%If NoActivityNagTimeOfDay_FS = "16:30" Then Response.Write(" selected ")%>>4:30 PM</option>
											<option value="16:45"<%If NoActivityNagTimeOfDay_FS = "16:45" Then Response.Write(" selected ")%>>4:45 PM</option>
											<option value="17:00"<%If NoActivityNagTimeOfDay_FS = "17:00" Then Response.Write(" selected ")%>>5:00 PM</option>
											<option value="17:15"<%If NoActivityNagTimeOfDay_FS = "17:15" Then Response.Write(" selected ")%>>5:15 PM</option>
											<option value="17:30"<%If NoActivityNagTimeOfDay_FS = "17:30" Then Response.Write(" selected ")%>>5:30 PM</option>
											<option value="17:45"<%If NoActivityNagTimeOfDay_FS = "17:45" Then Response.Write(" selected ")%>>5:45 PM</option>
											<option value="18:00"<%If NoActivityNagTimeOfDay_FS = "18:00" Then Response.Write(" selected ")%>>6:00 PM</option>
											<option value="18:15"<%If NoActivityNagTimeOfDay_FS = "18:15" Then Response.Write(" selected ")%>>6:15 PM</option>
											<option value="18:30"<%If NoActivityNagTimeOfDay_FS = "18:30" Then Response.Write(" selected ")%>>6:30 PM</option>
											<option value="18:45"<%If NoActivityNagTimeOfDay_FS = "18:45" Then Response.Write(" selected ")%>>6:45 PM</option>
											<option value="19:00"<%If NoActivityNagTimeOfDay_FS = "19:00" Then Response.Write(" selected ")%>>7:00 PM</option>
											<option value="19:15"<%If NoActivityNagTimeOfDay_FS = "19:15" Then Response.Write(" selected ")%>>7:15 PM</option>
											<option value="19:30"<%If NoActivityNagTimeOfDay_FS = "19:30" Then Response.Write(" selected ")%>>7:30 PM</option>
											<option value="19:45"<%If NoActivityNagTimeOfDay_FS = "19:45" Then Response.Write(" selected ")%>>7:45 PM</option>
											<option value="20:00"<%If NoActivityNagTimeOfDay_FS = "20:00" Then Response.Write(" selected ")%>>8:00 PM</option>
											<option value="20:15"<%If NoActivityNagTimeOfDay_FS = "20:15" Then Response.Write(" selected ")%>>8:15 PM</option>
											<option value="20:30"<%If NoActivityNagTimeOfDay_FS = "20:30" Then Response.Write(" selected ")%>>8:30 PM</option>
											<option value="20:45"<%If NoActivityNagTimeOfDay_FS = "20:45" Then Response.Write(" selected ")%>>8:45 PM</option>
											<option value="21:00"<%If NoActivityNagTimeOfDay_FS = "21:00" Then Response.Write(" selected ")%>>9:00 PM</option>
											<option value="21:15"<%If NoActivityNagTimeOfDay_FS = "21:15" Then Response.Write(" selected ")%>>9:15 PM</option>
											<option value="21:30"<%If NoActivityNagTimeOfDay_FS = "21:30" Then Response.Write(" selected ")%>>9:30 PM</option>
											<option value="21:45"<%If NoActivityNagTimeOfDay_FS = "21:45" Then Response.Write(" selected ")%>>9:45 PM</option>
											<option value="22:00"<%If NoActivityNagTimeOfDay_FS = "22:00" Then Response.Write(" selected ")%>>10:00 PM</option>
											<option value="22:15"<%If NoActivityNagTimeOfDay_FS = "22:15" Then Response.Write(" selected ")%>>10:15 PM</option>
											<option value="22:30"<%If NoActivityNagTimeOfDay_FS = "22:30" Then Response.Write(" selected ")%>>10:30 PM</option>
											<option value="22:45"<%If NoActivityNagTimeOfDay_FS = "22:45" Then Response.Write(" selected ")%>>10:45 PM</option>
											<option value="23:00"<%If NoActivityNagTimeOfDay_FS = "23:00" Then Response.Write(" selected ")%>>11:00 PM</option>
											<option value="23:15"<%If NoActivityNagTimeOfDay_FS = "23:15" Then Response.Write(" selected ")%>>11:15 PM</option>
											<option value="23:30"<%If NoActivityNagTimeOfDay_FS = "23:30" Then Response.Write(" selected ")%>>11:30 PM</option>
											<option value="23:45"<%If NoActivityNagTimeOfDay_FS = "23:45" Then Response.Write(" selected ")%>>11:45 PM</option>	
					 					</select>
									</div>
								</div>
								
								<div class="row">
			                    	<div class="col-lg-12">Send when there has been No Activity for 
										<select class="form-control custom-select" id="selNoActivityNagMinutes_FS" name="selNoActivityNagMinutes_FS">
											<%
												For x = 15 to 180 Step 5 ' 3 hours
													If x mod 60 = 0 Then
														If x = cint(NoActivityNagMinutes_FS) Then 
															Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
														else
															Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
														End If
													Else
														If x = cint(NoActivityNagMinutes_FS) Then 
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
										<select class="form-control custom-select" id="selNoActivityNagIntervalMinutes_FS" name="selNoActivityNagIntervalMinutes_FS">
											<%
												For x = 10 to 120 Step 5 ' 2 hours
													If x mod 60 = 0 Then
														If x = cint(NoActivityNagIntervalMinutes_FS) Then 
															Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
														else
															Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
														End If
													Else
														If x = cint(NoActivityNagIntervalMinutes_FS) Then 
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
										<select class="form-control custom-select" id="selNoActivityNagMessageMaxToSendPerStop_FS" name="selNoActivityNagMessageMaxToSendPerStop_FS">
											<%
												For x = 1 to 10
													If x = cint(NoActivityNagMessageMaxToSendPerStop_FS) Then 
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
										<select class="form-control custom-select"  id="selNoActivityNagMessageMaxToSendPerDriverPerDay_FS" name="selNoActivityNagMessageMaxToSendPerDriverPerDay_FS">
											<%
												For x = 1 to 25
													If x = cint(NoActivityNagMessageMaxToSendPerDriverPerDay_FS) Then 
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
										<select class="form-control custom-select select-large"   id="selNoActivityNagMessageSendMethod_FS" name="selNoActivityNagMessageSendMethod_FS">
											<option value="Text"<%If NoActivityNagMessageSendMethod_FS = "Text" Then Response.Write(" selected ")%>>Text Message Only</option>
										<!--	<option value="Email"<%If NoActivityNagMessageSendMethod_FS = "Email" Then Response.Write(" selected ")%>>Email Only</option>
											<option value="TextThenEmail"<%If NoActivityNagMessageSendMethod_FS = "TextThenEmail" Then Response.Write(" selected ")%>>Text - If no cell number, send email</option>
											<option value="EmailThenText"<%If NoActivityNagMessageSendMethod_FS = "EmailThenText" Then Response.Write(" selected ")%>>Email - If no valid email address, send text</option>
											<option value="Both"<%If NoActivityNagMessageSendMethod_FS = "Both" Then Response.Write(" selected ")%>>Both</option>-->
										</select>
									</div>
								</div>
			
			                </div>
			                <!-- eof no activity box -->
			                
			                </div>
			                <!-- eof nag boxes end here -->
					</div>
				</div>
			</div>
			
			<div class="col-md-4">
				&nbsp;
			</div>
			
		</div>
	
	
	
		<div class="row">
			<h3><i class="fad fa-file-pdf"></i>&nbsp;<%= GetTerm("Field Service") %> Report Settings</h3>
		</div>
	
		<div class="row">
		
			<div class="col-md-4">
				<% If FieldServiceNotesReportOnOff = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title"><%= GetTerm("Field Service") %> Notes Report (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title"><%= GetTerm("Field Service") %> Notes Report (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
					<div class="panel-body">

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON 
					      		<%
					      		If FieldServiceNotesReportOnOff = 0 Then
									Response.Write("<input type='checkbox' id='chkFieldServiceNotesReportOnOff' name='chkFieldServiceNotesReportOnOff'")
								Else
									Response.Write("<input type='checkbox' id='chkFieldServiceNotesReportOnOff' name='chkFieldServiceNotesReportOnOff' checked")
								End If
								Response.Write(">")
								%>
				            </div>
				            <!-- eof line -->
				         </div>  
				         					
					
					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
								
								<div class="text-element circles-list">
									<ol>
										<li>
											<p>Set the report send schedule:</p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalFieldServiceNotesReportScheduler" data-tooltip="true" data-title="<%= GetTerm("Field Service") %> Notes Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> <%= GetTerm("Field Service") %> Notes Report Scheduler</button>
										</li>
										<li>								
											<p>Specify the subject line to be used for the email:</p>
											<input type="text"class="form-control" style="width:100%;" name="txtFieldServiceNotesReportEmailSubject" id="txtFieldServiceNotesReportEmailSubject" value="<%= FieldServiceNotesReportEmailSubject %>">
										</li>
										<li>
											<p>Select users <i class="fad fa-user-friends"></i> to send the report to:</p>
											<input type="hidden" name="lstSelectedFieldServiceNotesReportUserIDs" id="lstSelectedFieldServiceNotesReportUserIDs" value="<%= FieldServiceNotesReportUserNos %>">
											<select id="lstExistingFieldServiceNotesReportUserIDs" multiple="multiple" name="lstExistingFieldServiceNotesReportUserIDs">
												<%	'Get list of all users not currently archived or disabled
													
												Set cnnUserList = Server.CreateObject("ADODB.Connection")
												cnnUserList.open Session("ClientCnnString")
								
												SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 and userEnabled <> 0 ORDER BY userFirstName,userLastName"
												
												Set rsUserList = Server.CreateObject("ADODB.Recordset")
												rsUserList.CursorLocation = 3 
												Set rsUserList = cnnUserList.Execute(SQLUserList)
												
												If Not rsUserList.EOF Then
													Do While Not rsUserList.EOF
													
														FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName") & " (" & rsUserList("userDisplayName") & ")"
														Response.Write("<option value='" & rsUserList("UserNo") & "'>" & FullName & "</option>")
												
														rsUserList.MoveNext
													Loop
												End If
									
												Set rsUserList = Nothing
												cnnUserList.Close
												Set cnnUserList = Nothing
													
												%>
											</select>				
										</li>
										<li>
											<p>Select additional email addresses to send the report to:</p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalNotesReportAdditionalEmails" data-tooltip="true" data-title="Additional emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails</button>						
				             				<% If FieldServiceNotesReportAdditionalEmails <> "" Then %>
				             					<p style="margin-top:20px;"><strong>Current Additional Emails:</strong> <%= FieldServiceNotesReportAdditionalEmails %></p>
				             				<% End If %>
										</li>
									</ol>
								</div>
					
							</div>
						</div>
					
					
					</div>
				</div>
			</div>
			
			
			<div class="col-md-4">
			
				<% If ServiceTicketCarryoverReportOnOff = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title">Service Ticket Carryover Report (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title">Service Ticket Carryover Report (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
				
					<div class="panel-body">
					
						<div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON
					      		<%
					      		If ServiceTicketCarryoverReportOnOff = 0 Then
									Response.Write("<input type='checkbox' id='chkServiceTicketCarryoverReportOnOff' name='chkServiceTicketCarryoverReportOnOff'")
									
								Else
									Response.Write("<input type='checkbox' id='chkServiceTicketCarryoverReportOnOff' name='chkServiceTicketCarryoverReportOnOff' checked")
								End If
								Response.Write(">")
								%>
				            </div>
				            <!-- eof line -->
					    </div>  
	
							
							
					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
								
								<div class="text-element circles-list">
									<ol>
										<li>
											<p>Set the report send schedule:</p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalServiceTicketCarryoverReportScheduler" data-tooltip="true" data-title="Service Ticket Carryover Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> Service Ticket Carryover Report Scheduler</button>
										</li>
										<li>								
											<p>Specify the subject line to be used for the email:</p>
											<input type="text"class="form-control" style="width:100%;" name="txtServiceTicketCarryoverReportEmailSubject" id="txtServiceTicketCarryoverReportEmailSubject" value="<%= ServiceTicketCarryoverReportEmailSubject %>">
										</li>
										<li>
											<p>Select users <i class="fad fa-user-friends"></i> to send an email summary report to:</p>
											<input type="hidden" name="lstSelectedServiceTicketCarryoverReportUserIDs" id="lstSelectedServiceTicketCarryoverReportUserIDs" value="<%= ServiceTicketCarryoverReportUserNos %>">
											<select id="lstExistingServiceTicketCarryoverReportUserIDs" multiple="multiple" name="lstExistingServiceTicketCarryoverReportUserIDs">
												<%	'Get list of all users not currently archived or disabled
													
												Set cnnUserList = Server.CreateObject("ADODB.Connection")
												cnnUserList.open Session("ClientCnnString")
								
												SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 and userEnabled <> 0 ORDER BY userFirstName,userLastName"
												
												Set rsUserList = Server.CreateObject("ADODB.Recordset")
												rsUserList.CursorLocation = 3 
												Set rsUserList = cnnUserList.Execute(SQLUserList)
												
												If Not rsUserList.EOF Then
													Do While Not rsUserList.EOF
													
														FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName") & " (" & rsUserList("userDisplayName") & ")"
														Response.Write("<option value='" & rsUserList("UserNo") & "'>" & FullName & "</option>")
												
														rsUserList.MoveNext
													Loop
												End If
									
												Set rsUserList = Nothing
												cnnUserList.Close
												Set cnnUserList = Nothing
													
												%>
											</select>				
										</li>
										<li>
											<p><strong><i class="fas fa-users-class"></i>&nbsp;Team Reports:</strong></p>
											<p>Includes all accounts owned by any of the team members:</p>
											<input type="hidden" name="lstSelectedServiceTicketCarryoverReportTeamIntRecIDs" id="lstSelectedServiceTicketCarryoverReportTeamIntRecIDs" value="<%= ServiceTicketCarryoverReportTeamIntRecIDs %>">
											<select id="lstExistingServiceTicketCarryoverReportTeamIntRecIDs" multiple="multiple" name="lstExistingServiceTicketCarryoverReportTeamIntRecIDs">
												<%	'Get list of teams users not currently archived or disabled
													
												Set cnnTeamList = Server.CreateObject("ADODB.Connection")
												cnnTeamList.open Session("ClientCnnString")
								
												SQLTeamList = "SELECT * FROM USER_Teams ORDER BY TeamName"
												
												Set rsTeamList = Server.CreateObject("ADODB.Recordset")
												rsTeamList.CursorLocation = 3 
												Set rsTeamList = cnnTeamList.Execute(SQLTeamList)
												
												If Not rsTeamList.EOF Then
													Do While Not rsTeamList.EOF
													
														Response.Write("<option value='" & rsTeamList("InternalRecordIdentifier") & "'>" & rsTeamList("TeamName") & "</option>")
												
														rsTeamList.MoveNext
													Loop
												End If
									
												Set rsTeamList = Nothing
												cnnTeamList.Close
												Set cnnTeamList = Nothing
													
												%>
											</select>											
										</li>
										<% If SalesmanFileExists =  True Then %>
										<li>
											<p><strong><i class="fas fa-user"></i>&nbsp;Individualized Reports:</strong></p>
							               	
							               	<p>
								      			<% If ServiceTicketCarryoverReportToPrimarySalesman = 0 Then %>
								      				<input type="checkbox" id="chkServiceTicketCarryoverReportToPrimarySalesman" name="chkServiceTicketCarryoverReportToPrimarySalesman">
								      			<% Else %>
								      				<input type="checkbox" id="chkServiceTicketCarryoverReportToPrimarySalesman" name="chkServiceTicketCarryoverReportToPrimarySalesman" checked="checked">
								      			<% End If %>
								      			Send Individualized Reports to <%= GetTerm("Primary Salesman") %>
											</p>
							               	<p>
								      			<% If ServiceTicketCarryoverReportToSecondarySalesman = 0 Then %>
								      				<input type="checkbox" id="chkServiceTicketCarryoverReportToSecondarySalesman" name="chkServiceTicketCarryoverReportToSecondarySalesman">
								      			<% Else %>
								      				<input type="checkbox" id="chkServiceTicketCarryoverReportToSecondarySalesman" name="chkServiceTicketCarryoverReportToSecondarySalesman" checked="checked">
								      			<% End If %>
								      			Send Individualized Reports to <%= GetTerm("Secondary Salesman") %>
								      		</p>
										</li>	
										<% End If %>									
										<li>
								      		<%
								      		If ServiceTicketCarryoverReportTextSummaryOnOff = 0 Then
												Response.Write("<input type='checkbox' id='chkServiceTicketCarryoverReportTextSummaryOnOff' name='chkServiceTicketCarryoverReportTextSummaryOnOff'")
											Else
												Response.Write("<input type='checkbox' id='chkServiceTicketCarryoverReportTextSummaryOnOff' name='chkServiceTicketCarryoverReportTextSummaryOnOff' checked")
											End If
											Response.Write(">")
											%>
											Send Text Summary <i class="fas fa-comment-alt-lines"></i> of Report To Selected Users <br><br>
											
											<p>Select users <i class="fad fa-user-friends"></i> to send a text message summary to:</p>
											<input type="hidden" name="lstSelectedServiceTicketCarryoverReportTextSummmaryUserIDs" id="lstSelectedServiceTicketCarryoverReportTextSummmaryUserIDs" value="<%= ServiceTicketCarryoverReportTextSummaryUserNos %>">
											<select id="lstExistingServiceTicketCarryoverReportTextSummmaryUserIDs" multiple="multiple" name="lstExistingServiceTicketCarryoverReportTextSummmaryUserIDs">
												<%	'Get list of all users not currently archived or disabled
													
												Set cnnUserList = Server.CreateObject("ADODB.Connection")
												cnnUserList.open Session("ClientCnnString")
								
												SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 and userEnabled <> 0 ORDER BY userFirstName,userLastName"
												
												Set rsUserList = Server.CreateObject("ADODB.Recordset")
												rsUserList.CursorLocation = 3 
												Set rsUserList = cnnUserList.Execute(SQLUserList)
												
												If Not rsUserList.EOF Then
													Do While Not rsUserList.EOF
													
														FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName") & " (" & rsUserList("userDisplayName") & ")"
														Response.Write("<option value='" & rsUserList("UserNo") & "'>" & FullName & "</option>")
												
														rsUserList.MoveNext
													Loop
												End If
									
												Set rsUserList = Nothing
												cnnUserList.Close
												Set cnnUserList = Nothing
													
												%>
											</select>
										</li>										
										<li>
											<p>Select additional email addresses to send the report to:</p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalCarryoverReportAdditionalEmails" data-tooltip="true" data-title="Additional emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails</button>						
				             				<% If ServiceTicketCarryoverReportAdditionalEmails <> "" Then %>
				             					<p style="margin-top:20px;"><strong>Current Additional Emails:</strong> <%= ServiceTicketCarryoverReportAdditionalEmails %></p>
				             				<% End If %>
										</li>
										<li>
											<p>Other Special Carryover Report Settings:</p>
											
											<p>
								      			<% If CarryoverReportInclTicketNum = 0 Then %>
								      				<input type="checkbox" id="chkCarryoverReportInclTicketNum" name="chkCarryoverReportInclTicketNum">
								      			<% Else %>
								      				<input type="checkbox" id="chkCarryoverReportInclTicketNum" name="chkCarryoverReportInclTicketNum" checked="checked">
								      			<% End If %>
								      			Include ticket number on report
											</p>
					               			<p>
								      			<% If CarryoverReportShowRedoBreakdown = 0 Then %>
								      				<input type="checkbox" id="chkCarryoverReportShowRedoBreakdown" name="chkCarryoverReportShowRedoBreakdown">
								      			<% Else %>
								      				<input type="checkbox" id="chkCarryoverReportShowRedoBreakdown" name="chkCarryoverReportShowRedoBreakdown" checked="checked">
								      			<% End If %>
								      			Show Redo Breakdown 
					               			</p>
					               			<p>
								      			<% If CarryoverReportInclCustType = 0 Then %>
								      				<input type="checkbox" id="chkCarryoverReportInclCustType" name="chkCarryoverReportInclCustType">
								      			<% Else %>
								      				<input type="checkbox" id="chkCarryoverReportInclCustType" name="chkCarryoverReportInclCustType" checked="checked">
								      			<% End If %>
								      			Include customer type on report 
					               			</p>
				               				<p>
								      			<% If ServiceTicketCarryoverReportIncludeRegions = 0 Then %>
								      				<input type="checkbox" id="chkCarryoverReportIncludeRegions" name="chkCarryoverReportIncludeRegions">
								      			<% Else %>
								      				<input type="checkbox" id="chkCarryoverReportIncludeRegions" name="chkCarryoverReportIncludeRegions" checked="checked">
								      			<% End If %>
								      			Show Breakdown By Region On Report 
					               			</p>
					               			
										</li>
									</ol>
								</div>
					
							</div>
						</div>
					
					
					</div>
				</div>
			</div>
			
			<div class="col-md-4">
			
				<% If ServiceTicketThresholdReportOnOff = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title">Service Ticket Threshold Report (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title">Service Ticket Threshold Report (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
			
					<div class="panel-body">
					
							
					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON 
					      		<%
					      		If ServiceTicketThresholdReportOnOff = 0 Then
									Response.Write("<input type='checkbox' id='chkServiceTicketThresholdReportOnOff' name='chkServiceTicketThresholdReportOnOff'")
									
								Else
									Response.Write("<input type='checkbox' id='chkServiceTicketThresholdReportOnOff' name='chkServiceTicketThresholdReportOnOff' checked")
								End If
								Response.Write(">")
								%>
				            </div>
				            <!-- eof line -->
				         </div>  
					         
						
					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
								
								<div class="text-element circles-list">
									<ol>
										<li>
											<p>Set the report send schedule:</p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalServiceTicketThresholdReportScheduler" data-tooltip="true" data-title="Service Ticket Threshold Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> Service Ticket Threshold Report Scheduler</button>
										</li>
										<li>
											<p>Select users <i class="fad fa-user-friends"></i> to send the report to:</p>
											<input type="hidden" name="lstSelectedServiceTicketThresholdReportUserNos" id="lstSelectedServiceTicketThresholdReportUserNos" value="<%= ServiceTicketThresholdReportUserNos %>">
											<select id="lstExistingServiceTicketThresholdReportUserNos" multiple="multiple" name="lstExistingServiceTicketThresholdReportUserNos">
												<%	'Get list of all users not currently archived or disabled
													
												Set cnnUserList = Server.CreateObject("ADODB.Connection")
												cnnUserList.open Session("ClientCnnString")
								
												SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 and userEnabled <> 0 ORDER BY userFirstName,userLastName"
												
												Set rsUserList = Server.CreateObject("ADODB.Recordset")
												rsUserList.CursorLocation = 3 
												Set rsUserList = cnnUserList.Execute(SQLUserList)
												
												If Not rsUserList.EOF Then
													Do While Not rsUserList.EOF
													
														FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName") & " (" & rsUserList("userDisplayName") & ")"
														Response.Write("<option value='" & rsUserList("UserNo") & "'>" & FullName & "</option>")
												
														rsUserList.MoveNext
													Loop
												End If
									
												Set rsUserList = Nothing
												cnnUserList.Close
												Set cnnUserList = Nothing
													
												%>
											</select>				
										</li>
										<li>
											<p>Select additional email addresses to send the report to:</p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalThresholdReportAdditionalEmails" data-tooltip="true" data-title="Additional emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails</button>						
											<% If ServiceTicketThresholdReportAdditionalEmails <> "" Then %>
												<p style="margin-top:20px;"><strong>Current Additional Emails:</strong> <%= ServiceTicketThresholdReportAdditionalEmails %></p>
											<% End If %>
										</li>
											<li>
											<p>Other Special Threshold Report Settings:</p>
											
											<p>
								      			<% If ServiceTicketThresholdReportOnlyUndispatched = 0 Then %>
								      				<input type="checkbox" id="chkServiceTicketThresholdReportOnlyUndispatched" name="chkServiceTicketThresholdReportOnlyUndispatched">
								      			<% Else %>
								      				<input type="checkbox" id="chkServiceTicketThresholdReportOnlyUndispatched" name="chkServiceTicketThresholdReportOnlyUndispatched" checked="checked">
								      			<% End If %>
								      			Only Show Awaiting Dispatch
											</p>
											<% If filterChangeModuleOn Then %>
						               			<p>
									      			<% If ServiceTicketThresholdReportOnlySkipFilterChanges = 0 Then %>
									      				<input type="checkbox" id="chkServiceTicketThresholdReportOnlySkipFilterChanges" name="chkServiceTicketThresholdReportOnlySkipFilterChanges">
									      			<% Else %>
									      				<input type="checkbox" id="chkServiceTicketThresholdReportOnlySkipFilterChanges" name="chkServiceTicketThresholdReportOnlySkipFilterChanges" checked="checked">
									      			<% End If %>
									      			Don't Include Filter Changes 
						               			</p>
						               		<% End If %>
					               			<p>
												Include Elapsed Time Over <select class="form-control custom-select" id="selServiceTicketThresholdReportThresholdHours" name="selServiceTicketThresholdReportThresholdHours">
													<%
														For x = 1 to 99
															If Not IsNumeric(ServiceTicketThresholdReportThresholdHours) Then ServiceTicketThresholdReportThresholdHours = 0
															If x = cint(ServiceTicketThresholdReportThresholdHours) Then 
																Response.Write("<option value='" & x & "' selected>" & x & "</option>")
															Else
																Response.Write("<option value='" & x & "'>" & x & "</option>")
															End If
														Next
													%>
												</select>
								      			Hours 
					               			</p>
										</li>
					         		</ol>	
					         	</div>				
						</div>
					</div>
				</div>
				
			</div>
			
			
		</div>
	
	
	
		<div class="row">
		
			<!-- cancel / save !-->
			<div class="row pull-right">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>admin/global/main.asp"><button type="button" class="btn btn-default btn-lg btn-huge"><i class="far fa-times-circle"></i> Cancel</button></a> 
					<button type="submit" class="btn btn-primary btn-lg btn-huge" onclick="showSavingChangesDiv()"><i class="far fa-save"></i> Save Changes</button>
				</div>
			</div>
			
			
		</div>
	
	</form>
	
	
	
<% End If %>
	 </div> <!-- container -->
<%
Function UserInList(UserToFind,UserList)

	result = False
	
	If len(UserList) > 1 Then
		UserNoList = Split(UserList,",")
		For x = 0 To UBound(UserNoList)
			If cint(UserToFind) = cint(UserNoList(x)) Then
				result = True
				Exit For
			End If
		Next
	End If
	UserInList = result
	
End Function
%>

<!--#include file="field-service-modals.asp"-->
<!--#include file="field-service-color-pickers.asp"-->

<!--#include file="../../../inc/footer-main.asp"-->

