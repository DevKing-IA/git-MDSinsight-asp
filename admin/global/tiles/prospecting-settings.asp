<!--#include file="../../../inc/header.asp"-->
<!--#include file="../../../inc/InsightFuncs.asp"-->

<!-- bootstrap timepicker !-->
<script type="text/javascript" src="http://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>	
<link href="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.css" rel="stylesheet" type="text/css">
<script src="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.min.js" type="text/javascript"></script>
<!-- eof bootstrap timepicker !-->

<!-- spectrum color picker !-->
<script src="<%= BaseURL %>/js/spectrum-color-picker/spectrum.js"></script>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>/js/spectrum-color-picker/spectrum.css">
<!-- eof spectrum color picker !-->

<!-- bootstrap multiselect !-->
<link rel="stylesheet" href="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.css" type="text/css">
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.js"></script>
<!-- eof bootstrap multiselect !-->

<%
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
		CRMTabAuditTrailColor = rs("CRMTabAuditTrailColor")
		CRMTileOfferingColor = rs("CRMTileOfferingColor")
		CRMTileCompetitorColor = rs("CRMTileCompetitorColor")				
		CRMTileDollarsColor = rs("CRMTileDollarsColor")		
		CRMTileActivityColor = rs("CRMTileActivityColor")	
		CRMTileStageColor = rs("CRMTileStageColor")	
		CRMTileOwnerColor = rs("CRMTileOwnerColor")
		CRMTileCommentsColor= rs("CRMTileCommentsColor")
		CRMMaxActivityDaysPermitted = rs("CRMMaxActivityDaysPermitted")			
		CRMMaxActivityDaysWarning = rs("CRMMaxActivityDaysWarning")
		EWSDefaultApptDuration = rs("EWSDefaultApptDuration")
		EWSDefaultMeetingDuration = rs("EWSDefaultMeetingDuration")
		EWSPostURL = rs("EWSPostURL")
		CRMAutoCoordinateColors = rs("CRMAutoCoordinateColors")
		CRMHideLocationTab = rs("CRMHideLocationTab")
		CRMHideProductsTab = rs("CRMHideProductsTab")
		CRMHideEquipmentTab = rs("CRMHideEquipmentTab")
		ProspSnapshotOnOff = rs("ProspSnapshotOnOff")			
		ProspSnapshotInsideSales = rs("ProspSnapshotInsideSales")	
		ProspSnapshotOutsideSales = rs("ProspSnapshotOutsideSales")	
		ProspSnapshotAdditionalEmails = rs("ProspSnapshotAdditionalEmails")	
		ProspSnapshotEmailSubject = rs("ProspSnapshotEmailSubject")	
		ProspSnapshotUserNos = rs("ProspSnapshotUserNos")	
		ProspSnapshotSalesRepDisplayUserNos = rs("ProspSnapshotSalesRepDisplayUserNos")		
		
	End If
			
	SQL = "SELECT * FROM Settings_Prospecting"
	
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		ShowLivePoolProspectSearchBox= rs("ShowLivePoolProspectSearchBox")
		ProspectActivityDefaultDaysToShow = rs("ProspectActivityDefaultDaysToShow")
		TabSocialMediaColor = rs("TabSocialMediaColor")
		ProspectingWeeklyAgendaReportOnOff = rs("ProspectingWeeklyAgendaReportOnOff")
		ProspectingWeeklyAgendaReportUserNos = rs("ProspectingWeeklyAgendaReportUserNos")
		ProspectingWeeklyAgendaReportEmailSubject = rs("ProspectingWeeklyAgendaReportEmailSubject")
		ProspectingWeeklyAgendaReportAdditionalEmails = rs("ProspectingWeeklyAgendaReportAdditionalEmails")	
	End If
			
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

%>



<style type="text/css">

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
	    padding-top:20px;
	    padding-bottom:20px;
	}		
</style>


<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;<%= GetTerm("Prospecting") %> Settings 
	<button id="toggle" class="btn btn-small btn-success"><i class="fas fa-arrows-v"></i>&nbsp;EXPAND/COLLAPSE ALL SETTINGS</button>
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>

<form method="post" action="prospecting-settings-submit.asp" name="frmProspectingSettings" id="frmProspectingSettings">

<div class="container">

<%
	Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
	Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
	Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
	Response.Write("</div>")
	Response.Flush()
%>


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
        

		$('#modalProspectingSnapshotReportScheduler').on('show.bs.modal', function(e) {
		    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
				cache: false,
				data: "action=GetContentForProspectingSnapshotReportScheduler",
				success: function(response)
				 {
	               	 $modal.find('#modalProspectingSnapshotReportSchedulerContent').html(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalProspectingSnapshotReportSchedulerContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
			
		});		
		
        $('#modalProspSnapshotAdditionalEmails').on('show.bs.modal', function (e) {
            var $modal = $(this);
        });
		        

		$('#modalProspectingWeeklyAgendaReportScheduler').on('show.bs.modal', function(e) {
		    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
				cache: false,
				data: "action=GetContentForProspectingWeeklyAgendaReportScheduler",
				success: function(response)
				 {
	               	 $modal.find('#modalProspectingWeeklyAgendaReportSchedulerContent').html(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalProspectingWeeklyAgendaReportSchedulerContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
			
		});		
		
        $('#modalProspectingWeeklyAgendaReportAdditionalEmails').on('show.bs.modal', function (e) {
            var $modal = $(this);
        });
		        

           
		$('#lstExistingProspectingSnapshotReportUserIDs').multiselect({
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
			nonSelectedText:'No Users Selected For Snapshot Report',
			numberDisplayed: 20,
		    onChange: function() {
		        var selected = this.$select.val();
		        $("#lstSelectedProspectingSnapshotReportUserIDs").val(selected);
		        console.log(selected);
		        // ...
		    }
    			
	    });	
	    
		//*************************************************************************************************
		//Load the bootstrap multiselect box with the current snapshot report users preselected
		//*************************************************************************************************
		var data= $("#lstSelectedProspectingSnapshotReportUserIDs").val();
		
		if (data) {
			//Make an array
			var dataarray=data.split(",");
			// Set the value
			$("#lstExistingProspectingSnapshotReportUserIDs").val(dataarray);
			// Then refresh
			$("#lstExistingProspectingSnapshotReportUserIDs").multiselect("refresh");
		}
		//*************************************************************************************************


           
		$('#lstExistingProspectingSnapshotReportSalesRepUserIDs').multiselect({
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
			nonSelectedText:'No Sales Reps Selected To Display In Report',
			numberDisplayed: 20,
		    onChange: function() {
		        var selected = this.$select.val();
		        $("#lstSelectedProspectingSnapshotReportSalesRepUserIDs").val(selected);
		        console.log(selected);
		        // ...
		    }
    			
	    });	
	    
		//*************************************************************************************************
		//Load the bootstrap multiselect box with the current sales reps to display preselected
		//*************************************************************************************************
		var data= $("#lstSelectedProspectingSnapshotReportSalesRepUserIDs").val();
		
		if (data) {
			//Make an array
			var dataarray=data.split(",");
			// Set the value
			$("#lstExistingProspectingSnapshotReportSalesRepUserIDs").val(dataarray);
			// Then refresh
			$("#lstExistingProspectingSnapshotReportSalesRepUserIDs").multiselect("refresh");
		}
		//*************************************************************************************************
		
		

           
		$('#lstExistingProspectingWeeklyAgendaReportUserIDs').multiselect({
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
			nonSelectedText:'No Sales Reps Selected To Display In Report',
			numberDisplayed: 20,
		    onChange: function() {
		        var selected = this.$select.val();
		        $("#lstSelectedProspectingWeeklyAgendaReportUserIDs").val(selected);
		        console.log(selected);
		        // ...
		    }
    			
	    });	
	    
		//*************************************************************************************************
		//Load the bootstrap multiselect box with the current sales reps to display preselected
		//*************************************************************************************************
		var data= $("#lstSelectedProspectingWeeklyAgendaReportUserIDs").val();
		
		if (data) {
			//Make an array
			var dataarray=data.split(",");
			// Set the value
			$("#lstExistingProspectingWeeklyAgendaReportUserIDs").val(dataarray);
			// Then refresh
			$("#lstExistingProspectingWeeklyAgendaReportUserIDs").multiselect("refresh");
		}
		//*************************************************************************************************
		
		
		
		        
	});
</script>


<% If MUV_READ("prospectingModuleOn") = "Disabled" Then %>
	<div class="col-lg-6">
		<br><br>
		Please contact support if you would like to activate the <%= GetTerm("Prospecting") %> module.
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
			<h3><i class="fad fa-sliders-h"></i>&nbsp;<%= GetTerm("Prospecting") %> General Settings</h3>
		</div>
	
	    <div class="row">
	    
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title"><%= GetTerm("Prospect") %> Activity Default Days To Show</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
		              		<!-- line !-->
		                	<div class="row schedule-info">
	               				<div class="col-lg-6"><strong>By default, show prospects that have activity </strong></div>
	                			<div class="col-lg-3">
									<select class="form-control" id="selProspectActivityDefaultDaysToShow" name="selProspectActivityDefaultDaysToShow">
										<%For x = 0 to 365
											If x = ProspectActivityDefaultDaysToShow Then
												Response.Write("<option selected >" & x & "</option>")
											Else
												Response.Write("<option>" & x & "</option>")												
											End If
										Next %>
									 </select>
								</div>
								<div class="col-lg-3">
	                				<strong>days and older</strong>
		                			</div>
			                 </div>
		                 	<!-- eof line !-->
					
					</div>
				</div>
			</div>
			
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">Next Activity Scheduling</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
		              
							<p>To turn off either of the functions below, select 0 as the number of days</p>		                
		              		<!-- line !-->
		                	<div class="row schedule-info">
		               			<div class="col-lg-6"><strong>Show a warning if an activity is scheduled more than</strong></div>
		                			<div class="col-lg-3">
										<select class="form-control" id="selCRMMaxActivityDaysWarning" name="selCRMMaxActivityDaysWarning">
											<%For x = 0 to 150
												If x = CRMMaxActivityDaysWarning Then
													Response.Write("<option selected >" & x & "</option>")
												Else
													Response.Write("<option>" & x & "</option>")												
												End If
											Next %>
										 </select>
									</div>
									
									<div class="col-lg-3">
		                				<strong>days in advance</strong>
		                			</div>
								</div>
		                 		<!-- eof line !-->
                 
								<!-- line !-->
								<div class="row schedule-info">
									<div class="col-lg-6"><strong>Don't allow an activity to be scheduled more than</strong></div>
								
									<div class="col-lg-3">
										<select class="form-control" id="selCRMMaxActivityDaysPermitted" name="selCRMMaxActivityDaysPermitted">
											<%For x = 0 to 365
												If x = CRMMaxActivityDaysPermitted Then
													Response.Write("<option selected >" & x & "</option>")
												Else
													Response.Write("<option>" & x & "</option>")												
												End If
											Next %>
										 </select>
									</div>
									<div class="col-lg-3">
										<strong>days in advance</strong>
									</div>
			                 	</div>
		                 		<!-- eof line !-->
					
					</div>
				</div>
			</div>


			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">Show/Hide Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
			
							<p>To hide tabs from displaying in the Prospect Detail, select show/hide below</p>		                
	              			<!-- line !-->
	                		<div class="row schedule-info">
	               				<div class="col-lg-6"><strong><%= GetTerm("Location") %> Tab</strong></div>
	                			<div class="col-lg-4">
									<select class="form-control" id="selCRMShowHideLocationTab" name="selCRMShowHideLocationTab">
										<%
											If CRMHideLocationTab = 0 Then
												Response.Write("<option value=0 selected>show</option>")
												Response.Write("<option value=1>hide</option>")
											ElseIf CRMHideLocationTab = 1 Then
												Response.Write("<option value=0>show</option>")
												Response.Write("<option value=1 selected>hide</option>")												
											End If
										%>
									 </select>
								</div>
							</div>
	                 		<!-- eof line !-->
	                 		
	                 		<div class="row schedule-info">
	                 		<div class="col-lg-6"><strong><%= GetTerm("Products") %> Tab</strong></div>
	                			<div class="col-lg-4">
									<select class="form-control" id="selCRMShowHideProductsTab" name="selCRMShowHideProductsTab">
										<%
											If CRMHideProductsTab = 0 Then
												Response.Write("<option value=0 selected>show</option>")
												Response.Write("<option value=1>hide</option>")
											ElseIf CRMHideProductsTab = 1 Then
												Response.Write("<option value=0>show</option>")
												Response.Write("<option value=1 selected>hide</option>")												
											End If
										%>
									 </select>
									</div>
								</div>
	                 		
	                 		
	                 		<div class="row schedule-info">
	               				<div class="col-lg-6"><strong><%= GetTerm("Equipment") %> Tab</strong></div>
	                			<div class="col-lg-4">
									<select class="form-control" id="selCRMShowHideEquipmentTab" name="selCRMShowHideEquipmentTab">
										<%
											If CRMHideEquipmentTab = 0 Then
												Response.Write("<option value=0 selected>show</option>")
												Response.Write("<option value=1>hide</option>")
											ElseIf CRMHideEquipmentTab = 1 Then
												Response.Write("<option value=0>show</option>")
												Response.Write("<option value=1 selected>hide</option>")												
											End If
										%>
									 </select>
								</div>
							</div>
	                 		<!-- eof line !-->
	                 		
	                 		<div class="row schedule-info">
	               				<div class="col-lg-6"><strong>(Live Pool) Show <u>all prospects</u> search box</strong></div>
	                			<div class="col-lg-4">
							      		<%
							      		If ShowLivePoolProspectSearchBox = 0 Then
											Response.Write("<input type='checkbox' id='chkShowLivePoolProspectSearchBox' name='chkShowLivePoolProspectSearchBox'")
											
										Else
											Response.Write("<input type='checkbox' id='chkShowLivePoolProspectSearchBox' name='chkShowLivePoolProspectSearchBox' checked")
										End If
										Response.Write(">")
										%>
									</div>
								</div>
							</div>
	                 		<!-- eof line !-->

				</div>
			</div>
			
		</div><!-- end first row of panels -->
		
		<!-- start second row of panels -->
	    <div class="row">

			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">Exchange Email &amp; Calendar Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
			
    
	              		<!-- line !-->
	                	<div class="row schedule-info">
	               			<div class="col-lg-10"><strong>EWS POST URL (Exchange)</strong></div>
						</div>
	                 	<!-- eof line !-->
	              		<!-- line !-->
	                	<div class="row schedule-info">
	               			<div class="col-lg-12"><input type="text"class="form-control" style="width:100%;" name="txtEWSPostURL" id="txtEWSPostURL" value="<%= EWSPostURL %>"></div>
						</div>
	                 		
	              		<!-- line !-->
	                	<div class="row schedule-info">
	               			<div class="col-lg-6"><strong>Default appointment duration in minutes</strong></div>
	                			<div class="col-lg-4">
									<select class="form-control" style="width:100%;" id="selEWSDefaultApptDuration" name="selEWSDefaultApptDuration">
										<%For x = 15 to 180 Step 5
											If x mod 60 = 0 Then
												If x = cint(EWSDefaultApptDuration) Then 
													Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
												else
													Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
												End If
											Else
												If x = cint(EWSDefaultApptDuration) Then 
													Response.Write("<option value='" & x & "' selected>" & x & "</option>")
												Else
													Response.Write("<option value='" & x & "'>" & x & "</option>")
												End If
											End If
										Next %>
									 </select>
								</div>
						</div>
                 		<!-- eof line !-->


	              		<!-- line !-->
	                	<div class="row schedule-info">
	               			<div class="col-lg-6"><strong>Default meeting duration in minutes</strong></div>
	                			<div class="col-lg-4">
									<select class="form-control" style="width:100%;" id="selEWSDefaultMeetingDuration" name="selEWSDefaultMeetingDuration">
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
                 		<!-- eof line !-->
					
					</div>
				</div>
			</div>
						
		</div>
		
	
		<div class="row">
			<h3><i class="fad fa-palette"></i>&nbsp;<%= GetTerm("Prospecting") %> Colors &amp; Settings</h3>
		</div>
		
		<div class="row">
		
			<div class="col-md-4">
				<div class="panel panel-info">
					<div class="panel-heading">
						<h3 class="panel-title">Tab Colors</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
		 	         	<div class="row">
			         	   	<div class="col-lg-9"><strong>Automatically link Tab and Tile Colors</strong><br>(If selected, Tab colors will dictate Tile colors)</div>
					       	<div class="col-lg-3">
					      		<%
					      		If CRMAutoCoordinateColors = 0 Then
									Response.Write("<input type='checkbox' id='chkCRMAutoCoordinateColors' name='chkCRMAutoCoordinateColors'")
									
								Else
									Response.Write("<input type='checkbox' id='chkCRMAutoCoordinateColors' name='chkCRMAutoCoordinateColors' checked")
								End If
								Response.Write(">")
								%>
							</div>
						</div>
		           
		            
		 	         	<div class="row">
			         	   	<div class="col-lg-9">Background color for <%= GetTerm("Journal") %> tab </div>
					       	<div class="col-lg-3">
								<input type="text" id="txtCRMTabLogColor" name="txtCRMTabLogColor"  value="<%= CRMTabLogColor %>">
							</div>
						</div>
			         	<div class="row">
			         	   	<div class="col-lg-9">Background color for <%= GetTerm("Products") %> tab </div>
					       	<div class="col-lg-3">
								<input type="text" id="txtCRMTabProductsColor" name="txtCRMTabProductsColor"  value="<%= CRMTabProductsColor %>">
							</div>
						</div>
			         	<div class="row">
			         	   	<div class="col-lg-9">Background color for <%= GetTerm("Equipment") %> tab </div>
					       	<div class="col-lg-3">
								<input type="text" id="txtCRMTabEquipmentColor" name="txtCRMTabEquipmentColor"  value="<%= CRMTabEquipmentColor %>">
							</div>
						</div>
			         	<div class="row">
			         	   	<div class="col-lg-9">Background color for <%= GetTerm("Documents") %> tab </div>
					       	<div class="col-lg-3">
								<input type="text" id="txtCRMTabDocumentsColor" name="txtCRMTabDocumentsColor"  value="<%= CRMTabDocumentsColor %>">
							</div>
						</div>	
			         	<div class="row">
			         	   	<div class="col-lg-9">Background color for <%= GetTerm("Contacts") %> tab </div>
					       	<div class="col-lg-3">
								<input type="text" id="txtCRMTabContactsColor" name="txtCRMTabContactsColor"  value="<%= CRMTabContactsColor %>">
							</div>
						</div>
			         	<div class="row">
			         	   	<div class="col-lg-9">Background color for <%= GetTerm("Competitors") %> tab </div>
					       	<div class="col-lg-3">
								<input type="text" id="txtCRMTabCompetitorsColor" name="txtCRMTabCompetitorsColor"  value="<%= CRMTabCompetitorsColor %>">
							</div>
						</div>
						<!--
			         	<div class="row">
			         	   	<div class="col-lg-9">Background color for opportunities tab </div>
					       	<div class="col-lg-3">
								<input type="text" id="txtCRMTabOpportunityColor" name="txtCRMTabOpportunityColor"  value="<%= CRMTabOpportunityColor %>">
							</div>
						</div>-->
						<div class="row">
			         	   	<div class="col-lg-9">Background color for <%= GetTerm("Location") %> tab </div>
					       	<div class="col-lg-3">
								<input type="text" id="txtCRMTabLocationColor" name="txtCRMTabLocationColor"  value="<%= CRMTabLocationColor %>">
							</div>
						</div>	
			         	<div class="row">
			         	   	<div class="col-lg-9">Background color for <%= GetTerm("Audit Trail") %> tab </div>
					       	<div class="col-lg-3">
								<input type="text" id="txtCRMTabAuditTrailColor" name="txtCRMTabAuditTrailColor"  value="<%= CRMTabAuditTrailColor %>">
							</div>
						</div>
		   	         	<div class="row">
			         	   	<div class="col-lg-9">Background color for Social Media tab </div>
					       	<div class="col-lg-3">
								<input type="text" id="txtTabSocialMediaColor" name="txtTabSocialMediaColor"  value="<%= TabSocialMediaColor %>">
							</div>
						</div>
					
					</div>
				</div>
			</div>
			
			<div class="col-md-4">
				<div class="panel panel-info">
					<div class="panel-heading">
						<h3 class="panel-title">Tile Colors</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					

			 	         	<div class="row">
				         	   	<div class="col-lg-9"><%= GetTerm("Owner") %> tile color</div>
						       	<div class="col-lg-3">
									<input type="text" id="txtCRMTileOwnerColor" name="txtCRMTileOwnerColor"  value="<%= CRMTileOwnerColor %>">
								</div>
							</div>
 
			 	         	<div class="row">
				         	   	<div class="col-lg-9"><%= GetTerm("Comments") %> tile color</div>
						       	<div class="col-lg-3">
									<input type="text" id="txtCRMTileCommentsColor" name="txtCRMTileCommentsColor"  value="<%= CRMTileCommentsColor %>">
								</div>
							</div>
         
			 	         	<div class="row">
				         	   	<div class="col-lg-9"><%= GetTerm("Current Supplier Info") %> tile color</div>
						       	<div class="col-lg-3">
									<input type="text" id="txtCRMTileOfferingColor" name="txtCRMTileOfferingColor"  value="<%= CRMTileOfferingColor%>">
								</div>
							</div>
				
				
				         	<div class="row">
				         	   	<div class="col-lg-9"><%= GetTerm("Primary Competitor") %> tile color </div>
						       	<div class="col-lg-3">
									<input type="text" id="txtCRMTileCompetitorColor" name="txtCRMTileCompetitorColor"  value="<%= CRMTileCompetitorColor%>">
								</div>
							</div>
				
				         	<div class="row">
				         	   	<div class="col-lg-9"><%= GetTerm("Opportunity") %> tile color </div>
						       	<div class="col-lg-3">
									<input type="text" id="txtCRMTileDollarsColor" name="txtCRMTileDollarsColor"  value="<%= CRMTileDollarsColor%>">
								</div>
							</div>
							
				         	<div class="row">
				         	   	<div class="col-lg-9"><%= GetTerm("Next Activity") %> tile color</div>
						       	<div class="col-lg-3">
									<input type="text" id="txtCRMTileActivityColor" name="txtCRMTileActivityColor"  value="<%= CRMTileActivityColor%>">
								</div>
							</div>
							
					        <div class="row">
				         	   	<div class="col-lg-9"><%= GetTerm("Stage") %> tile color </div>
						       	<div class="col-lg-3">
									<input type="text" id="txtCRMTileStageColor" name="txtCRMTileStageColor"  value="<%= CRMTileStageColor%>">
								</div>
							</div>		
					</div>
				</div>
			</div>

			
			<div class="col-md-4">
				&nbsp;
			</div>	
					
		</div>
		
	
		<div class="row">
			<h3><i class="fad fa-file-pdf"></i>&nbsp;<%= GetTerm("Prospecting") %> Report Settings</h3>
		</div>
	
		<div class="row">
		
			
		
			<div class="col-md-4">
				<% If ProspSnapshotOnOff = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title"><%= GetTerm("Prospecting") %> Weekly Snapshot Report (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title"><%= GetTerm("Prospecting") %> Weekly Snapshot Report (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
					<div class="panel-body">

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON 
					      		<%
					      		If ProspSnapshotOnOff = 0 Then
									Response.Write("<input type='checkbox' id='chkProspSnapshotOnOff' name='chkProspSnapshotOnOff'")
								Else
									Response.Write("<input type='checkbox' id='chkProspSnapshotOnOff' name='chkProspSnapshotOnOff' checked")
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalProspectingSnapshotReportScheduler" data-tooltip="true" data-title="<%= GetTerm("Prospecting") %> Weekly Snapshot Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> <%= GetTerm("Prospecting") %> Weekly Snapshot Report Scheduler</button>
										</li>
										<li>								
											<p>Specify the subject line to be used for the email:</p>
											<input type="text"class="form-control" style="width:100%;" name="txtProspSnapshotEmailSubject" id="txtProspSnapshotEmailSubject" value="<%= ProspSnapshotEmailSubject %>">
										</li>
										<li>
											<p>Select Sales Reps <i class="fas fa-user-tie"></i> To Display In Report:</p>
											<input type="hidden" name="lstSelectedProspectingSnapshotReportSalesRepUserIDs" id="lstSelectedProspectingSnapshotReportSalesRepUserIDs" value="<%= ProspSnapshotSalesRepDisplayUserNos %>">
											<select id="lstExistingProspectingSnapshotReportSalesRepUserIDs" multiple="multiple" name="lstExistingProspectingSnapshotReportSalesRepUserIDs">
												<%	'Get list of all users not currently archived or disabled
													
												Set cnnUserList = Server.CreateObject("ADODB.Connection")
												cnnUserList.open Session("ClientCnnString")
								
												SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 AND userEnabled <> 0 ORDER BY userFirstName,userLastName"
												
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
											<p>Select users <i class="fad fa-user-friends"></i> to send the report to:</p>
											<input type="hidden" name="lstSelectedProspectingSnapshotReportUserIDs" id="lstSelectedProspectingSnapshotReportUserIDs" value="<%= ProspSnapshotUserNos %>">
											<select id="lstExistingProspectingSnapshotReportUserIDs" multiple="multiple" name="lstExistingProspectingSnapshotReportUserIDs">
												<%	'Get list of all users not currently archived or disabled
													
												Set cnnUserList = Server.CreateObject("ADODB.Connection")
												cnnUserList.open Session("ClientCnnString")
								
												SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 AND userEnabled <> 0 ORDER BY userFirstName,userLastName"
												
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
										
			                            <!-- line -->
			                            <li>
			                            	Send report to all Inside Sales managers
			      				      		<%
								      		If ProspSnapshotInsideSales = 0 Then
												Response.Write("<input type='checkbox' id='chkProspSnapshotInsideSales' name='chkProspSnapshotInsideSales'")
												
											Else
												Response.Write("<input type='checkbox' id='chkProspSnapshotInsideSales' name='chkProspSnapshotInsideSales' checked")
											End If
											Response.Write(">")
											%>
			                            </li>
			                            <!-- eof line -->
			                            
			                            <!-- line -->
			                            <li>
			                            	Send report to all Outside Sales managers
			      				      		<%
								      		If ProspSnapshotOutsideSales = 0 Then
												Response.Write("<input type='checkbox' id='chkProspSnapshotOutsideSales' name='chkProspSnapshotOutsideSales'")
												
											Else
												Response.Write("<input type='checkbox' id='chkProspSnapshotOutsideSales' name='chkProspSnapshotOutsideSales' checked")
											End If
											Response.Write(">")
											%>
			                            </li>
			                            <!-- eof line -->
                            
										
										<li>
											<p>Select additional email addresses to send the report to:</p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalProspSnapshotAdditionalEmails" data-tooltip="true" data-title="Additional emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails</button>						
				             				<% If ProspSnapshotAdditionalEmails <> "" Then %>
				             					<p style="margin-top:20px;"><strong>Current Additional Emails:</strong> <%= ProspSnapshotAdditionalEmails %></p>
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
				<% If ProspectingWeeklyAgendaReportOnOff = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title"><%= GetTerm("Prospecting") %> Weekly Agenda Report (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title"><%= GetTerm("Prospecting") %> Weekly Agenda Report (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
					<div class="panel-body">

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON 
					      		<%
					      		If ProspectingWeeklyAgendaReportOnOff = 0 Then
									Response.Write("<input type='checkbox' id='chkProspectingWeeklyAgendaReportOnOff' name='chkProspectingWeeklyAgendaReportOnOff'")
								Else
									Response.Write("<input type='checkbox' id='chkProspectingWeeklyAgendaReportOnOff' name='chkProspectingWeeklyAgendaReportOnOff' checked")
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalProspectingWeeklyAgendaReportScheduler" data-tooltip="true" data-title="<%= GetTerm("Prospecting") %> Weekly Agenda Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> <%= GetTerm("Prospecting") %> Weekly Agenda Report Scheduler</button>
										</li>
										<li>								
											<p>Specify the subject line to be used for the email:</p>
											<input type="text"class="form-control" style="width:100%;" name="txtProspectingWeeklyAgendaReportEmailSubject" id="txtProspectingWeeklyAgendaReportEmailSubject" value="<%= ProspectingWeeklyAgendaReportEmailSubject %>">
										</li>
										<li>
											<p>Select users <i class="fad fa-user-friends"></i> to send the report to:</p>
											<input type="hidden" name="lstSelectedProspectingWeeklyAgendaReportUserIDs" id="lstSelectedProspectingWeeklyAgendaReportUserIDs" value="<%= ProspectingWeeklyAgendaReportUserNos %>">
											<select id="lstExistingProspectingWeeklyAgendaReportUserIDs" multiple="multiple" name="lstExistingProspectingWeeklyAgendaReportUserIDs">
												<%	'Get list of all users not currently archived or disabled
													
												Set cnnUserList = Server.CreateObject("ADODB.Connection")
												cnnUserList.open Session("ClientCnnString")
								
												SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 AND userEnabled <> 0 ORDER BY userFirstName,userLastName"
												
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
											<p>Select additional email addresses to send each report to:</p>
											<p><strong>NOTE</strong>: <em>Each address will be copied on <strong>EVERY</strong> agenda mailed out</em></p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalProspectingWeeklyAgendaReportAdditionalEmails" data-tooltip="true" data-title="Additional emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails</button>						
				             				<% If ProspectingWeeklyAgendaReportAdditionalEmails <> "" Then %>
				             					<p style="margin-top:20px;"><strong>Current Additional Emails:</strong> <%= ProspectingWeeklyAgendaReportAdditionalEmails %></p>
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
				&nbsp;				
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


<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR REPORT SCHEDULERS START HERE !-->
<!-- **************************************************************************************************************************** -->

<!-- pencil Modal -->
<div class="modal fade" id="modalProspectingSnapshotReportScheduler" tabindex="-1" role="dialog" aria-labelledby="modalProspectingSnapshotReportSchedulerLabel">
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="titleProspectingSnapshotReportSchedulerLabel"><%= GetTerm("Prospecting") %> Snapshot Report Generation Scheduler</h4>
		    </div>

			<form name="frmEditProspectingSnapshotReportSchedulerModal" id="frmEditProspectingSnapshotReportSchedulerModal" action="prospecting-settings-snapshot-report-scheduler-submit.asp" method="POST">

				<div class="modal-body">
				    
					<div id="modalProspectingSnapshotReportSchedulerContent">
						<!-- Content for the modal will be generated and written here -->
						<!-- Content generated by Sub GetContentForProspectingSnapshotReportScheduler() in InSightFuncs_AjaxForAdminTimepickerModals.asp -->
					</div>
						
				</div>
				<!-- eof modal body !-->
				
				 <div class="clearfix"></div>
			      
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
					<button type="submit" id="btnProspectingSnapshotReportScheduleSave" class="btn btn-primary">Save Schedule Changes</button>
				</div>
				
			</form>

		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->


<div class="modal fade" id="modalProspSnapshotAdditionalEmails" tabindex="-1" role="dialog" aria-labelledby="modalProspSnapshotAdditionalEmailsLabel">
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="H6">Send report to the following additional email addresses</h4>
		    </div>

			<form name="frmEditProspSnapshotAdditionalEmails" id="frmEditProspSnapshotAdditionalEmails" action="users-list-update.asp" method="POST">
                <input type="hidden" name="userListName" value="ProspSnapshotAdditionalEmails" />
				<div class="modal-body">
				    
					<div id="Div6">
						<textarea class="form-control email-alert-line" rows="5" id="txtProspSnapshotAdditionalEmails" name="txtProspSnapshotAdditionalEmails"><%=ProspSnapshotAdditionalEmails%></textarea>
						<strong>Separate multiple email addresses with a semicolon</strong>
					</div>
						
				</div>
				<!-- eof modal body !-->
				
				 <div class="clearfix"></div>
			      
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
					<button type="submit" id="Button5" class="btn btn-primary">Save Changes</button>
				</div>
				
			</form>

		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>



<!-- pencil Modal -->
<div class="modal fade" id="modalProspectingWeeklyAgendaReportScheduler" tabindex="-1" role="dialog" aria-labelledby="modalProspectingWeeklyAgendaReportSchedulerLabel">
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="titleProspectingWeeklyAgendaReportSchedulerLabel"><%= GetTerm("Prospecting") %> Weekly Agenda Report Generation Scheduler</h4>
		    </div>

			<form name="frmEditProspectingWeeklyAgendaReportSchedulerModal" id="frmEditProspectingWeeklyAgendaReportSchedulerModal" action="prospecting-settings-weeklay-agenda-report-scheduler-submit.asp" method="POST">

				<div class="modal-body">
				    
					<div id="modalProspectingWeeklyAgendaReportSchedulerContent">
						<!-- Content for the modal will be generated and written here -->
						<!-- Content generated by Sub GetContentForProspectingWeeklyAgendaReportScheduler() in InSightFuncs_AjaxForAdminTimepickerModals.asp -->
					</div>
						
				</div>
				<!-- eof modal body !-->
				
				 <div class="clearfix"></div>
			      
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
					<button type="submit" id="btnProspectingWeeklyAgendaReportScheduleSave" class="btn btn-primary">Save Schedule Changes</button>
				</div>
				
			</form>

		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->


<div class="modal fade" id="modalProspectingWeeklyAgendaReportAdditionalEmails" tabindex="-1" role="dialog" aria-labelledby="modalProspectingWeeklyAgendaReportAdditionalEmailsLabel">
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="H6">Send report to the following additional email addresses</h4>
		    </div>

			<form name="frmEditProspectingWeeklyAgendaReportAdditionalEmails" id="frmEditProspectingWeeklyAgendaReportAdditionalEmails" action="users-list-update.asp" method="POST">
                <input type="hidden" name="userListName" value="ProspectingWeeklyAgendaReportAdditionalEmails" />
				<div class="modal-body">
				    
					<div id="Div6">
						<textarea class="form-control email-alert-line" rows="5" id="txtProspectingWeeklyAgendaReportAdditionalEmails" name="txtProspectingWeeklyAgendaReportAdditionalEmails"><%=ProspectingWeeklyAgendaReportAdditionalEmails%></textarea>
						<strong>Separate multiple email addresses with a semicolon</strong>
					</div>
						
				</div>
				<!-- eof modal body !-->
				
				 <div class="clearfix"></div>
			      
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
					<button type="submit" id="Button5" class="btn btn-primary">Save Changes</button>
				</div>
				
			</form>

		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>



<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR REPORT SCHEDULERS END HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="prospecting-settings-color-pickers.asp"-->

<!--#include file="../../../inc/footer-main.asp"-->

