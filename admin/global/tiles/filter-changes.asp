<!--#include file="../../../inc/header.asp"-->
	
	<!-- bootstrap timepicker !-->
	<script type="text/javascript" src="http://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>	
	<link href="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.css" rel="stylesheet" type="text/css">
	<script src="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.js" type="text/javascript"></script>
	<!-- eof bootstrap timepicker !-->

	<!-- spectrum color picker !-->
	<script src="<%= BaseURL %>/js/spectrum-color-picker/spectrum.js"></script>
	<link rel="stylesheet" type="text/css" href="<%= BaseURL %>/js/spectrum-color-picker/spectrum.css">
	<!-- eof spectrum color picker !-->

<%
	SQL = "SELECT * FROM Settings_Global"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		FilterChangeDays = rs("FilterChangeDays")
		FilterChangeDaysFieldService = rs("FilterChangeDaysFieldService")			
	End If
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	
	
	SQL = "SELECT * FROM Settings_EmailService"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		FilterChangePDFIncludeServiceNotes = rs("FilterChangePDFIncludeServiceNotes")
		CompletedFilterChangeEmailOn = rs("CompletedFilterChangeEmailOn")
		DoNotSendClientCompletedFilter= rs("DoNotSendClientCompletedFilter")
		SendCompletedFilterChangesTo = rs("SendCompletedFilterChangesTo")
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	


	SQL = "SELECT * FROM Settings_FieldService"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		FilterChangeIndicatorAndButtonColor = rs("FilterChangeIndicatorAndButtonColor")
		ShowSeparateFilterChangesTabOnServiceScreen = rs("ShowSeparateFilterChangesTabOnServiceScreen")
		AutoFilterChangeGenerationONOFF = rs("AutoFilterChangeGenerationONOFF")
		AutoFilterChangeUseRegions = rs("AutoFilterChangeUseRegions")
		AutoFilterChangeMaxNumTicketsPerDay = rs("AutoFilterChangeMaxNumTicketsPerDay")
	End If
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
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
					
		$('#modalAutoFilterGenerationScheduler').on('show.bs.modal', function(e) {
		    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
				cache: false,
				data: "action=GetContentForAutoFilterGenerationScheduler",
				success: function(response)
				 {
	               	 $modal.find('#modalAutoFilterGenerationSchedulerContent').html(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalAutoFilterGenerationSchedulerContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
			
		});
				
		
	});
</script>



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
	    padding: 18px 28px;
	    font-size: 22px;	    
	}	
</style>


<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;Filter Changes 
	<button id="toggle" class="btn btn-small btn-success"><i class="fas fa-arrows-v"></i>&nbsp;EXPAND/COLLAPSE ALL SETTINGS</button>
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>


<form method="post" action="filter-changes-submit.asp" name="frmFilterChanges" id="frmFilterChanges">



	<div class="container">
		
		<%
			Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
			Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
			Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
			Response.Write("</div>")
			Response.Flush()
		%>
	
		<div class="row">
			<h3><i class="fad fa-sliders-h"></i>&nbsp;<%= GetTerm("Filter Changes") %> General Settings</h3>
		</div>
	
	    <div class="row">
	    
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">Service Screen Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:2px;">
	               			
							<div class="col-lg-6">
								Show filter changes within 
							</div>
							<div class="col-lg-4">
								<select class="form-control" name="selFilterChangeDays">
									<% For x = 1 to 90
										If x = FilterChangeDays Then
										 	Response.Write("<option selected>" & x & "</option>")
										Else
										 	Response.Write("<option>" & x & "</option>")
										End If
									Next %>
								</select> 
							</div>
							
							<div class="col-lg-2">
								days
							</div>
				               			
						</div>
	                 	<!-- eof line !-->
	                 	
	                 	
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:2px;">
	               			
							<div class="col-lg-12">
								<strong>Field Service</strong> 
							</div>

				               			
						</div>
	                 	<!-- eof line !-->
                
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:2px;">
	               			
							<div class="col-lg-6">
								Show filter changes within 
							</div>
							<div class="col-lg-4">
								<select class="form-control" name="selFilterChangeDaysFieldService">
									<% For x = 1 to 90
										If x = FilterChangeDaysFieldService Then
											Response.Write("<option selected>" & x & "</option>")
										Else
											Response.Write("<option>" & x & "</option>")
										End If
									Next %>
								</select> 
							</div>
							
							<div class="col-lg-2">
								days
							</div>
				               			
						</div>
	                 	<!-- eof line !-->


	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:10px;margin-bottom:2px;">
							<div class="col-lg-8">
								Filter change button and indicator
							</div>
				         	<div class="col-lg-4">
								<input type='text' id="txtFilterChangeIndicatorAndButtonColor" name="txtFilterChangeIndicatorAndButtonColor" value="<%= FilterChangeIndicatorAndButtonColor %>">
							</div>      			
						</div>
	                 	<!-- eof line !-->
	                 	

	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:10px;margin-bottom:2px;">

				     	   	<div class="col-lg-10">Show Separate Filter Change Tab On <%= GetTerm("Service") %> Screen</div>
				         	<div class="col-lg-2">
								<%
								
								If Not IsNumeric(ShowSeparateFilterChangesTabOnServiceScreen) Then ShowSeparateFilterChangesTabOnServiceScreen = 0
								If cInt(ShowSeparateFilterChangesTabOnServiceScreen) = 0 Then
									Response.Write("<input type='checkbox' id='chkShowSeparateFilterChangesTabOnServiceScreen' name='chkShowSeparateFilterChangesTabOnServiceScreen' ")
								Else
									Response.Write("<input type='checkbox' id='chkShowSeparateFilterChangesTabOnServiceScreen' name='chkShowSeparateFilterChangesTabOnServiceScreen' checked ")
								End If
								Response.Write(">")%>
							</div>
						</div>
	                 	<!-- eof line !-->
					
					</div>
				</div>
			</div>
			

	    
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">Automatic Ticket Generation Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
							<div class="row schedule-info" style="margin-top:0px;margin-bottom:2px;">
	
					     	   	<div class="col-lg-10">Turn ON automatic filter change ticket generation</div>
					         	<div class="col-lg-2">
									<%
									If Not IsNumeric(AutoFilterChangeGenerationONOFF) Then AutoFilterChangeGenerationONOFF = 0
									If cInt(AutoFilterChangeGenerationONOFF) = 0 Then
										Response.Write("<input type='checkbox' id='chkAutoFilterChangeGenerationONOFF' name='chkAutoFilterChangeGenerationONOFF'>")
									Else
										Response.Write("<input type='checkbox' id='chkAutoFilterChangeGenerationONOFF' name='chkAutoFilterChangeGenerationONOFF' checked>")
									End If
									%>
								</div>
	
							</div>
							
							
							<div class="row schedule-info" style="margin-top:20px;margin-bottom:2px;">
	
					     	   	<div class="col-lg-10">Use Regions</div>
					         	<div class="col-lg-2">
									<%
									If Not IsNumeric(AutoFilterChangeUseRegions) Then AutoFilterChangeUseRegions = 0
									
									If cInt(AutoFilterChangeUseRegions) = 0 Then
										Response.Write("<input type='checkbox' id='chkAutoFilterChangeUseRegions' name='chkAutoFilterChangeUseRegions'>")
									Else
										Response.Write("<input type='checkbox' id='chkAutoFilterChangeUseRegions' name='chkAutoFilterChangeUseRegions' checked>")
									End If
									%>
								</div>
	
							</div>
							
							
							<div class="row schedule-info" style="margin-top:20px;margin-bottom:2px;">
	
					     	   	<div class="col-lg-8">
					     	   		Maximum # of Tickets to Generate Per Day
					     	   		<br>*<strong>ZERO DENOTES NO LIMIT</strong>*
					     	   	</div>
					         	<div class="col-lg-4">								
									<select class="form-control" name="selAutoFilterChangeMaxNumTicketsPerDay" id="selAutoFilterChangeMaxNumTicketsPerDay">
										<% 
											
											For x = 0 to 250
											If x = AutoFilterChangeMaxNumTicketsPerDay Then
												Response.Write("<option selected>" & x & "</option>")
											Else
												Response.Write("<option>" & x & "</option>")
											End If
											x = x + 4
										Next %>
									</select> 
								</div>
	
							</div>
							
	
							<div class="row schedule-info" style="margin-top:20px;margin-bottom:2px;">
					     	   	<div class="col-lg-12">
									<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalAutoFilterGenerationScheduler" data-tooltip="true" data-title="Filter Ticket Scheduler" style="cursor:pointer;"><i class="fa fa-calendar" aria-hidden="true"></i> Automatic Filter Ticket Generation Scheduler</button>						
								</div>
							</div>
					
					</div>
				</div>
			</div>


	    
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">Email Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
						
						<div class="row schedule-info" style="margin-top:0px;margin-bottom:2px;">
				       	   	<div class="col-lg-10">Completed filter change triggers email</div>
		        		 	<div class="col-lg-2">
			    				<%
								Response.Write("<input type='checkbox' class='check' id='chkFilterChangeEmail' name='chkFilterChangeEmail'")
								If CompletedFilterChangeEmailOn = vbTrue Then Response.Write(" checked ")
								Response.Write(">")
								%>
			    			</div>
						</div>
	
						<div class="row schedule-info" style="margin-top:20px;margin-bottom:2px;">
				       	   	<div class="col-lg-12">
				       	   		Send completed filter change email to the following additional addresses:
				       	   		<br>
				       	   		<strong>Separate multiple email addresses with a semicolon.</strong>
				       	   	</div>
						</div>
						
	
						<div class="row schedule-info" style="margin-top:20px;margin-bottom:2px;">
		        		 	<div class="col-lg-12">
								<textarea name="txtFilterChangeEmailsTo" id="txtFilterChangeEmailsTo" rows="4" class="form-control"><%= SendCompletedFilterChangesTo %></textarea>
			    			</div>
						</div>
						
						<div class="row schedule-info" style="margin-top:10px;margin-bottom:2px;">
				       	   	<div class="col-lg-10">Do not send an email to <%=GetTerm("clients")%> for completed filter changes </div>
		        		 	<div class="col-lg-2">
		    				<%
							Response.Write("<input type='checkbox' class='check' id='chkDoNotSendClientCompletedFilter' name='chkDoNotSendClientCompletedFilter'")
							If DoNotSendClientCompletedFilter = vbTrue Then Response.Write(" checked ")
							Response.Write(">")
							%>
			    			</div>
						</div>
	
						<div class="row schedule-info" style="margin-top:20px;margin-bottom:2px;">
							<div class="col-lg-10">Include service notes in filter change emailed .pdf </div>
		        		 	<div class="col-lg-2">
		    				<%
							Response.Write("<input type='checkbox' class='check' id='chkFilterChangePDFIncludeServiceNotes' name='chkFilterChangePDFIncludeServiceNotes'")
							If FilterChangePDFIncludeServiceNotes = vbTrue Then Response.Write(" checked ")
							Response.Write(">")
							%>
			    			</div>
						</div>
	
					
					</div>
				</div>
			</div>

	</div>
				
	<!-- cancel / save !-->
	<div class="row pull-right">
		<div class="col-lg-12">
			<a href="<%= BaseURL %>admin/global/main.asp"><button type="button" class="btn btn-default btn-lg btn-huge"><i class="far fa-times-circle"></i> Cancel</button></a> 
			<button type="submit" class="btn btn-primary btn-lg btn-huge" onclick="showSavingChangesDiv()"><i class="far fa-save"></i> Save Changes</button>
		</div>
	</div>
	
</div><!-- container -->

</form>


<!-- color picker js !-->
<!-- spectrum color picker js !-->
<script>
$("#txtFilterChangeIndicatorAndButtonColor").spectrum({
    color: '<%= FilterChangeIndicatorAndButtonColor %>',
    showInput: true,
    className: "full-spectrum",
    showInitial: true,
    showPalette: true,
    showSelectionPalette: true,
    maxSelectionSize: 10,
    preferredFormat: "hex",
    localStorageKey: "spectrum.demo",
    move: function (color) {
        
    },
    show: function () {
    
    },
    beforeShow: function () {
    
    },
    hide: function () {
    
    },
    change: function() {
        
    },
    palette: [
        ["rgb(0, 0, 0)", "rgb(67, 67, 67)", "rgb(102, 102, 102)",
        "rgb(204, 204, 204)", "rgb(217, 217, 217)","rgb(255, 255, 255)"],
        ["rgb(152, 0, 0)", "rgb(255, 0, 0)", "rgb(255, 153, 0)", "rgb(255, 255, 0)", "rgb(0, 255, 0)",
        "rgb(0, 255, 255)", "rgb(74, 134, 232)", "rgb(0, 0, 255)", "rgb(153, 0, 255)", "rgb(255, 0, 255)"], 
        ["rgb(230, 184, 175)", "rgb(244, 204, 204)", "rgb(252, 229, 205)", "rgb(255, 242, 204)", "rgb(217, 234, 211)", 
        "rgb(208, 224, 227)", "rgb(201, 218, 248)", "rgb(207, 226, 243)", "rgb(217, 210, 233)", "rgb(234, 209, 220)", 
        "rgb(221, 126, 107)", "rgb(234, 153, 153)", "rgb(249, 203, 156)", "rgb(255, 229, 153)", "rgb(182, 215, 168)", 
        "rgb(162, 196, 201)", "rgb(164, 194, 244)", "rgb(159, 197, 232)", "rgb(180, 167, 214)", "rgb(213, 166, 189)", 
        "rgb(204, 65, 37)", "rgb(224, 102, 102)", "rgb(246, 178, 107)", "rgb(255, 217, 102)", "rgb(147, 196, 125)", 
        "rgb(118, 165, 175)", "rgb(109, 158, 235)", "rgb(111, 168, 220)", "rgb(142, 124, 195)", "rgb(194, 123, 160)",
        "rgb(166, 28, 0)", "rgb(204, 0, 0)", "rgb(230, 145, 56)", "rgb(241, 194, 50)", "rgb(106, 168, 79)",
        "rgb(69, 129, 142)", "rgb(60, 120, 216)", "rgb(61, 133, 198)", "rgb(103, 78, 167)", "rgb(166, 77, 121)",
        "rgb(91, 15, 0)", "rgb(102, 0, 0)", "rgb(120, 63, 4)", "rgb(127, 96, 0)", "rgb(39, 78, 19)", 
        "rgb(12, 52, 61)", "rgb(28, 69, 135)", "rgb(7, 55, 99)", "rgb(32, 18, 77)", "rgb(76, 17, 48)"]
    ]
});

</script>


<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR FILTER OPTIONS START HERE !-->
<!-- **************************************************************************************************************************** -->

<!-- pencil Modal -->
<div class="modal fade" id="modalAutoFilterGenerationScheduler" tabindex="-1" role="dialog" aria-labelledby="modalAutoFilterGenerationSchedulerLabel">
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="titleAutoFilterGenerationSchedulerLabel">Automatic Filter Ticket Generation Scheduler</h4>
		    </div>

			<form name="frmEditAutoFilterGenerationSchedulerModal" id="frmEditAutoFilterGenerationSchedulerModal" action="filter-changes-scheduler-submit.asp" method="POST">

				<div class="modal-body">
				    
					<div id="modalAutoFilterGenerationSchedulerContent">
						<!-- Content for the modal will be generated and written here -->
						<!-- Content generated by Sub GetContentForAutoFilterGenerationScheduler() in InSightFuncs_AjaxForAdminTimepickerModals.asp -->
					</div>
						
				</div>
				<!-- eof modal body !-->
				
				 <div class="clearfix"></div>
			      
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
					<button type="submit" id="btnAutoFilterGenerationScheduleSave" class="btn btn-primary">Save Schedule Changes</button>
				</div>
				
			</form>

		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->
            


<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR FILTER OPTIONS END HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="../../../inc/footer-main.asp"-->
