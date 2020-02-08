<!--#include file="../../../inc/header.asp"-->

<!-- bootstrap timepicker !-->
<script type="text/javascript" src="http://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>	
<link href="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.css" rel="stylesheet" type="text/css">
<script src="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.min.js" type="text/javascript"></script>
<!-- eof bootstrap timepicker !-->

<!-- bootstrap multiselect !-->
<link rel="stylesheet" href="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.css" type="text/css">
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.js"></script>
<!-- eof bootstrap multiselect !-->

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

					
		$('#modalDailyAPIActivityByPartnerReportScheduler').on('show.bs.modal', function(e) {
		    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
				cache: false,
				data: "action=GetContentForDailyInventoryAPIActivityByPartnerReportScheduler",
				success: function(response)
				 {
	               	 $modal.find('#modalDailyAPIActivityByPartnerReportSchedulerContent').html(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalDailyAPIActivityByPartnerReportSchedulerContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
			
		});		

		$('#modalInventoryProductChangesReportScheduler').on('show.bs.modal', function(e) {
		    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
				cache: false,
				data: "action=GetContentForInventoryProductChangesReportScheduler",
				success: function(response)
				 {
	               	 $modal.find('#modalInventoryProductChangesReportSchedulerContent').html(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalInventoryProductChangesReportSchedulerContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
			
		});	
		

        $('#modalInventoryAPIDailyActivityReportAdditionalEmails').on('show.bs.modal', function (e) {
            var $modal = $(this);
        });


        $('#modalInventoryProductChangesReportAdditionalEmails').on('show.bs.modal', function (e) {
            var $modal = $(this);
        });
        
        
		$('#lstExistingInventoryAPIDailyActivityReportUserIDs').multiselect({
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
			nonSelectedText:'No Users Selected For Daily API Activity Report',
			numberDisplayed: 20,
		    onChange: function() {
		        var selected = this.$select.val();
		        $("#lstSelectedInventoryAPIDailyActivityReportUserIDs").val(selected);
		        console.log(selected);
		        // ...
		    }
    			
	    });	
	    
		//*************************************************************************************************
		//Load the bootstrap multiselect box with the current daily API activity report users preselected
		//*************************************************************************************************
		var data= $("#lstSelectedInventoryAPIDailyActivityReportUserIDs").val();
		
		if (data) {
			//Make an array
			var dataarray=data.split(",");
			// Set the value
			$("#lstExistingInventoryAPIDailyActivityReportUserIDs").val(dataarray);
			// Then refresh
			$("#lstExistingInventoryAPIDailyActivityReportUserIDs").multiselect("refresh");
		}
		//*************************************************************************************************
        


		$('#lstExistingInventoryProductChangesReportUserIDs').multiselect({
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
			nonSelectedText:'No Users Selected For Product Changes Report',
			numberDisplayed: 20,
		    onChange: function() {
		        var selected = this.$select.val();
		        $("#lstSelectedInventoryProductChangesReportUserIDs").val(selected);
		        console.log(selected);
		        // ...
		    }
    			
	    });	
	    
		//*************************************************************************************************
		//Load the bootstrap multiselect box with the current daily API activity report users preselected
		//*************************************************************************************************
		var data= $("#lstSelectedInventoryProductChangesReportUserIDs").val();
		
		if (data) {
			//Make an array
			var dataarray=data.split(",");
			// Set the value
			$("#lstExistingInventoryProductChangesReportUserIDs").val(dataarray);
			// Then refresh
			$("#lstExistingInventoryProductChangesReportUserIDs").multiselect("refresh");
		}
		//*************************************************************************************************

		
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
	
	.line-full{
	 	margin-bottom:20px;
	}
	
	.multi-select{
		min-height:200px;
		min-width:170px;
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



<%
	SQL = "SELECT * FROM Settings_Global"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		InventoryAPIRepostONOFF = rs("InventoryAPIRepostONOFF")
		InventoryAPIRepostMode = rs("InventoryAPIRepostMode")	
		InventoryAPIRepostURL = rs("InventoryAPIRepostURL")	
		InventoryAPIDailyActivityReportOnOff	= rs("InventoryAPIDailyActivityReportOnOff")			
		InventoryAPIDailyActivityReportAdditionalEmails = rs("InventoryAPIDailyActivityReportAdditionalEmails")	
		InventoryAPIDailyActivityReportEmailSubject = rs("InventoryAPIDailyActivityReportEmailSubject")	
		InventoryAPIDailyActivityReportUserNos = rs("InventoryAPIDailyActivityReportUserNos")	
		InventoryAPIRepostOnHandONOFF = rs("InventoryAPIRepostOnHandONOFF")
		InventoryAPIRepostOnHandMode = rs("InventoryAPIRepostOnHandMode")	
		InventoryAPIRepostOnHandURL = rs("InventoryAPIRepostOnHandURL")	
		InventoryWebAppPostOnHandMode = rs("InventoryWebAppPostOnHandMode")	
		InventoryWebAppPostOnHandURL = rs("InventoryWebAppPostOnHandURL")	
	End If
	
	SQL = "SELECT * FROM Settings_InventoryControl"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		InventoryProductChangesReportOnOff = rs("InventoryProductChangesReportOnOff")			
		InventoryProductChangesReportAdditionalEmails = rs("InventoryProductChangesReportAdditionalEmails")	
		InventoryProductChangesReportEmailSubject = rs("InventoryProductChangesReportEmailSubject")	
		InventoryProductChangesReportUserNos = rs("InventoryProductChangesReportUserNos")	
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
%>


<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;Inventory Control 
	<button id="toggle" class="btn btn-small btn-success"><i class="fas fa-arrows-v"></i>&nbsp;EXPAND/COLLAPSE ALL SETTINGS</button>
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>


<form method="post" action="inventory-submit.asp" name="frmInventoryControl" id="frmInventoryControl">


<% If MUV_Read("InventoryControlModuleOn")  = "Disabled" Then %>
	<div class="col-lg-6">
		<br><br>
		Please contact support if you would like to activate the Inventory module.
	</div>
<% ElseIf MUV_Read("InventoryControlModuleOn")  = "Enabled" Then  %>


	<div class="container">
	
		<%
			Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
			Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
			Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
			Response.Write("</div>")
			Response.Flush()
		%>
	
		<div class="row">
			<h3><i class="fad fa-sliders-h"></i>&nbsp;Inventory General Settings</h3>
		</div>
	
	    <div class="row">
	    
			<div class="col-md-4">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">INVENTORY API RePost/Post Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
					
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
	               			<div class="col-lg-6">
	               			<%
								If InventoryAPIRepostONOFF = 0 Then
									Response.Write("<input type='checkbox' class='check' id='chkInventoryOrderAPIRepostONOFF' name='chkInventoryOrderAPIRepostONOFF'")
								Else
									Response.Write("<input type='checkbox' class='check' id='chkInventoryOrderAPIRepostONOFF' name='chkInventoryOrderAPIRepostONOFF' checked")
								End If
								Response.Write("> Check to turn on")
							%>
	               			</div>
	               			
							<div class="col-lg-6">
								Mode:<br>
								<select class="form-control pull-left" name="selInventoryOrderAPIRepostMode">
									<option value="TEST" <% If InventoryAPIRepostMode = "TEST" Then Response.Write("selected") %>>TEST</option>
									<option value="LIVE" <% If InventoryAPIRepostMode = "LIVE" Then Response.Write("selected") %>>LIVE</option>
								</select>
							</div>
						</div>
						<!-- eof line !-->
					
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:10px;">
	               			<div class="col-lg-10"><strong>Re-post</strong> received adjustments to URL</div>
						</div>
	                 	<!-- eof line !-->
	                 	
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
	               			<div class="col-lg-12"><input type="text"class="form-control" style="width:100%;" name="txtInventoryAPIRepostURL" id="txtInventoryAPIRepostURL" value="<%= InventoryAPIRepostURL %>"></div>
						</div>
						
						
		              	<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:10px;">
							<div class="col-lg-12"><hr style="border-top: 2px dotted #337AB7"></div>
						</div>

					
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
	               			<div class="col-lg-6">
	               			<%
								If InventoryAPIRepostOnHandONOFF = 0 Then
									Response.Write("<input type='checkbox' class='check' id='chkInventoryAPIRepostOnHandONOFF' name='chkInventoryAPIRepostOnHandONOFF'")
								Else
									Response.Write("<input type='checkbox' class='check' id='chkInventoryAPIRepostOnHandONOFF' name='chkInventoryAPIRepostOnHandONOFF' checked")
								End If
								Response.Write("> Check to turn on")
							%>
	               			</div>
	               			
							<div class="col-lg-6">
								Mode:<br>
								<select class="form-control pull-left" name="selInventoryOrderAPIRepostOnHandMode">
									<option value="TEST" <% If InventoryAPIRepostOnHandMode = "TEST" Then Response.Write("selected") %>>TEST</option>
									<option value="LIVE" <% If InventoryAPIRepostOnHandMode = "LIVE" Then Response.Write("selected") %>>LIVE</option>
								</select>
							</div>
						</div>
						<!-- eof line !-->
					
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:10px;">
	               			<div class="col-lg-10"><strong>Re-post</strong> received on-hand replacement posts to URL</div>
						</div>
	                 	<!-- eof line !-->
	                 	
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
	               			<div class="col-lg-12"><input type="text"class="form-control" style="width:100%;" name="txtInventoryAPIRepostOnHandURL" id="txtInventoryAPIRepostOnHandURL" value="<%= InventoryAPIRepostOnHandURL %>"></div>
						</div>
						
						
		              	<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:10px;">
							<div class="col-lg-12"><hr style="border-top: 2px dotted #337AB7"></div>
						</div>
					
					
					
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:10px;">
	               			<div class="col-lg-10"><strong>Web App</strong> <strong>post</strong> on-hand posts to URL</div>
						</div>
	                 	<!-- eof line !-->
	                 	
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">
	               			<div class="col-lg-12"><input type="text"class="form-control" style="width:100%;" name="txtInventoryWebAppPostOnHandURL" id="txtInventoryWebAppPostOnHandURL" value="<%= InventoryWebAppPostOnHandURL %>"></div>
						</div>
						
						
	              		<!-- line !-->
	                	<div class="row schedule-info" style="margin-top:0px;margin-bottom:20px;">	               			
							<div class="col-lg-6">
								Mode:<br>
								<select class="form-control pull-left" name="selInventoryWebAppPostOnHandMode">
									<option value="TEST" <% If InventoryWebAppPostOnHandMode = "TEST" Then Response.Write("selected") %>>TEST</option>
									<option value="LIVE" <% If InventoryWebAppPostOnHandMode = "LIVE" Then Response.Write("selected") %>>LIVE</option>
								</select>
							</div>
						</div>
						<!-- eof line !-->
						

				
						
					</div>
				</div>
			</div>
			
			<div class="col-md-4">
				&nbsp;
			</div>
			
			<div class="col-md-4">
				&nbsp;
			</div>
			
		</div><!-- eof row -->
		



	
	
		<div class="row">
			<h3><i class="fad fa-file-pdf"></i>&nbsp;<%= GetTerm("Inventory") %> Report Settings</h3>
		</div>
	
		<div class="row">
		
			<div class="col-md-4">
				<% If InventoryAPIDailyActivityReportOnOff = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title">Daily <%= GetTerm("Inventory") %> API Activity By Partner Report Scheduler (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title">Daily <%= GetTerm("Inventory") %> API Activity By Partner Report Scheduler (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
					<div class="panel-body">

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON 
					      		<%
					      		If InventoryAPIDailyActivityReportOnOff = 0 Then
									Response.Write("<input type='checkbox' id='chkInventoryAPIDailyActivityReportOnOff' name='chkInventoryAPIDailyActivityReportOnOff'")
								Else
									Response.Write("<input type='checkbox' id='chkInventoryAPIDailyActivityReportOnOff' name='chkInventoryAPIDailyActivityReportOnOff' checked")
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalDailyAPIActivityByPartnerReportScheduler" data-tooltip="true" data-title="Daily <%= GetTerm("Inventory") %> API Activity By Partner Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> Daily <%= GetTerm("Inventory") %> API Activity By Partner Report Scheduler</button>
										</li>
										<li>								
											<p>Specify the subject line to be used for the email:</p>
											<input type="text"class="form-control" style="width:100%;" name="txtInventoryAPIDailyActivityReportEmailSubject" id="txtInventoryAPIDailyActivityReportEmailSubject" value="<%= InventoryAPIDailyActivityReportEmailSubject %>">
										</li>
										<li>
											<p>Select users <i class="fad fa-user-friends"></i> to send the report to:</p>
											<input type="hidden" name="lstSelectedInventoryAPIDailyActivityReportUserIDs" id="lstSelectedInventoryAPIDailyActivityReportUserIDs" value="<%= InventoryAPIDailyActivityReportUserNos %>">
											<select id="lstExistingInventoryAPIDailyActivityReportUserIDs" multiple="multiple" name="lstExistingInventoryAPIDailyActivityReportUserIDs">
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalInventoryAPIDailyActivityReportAdditionalEmails" data-tooltip="true" data-title="Additional emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails</button>						
				             				<% If InventoryAPIDailyActivityReportAdditionalEmails <> "" Then %>
				             					<p style="margin-top:20px;"><strong>Current Additional Emails:</strong> <%= InventoryAPIDailyActivityReportAdditionalEmails %></p>
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
				<% If InventoryProductChangesReportOnOff = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title">Inventory Product Changes Report (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title"><%= GetTerm("Inventory") %> Product Changes Report (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
					<div class="panel-body">

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON 
					      		<%
					      		If InventoryProductChangesReportOnOff = 0 Then
									Response.Write("<input type='checkbox' id='chkInventoryProductChangesReportOnOff' name='chkInventoryProductChangesReportOnOff'")
								Else
									Response.Write("<input type='checkbox' id='chkInventoryProductChangesReportOnOff' name='chkInventoryProductChangesReportOnOff' checked")
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalInventoryProductChangesReportScheduler" data-tooltip="true" data-title="<%= GetTerm("Inventory") %> Product Changes Report Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> <%= GetTerm("Inventory") %> Product Changes Report Report Scheduler</button>
										</li>
										<li>								
											<p>Specify the subject line to be used for the email:</p>
											<input type="text"class="form-control" style="width:100%;" name="txtInventoryProductChangesReportEmailSubject" id="txtInventoryProductChangesReportEmailSubject" value="<%= InventoryProductChangesReportEmailSubject %>">
										</li>
										<li>
											<p>Select users <i class="fad fa-user-friends"></i> to send the report to:</p>
											<input type="hidden" name="lstSelectedInventoryProductChangesReportUserIDs" id="lstSelectedInventoryProductChangesReportUserIDs" value="<%= InventoryProductChangesReportUserNos %>">
											<select id="lstExistingInventoryProductChangesReportUserIDs" multiple="multiple" name="lstExistingInventoryProductChangesReportUserIDs">
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalInventoryProductChangesReportAdditionalEmails" data-tooltip="true" data-title="Additional emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails</button>						
				             				<% If InventoryProductChangesReportAdditionalEmails <> "" Then %>
				             					<p style="margin-top:20px;"><strong>Current Additional Emails:</strong> <%= InventoryProductChangesReportAdditionalEmails %></p>
				             				<% End If %>
										</li>
									</ol>
								</div>
					
							</div>
						</div>
					
					
					</div>
				</div>	
				

				<div class="col-md-4">
					&nbsp;
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
	
	<% End If %>

</div><!-- container -->

</form>

<!--#include file="inventory-modals.asp"-->

<!--#include file="../../../inc/footer-main.asp"-->
