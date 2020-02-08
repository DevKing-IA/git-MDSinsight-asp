<!--#include file="../../../inc/header.asp"-->
<!--#include file="../../../inc/InSightFuncs_Routing.asp"-->

<!-- spectrum color picker !-->
<script src="<%= BaseURL %>/js/spectrum-color-picker/spectrum.js"></script>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>/js/spectrum-color-picker/spectrum.css">

    		
<script>

	function showSavingChangesDiv() {
	  document.getElementById('PleaseWaitPanel').style.display = "block";
	  setTimeout(function() {
	    document.getElementById('PleaseWaitPanel').style.display = "none";
	  },1500);
	   
	}

	$(document).ready(function () {
	
	    var ckbox = $('#txtDelBoardRoutesToIgnore');
	
	    $('input').on('click',function () {
	        if (ckbox.is(':checked')) {
	           $(this).parent().toggleClass('redBg');
	        } else {
	           $(this).parent().toggleClass('redBg');
	        }
	    });
	    
		$(".routes-to-ignore label input[type='checkbox']:checked").each(
		    function() {
		      $(this).parent().addClass('redBg');
		    }
		
		);
		
		$(".ups-routes label input[type='checkbox']:checked").each(
		    function() {
		      $(this).parent().addClass('brownBg');
		    }
		
		);
		
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
		
	    
	});
	
</script>
<!-- eof green / red background jQuery -->


<style>


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
		
	.custom-select{
		width: auto !important;
		display:inline-block;
	}
	
	.select-large{
		min-width:40% !important;
	}
	
	.nag-box{
	 	background:#f5f5f5;
		width:100%;
		float:left;
		padding:10px;
		margin-bottom:10px;
	}
	
	.nag-box2{
	 	background:#fff;
		width:100%;
		float:left;
		padding:10px;
		margin-bottom:10px;
	}
	
	.greenBg{
		background: #e3ecdf !important;
	 }
	
	.redBg{
		background: #f9d9d9 !important;
	 }
	
	.brownBg{
		background: #9c7705 !important
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
	
		DelBoardScheduledColor = rs("DelBoardScheduledColor")	
		DelBoardCompletedColor = rs("DelBoardCompletedColor")				
		DelBoardInProgressColor = rs("DelBoardInProgressColor")						
		DelBoardSkippedColor = rs("DelBoardSkippedColor")		
		DelBoardNextStopColor = rs("DelBoardNextStopColor")	
		DelBoardAMColor = rs("DelBoardAMColor")	
		DelBoardPriorityColor = rs("DelBoardPriorityColor")
		DelBoardPieTimerColor = rs("DelBoardPieTimerColor")	
		DelBoardTitleGradientColor = rs("DelBoardTitleGradientColor")	
		DelBoardTitleText = rs("DelBoardTitleText")	
		DelBoardTitleTextFontColor = rs("DelBoardTitleTextFontColor")					
		DelBoardProfitDollars = rs("DelBoardProfitDollars")			
		DelBoardAtOrAboveProfitColor = rs("DelBoardAtOrAboveProfitColor")			
		DelBoardBelowProfitColor = rs("DelBoardBelowProfitColor")			
		DelBoardUserAlertColor = rs("DelBoardUserAlertColor")	
		AutoPromptNextStop = rs("AutoPromptNextStop")
		AutoForceSelectNextStop = rs("AutoForceSelectNextStop")
		DoNotShowDeliveryLineItems = rs("DoNotShowDeliveryLineItems")
		DelBoardDontUseStopSequence = rs("DelBoardDontUseStopSequencing")
		DelBoardRoutesToIgnore = rs("DelBoardRoutesToIgnore")
		DelBoardUPSRoutes = rs("DelBoardUPSRoutes")	
		DelBoardPriorityColor = rs("DelBoardPriorityColor")	
			
		If MUV_Read("routingModuleOn") = "Enabled" Then	
		
			NextStopNagMessageONOFF = rs("NextStopNagMessageONOFF")
			NextStopNagMinutes = rs("NextStopNagMinutes")
			NextStopNagIntervalMinutes = rs("NextStopNagIntervalMinutes")
			NextStopNagMessageMaxToSendPerStop = rs("NextStopNagMessageMaxToSendPerStop")
			NextStopNagMessageMaxToSendPerDriverPerDay = rs("NextStopNagMessageMaxToSendPerDriverPerDay")
			NextStopNagMessageSendMethod = rs("NextStopNagMessageSendMethod")
			NoActivityNagMessageONOFF = rs("NoActivityNagMessageONOFF")
			NoActivityNagMinutes = rs("NoActivityNagMinutes")
			NoActivityNagIntervalMinutes = rs("NoActivityNagIntervalMinutes")
			NoActivityNagMessageMaxToSendPerStop = rs("NoActivityNagMessageMaxToSendPerStop")
			NoActivityNagMessageMaxToSendPerDriverPerDay = rs("NoActivityNagMessageMaxToSendPerDriverPerDay")
			NoActivityNagMessageSendMethod = rs("NoActivityNagMessageSendMethod")
			NoActivityNagTimeOfDay = rs("NoActivityNagTimeOfDay")
		End If   		
		
	End If
	
	If IsNull(DoNotShowDeliveryLineItems) Then DoNotShowDeliveryLineItems  = 0
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	


%>


<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;Routing Settings 
	<button id="toggle" class="btn btn-small btn-success"><i class="fas fa-arrows-v"></i>&nbsp;EXPAND/COLLAPSE ALL SETTINGS</button>
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>


<form method="post" action="delivery-board-submit.asp" name="frmDeliveryBoard" id="frmDeliveryBoard">


	<div class="container">
		
		<%
			Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
			Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
			Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
			Response.Write("</div>")
			Response.Flush()
		%>
	
		<div class="row">
			<h3><i class="fad fa-sliders-h"></i>&nbsp;<%= GetTerm("Routing") %> General Settings</h3>
		</div>
	
	    <div class="row">
	    
			<div class="col-md-6">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">Service Screen Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					

			         	<div class="row">
				         	<div class="col-lg-3">Do Not Use Stop Sequencing</div>
								<%
								If DelBoardDontUseStopSequence = 0 Then
									Response.Write("<input type='checkbox' id='chkDelBoardDontUseStopSequence' name='chkDelBoardDontUseStopSequence'>")
								Else
									Response.Write("<input type='checkbox' id='chkDelBoardDontUseStopSequence' name='chkDelBoardDontUseStopSequence' checked>")
								End If
								%>
						</div>
					
					</div>
				</div>
			</div>
			
			<div class="col-md-6">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">Route Profitability Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
			         	<div class="row">
			         	   	<div class="col-lg-8">Route profitabilty $ target </div>
				         	<div class="col-lg-2">
					         	<select class="form-control" name="selProfitDollars" id="selProfitDollars">
						         	<% For x = 0 to 3500 Step 5
						         			If x = DelBoardProfitDollars Then
		   							         	Response.Write("<option selected>$" & x & "</option>")
						         			Else
		   							         	Response.Write("<option>$" & x & "</option>")
								         	End If
							         	Next %>
					         	</select>
				         	</div>
			         	</div>
			         					    
			         	<div class="row">
			         	   	<div class="col-lg-8">Highlight color for routes at or above profit target</div>
				         	<div class="col-lg-2">
								<input type='text' id="txtAboveProfit" name="txtAboveProfit"  value="<%= DelBoardAtOrAboveProfitColor %>">
							</div>
						</div>
					    
			         	<div class="row">
			         	   	<div class="col-lg-8">Highlight color for routes below profit target </div>
				         	<div class="col-lg-2">
								<input type='text' id="txtBelowProfit" name="txtBelowProfit"  value="<%= DelBoardBelowProfitColor %>">
							</div>
						</div>
					
					</div>
				</div>
			</div>
			
			
		</div>
		
		
		
		<div class="row">
			<h3><i class="fad fa-mobile-alt"></i>&nbsp;<%= GetTerm("Routing") %> Mobile Web App Settings</h3>
		</div>
	
	    <div class="row">
	    
			<div class="col-md-6">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">General Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
				         	<div class="row">
					         	<div class="col-lg-8">Do Not Show Order Line Items For Deliveries</div>
									<%
									If DoNotShowDeliveryLineItems = 0 Then
										Response.Write("<input type='checkbox' id='chkDoNotShowDeliveryLineItems' name='chkDoNotShowDeliveryLineItems'>")
									Else
										Response.Write("<input type='checkbox' id='chkDoNotShowDeliveryLineItems' name='chkDoNotShowDeliveryLineItems' checked>")
									End If
									%>
							</div>
			
				         	<div class="row">
					         	<div class="col-lg-8">Automatically Prompt for Next Stop</div>
									<%
									If AutoPromptNextStop = 0 Then
										Response.Write("<input type='checkbox' id='chkAutoPromptNextStop' name='chkAutoPromptNextStop'>")
									Else
										Response.Write("<input type='checkbox' id='chkAutoPromptNextStop' name='chkAutoPromptNextStop' checked>")
									End If
									%>
							</div>
			
			
				         	<div class="row">
					         	<div class="col-lg-8">Force Driver to Set Next Stop<br>
						         	<font color="Red"><small><strong><i>This option can also be set for individual users via Manage Users. Individual user settings will override what is set here unless the user is set to Use Global.</i></small></strong></font>
					         	</div>
								<%
								If AutoForceSelectNextStop = 0 Then
									Response.Write("<input type='checkbox' id='chkAutoForceSelectNextStop' name='chkAutoForceSelectNextStop'>")
								Else
									Response.Write("<input type='checkbox' id='chkAutoForceSelectNextStop' name='chkAutoForceSelectNextStop' checked>")
								End If
								%>
							</div>
					
					</div>
				</div>
			</div>
			
			<div class="col-md-6">
				&nbsp;
			</div>
			
			
		</div>
		
		
		

		<div class="row">
			<h3><i class="fad fa-palette"></i>&nbsp;<%= GetTerm("Routing") %> Colors &amp; Settings</h3>
		</div>

		<div class="row">
				
			<div class="col-md-6">
				<div class="panel panel-info">
					<div class="panel-heading">
						<h3 class="panel-title">Delivery Board Settings</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
			         	<div class="row">
			         	   	<div class="col-lg-8">Title bar and border gradient color (in kiosk mode) </div>
					       	<div class="col-lg-2">
								<input type='text' id="txtDelBoardTitleGradientColor" name="txtDelBoardTitleGradientColor"  value="<%= DelBoardTitleGradientColor %>">
							</div>
						</div>
		
			         	<div class="row">
			         	   	<div class="col-lg-8">Title text font color (in kiosk mode) </div>
					       	<div class="col-lg-2">
								<input type='text' id="txtDelBoardTitleTextFontColor" name="txtDelBoardTitleTextFontColor"  value="<%= DelBoardTitleTextFontColor %>">
							</div>
						</div>
		
			         	<div class="row">
			         	   	<div class="col-lg-7">Title text (in kiosk mode)
			         	   		<br>&nbsp;&nbsp;&nbsp;<small>*To include the current date, use <b><i>~today~</i></b>.
			         	   		<br>&nbsp;&nbsp;&nbsp;*To include the day of week name, use <b><i>~dow~</i></b>.
			         	   		<br>&nbsp;&nbsp;&nbsp;Example: Deliveries for ~dow~ ~today~ will show<b>&nbsp;&nbsp;Deliveries for <%=WeekDayName(Datepart("w",Now()))%>&nbsp;<%=FormatDateTime(Now(),2)%></b></small>
							</div>
					       	<div class="col-lg-4">
								<input type='text' id="txtDelBoardTitleText" name="txtDelBoardTitleText"  value="<%= DelBoardTitleText%>" class="form-control" style="width:300px;">
								
							</div>
						</div>
						
		
			         	<div class="row">
			         	   	<div class="col-lg-8">Pie timer color </div>
					       	<div class="col-lg-2">
								<input type='text' id="txtDelBoardPieTimerColor" name="txtDelBoardPieTimerColor"  value="<%= DelBoardPieTimerColor %>">
							</div>
						</div>
						

			         	<div class="row">
			         	   	<div class="col-lg-8">Highlight color for scheduled deliveries </div>
					       	<div class="col-lg-2">
						       	<input type='text'  id="txtScheduledColor" name="txtScheduledColor" value="<%= DelBoardScheduledColor %>" > 
							</div>
						</div>
					    
					    
			         	<div class="row">
			         	   	<div class="col-lg-8">Highlight color for completed deliveries </div>
					       	<div class="col-lg-2">
								<input type='text' id="txtCompletedDeliveries" name="txtCompletedDeliveries"  value="<%= DelBoardCompletedColor%>">
							</div>
						</div>
					    
					    
			         	<div class="row">
			         	   	<div class="col-lg-8">Highlight color for in progress deliveries </div>
					       	<div class="col-lg-2">
								<input type='text' id="txtInProgress" name="txtInProgress"  value="<%= DelBoardInProgressColor%>">
							</div>
						</div>
					    				    
			         	<div class="row">
			         	   	<div class="col-lg-8">Highlight color for No Delivery / Skipped </div>
					       	<div class="col-lg-2">
								<input type='text' id="txtSkippedDeliveries" name="txtSkippedDeliveries"  value="<%= DelBoardSkippedColor%>">
							</div>
						</div>
		
			         	<div class="row">
			         	   	<div class="col-lg-8">Highlight color for next stop </div>
					       	<div class="col-lg-2">
								<input type='text' id="txtNextStop" name="txtNextStop"  value="<%= DelBoardNextStopColor%>">
							</div>
						</div>
		
			         	<div class="row">
			         	   	<div class="col-lg-8">Highlight color for AM deliveries </div>
					       	<div class="col-lg-2">
								<input type='text' id="txtDelBoardAMColor" name="txtDelBoardAMColor"  value="<%= DelBoardAMColor %>">
							</div>
						</div>
		
			         	<div class="row">
			         	   	<div class="col-lg-8">Highlight color for Priority deliveries </div>
					       	<div class="col-lg-2">
								<input type='text' id="txtDelBoardPriorityColor" name="txtDelBoardPriorityColor"  value="<%= DelBoardPriorityColor %>">
							</div>
						</div>
								
					    
			         	<div class="row">
			         	   	<div class="col-lg-8">Highlight color for deliveries with user defined alert </div>
				         	<div class="col-lg-2">
								<input type='text' id="txtUserDefinedAlert" name="txtUserDefinedAlert"  value="<%= DelBoardUserAlertColor %>">
							</div>
						</div>
						
		
					
					</div>
				</div>
			</div>
			
			<div class="col-md-6">
				&nbsp;
			</div>
			
		</div>
		
		
		<div class="row">
			<h3><i class="fad fa-comment-exclamation"></i>&nbsp;<%= GetTerm("Routing") %> Nag Alert Settings</h3>
		</div>

		<div class="row">
				
			<div class="col-md-6">
			
					<% If NextStopNagMessageONOFF = 0 Then %>
						<div class="panel panel-danger">
							<div class="panel-heading">
								<h3 class="panel-title">Nag Alert Settings - No Next Stop Selected (OFF)</h3>
					<% Else %>
						<div class="panel panel-success">
							<div class="panel-heading">
								<h3 class="panel-title">Nag Alert Settings - No Next Stop Selected (ON)</h3>
					<% End If %>
			
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
					
					
	                    <font color="Red"><strong><i>These options can also be set for individual users via Manage Users. Individual user settings will override what is set here unless the user is set to Use Global.<br><br>
	                    Insight will begin sending these nag messages after the driver has marked their first delivery.</i></strong></font>
						<div class="row">
						 	<div class="col-lg-12">Send 'nag' messages when a driver does not select the <strong>Next Stop</strong>
								<%
								If NextStopNagMessageONOFF = 0 Then
									Response.Write("<input type='checkbox' id='chkNextStopNagMessageONOFF' name='chkNextStopNagMessageONOFF'>")
								Else
									Response.Write("<input type='checkbox' id='chkNextStopNagMessageONOFF' name='chkNextStopNagMessageONOFF' checked>")
								End If
								%>
							</div>
						</div>
	                    
						<div class="row">
	                    	<div class="col-lg-12">Send when the Next Stop has not been set for 
								<select class="form-control custom-select" id="selNextStopNagMinutes" name="selNextStopNagMinutes">
									<%
										For x = 5 to 180 Step 5 ' 3 hours
											If x mod 60 = 0 Then
												If x = cint(NextStopNagMinutes) Then 
													Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
												else
													Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
												End If
											Else
												If x = cint(NextStopNagMinutes) Then 
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
								<select class="form-control custom-select" id="selNextStopNagIntervalMinutes" name="selNextStopNagIntervalMinutes">
									<%
										For x = 10 to 120 Step 5 ' 2 hours
											If x mod 60 = 0 Then
												If x = cint(NextStopNagIntervalMinutes) Then 
													Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
												else
													Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
												End If
											Else
												If x = cint(NextStopNagIntervalMinutes) Then 
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
								<select class="form-control custom-select" id="selNextStopNagMessageMaxToSendPerStop" name="selNextStopNagMessageMaxToSendPerStop">
									<%
										For x = 1 to 10
											If x = cint(NextStopNagMessageMaxToSendPerStop) Then 
												Response.Write("<option value='" & x & "' selected>" & x & "</option>")
											Else
												Response.Write("<option value='" & x & "'>" & x & "</option>")
											End If
										Next
									%>
								</select>&nbsp;messages each time a 'No Next Stop' event occurs
							</div>
						</div>
					
	
						<div class="row">
							<div class="col-lg-12">Send a maxium of 
								<select class="form-control custom-select"  id="selNextStopNagMessageMaxToSendPerDriverPerDay" name="selNextStopNagMessageMaxToSendPerDriverPerDay">
									<%
										For x = 1 to 25
											If x = cint(NextStopNagMessageMaxToSendPerDriverPerDay) Then 
												Response.Write("<option value='" & x & "' selected>" & x & "</option>")
											Else
												Response.Write("<option value='" & x & "'>" & x & "</option>")
											End If
										Next
									%>
								</select>&nbsp;messages to a driver on any given day
							</div>
						</div>
	
	
						<div class="row">
							<div class="col-lg-12">Send method 
								<select class="form-control custom-select select-large"   id="selNextStopNagMessageSendMethod" name="selNextStopNagMessageSendMethod">
									<option value="Text"<%If NextStopNagMessageSendMethod = "Text" Then Response.Write(" selected ")%>>Text Message Only</option>
									<!--<option value="Email"<%If NextStopNagMessageSendMethod = "Email" Then Response.Write(" selected ")%>>Push Only</option>
									<option value="TextThenEmail"<%If NextStopNagMessageSendMethod = "TextThenEmail" Then Response.Write(" selected ")%>>Text - If no cell number, then Push</option>
									<option value="EmailThenText"<%If NextStopNagMessageSendMethod = "EmailThenText" Then Response.Write(" selected ")%>>Push - If unable, send text</option>
									<option value="Both"<%If NextStopNagMessageSendMethod = "Both" Then Response.Write(" selected ")%>>Both</option> -->
								</select>
							</div>
						</div>
						
					</div>
				</div>
			</div>
			
			<div class="col-md-6">
			
				<% If NoActivityNagMessageONOFF = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title">Nag Alert Settings - No Activity (OFF)</h3>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title">Nag Alert Settings - No Activity (ON)</h3>
				<% End If %>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
							
		                    <font color="Red"><strong><i>*These options can also be set for individual users via Manage Users. Individual user settings will override what is set here unless the user is set to Use Global.<br><br>
		                    Insight will begin sending these nag messages after the driver has marked their first delivery, selected their first stop or the time of day has reached the time set below.</i></strong></font>
		
							<div class="row">
							 	<div class="col-lg-12">Send 'nag' messages when there has been a period of <strong>No Activity</strong>
									<%
									If NoActivityNagMessageONOFF = 0 Then
										Response.Write("<input type='checkbox' id='chkNoActivityNagMessageONOFF' name='chkNoActivityNagMessageONOFF'>")
									Else
										Response.Write("<input type='checkbox' id='chkNoActivityNagMessageONOFF' name='chkNoActivityNagMessageONOFF' checked>")
									End If
									%>
								</div>
							</div>
							
							<div class="row">
								<div class="col-lg-12">Start sending messages if there has been No Activity by 
									<select class="form-control custom-select" id="selNoActivityNagTimeOfDay" name="selNoActivityNagTimeOfDay">			
										<option value="12:00"<%If NoActivityNagTimeOfDay = "12:00" Then Response.Write(" selected ")%>>-Midnight-</option>
										<option value="12:15"<%If NoActivityNagTimeOfDay = "12:15" Then Response.Write(" selected ")%>>12:15 AM</option>
										<option value="12:30"<%If NoActivityNagTimeOfDay = "12:30" Then Response.Write(" selected ")%>>12:30 AM</option>
										<option value="12:45"<%If NoActivityNagTimeOfDay = "12:45" Then Response.Write(" selected ")%>>12:45 AM</option>
										<option value="1:00"<%If NoActivityNagTimeOfDay = "1:00" Then Response.Write(" selected ")%>>1:00 AM</option>
										<option value="1:15"<%If NoActivityNagTimeOfDay = "1:15" Then Response.Write(" selected ")%>>1:15 AM</option>
										<option value="1:30"<%If NoActivityNagTimeOfDay = "1:30" Then Response.Write(" selected ")%>>1:30 AM</option>
										<option value="1:45"<%If NoActivityNagTimeOfDay = "1:45" Then Response.Write(" selected ")%>>1:45 AM</option>
										<option value="2:00"<%If NoActivityNagTimeOfDay = "2:00" Then Response.Write(" selected ")%>>2:00 AM</option>
										<option value="2:15"<%If NoActivityNagTimeOfDay = "2:15" Then Response.Write(" selected ")%>>2:15 AM</option>
										<option value="2:30"<%If NoActivityNagTimeOfDay = "2:30" Then Response.Write(" selected ")%>>2:30 AM</option>
										<option value="2:45"<%If NoActivityNagTimeOfDay = "2:45" Then Response.Write(" selected ")%>>2:45 AM</option>
										<option value="3:00"<%If NoActivityNagTimeOfDay = "3:00" Then Response.Write(" selected ")%>>3:00 AM</option>
										<option value="3:15"<%If NoActivityNagTimeOfDay = "3:15" Then Response.Write(" selected ")%>>3:15 AM</option>
										<option value="3:30"<%If NoActivityNagTimeOfDay = "3:30" Then Response.Write(" selected ")%>>3:30 AM</option>
										<option value="3:45"<%If NoActivityNagTimeOfDay = "3:45" Then Response.Write(" selected ")%>>3:45 AM</option>
										<option value="4:00"<%If NoActivityNagTimeOfDay = "4:00" Then Response.Write(" selected ")%>>4:00 AM</option>
										<option value="4:15"<%If NoActivityNagTimeOfDay = "4:15" Then Response.Write(" selected ")%>>4:15 AM</option>
										<option value="4:30"<%If NoActivityNagTimeOfDay = "4:30" Then Response.Write(" selected ")%>>4:30 AM</option>
										<option value="4:45"<%If NoActivityNagTimeOfDay = "4:45" Then Response.Write(" selected ")%>>4:45 AM</option>
										<option value="5:00"<%If NoActivityNagTimeOfDay = "5:00" Then Response.Write(" selected ")%>>5:00 AM</option>
										<option value="5:15"<%If NoActivityNagTimeOfDay = "5:15" Then Response.Write(" selected ")%>>5:15 AM</option>
										<option value="5:30"<%If NoActivityNagTimeOfDay = "5:30" Then Response.Write(" selected ")%>>5:30 AM</option>
										<option value="5:45"<%If NoActivityNagTimeOfDay = "5:45" Then Response.Write(" selected ")%>>5:45 AM</option>
										<option value="6:00"<%If NoActivityNagTimeOfDay = "6:00" Then Response.Write(" selected ")%>>6:00 AM</option>
										<option value="6:15"<%If NoActivityNagTimeOfDay = "6:15" Then Response.Write(" selected ")%>>6:15 AM</option>
										<option value="6:30"<%If NoActivityNagTimeOfDay = "6:30" Then Response.Write(" selected ")%>>6:30 AM</option>
										<option value="6:45"<%If NoActivityNagTimeOfDay = "6:45" Then Response.Write(" selected ")%>>6:45 AM</option>
										<option value="7:00"<%If NoActivityNagTimeOfDay = "7:00" Then Response.Write(" selected ")%>>7:00 AM</option>
										<option value="7:15"<%If NoActivityNagTimeOfDay = "7:15" Then Response.Write(" selected ")%>>7:15 AM</option>
										<option value="7:30"<%If NoActivityNagTimeOfDay = "7:30" Then Response.Write(" selected ")%>>7:30 AM</option>
										<option value="7:45"<%If NoActivityNagTimeOfDay = "7:45" Then Response.Write(" selected ")%>>7:45 AM</option>
										<option value="8:00"<%If NoActivityNagTimeOfDay = "8:00" Then Response.Write(" selected ")%>>8:00 AM</option>
										<option value="8:15"<%If NoActivityNagTimeOfDay = "8:15" Then Response.Write(" selected ")%>>8:15 AM</option>
										<option value="8:30"<%If NoActivityNagTimeOfDay = "8:30" Then Response.Write(" selected ")%>>8:30 AM</option>
										<option value="8:45"<%If NoActivityNagTimeOfDay = "8:45" Then Response.Write(" selected ")%>>8:45 AM</option>
										<option value="9:00"<%If NoActivityNagTimeOfDay = "9:00" Then Response.Write(" selected ")%>>9:00 AM</option>
										<option value="9:15"<%If NoActivityNagTimeOfDay = "9:15" Then Response.Write(" selected ")%>>9:15 AM</option>
										<option value="9:30"<%If NoActivityNagTimeOfDay = "9:30" Then Response.Write(" selected ")%>>9:30 AM</option>
										<option value="9:45"<%If NoActivityNagTimeOfDay = "9:45" Then Response.Write(" selected ")%>>9:45 AM</option>
										<option value="10:00"<%If NoActivityNagTimeOfDay = "10:00" Then Response.Write(" selected ")%>>10:00 AM</option>
										<option value="10:15"<%If NoActivityNagTimeOfDay = "10:15" Then Response.Write(" selected ")%>>10:15 AM</option>
										<option value="10:30"<%If NoActivityNagTimeOfDay = "10:30" Then Response.Write(" selected ")%>>10:30 AM</option>
										<option value="10:45"<%If NoActivityNagTimeOfDay = "10:45" Then Response.Write(" selected ")%>>10:45 AM</option>
										<option value="11:00"<%If NoActivityNagTimeOfDay = "11:00" Then Response.Write(" selected ")%>>11:00 AM</option>
										<option value="11:15"<%If NoActivityNagTimeOfDay = "11:15" Then Response.Write(" selected ")%>>11:15 AM</option>
										<option value="11:30"<%If NoActivityNagTimeOfDay = "11:30" Then Response.Write(" selected ")%>>11:30 AM</option>
										<option value="11:45"<%If NoActivityNagTimeOfDay = "11:45" Then Response.Write(" selected ")%>>11:45 AM</option>
										<option value="12:00"<%If NoActivityNagTimeOfDay = "12:00" Then Response.Write(" selected ")%>>-Noon-</option>
										<option value="12:15"<%If NoActivityNagTimeOfDay = "12:15" Then Response.Write(" selected ")%>>12:15 PM</option>
										<option value="12:30"<%If NoActivityNagTimeOfDay = "12:30" Then Response.Write(" selected ")%>>12:30 PM</option>
										<option value="12:45"<%If NoActivityNagTimeOfDay = "12:45" Then Response.Write(" selected ")%>>12:45 PM</option>
										<option value="13:00"<%If NoActivityNagTimeOfDay = "13:00" Then Response.Write(" selected ")%>>1:00 PM</option>
										<option value="13:15"<%If NoActivityNagTimeOfDay = "13:15" Then Response.Write(" selected ")%>>1:15 PM</option>
										<option value="13:30"<%If NoActivityNagTimeOfDay = "13:30" Then Response.Write(" selected ")%>>1:30 PM</option>
										<option value="13:45"<%If NoActivityNagTimeOfDay = "13:45" Then Response.Write(" selected ")%>>1:45 PM</option>
										<option value="14:00"<%If NoActivityNagTimeOfDay = "14:00" Then Response.Write(" selected ")%>>2:00 PM</option>
										<option value="14:15"<%If NoActivityNagTimeOfDay = "14:15" Then Response.Write(" selected ")%>>2:15 PM</option>
										<option value="14:30"<%If NoActivityNagTimeOfDay = "14:30" Then Response.Write(" selected ")%>>2:30 PM</option>
										<option value="14:45"<%If NoActivityNagTimeOfDay = "14:45" Then Response.Write(" selected ")%>>2:45 PM</option>
										<option value="15:00"<%If NoActivityNagTimeOfDay = "15:00" Then Response.Write(" selected ")%>>3:00 PM</option>
										<option value="15:15"<%If NoActivityNagTimeOfDay = "15:15" Then Response.Write(" selected ")%>>3:15 PM</option>
										<option value="15:30"<%If NoActivityNagTimeOfDay = "15:30" Then Response.Write(" selected ")%>>3:30 PM</option>
										<option value="15:45"<%If NoActivityNagTimeOfDay = "15:45" Then Response.Write(" selected ")%>>3:45 PM</option>
										<option value="16:00"<%If NoActivityNagTimeOfDay = "16:00" Then Response.Write(" selected ")%>>4:00 PM</option>
										<option value="16:15"<%If NoActivityNagTimeOfDay = "16:15" Then Response.Write(" selected ")%>>4:15 PM</option>
										<option value="16:30"<%If NoActivityNagTimeOfDay = "16:30" Then Response.Write(" selected ")%>>4:30 PM</option>
										<option value="16:45"<%If NoActivityNagTimeOfDay = "16:45" Then Response.Write(" selected ")%>>4:45 PM</option>
										<option value="17:00"<%If NoActivityNagTimeOfDay = "17:00" Then Response.Write(" selected ")%>>5:00 PM</option>
										<option value="17:15"<%If NoActivityNagTimeOfDay = "17:15" Then Response.Write(" selected ")%>>5:15 PM</option>
										<option value="17:30"<%If NoActivityNagTimeOfDay = "17:30" Then Response.Write(" selected ")%>>5:30 PM</option>
										<option value="17:45"<%If NoActivityNagTimeOfDay = "17:45" Then Response.Write(" selected ")%>>5:45 PM</option>
										<option value="18:00"<%If NoActivityNagTimeOfDay = "18:00" Then Response.Write(" selected ")%>>6:00 PM</option>
										<option value="18:15"<%If NoActivityNagTimeOfDay = "18:15" Then Response.Write(" selected ")%>>6:15 PM</option>
										<option value="18:30"<%If NoActivityNagTimeOfDay = "18:30" Then Response.Write(" selected ")%>>6:30 PM</option>
										<option value="18:45"<%If NoActivityNagTimeOfDay = "18:45" Then Response.Write(" selected ")%>>6:45 PM</option>
										<option value="19:00"<%If NoActivityNagTimeOfDay = "19:00" Then Response.Write(" selected ")%>>7:00 PM</option>
										<option value="19:15"<%If NoActivityNagTimeOfDay = "19:15" Then Response.Write(" selected ")%>>7:15 PM</option>
										<option value="19:30"<%If NoActivityNagTimeOfDay = "19:30" Then Response.Write(" selected ")%>>7:30 PM</option>
										<option value="19:45"<%If NoActivityNagTimeOfDay = "19:45" Then Response.Write(" selected ")%>>7:45 PM</option>
										<option value="20:00"<%If NoActivityNagTimeOfDay = "20:00" Then Response.Write(" selected ")%>>8:00 PM</option>
										<option value="20:15"<%If NoActivityNagTimeOfDay = "20:15" Then Response.Write(" selected ")%>>8:15 PM</option>
										<option value="20:30"<%If NoActivityNagTimeOfDay = "20:30" Then Response.Write(" selected ")%>>8:30 PM</option>
										<option value="20:45"<%If NoActivityNagTimeOfDay = "20:45" Then Response.Write(" selected ")%>>8:45 PM</option>
										<option value="21:00"<%If NoActivityNagTimeOfDay = "21:00" Then Response.Write(" selected ")%>>9:00 PM</option>
										<option value="21:15"<%If NoActivityNagTimeOfDay = "21:15" Then Response.Write(" selected ")%>>9:15 PM</option>
										<option value="21:30"<%If NoActivityNagTimeOfDay = "21:30" Then Response.Write(" selected ")%>>9:30 PM</option>
										<option value="21:45"<%If NoActivityNagTimeOfDay = "21:45" Then Response.Write(" selected ")%>>9:45 PM</option>
										<option value="22:00"<%If NoActivityNagTimeOfDay = "22:00" Then Response.Write(" selected ")%>>10:00 PM</option>
										<option value="22:15"<%If NoActivityNagTimeOfDay = "22:15" Then Response.Write(" selected ")%>>10:15 PM</option>
										<option value="22:30"<%If NoActivityNagTimeOfDay = "22:30" Then Response.Write(" selected ")%>>10:30 PM</option>
										<option value="22:45"<%If NoActivityNagTimeOfDay = "22:45" Then Response.Write(" selected ")%>>10:45 PM</option>
										<option value="23:00"<%If NoActivityNagTimeOfDay = "23:00" Then Response.Write(" selected ")%>>11:00 PM</option>
										<option value="23:15"<%If NoActivityNagTimeOfDay = "23:15" Then Response.Write(" selected ")%>>11:15 PM</option>
										<option value="23:30"<%If NoActivityNagTimeOfDay = "23:30" Then Response.Write(" selected ")%>>11:30 PM</option>
										<option value="23:45"<%If NoActivityNagTimeOfDay = "23:45" Then Response.Write(" selected ")%>>11:45 PM</option>	
				 					</select>
								</div>
							</div>
							
							<div class="row">
		                    	<div class="col-lg-12">Send when there has been No Activity for 
									<select class="form-control custom-select" id="selNoActivityNagMinutes" name="selNoActivityNagMinutes">
										<%
											For x = 15 to 180 Step 5 ' 3 hours
												If x mod 60 = 0 Then
													If x = cint(NoActivityNagMinutes) Then 
														Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
													else
														Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
													End If
												Else
													If x = cint(NoActivityNagMinutes) Then 
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
									<select class="form-control custom-select" id="selNoActivityNagIntervalMinutes" name="selNoActivityNagIntervalMinutes">
										<%
											For x = 10 to 120 Step 5 ' 2 hours
												If x mod 60 = 0 Then
													If x = cint(NoActivityNagIntervalMinutes) Then 
														Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
													else
														Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
													End If
												Else
													If x = cint(NoActivityNagIntervalMinutes) Then 
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
									<select class="form-control custom-select" id="selNoActivityNagMessageMaxToSendPerStop" name="selNoActivityNagMessageMaxToSendPerStop">
										<%
											For x = 1 to 10
												If x = cint(NoActivityNagMessageMaxToSendPerStop) Then 
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
									<select class="form-control custom-select"  id="selNoActivityNagMessageMaxToSendPerDriverPerDay" name="selNoActivityNagMessageMaxToSendPerDriverPerDay">
										<%
											For x = 1 to 25
												If x = cint(NoActivityNagMessageMaxToSendPerDriverPerDay) Then 
													Response.Write("<option value='" & x & "' selected>" & x & "</option>")
												Else
													Response.Write("<option value='" & x & "'>" & x & "</option>")
												End If
											Next
										%>
									</select>&nbsp;messages to a driver on any given day
								</div>
							</div>
		
							<div class="row">
								<div class="col-lg-12">Send method 
									<select class="form-control custom-select select-large"   id="selNoActivityNagMessageSendMethod" name="selNoActivityNagMessageSendMethod">
										<option value="Text"<%If NoActivityNagMessageSendMethod = "Text" Then Response.Write(" selected ")%>>Text Message Only</option>
									<!--	<option value="Email"<%If NoActivityNagMessageSendMethod = "Email" Then Response.Write(" selected ")%>>Email Only</option>
										<option value="TextThenEmail"<%If NoActivityNagMessageSendMethod = "TextThenEmail" Then Response.Write(" selected ")%>>Text - If no cell number, send email</option>
										<option value="EmailThenText"<%If NoActivityNagMessageSendMethod = "EmailThenText" Then Response.Write(" selected ")%>>Email - If no valid email address, send text</option>
										<option value="Both"<%If NoActivityNagMessageSendMethod = "Both" Then Response.Write(" selected ")%>>Both</option>-->
									</select>
								</div>
							</div>

					
					</div>
				</div>
			</div>
			
		</div>
		
		<div class="row">
			<h3><i class="fad fa-truck"></i>&nbsp;<%= GetTerm("Route") %> Selection Settings</h3>
		</div>
		
			
	    <div class="row">
	    
			<div class="col-md-6">
				<div class="panel panel-warning">
					<div class="panel-heading">
						<h3 class="panel-title">Routes To Ignore</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
							
			         	<div class="row" style="margin-right:0px">
							<div class="form-control routes-to-ignore" style="height: auto; margin:10px;">
								<% 'Get all trucks
									SQL9 = "SELECT DISTINCT TruckID FROM RT_Truck"
									Set cnn9 = Server.CreateObject("ADODB.Connection")
									cnn9.open (Session("ClientCnnString"))
									Set rs9 = Server.CreateObject("ADODB.Recordset")
									rs9.CursorLocation = 3 
									Set rs9 = cnn9.Execute(SQL9)
									If not rs9.EOF Then
										If DelBoardRoutesToIgnore <> "" Then chkIgnoreArray = Split(DelBoardRoutesToIgnore,",")
										Do
											checked = ""
											If DelBoardRoutesToIgnore <> "" Then
												For i = 0 to Ubound(chkIgnoreArray) 
													If trim(rs9("TruckID")) = trim(chkIgnoreArray(i)) Then checked = " checked "
												Next
											End If
											Response.Write( "<label class='btn btn-default btn-xs greenBg' style='width: 85px; text-align: left; font-size: 14px; margin: 0 3px 3px 0;'><input "& checked &" type='checkbox' name='txtDelBoardRoutesToIgnore' id='txtDelBoardRoutesToIgnore' value='"&rs9("TruckID")&"'> "& rs9("TruckID") & "<br>" & GetDriverNameByTruckID(rs9("TruckID"))&"</label>")
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
				</div>
			</div>
			
			<div class="col-md-6">
				<div class="panel panel-warning">
					<div class="panel-heading">
						<h3 class="panel-title">UPS Routes</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">
							
			         	<div class="row" style="margin-right:0px">
							<div class="form-control ups-routes" style="height: auto; margin:10px;">
								<% 'Get all trucks
									SQL9 = "SELECT DISTINCT TruckID FROM RT_Truck"
									Set cnn9 = Server.CreateObject("ADODB.Connection")
									cnn9.open (Session("ClientCnnString"))
									Set rs9 = Server.CreateObject("ADODB.Recordset")
									rs9.CursorLocation = 3 
									Set rs9 = cnn9.Execute(SQL9)
									If not rs9.EOF Then
										If DelBoardUPSRoutes <> "" Then chkUPSArray = Split(DelBoardUPSRoutes,",")
										Do
											checked = ""
											If DelBoardUPSRoutes <> "" Then
												For i = 0 to Ubound(chkUPSArray) 
													If trim(rs9("TruckID")) = trim(chkUPSArray(i)) Then checked = " checked "
												Next
											End If
											Response.Write( "<label class='btn btn-default btn-xs' style='width: 85px; text-align: left; font-size: 14px; margin: 0 3px 3px 0;'><input "& checked &" type='checkbox' name='txtDelBoardUPSRoutes' id='txtDelBoardUPSRoutes' value='"&rs9("TruckID")&"'> "& rs9("TruckID") & "<br>" & GetDriverNameByTruckID(rs9("TruckID"))&"</label>")
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


<!-- spectrum color picker js !-->
<script>

$("#txtScheduledColor").spectrum({
    color: '<%= DelBoardScheduledColor %>',
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

$("#txtCompletedDeliveries").spectrum({
    color: '<%= DelBoardCompletedColor %>',
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


$("#txtInProgress").spectrum({
    color: '<%= DelBoardInProgressColor %>',
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

$("#txtSkippedDeliveries").spectrum({
    color: '<%= DelBoardSkippedColor %>',
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

$("#txtAboveProfit").spectrum({
    color: '<%=DelBoardAtOrAboveProfitColor%>',
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

$("#txtNextStop").spectrum({
    color: '<%=DelBoardNextStopColor%>',
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


$("#txtDelBoardAMColor").spectrum({
    color: '<%=DelBoardAMColor%>',
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



$("#txtDelBoardPriorityColor").spectrum({
    color: '<%=DelBoardPriorityColor%>',
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


$("#txtDelBoardPieTimerColor").spectrum({
    color: '<%= DelBoardPieTimerColor %>',
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

$("#txtBelowProfit").spectrum({
    color: '<%=DelBoardBelowProfitColor%>',
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


$("#txtUserDefinedAlert").spectrum({
    color: '<%=DelBoardUserAlertColor%>',
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

$("#txtDelBoardTitleGradientColor").spectrum({
    color: '<%=DelBoardTitleGradientColor%>',
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

$("#txtDelBoardTitleTextFontColor").spectrum({
    color: '<%=DelBoardTitleTextFontColor%>',
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


<!--#include file="../../../inc/footer-main.asp"-->
