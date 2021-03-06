<!--#include file="../../../inc/header.asp"-->

<!-- bootstrap timepicker !-->
<script type="text/javascript" src="http://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>	
<link href="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.css" rel="stylesheet" type="text/css">
<script src="<%= baseURL %>js/bootstrap-timepicker/bootstrap-timepicker.js" type="text/javascript"></script>
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
						
			$('#modalAutomaticCustomerAnalysisSummary1ReportScheduler').on('show.bs.modal', function(e) {
			    	    
			    var $modal = $(this);
		
		    	$.ajax({
					type:"POST",
					url: "../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
					cache: false,
					data: "action=GetContentForAutomaticCustomerAnalysisSummary1ReportScheduler",
					success: function(response)
					 {
		               	 $modal.find('#modalAutomaticCustomerAnalysisSummary1ReportSchedulerContent').html(response);               	 
		             },
		             failure: function(response)
					 {
					  	$modal.find('#modalAutomaticCustomerAnalysisSummary1ReportSchedulerContent').html("Failed");
			            //var height = $(window).height() - 600;
			            //$(this).find(".modal-body").css("max-height", height);
		             }
				});
				
			});	
			
			$('#modalMCSActivityReportScheduler').on('show.bs.modal', function(e) {
			    	    
			    var $modal = $(this);
		
		    	$.ajax({
					type:"POST",
					url: "../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
					cache: false,
					data: "action=GetContentForMCSActivityReportScheduler",
					success: function(response)
					 {
		               	 $modal.find('#modalMCSActivityReportSchedulerContent').html(response);               	 
		             },
		             failure: function(response)
					 {
					  	$modal.find('#modalMCSActivityReportSchedulerContent').html("Failed");
			            //var height = $(window).height() - 600;
			            //$(this).find(".modal-body").css("max-height", height);
		             }
				});
				
			});	
			
           $('#modalCustAnalSum1EmailAddressesToCC').on('show.bs.modal', function (e) {
                var $modal = $(this);
            });


            $('#modalMCSActivitySummaryEmailAddressesToCC').on('show.bs.modal', function (e) {
                var $modal = $(this);
            });
	            
	            
			$('#lstExistingCustAnalSum1EmailToUserIDs').multiselect({
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
				nonSelectedText:'No Users Selected For Automatic Customer Analysis Summary 1',
				numberDisplayed: 20,
			    onChange: function() {
			        var selected = this.$select.val();
			        $("#lstSelectedCustAnalSum1EmailToUserIDs").val(selected);
			        console.log(selected);
			        // ...
			    }
	    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current daily API activity report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedCustAnalSum1EmailToUserIDs").val();
			
			if (data) {
				//Make an array
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingCustAnalSum1EmailToUserIDs").val(dataarray);
				// Then refresh
				$("#lstExistingCustAnalSum1EmailToUserIDs").multiselect("refresh");
			}
			//*************************************************************************************************
	            
	            
			$('#lstExistingCustAnalSum1EmailToUserIDsCC').multiselect({
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
				nonSelectedText:'No Users Selected To CC Customer Analysis Summary 1',
				numberDisplayed: 20,
			    onChange: function() {
			        var selected = this.$select.val();
			        $("#lstSelectedCustAnalSum1EmailToUserIDsCC").val(selected);
			        console.log(selected);
			        // ...
			    }
	    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current daily API activity report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedCustAnalSum1EmailToUserIDsCC").val();
			
			if (data) {
				//Make an array
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingCustAnalSum1EmailToUserIDsCC").val(dataarray);
				// Then refresh
				$("#lstExistingCustAnalSum1EmailToUserIDsCC").multiselect("refresh");
			}
			//*************************************************************************************************
			
			
			$('#lstExistingMCSAnalysisCCUserIDs').multiselect({
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
				nonSelectedText:'No Users Selected To CC On MCS Analysis',
				numberDisplayed: 20,
			    onChange: function() {
			        var selected = this.$select.val();
			        $("#lstSelectedMCSAnalysisCCUserIDs").val(selected);
			        console.log(selected);
			        // ...
			    }
	    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current daily API activity report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedMCSAnalysisCCUserIDs").val();
			
			if (data) {
				//Make an array
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingMCSAnalysisCCUserIDs").val(dataarray);
				// Then refresh
				$("#lstExistingMCSAnalysisCCUserIDs").multiselect("refresh");
			}
			//*************************************************************************************************
			
	 

			$('#lstExistingMCSActivitySummaryEmailToUserIDs').multiselect({
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
				nonSelectedText:'No Users Selected To CC On MCS Activity Summary',
				numberDisplayed: 20,
			    onChange: function() {
			        var selected = this.$select.val();
			        $("#lstSelectedMCSActivitySummaryEmailToUserIDs").val(selected);
			        console.log(selected);
			        // ...
			    }
	    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current daily API activity report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedMCSActivitySummaryEmailToUserIDs").val();
			
			if (data) {
				//Make an array
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingMCSActivitySummaryEmailToUserIDs").val(dataarray);
				// Then refresh
				$("#lstExistingMCSActivitySummaryEmailToUserIDs").multiselect("refresh");
			}
			//*************************************************************************************************
   

	 

			$('#lstExistingMCSActivitySummaryCCToUserIDs').multiselect({
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
				nonSelectedText:'No Users Selected To CC On MCS Activity Summary',
				numberDisplayed: 20,
			    onChange: function() {
			        var selected = this.$select.val();
			        $("#lstSelectedMCSActivitySummaryCCToUserIDs").val(selected);
			        console.log(selected);
			        // ...
			    }
	    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current daily API activity report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedMCSActivitySummaryCCToUserIDs").val();
			
			if (data) {
				//Make an array
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingMCSActivitySummaryCCToUserIDs").val(dataarray);
				// Then refresh
				$("#lstExistingMCSActivitySummaryCCToUserIDs").multiselect("refresh");
			}
			//*************************************************************************************************
        

			
		});
	</script>

<%
	SQL = "SELECT * FROM Settings_BizIntel"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		CustAnalSum1OnOff = rs("CustAnalSum1OnOff")
		CustAnalSum1EmailToUserNos = rs("CustAnalSum1EmailToUserNos")
		CustAnalSum1UserNosToCC = rs("CustAnalSum1UserNosToCC")
		CustAnalSum1EmailAddressesToCC = rs("CustAnalSum1EmailAddressesToCC")
		MCSUserNosToCC = rs("MCSUserNosToCC")
		MCSActivitySummaryOnOff = rs("MCSActivitySummaryOnOff")
		MCSActivitySummaryEmailToUserNos = rs("MCSActivitySummaryEmailToUserNos")
		MCSActivitySummaryUserNosToCC = rs("MCSActivitySummaryUserNosToCC")
		MCSActivitySummaryEmailAddressesToCC = rs("MCSActivitySummaryEmailAddressesToCC")
		MCSUseAlternateHeader = rs("MCSUseAlternateHeader")
	End If
				
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
%>

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
	
	.line-full{
	 	margin-bottom:20px;
	}
	
	.multi-select{
		min-height:200px;
		min-width:170px;
	}
	.form-control {
		padding: 6px 5px;
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


<h1 class="page-header"><i class="fa fa-graduation-cap"></i>&nbsp;Business Intelligence 
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>


<form method="post" action="bizintel-submit.asp" name="frmBizIntel" id="frmBizIntel">


	<div class="container">
		
		<%
			Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
			Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
			Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
			Response.Write("</div>")
			Response.Flush()
		%>
	
		<div class="row">
			<h3><i class="fad fa-file-pdf"></i>&nbsp;Business Intelligence Report Settings</h3>
		</div>
	
		<div class="row">
		
			<div class="col-md-4">
				<% If CustAnalSum1OnOff = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title">Automatic Customer Analysis Summary 1 (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title">Automatic Customer Analysis Summary 1 (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
					<div class="panel-body">

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON 
					      		<%
					      		If CustAnalSum1OnOff = 0 Then
									Response.Write("<input type='checkbox' id='chkCustAnalSum1OnOff' name='chkCustAnalSum1OnOff'")
								Else
									Response.Write("<input type='checkbox' id='chkCustAnalSum1OnOff' name='chkCustAnalSum1OnOff' checked")
								End If
								Response.Write(">")
								%>
								<br><small>If turned on, this report is automatically run every Monday.</small>
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalAutomaticCustomerAnalysisSummary1ReportScheduler" data-tooltip="true" data-title="Automatic Customer Analysis Summary 1 Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> Automatic Customer Analysis Summary 1 Report Scheduler</button>
										</li>
										<li>
											<p>Select users <i class="fad fa-user-friends"></i> to send the report to:</p>
											<input type="hidden" name="lstSelectedCustAnalSum1EmailToUserIDs" id="lstSelectedCustAnalSum1EmailToUserIDs" value="<%= CustAnalSum1EmailToUserNos %>">
											<select id="lstExistingCustAnalSum1EmailToUserIDs" multiple="multiple" name="lstExistingCustAnalSum1EmailToUserIDs">
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
											<p>Select additional email addresses to CC the report to:</p>
											<p><small>&nbsp;(CC:'s will receive a separate email for each recipient of this report)</small></p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalCustAnalSum1EmailAddressesToCC" data-tooltip="true" data-title="Additional emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails To CC</button>						
				             				<% If CustAnalSum1EmailAddressesToCC <> "" Then %>
				             					<p style="margin-top:20px;"><strong>Current Additional CC Emails:</strong> <%= CustAnalSum1EmailAddressesToCC %></p>
				             				<% End If %>
										</li>
										<li>
											<p>Select additonal users <i class="fad fa-user-friends"></i> to CC the report to:</p>
											<p><small>&nbsp;(CC:'s will receive a separate email for each recipient of this report)</small></p>
											<input type="hidden" name="lstSelectedCustAnalSum1EmailToUserIDsCC" id="lstSelectedCustAnalSum1EmailToUserIDsCC" value="<%= CustAnalSum1UserNosToCC %>">
											<select id="lstExistingCustAnalSum1EmailToUserIDsCC" multiple="multiple" name="lstExistingCustAnalSum1EmailToUserIDsCC">
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
										
									</ol>
								</div>
					
							</div>
						</div>
					
					
					</div>
				</div>
			</div>	
				
				
				
				
				
		
			<div class="col-md-4">
				<div class="panel panel-success">
					<div class="panel-heading">
						<h3 class="panel-title">MCS Analysis Report</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>				

					<div class="panel-body">
			
					
					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
								
								<div class="text-element circles-list">
									<ol>
										<li>
											<p><strong>When any actions results in the MCS system sending an email, automatically CC: the following people on the email.</strong></p> 
											<p><strong>(The user performing the MCS action will <em>NOT</em>&nbsp;&nbsp;be able to remove these CC's.)</strong></p>
											<input type="hidden" name="lstSelectedMCSAnalysisCCUserIDs" id="lstSelectedMCSAnalysisCCUserIDs" value="<%= MCSUserNosToCC %>">
											<select id="lstExistingMCSAnalysisCCUserIDs" multiple="multiple" name="lstExistingMCSAnalysisCCUserIDs">
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
									</ol>
								</div>
					
							</div>
						</div>
					
					
					</div>
				</div>	
				
			</div>

			<div class="col-md-4">
				<% If MCSActivitySummaryOnOff = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title">MCS Activity Report (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title">MCS Activity Report (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
					<div class="panel-body">

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON 
					      		<%
					      		If MCSActivitySummaryOnOff = 0 Then
									Response.Write("<input type='checkbox' id='chkMCSActivitySummaryOnOff' name='chkMCSActivitySummaryOnOff'")
								Else
									Response.Write("<input type='checkbox' id='chkMCSActivitySummaryOnOff' name='chkMCSActivitySummaryOnOff' checked")
								End If
								Response.Write(">")
								%>
				            </div>
				            <!-- eof line -->
				         </div>  
				         					

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	USE ALTERNATE HEADER FORMAT 
					      		<%
					      		If MCSUseAlternateHeader = 0 Then
									Response.Write("<input type='checkbox' id='chkMCSUseAlternateHeader' name='chkMCSUseAlternateHeader'")
								Else
									Response.Write("<input type='checkbox' id='chkMCSUseAlternateHeader' name='chkMCSUseAlternateHeader' checked")
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalMCSActivityReportScheduler" data-tooltip="true" data-title="MCS Activity Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> MCS Activity Report Scheduler</button>
										</li>
										<li>
											<p>Select users <i class="fad fa-user-friends"></i> to send the report to:</p>
											<input type="hidden" name="lstSelectedMCSActivitySummaryEmailToUserIDs" id="lstSelectedMCSActivitySummaryEmailToUserIDs" value="<%= MCSActivitySummaryEmailToUserNos %>">
											<select id="lstExistingMCSActivitySummaryEmailToUserIDs" multiple="multiple" name="lstExistingMCSActivitySummaryEmailToUserIDs">
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
											<p>Select additional email addresses to CC the report to:</p>
											<p><small>&nbsp;(CC:'s will receive a separate email for each recipient of this report)</small></p>
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalMCSActivitySummaryEmailAddressesToCC" data-tooltip="true" data-title="Additional CC Emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails To CC</button>						
				             				<% If MCSActivitySummaryEmailAddressesToCC <> "" Then %>
				             					<p style="margin-top:20px;"><strong>Current Additional CC Emails:</strong> <%= MCSActivitySummaryEmailAddressesToCC %></p>
				             				<% End If %>
										</li>
										<li>
											<p>Select additonal users <i class="fad fa-user-friends"></i> to CC the report to:</p>
											<p><small>&nbsp;(CC:'s will receive a separate email for each recipient of this report)</small></p>
											<input type="hidden" name="lstSelectedMCSActivitySummaryCCToUserIDs" id="lstSelectedMCSActivitySummaryCCToUserIDs" value="<%= MCSActivitySummaryUserNosToCC %>">
											<select id="lstExistingMCSActivitySummaryCCToUserIDs" multiple="multiple" name="lstExistingMCSActivitySummaryCCToUserIDs">
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
										
									</ol>
								</div>
					
							</div>
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
		</div>				
		
	</div>
</form>


<!--#include file="bizintel-modals.asp"-->

<!--#include file="../../../inc/footer-main.asp"-->

