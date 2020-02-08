<!--#include file="../../../../inc/header.asp"-->

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
					
		$('#modalEquipmentNeedToKnowReportScheduler').on('show.bs.modal', function(e) {
		    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
				cache: false,
				data: "action=GetContentForEquipmentN2KReportScheduler",
				success: function(response)
				 {
	               	 $modal.find('#modalEquipmentNeedToKnowReportSchedulerContent').html(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalEquipmentNeedToKnowReportSchedulerContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
			
		});
		

        $('#modalN2KEquipmentEmailAddressesToCC').on('show.bs.modal', function (e) {
            var $modal = $(this);
        });
		

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
						

            
		$('#lstExistingN2KAPIEmailToUserNos').multiselect({
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
			nonSelectedText:'No Users Selected For Need To Know Report',
			numberDisplayed: 20,
		    onChange: function() {
		        var selected = this.$select.val();
		        $("#lstSelectedN2KAPIEmailToUserNos").val(selected);
		        console.log(selected);
		        // ...
		    }
    			
	    });	
	    
		//*************************************************************************************************
		//Load the bootstrap multiselect box with the current daily API activity report users preselected
		//*************************************************************************************************
		var data= $("#lstSelectedN2KAPIEmailToUserNos").val();
		
		if (data) {
			//Make an array
			var dataarray=data.split(",");
			// Set the value
			$("#lstExistingN2KAPIEmailToUserNos").val(dataarray);
			// Then refresh
			$("#lstExistingN2KAPIEmailToUserNos").multiselect("refresh");
		}
		//*************************************************************************************************


            
		$('#lstExistingN2KAPIUserNosToCC').multiselect({
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
			nonSelectedText:'No Users Selected To CC Report To',
			numberDisplayed: 20,
		    onChange: function() {
		        var selected = this.$select.val();
		        $("#lstSelectedN2KAPIUserNosToCC").val(selected);
		        console.log(selected);
		        // ...
		    }
    			
	    });	
	    
		//*************************************************************************************************
		//Load the bootstrap multiselect box with the current daily API activity report users preselected
		//*************************************************************************************************
		var data= $("#lstSelectedN2KAPIUserNosToCC").val();
		
		if (data) {
			//Make an array
			var dataarray=data.split(",");
			// Set the value
			$("#lstExistingN2KAPIUserNosToCC").val(dataarray);
			// Then refresh
			$("#lstExistingN2KAPIUserNosToCC").multiselect("refresh");
		}
		//*************************************************************************************************

		
		
	});
</script>


<%
	SQL = "SELECT * FROM Settings_NeedToKnow"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		N2KEquipmentEmailToUserNos = rs("N2KEquipmentEmailToUserNos")
		N2KEquipmentUserNosToCC = rs("N2KEquipmentUserNosToCC")
		N2KEquipmentEmailAddressesToCC = rs("N2KEquipmentEmailAddressesToCC")
		N2KEquipmentReportONOFF = rs("N2KEquipmentReportONOFF")
		N2KEqpIncludeBlankInsightAssetTagBrandPrefix = rs("N2KEqpIncludeBlankInsightAssetTagBrandPrefix")
		N2KEqpIncludeBlankInsightAssetTagClassPrefix = rs("N2KEqpIncludeBlankInsightAssetTagClassPrefix")
		N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix = rs("N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix")
		N2KEqpIncludeBlankInsightAssetTagModelPrefix = rs("N2KEqpIncludeBlankInsightAssetTagModelPrefix")
		N2KEqpIncludeUndefinedBrandExistsforEqp = rs("N2KEqpIncludeUndefinedBrandExistsforEqp")
		N2KEqpIncludeUndefinedClassExistsforEqp = rs("N2KEqpIncludeUndefinedClassExistsforEqp")
		N2KEqpIncludeUndefinedConditionCodeExistsforEqp = rs("N2KEqpIncludeUndefinedConditionCodeExistsforEqp")
		N2KEqpIncludeUndefinedGroupExistsforEqp = rs("N2KEqpIncludeUndefinedGroupExistsforEqp")
		N2KEqpIncludeUndefinedManufacturerExistsforEqp = rs("N2KEqpIncludeUndefinedManufacturerExistsforEqp")
		N2KEqpIncludeUndefinedModelExistsforEqp = rs("N2KEqpIncludeUndefinedModelExistsforEqp")
		N2KEqpIncludeUndefinedStatusCodeExistsforEqp = rs("N2KEqpIncludeUndefinedStatusCodeExistsforEqp")
		N2KEqpIncludeZeroDollarRentalsExistforEqp = rs("N2KEqpIncludeZeroDollarRentalsExistforEqp")
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
<h1 class="page-header"><i class="fa fa-tint"></i><i class="fa fa-coffee"></i>&nbsp;Need To Know - Equipment 
	<button id="toggle" class="btn btn-small btn-success"><i class="fas fa-arrows-v"></i>&nbsp;EXPAND/COLLAPSE ALL SETTINGS</button>
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
	<a href="<%= BaseURL %>admin/global/tiles/needtoknow/main.asp"><button class="btn btn-small btn-secondary pull-right"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-lightbulb-on"></i>&nbsp;NEED TO KNOW MAIN</button></a>
</h1>

<form method="post" action="<%= BaseURL %>admin/global/tiles/needtoknow/equipment-submit.asp" name="frmN2KEquipment" id="frmN2KEquipment">

	<div class="container">
		
		<%
			Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
			Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
			Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
			Response.Write("</div>")
			Response.Flush()
		%>
		
		<div class="row">
			<h3><i class="fad fa-sliders-h"></i>&nbsp;<%= GetTerm("Equipment") %> Need To Know Report General Settings</h3>
		</div>
	
		<div class="row">
		
			<div class="col-md-6">
				<div class="panel panel-primary">
					<div class="panel-heading">
						<h3 class="panel-title">Sections To Include In Report</h3>
						<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
					</div>
					<div class="panel-body">

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">

								<p><strong>Check off each section to be included in the Need to Know report</strong></p>

									<%
										If N2KEqpIncludeBlankInsightAssetTagBrandPrefix = 0 Then
											BlankInsightAssetTagBrandPrefix = ""
										Else
											BlankInsightAssetTagBrandPrefix = "checked"
										End If
		
										If N2KEqpIncludeBlankInsightAssetTagClassPrefix = 0 Then
											BlankInsightAssetTagClassPrefix = ""
										Else
											BlankInsightAssetTagClassPrefix = "checked"
										End If
		
										If N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix = 0 Then
											BlankInsightAssetTagManufacturerPrefix = ""
										Else
											BlankInsightAssetTagManufacturerPrefix = "checked"
										End If
		
										If N2KEqpIncludeBlankInsightAssetTagModelPrefix = 0 Then
											BlankInsightAssetTagModelPrefix = ""
										Else
											BlankInsightAssetTagModelPrefix = "checked"
										End If
		
										If N2KEqpIncludeUndefinedBrandExistsforEqp = 0 Then
											UndefinedBrandExistsforEqp = ""
										Else
											UndefinedBrandExistsforEqp = "checked"
										End If
		
										If N2KEqpIncludeUndefinedClassExistsforEqp = 0 Then
											UndefinedClassExistsforEqp = ""
										Else
											UndefinedClassExistsforEqp = "checked"
										End If
		
										If N2KEqpIncludeUndefinedConditionCodeExistsforEqp = 0 Then
											UndefinedConditionCodeExistsforEqp = ""
										Else
											UndefinedConditionCodeExistsforEqp = "checked"
										End If
										
										
										If N2KEqpIncludeUndefinedGroupExistsforEqp = 0 Then
											UndefinedGroupExistsforEqp = ""
										Else
											UndefinedGroupExistsforEqp = "checked"
										End If
		
										If N2KEqpIncludeUndefinedManufacturerExistsforEqp = 0 Then
											UndefinedManufacturerExistsforEqp = ""
										Else
											UndefinedManufacturerExistsforEqp = "checked"
										End If
		
										If N2KEqpIncludeUndefinedModelExistsforEqp = 0 Then
											UndefinedModelExistsforEqp = ""
										Else
											UndefinedModelExistsforEqp = "checked"
										End If
		
										If N2KEqpIncludeUndefinedStatusCodeExistsforEqp = 0 Then
											UndefinedStatusCodeExistsforEqp = ""
										Else
											UndefinedStatusCodeExistsforEqp = "checked"
										End If
		
										If N2KEqpIncludeZeroDollarRentalsExistforEqp = 0 Then
											ZeroDollarRentalsExistforEqp = ""
										Else
											ZeroDollarRentalsExistforEqp = "checked"
										End If
									%>
									
									<table cellspacing="5" cellpadding="5" width="100%">
									<tr>
									<td>
										<input type="checkbox" id="chkBlankInsightAssetTagBrandPrefix" name="chkBlankInsightAssetTagBrandPrefix" <%=BlankInsightAssetTagBrandPrefix%>>&nbsp;&nbsp;Blank Insight Asset Tag Brand Prefix
									</td>
									<td>
										<input type="checkbox" id="chkBlankInsightAssetTagClassPrefix" name="chkBlankInsightAssetTagClassPrefix" <%=BlankInsightAssetTagClassPrefix%>>&nbsp;&nbsp;Blank Insight Asset Tag Class Prefix
									</td>
									<td>
										<input type="checkbox" id="chkBlankInsightAssetTagManufacturerPrefix" name="chkBlankInsightAssetTagManufacturerPrefix" <%=BlankInsightAssetTagManufacturerPrefix%>>&nbsp;&nbsp;Blank Insight Asset Tag Manufacturer Prefix
									</td>
									</tr>
		
									<tr>
									<td>
										<input type="checkbox" id="chkBlankInsightAssetTagModelPrefix" name="chkBlankInsightAssetTagModelPrefix" <%=BlankInsightAssetTagModelPrefix%>>&nbsp;&nbsp;Blank Insight Asset Tag Model Prefix
									</td>
									<td>
										<input type="checkbox" id="chkUndefinedBrandExistsforEqp" name="chkUndefinedBrandExistsforEqp" <%=UndefinedBrandExistsforEqp%>>&nbsp;&nbsp;Undefined Brand Exists for Eqp
									</td>
									<td>
										<input type="checkbox" id="chkUndefinedClassExistsforEqp" name="chkUndefinedClassExistsforEqp" <%=UndefinedClassExistsforEqp%>>&nbsp;&nbsp;Undefined Class Exists for Eqp
									</td>							
									</tr>
		
									<tr>
									<td>
										<input type="checkbox" id="chkUndefinedConditionCodeExistsforEqp" name="chkUndefinedConditionCodeExistsforEqp" <%=UndefinedConditionCodeExistsforEqp%>>&nbsp;&nbsp;Undefined Condition Code Exists for Eqp
									</td>
									<td>
										<input type="checkbox" id="chkUndefinedGroupExistsforEqp" name="chkUndefinedGroupExistsforEqp" <%=UndefinedGroupExistsforEqp%>>&nbsp;&nbsp;Undefined Group Exists for Eqp
									</td>
									<td>
										<input type="checkbox" id="chkUndefinedManufacturerExistsforEqp" name="chkUndefinedManufacturerExistsforEqp" <%=UndefinedManufacturerExistsforEqp%>>&nbsp;&nbsp;Undefined Manufacturer Exists for Eqp
									</td>							
									</tr>
									
									<tr>
									<td>
										<input type="checkbox" id="chkUndefinedModelExistsforEqp" name="chkUndefinedModelExistsforEqp" <%=UndefinedModelExistsforEqp%>>&nbsp;&nbsp;Undefined Model Exists for Eqp
									</td>
									<td>
										<input type="checkbox" id="chkUndefinedStatusCodeExistsforEqp" name="chkUndefinedStatusCodeExistsforEqp" <%=UndefinedStatusCodeExistsforEqp%>>&nbsp;&nbsp;Undefined Status Code Exists for Eqp
									</td>
									<td>
										<input type="checkbox" id="chkZeroDollarRentalsExistforEqp" name="chkZeroDollarRentalsExistforEqp" <%=ZeroDollarRentalsExistforEqp%>>&nbsp;&nbsp;Zero Dollar Rentals Exist for Eqp
									</td>
									</tr>
									
									</table>

				            </div>
				            <!-- eof line -->
				         </div> 
					</div>
				</div>
			</div>	
			
			<div class="col-md-6">
				&nbsp;
			</div>	
						
			
		</div> 
		
	
	
		<div class="row">
			<h3><i class="fad fa-file-pdf"></i>&nbsp;<%= GetTerm("Equipment") %> Need To Know Report Settings</h3>
		</div>
	
		<div class="row">
		
			<div class="col-md-4">
				<% If N2KEquipmentReportONOFF = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title"><%= GetTerm("Equipment") %> Need To Know Report (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title"><%= GetTerm("Equipment") %> Need To Know Report (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
					<div class="panel-body">

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON 
					      		<%
					      		If N2KEquipmentReportONOFF = 0 Then
									Response.Write("<input type='checkbox' id='chkN2KEquipmentReportONOFF' name='chkN2KEquipmentReportONOFF'")
								Else
									Response.Write("<input type='checkbox' id='chkN2KEquipmentReportONOFF' name='chkN2KEquipmentReportONOFF' checked")
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalEquipmentNeedToKnowReportScheduler" data-tooltip="true" data-title="<%= GetTerm("Equipment") %> Need To Know Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> <%= GetTerm("Equipment") %> Need To Know Report Scheduler</button>
										</li>
										<li>
											<p>Select users <i class="fad fa-user-friends"></i> to send the report to:</p>
											<input type="hidden" name="lstSelectedN2KAPIEmailToUserNos" id="lstSelectedN2KAPIEmailToUserNos" value="<%= N2KEquipmentEmailToUserNos %>">
											<select id="lstExistingN2KAPIEmailToUserNos" multiple="multiple" name="lstExistingN2KAPIEmailToUserNos">
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalN2KEquipmentEmailAddressesToCC" data-tooltip="true" data-title="Additional emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails To CC</button>						
				             				<% If N2KEquipmentEmailAddressesToCC <> "" Then %>
				             					<p style="margin-top:20px;"><strong>Current Additional CC Emails:</strong> <%= N2KEquipmentEmailAddressesToCC %></p>
				             				<% End If %>
										</li>
										<li>
											<p>Select additonal users <i class="fad fa-user-friends"></i> to CC the report to:</p>
											<p><small>&nbsp;(CC:'s will receive a separate email for each recipient of this report)</small></p>
											<input type="hidden" name="lstSelectedN2KAPIUserNosToCC" id="lstSelectedN2KAPIUserNosToCC" value="<%= N2KEquipmentUserNosToCC %>">
											<select id="lstExistingN2KAPIUserNosToCC" multiple="multiple" name="lstExistingN2KAPIUserNosToCC">
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
				&nbsp;
			</div>

		
			<div class="col-md-4">
				&nbsp;
			</div>

		</div>
	</div>


	<!-- cancel / save !-->
	<div class="row pull-right">
		<div class="col-lg-12">
			<a href="<%= BaseURL %>admin/global/tiles/needtoknow/main.asp"><button type="button" class="btn btn-default btn-lg btn-huge"><i class="far fa-times-circle"></i> Cancel</button></a> 
			<button type="submit" class="btn btn-primary btn-lg btn-huge" onclick="showSavingChangesDiv()"><i class="far fa-save"></i> Save Changes</button>
		</div>
	</div>
	<!-- eof cancel / save !-->

		
	</div>
</form>

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR REPORT SCHEDULERS START HERE !-->
<!-- **************************************************************************************************************************** -->

<!-- pencil Modal -->
<div class="modal fade" id="modalEquipmentNeedToKnowReportScheduler" tabindex="-1" role="dialog" aria-labelledby="modalEquipmentNeedToKnowReportSchedulerLabel">
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="titleEquipmentNeedToKnowReportSchedulerLabel">Equipment Need To Know Report Generation Scheduler</h4>
		    </div>

			<form name="frmEditEquipmentNeedToKnowReportSchedulerModal" id="frmEditEquipmentNeedToKnowReportSchedulerModal" action="equipment-n2k-report-scheduler-submit.asp" method="POST">

				<div class="modal-body">
				    
					<div id="modalEquipmentNeedToKnowReportSchedulerContent">
						<!-- Content for the modal will be generated and written here -->
						<!-- Content generated by Sub GetContentForEquipmentNeedToKnowReportScheduler() in InSightFuncs_AjaxForAdminTimepickerModals.asp -->
					</div>
						
				</div>
				<!-- eof modal body !-->
				
				 <div class="clearfix"></div>
			      
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
					<button type="submit" id="btnEquipmentNeedToKnowReportScheduleSave" class="btn btn-primary">Save Schedule Changes</button>
				</div>
				
			</form>

		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->



<!-- Modal for Selecting CC Emails -->
<div class="modal fade" id="modalN2KEquipmentEmailAddressesToCC" tabindex="-1" role="dialog" aria-labelledby="modalN2KEquipmentEmailAddressesToCCLabel">
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                <h4 class="modal-title" id="H6">Cc: the following email addresses on this report:</h4>
               	<small>&nbsp;(Cc:'s will receive a separate email for each recipient of this report)</small>				
		    </div>

			<form  name="frmEditUsersList" id="Form5" action="users-list-update-needtoknow.asp" method="POST">
                <input type="hidden" name="userListName" value="N2KEquipmentEmailAddressesToCC" />
				<div class="modal-body">
				    
					<div id="Div2">
						<textarea class="form-control email-alert-line" rows="5" id="Textarea1" name="txtN2KEquipmentEmailAddressesToCC"><%= N2KEquipmentEmailAddressesToCC %></textarea>
						<strong>Separate multiple email addresses with a semicolon</strong>
					</div>
						
				</div>
				<!-- eof modal body !-->
				
				 <div class="clearfix"></div>
			      
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
					<button type="submit" id="Button6" class="btn btn-primary">Save Changes</button>
				</div>
				
			</form>

		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->


<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR REPORT SCHEDULERS END HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="../../../../inc/footer-main.asp"-->

