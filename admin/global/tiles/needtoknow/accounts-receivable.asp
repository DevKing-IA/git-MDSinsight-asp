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
					
		$('#modalFinanceNeedToKnowReportScheduler').on('show.bs.modal', function(e) {
		    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../../../inc/InSightFuncs_AjaxForAdminTimepickerModals.asp",
				cache: false,
				data: "action=GetContentForAccountsReceivableN2KReportScheduler",
				success: function(response)
				 {
	               	 $modal.find('#modalFinanceNeedToKnowReportSchedulerContent').html(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalFinanceNeedToKnowReportSchedulerContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
			
		});	
		

        $('#modalN2KAccountsReceivableEmailAddressesToCC').on('show.bs.modal', function (e) {
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
		N2KAccountsReceivableEmailToUserNos = rs("N2KAREmailToUserNos")
		N2KAccountsReceivableUserNosToCC = rs("N2KARUserNosToCC")
		N2KAccountsReceivableEmailAddressesToCC = rs("N2KAREmailAddressesToCC")
		N2KARReportONOFF = rs("N2KARReportONOFF")
		N2KARIncludeEmptyAddress2 = rs("N2KARIncludeEmptyAddress2")
		N2KARIncludeEmptyCity = rs("N2KARIncludeEmptyCity")
		N2KARIncludeEmptyCityStateZip = rs("N2KARIncludeEmptyCityStateZip")
		N2KARIncludeEmptyCustomerName = rs("N2KARIncludeEmptyCustomerName")
		N2KARIncludeEmptyPhoneNumber = rs("N2KARIncludeEmptyPhoneNumber")
		N2KARIncludeEmptyState = rs("N2KARIncludeEmptyState")
		N2KARIncludeEmptyZip = rs("N2KARIncludeEmptyZip")
		N2KARIncludeInvalidCityStateZip = rs("N2KARIncludeInvalidCityStateZip")
		N2KARIncludeInvalidPhoneNumber = rs("N2KARIncludeInvalidPhoneNumber")
		N2KARIncludeInvalidState = rs("N2KARIncludeInvalidState")
		N2KARIncludeInvalidZipCode = rs("N2KARIncludeInvalidZipCode")
		N2KARIncludeMissingcustomertype = rs("N2KARIncludeMissingcustomertype")
		N2KARIncludeMissingprimarysalesman = rs("N2KARIncludeMissingprimarysalesman")
		N2KARIncludeMissingsecondarysalesman = rs("N2KARIncludeMissingsecondarysalesman")
		N2KARIncludeNotAssignedToRegion = rs("N2KARIncludeNotAssignedToRegion")
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

<h1 class="page-header"><i class="fas fa-fw fa-file-invoice-dollar"></i>&nbsp;Need To Know - <%= GetTerm("Accounts Receivable") %> 
	<button id="toggle" class="btn btn-small btn-success"><i class="fas fa-arrows-v"></i>&nbsp;EXPAND/COLLAPSE ALL SETTINGS</button>
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
	<a href="<%= BaseURL %>admin/global/tiles/needtoknow/main.asp"><button class="btn btn-small btn-secondary pull-right"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-lightbulb-on"></i>&nbsp;NEED TO KNOW MAIN</button></a>
</h1>


<form method="post" action="<%= BaseURL %>admin/global/tiles/needtoknow/accounts-receivable-submit.asp" name="frrmN2KAccountsReceivable" id="frrmN2KAccountsReceivable">

	<div class="container">
	
		<%
			Response.Write("<div id='PleaseWaitPanel' style='display:none;'>")
			Response.Write("<br><br>Saving your recent changes, please wait...<br><br>")
			Response.Write("<img src=""" & baseURL & "/img/loading.gif"" />")
			Response.Write("</div>")
			Response.Flush()
		%>
	
		<div class="row">
			<h3><i class="fad fa-sliders-h"></i>&nbsp;<%= GetTerm("Accounts Receivable") %> Need To Know Report General Settings</h3>
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
									If N2KARIncludeEmptyAddress2 = 0 Then
										EmptyAddress2 = ""
									Else
										EmptyAddress2 = "checked"
									End If
	
									If N2KARIncludeEmptyCity = 0 Then
										EmptyCity = ""
									Else
										EmptyCity = "checked"
									End If
	
									If N2KARIncludeEmptyCityStateZip = 0 Then
										EmptyCityStateZip = ""
									Else
										EmptyCityStateZip = "checked"
									End If
	
									If N2KARIncludeEmptyCustomerName = 0 Then
										EmptyCustomerName = ""
									Else
										EmptyCustomerName = "checked"
									End If
	
									If N2KARIncludeEmptyPhoneNumber = 0 Then
										EmptyPhoneNumber = ""
									Else
										EmptyPhoneNumber = "checked"
									End If
	
									If N2KARIncludeEmptyState = 0 Then
										EmptyState = ""
									Else
										EmptyState = "checked"
									End If
	
									If N2KARIncludeEmptyZip = 0 Then
										EmptyZip = ""
									Else
										EmptyZip = "checked"
									End If
									
									
									If N2KARIncludeInvalidCityStateZip = 0 Then
										InvalidCityStateZip = ""
									Else
										InvalidCityStateZip = "checked"
									End If
	
									If N2KARIncludeInvalidPhoneNumber = 0 Then
										InvalidPhoneNumber = ""
									Else
										InvalidPhoneNumber = "checked"
									End If
	
									If N2KARIncludeInvalidState = 0 Then
										InvalidState = ""
									Else
										InvalidState = "checked"
									End If
	
									If N2KARIncludeInvalidZipCode = 0 Then
										InvalidZipCode = ""
									Else
										InvalidZipCode = "checked"
									End If
	
									If N2KARIncludeMissingcustomertype = 0 Then
										Missingcustomertype = ""
									Else
										Missingcustomertype = "checked"
									End If
	
									If N2KARIncludeMissingprimarysalesman = 0 Then
										Missingprimarysalesman = ""
									Else
										Missingprimarysalesman = "checked"
									End If
	
									If N2KARIncludeMissingsecondarysalesman = 0 Then
										Missingsecondarysalesman = ""
									Else
										Missingsecondarysalesman = "checked"
									End If		
									
									If N2KARIncludeNotAssignedToRegion = 0 Then
										NotAssignedToRegion = ""
									Else
										NotAssignedToRegion = "checked"
									End If									
									
								%>
								
								<table cellspacing="5" cellpadding="5" width="100%">
								<tr>
								<td>
									<input type="checkbox" id="chkEmptyAddress2" name="chkEmptyAddress2" <%=EmptyAddress2%>>&nbsp;&nbsp;Empty Address2
								</td>
								<td>
									<input type="checkbox" id="chkEmptyCity" name="chkEmptyCity" <%=EmptyCity%>>&nbsp;&nbsp;Empty City
								</td>
								<td>
									<input type="checkbox" id="chkEmptyCityStateZip" name="chkEmptyCityStateZip" <%=EmptyCityStateZip%>>&nbsp;&nbsp;Empty City State Zip
								</td>
								</tr>
	
								<tr>
								<td>
									<input type="checkbox" id="chkEmptyCustomerName" name="chkEmptyCustomerName" <%=EmptyCustomerName%>>&nbsp;&nbsp;Empty Customer Name
								</td>
								<td>
									<input type="checkbox" id="chkEmptyPhoneNumber" name="chkEmptyPhoneNumber" <%=EmptyPhoneNumber%>>&nbsp;&nbsp;Empty Phone Number
								</td>
								<td>
									<input type="checkbox" id="chkEmptyState" name="chkEmptyState" <%=EmptyState%>>&nbsp;&nbsp;Empty State
								</td>							
								</tr>
	
								<tr>
								<td>
									<input type="checkbox" id="chkEmptyZip" name="chkEmptyZip" <%=EmptyZip%>>&nbsp;&nbsp;Empty Zip
								</td>
								<td>
									<input type="checkbox" id="chkInvalidCityStateZip" name="chkInvalidCityStateZip" <%=InvalidCityStateZip%>>&nbsp;&nbsp;Invalid City State Zip
								</td>
								<td>
									<input type="checkbox" id="chkInvalidPhoneNumber" name="chkInvalidPhoneNumber" <%=InvalidPhoneNumber%>>&nbsp;&nbsp;Invalid Phone Number
								</td>							
								</tr>
								
								<tr>
								<td>
									<input type="checkbox" id="chkInvalidState" name="chkInvalidState" <%=InvalidState%>>&nbsp;&nbsp;Invalid State
								</td>
								<td>
									<input type="checkbox" id="chkInvalidZipCode" name="chkInvalidZipCode" <%=InvalidZipCode%>>&nbsp;&nbsp;Invalid Zip Code
								</td>
								</tr>
	
								<tr>
								<td>
									<input type="checkbox" id="chkMissingcustomertype" name="chkMissingcustomertype" <%=Missingcustomertype%>>&nbsp;&nbsp;Missing Customer Type
								</td>
								<td>
									<input type="checkbox" id="chkMissingprimarysalesman" name="chkMissingprimarysalesman" <%=Missingprimarysalesman%>>&nbsp;&nbsp;Missing Primary Salesman
								</td>
								<td>
									<input type="checkbox" id="chkMissingsecondarysalesman" name="chkMissingsecondarysalesman" <%=Missingsecondarysalesman%>>&nbsp;&nbsp;Missing Secondary Salesman
								</td>							
								</tr>
								
								<tr>
									<td>
										<input type="checkbox" id="chkNotAssignedToRegion" name="chkNotAssignedToRegion" <%= NotAssignedToRegion %>>&nbsp;&nbsp;Not Assigned To Region
									</td>
									<td>
										&nbsp;
									</td>
									<td>
										&nbsp;
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
			<h3><i class="fad fa-file-pdf"></i>&nbsp;<%= GetTerm("Accounts Receivable") %> Need To Know Report Settings</h3>
		</div>
	
		<div class="row">
		
			<div class="col-md-4">
				<% If N2KARReportONOFF = 0 Then %>
					<div class="panel panel-danger">
						<div class="panel-heading">
							<h3 class="panel-title"><%= GetTerm("Accounts Receivable") %> Need To Know Report (OFF)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>
				<% Else %>
					<div class="panel panel-success">
						<div class="panel-heading">
							<h3 class="panel-title"><%= GetTerm("Accounts Receivable") %> Need To Know Report (ON)</h3>
							<span class="pull-right clickable panel-collapsed"><i class="glyphicon glyphicon-chevron-down"></i></span> 
						</div>				
				<% End If %>
					<div class="panel-body">

					    <div class="row">
				            <!-- line -->
				            <div class="col-lg-12 line-full">
				               	TURN THIS REPORT ON 
					      		<%
					      		If N2KARReportONOFF = 0 Then
									Response.Write("<input type='checkbox' id='chkN2KARReportONOFF' name='chkN2KARReportONOFF'")
								Else
									Response.Write("<input type='checkbox' id='chkN2KARReportONOFF' name='chkN2KARReportONOFF' checked")
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalFinanceNeedToKnowReportScheduler" data-tooltip="true" data-title="<%= GetTerm("Accounts Receivable") %> Need To Know Report Scheduler" style="cursor:pointer;"><i class="far fa-calendar-alt"></i> <%= GetTerm("Accounts Receivable") %> Need To Know Report Scheduler</button>
										</li>
										<li>
											<p>Select users <i class="fad fa-user-friends"></i> to send the report to:</p>
											<input type="hidden" name="lstSelectedN2KAPIEmailToUserNos" id="lstSelectedN2KAPIEmailToUserNos" value="<%= N2KAccountsReceivableEmailToUserNos %>">
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
											<button type="button" class="btn btn-primary" data-toggle="modal" data-show="true" href="#" data-target="#modalN2KAccountsReceivableEmailAddressesToCC" data-tooltip="true" data-title="Additional emails" style="cursor:pointer;"><i class="fas fa-at"></i> Add Additional Emails To CC</button>						
				             				<% If N2KAccountsReceivableEmailAddressesToCC <> "" Then %>
				             					<p style="margin-top:20px;"><strong>Current Additional CC Emails:</strong> <%= N2KAccountsReceivableEmailAddressesToCC %></p>
				             				<% End If %>
										</li>
										<li>
											<p>Select additonal users <i class="fad fa-user-friends"></i> to CC the report to:</p>
											<p><small>&nbsp;(CC:'s will receive a separate email for each recipient of this report)</small></p>
											<input type="hidden" name="lstSelectedN2KAPIUserNosToCC" id="lstSelectedN2KAPIUserNosToCC" value="<%= N2KAccountsReceivableUserNosToCC %>">
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
<div class="modal fade" id="modalFinanceNeedToKnowReportScheduler" tabindex="-1" role="dialog" aria-labelledby="modalFinanceNeedToKnowReportSchedulerLabel">
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="titleFinanceNeedToKnowReportSchedulerLabel"><%= GetTerm("Accounts Receivable") %> Need To Know Report Generation Scheduler</h4>
		    </div>

			<form name="frmEditFinanceNeedToKnowReportSchedulerModal" id="frmEditFinanceNeedToKnowReportSchedulerModal" action="accounts-receivable-n2k-report-scheduler-submit.asp" method="POST">

				<div class="modal-body">
				    
					<div id="modalFinanceNeedToKnowReportSchedulerContent">
						<!-- Content for the modal will be generated and written here -->
						<!-- Content generated by Sub GetContentForFinanceNeedToKnowReportScheduler() in InSightFuncs_AjaxForAdminTimepickerModals.asp -->
					</div>
						
				</div>
				<!-- eof modal body !-->
				
				 <div class="clearfix"></div>
			      
				<div class="modal-footer">
					<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
					<button type="submit" id="btnFinanceNeedToKnowReportScheduleSave" class="btn btn-primary">Save Schedule Changes</button>
				</div>
				
			</form>

		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->

<!-- Modal for Selecting CC Emails -->
<div class="modal fade" id="modalN2KAccountsReceivableEmailAddressesToCC" tabindex="-1" role="dialog" aria-labelledby="modalN2KAccountsReceivableEmailAddressesToCCLabel">
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	
	
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                <h4 class="modal-title" id="H6">Cc: the following email addresses on this report:</h4>
               	<small>&nbsp;(Cc:'s will receive a separate email for each recipient of this report)</small>				
		    </div>

			<form  name="frmEditUsersList" id="Form5" action="users-list-update-needtoknow.asp" method="POST">
                <input type="hidden" name="userListName" value="N2KAccountsReceivableEmailAddressesToCC" />
				<div class="modal-body">
				    
					<div id="Div2">
						<textarea class="form-control email-alert-line" rows="5" id="txtN2KAccountsReceivableEmailAddressesToCC" name="txtN2KAccountsReceivableEmailAddressesToCC"><%= N2KAccountsReceivableEmailAddressesToCC %></textarea>
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

