<!--#include file="../../inc/header.asp"-->

<!-- parsley validation !-->


<style type="text/css">
 
.bt-flabels > *:not(:first-child).bt-flabels__wrapper,
.bt-flabels > *:not(:first-child) .bt-flabels__wrapper {
  border-top: none;
}
.bt-flabels__wrapper {
  position: relative;
 }
.bt-flabels__error-desc {
  position: absolute;
  top: 0;
  right: 6px;
  opacity: 0;
  font-weight: bold;
  color: #f44545;
  font-size: 10px;
  text-transform: uppercase;
  z-index: 3;
  pointer-events: none;
}
.bt-flabels__error input[type] {
  background: #feeeee;
}
.bt-flabels__error input[type]:focus {
  background: #feeeee;
}
.bt-flabels__error .bt-flabels__error-desc {
  opacity: 1;
  transform: translateY(0);
}
.bt-flabels--right {
  border-left: none;
}
.bt-flabel__float label {
  opacity: 1;
  transform: translateY(0);
}
.bt-flabel__float input[type] {
  padding-top: 9px;
}
	
</style>
<!-- eof parsley validation !-->


<script type="text/javascript">
	function cndChanged() {
		$("#pnlTimeOfDay").hide();
		$("#pnlLog").hide();
		$("#pnlLimits").hide();
		$("#pnlMinutes").hide();
		$("#pnlDays").hide();
		if (document.getElementById('selCond').selectedIndex == '1'){
			$("#pnlTimeOfDay").show();
			}
		if (document.getElementById('selCond').selectedIndex == '5'){
			$("#pnlTimeOfDay").show();
			}
		if (document.getElementById('selCond').selectedIndex == '9'){
			$("#pnlTimeOfDay").show();
			}
		if (document.getElementById('selCond').selectedIndex == '2'){
			$("#pnlMinutes").show();
			}
		if (document.getElementById('selCond').selectedIndex == '6'){
			$("#pnlMinutes").show();
			}
		if (document.getElementById('selCond').selectedIndex == '13'){
			$("#pnlTimeOfDay").show();
			}
		if (document.getElementById('selCond').selectedIndex == '14'){
			$("#pnlDays").show();
			}
			
	}
	$(function () {
		cndChanged();
	});
</script>

<SCRIPT LANGUAGE="JavaScript">
<!--
    function checkform()
    {
	
        if (document.getElementById('selCond').selectedIndex == '3'){

   		var minut = document.getElementById('selLimitMinutes').value;
		var maxim = document.getElementById('selLimitMaxTimes').value;

       	 if ( minut * maxim > 1200){
	        swal("The combination of [minutes to wait] * [max times to send] cannot exceed 20 hours. Please adjust your entries before saving.")
            return false;
         }
         return true;
         
        }

        return true;

    }
    
   
// -->
</SCRIPT>   
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


<!-- password strength meter !-->

<style type="text/css">

.pass-strength h5{
	margin-top: 0px;
	color: #000;
}
.popover.primary {
    border-color:#337ab7;
}
.popover.primary>.arrow {
    border-top-color:#337ab7;
}
.popover.primary>.popover-title {
    color:#fff;
    background-color:#337ab7;
    border-color:#337ab7;
}
.popover.success {
    border-color:#d6e9c6;
}
.popover.success>.arrow {
    border-top-color:#d6e9c6;
}
.popover.success>.popover-title {
    color:#3c763d;
    background-color:#dff0d8;
    border-color:#d6e9c6;
}
.popover.info {
    border-color:#bce8f1;
}
.popover.info>.arrow {
    border-top-color:#bce8f1;
}
.popover.info>.popover-title {
    color:#31708f;
    background-color:#d9edf7;
    border-color:#bce8f1;
}
.popover.warning {
    border-color:#faebcc;
}
.popover.warning>.arrow {
    border-top-color:#faebcc;
}
.popover.warning>.popover-title {
    color:#8a6d3b;
    background-color:#fcf8e3;
    border-color:#faebcc;
}
.popover.danger {
    border-color:#ebccd1;
}
.popover.danger>.arrow {
    border-top-color:#ebccd1;
}
.popover.danger>.popover-title {
    color:#a94442;
    background-color:#f2dede;
    border-color:#ebccd1;
}

.select-line{
	margin-bottom: 15px;
}

.enable-disable{
	margin-top:20px;
}

.row-line{
	margin-bottom: 25px;
}

.table th, tr, td{
	font-weight: normal;
}

.table>thead>tr>th{
	border: 0px;
}
.table thead>tr>th,.table tbody>tr>th,.table tfoot>tr>th,.table thead>tr>td,.table tbody>tr>td,.table tfoot>tr>td{
border:0px;
}

.when-col{
	width: 10%;
}

.reference-col{
	width: 45%;
}

.has-more-col{
	width: 12%;
}

.form-control{
	min-width: 100px;
}

.textarea-box{
	min-width: 260px;
}

.alert-checkbox{
	margin-top: 15px;
}

.when-line{
	margin-top: 20px;
}

.when-line strong{
	padding-top: 7px;
	display: inline-block;
}


.email-alert-line{
	margin-top: 20px;
}

.incl-message-input {
    border: 1px solid #ccc;
    box-shadow: 0px 0px 0px 0px;
    max-width: 160px;
}

[class^="col-"]{
	 padding:5px;
  }
  
  .email-multi-select{
	  min-height: 160px;
  }
  
  .limit-alerts{
	  display: inline-block;
	  padding-top: 15px;
  }
  
 </style>


<h1 class="page-header"><i class="fa fa-fw fa-exclamation"></i>New System Alert</h1>


<form method="POST" action="addAlertSystem_submit.asp"  onsubmit="return checkform();" name="frmAddAlert" id="frmAddAlert" class="uk-form bt-flabels js-flabels" >


	<div class="row row-line">
		<div class="col-lg-2">
			<div class="bt-flabels__wrapper">
			<strong>Alert Name</strong><input type="text" id="txtAlertName" name="txtAlertName" value=""  class="form-control" data-parsley-required>
			<span class="bt-flabels__error-desc">Required</span>
			</div>
		</div>

		<!-- enabled !-->
		<div class="col-lg-1 alert-checkbox">
			<div class="checkbox">
				<label>
					<input type="checkbox" id="chkEnabled" name="chkEnabled" checked><strong> Enabled</strong>
				</label>
			</div>
		</div>
		<!-- eof enabled !-->
		
		<!-- alert or notification !-->
		<div class="col-lg-1 alert-checkbox">
			<div class="checkbox">
				<label>
					<input type="radio" name="optNotificationType" id="optNotificationType" value="Alert" checked><strong> Alert</strong><br>
					<input type="radio" name="optNotificationType" id="optNotificationType" value="Notification" ><strong> Notification</strong>

				</label>
			</div>
		</div>
		<!-- eof alert or notification !-->

		<!-- public or private !-->
		<div class="col-lg-1 alert-checkbox">
			<div class="checkbox">
				<label>
					<input type="radio" name="optPublicOrPrivate" id="optPublicOrPrivate" value="Private" checked><strong> Private</strong><br>
					<input type="radio" name="optPublicOrPrivate" id="optPublicOrPrivate" value="Public" ><strong> Public</strong>
				</label>
			</div>
		</div>
		<!-- eof public or private !-->

		<!-- when label !-->
		<div class="col-lg-1">
			<p class="when-line" align="right">
				<strong>When</strong>
			</p>
		</div>
		<!-- eof when label !-->
	
		<!-- when select !-->
		<div class="col-lg-2">
			<select class="form-control when-line" name="selCond" id="selCond" onchange="cndChanged();">
				<option selected value="">-- Nothing Selected --</option>
				<option value="BackendNoStart">Backend data import did not start</option>
				<option value="BackendRunTooLong">Backend data import has been running longer than</option>
				<option value="BackendStarted">Backend data import started</option>
				<option value="BackendFinished">Backend data import finished</option>
				<option value="RebuildNotRun">Daily data rebuild did not start</option>
				<option value="RebuildRunTooLong">Daily data rebuild has been running longer than</option>
				<option value="RebuildStarted">Daily data rebuild started</option>
				<option value="RebuildFinished">Daily data rebuild finished</option>
				<option value="DBoardNotRun">Nightly delivery board did not run</option>
				<option value="DBoardSkipped">Nightly delivery board update skipped the update</option>
				<option value="DBoardFinished">Nightly delivery board update finished</option>
				<option value="DBoardOnDemandRun">Delivery board update on demand was run</option>
				<option value="AutoCompJSONNotRun">Auto-complete JSON file not rebuilt</option>	
				<option value="HistOldInvoice">Most recent history invoice older than</option>
				<option value="RouteFileEmpty">Route file empty</option>
				<% If MUV_Read("prospectingModuleOn") = "Enabled" Then %>
					<option value="ProspectNoNextActivity">Prospect found with no next activity</option>				
				<% End If %>
			</select>
		</div>
		<!-- eof when select !-->

		<!-- time of day !-->
		<div class="col-lg-2" id="pnlTimeOfDay" style="display: none;">
			<strong>Time Of Day</strong>
			<select class="form-control" id="selTimeOfDay" name="selTimeOfDay">
					<option value="0000">-Midnight-</option>
					<option value="0015">12:15 AM</option>
					<option value="0030">12:30 AM</option>
					<option value="0045">12:45 AM</option>
					<option value="100">1:00 AM</option>
					<option value="115">1:15 AM</option>
					<option value="130">1:30 AM</option>
					<option value="145">1:45 AM</option>
					<option value="200">2:00 AM</option>
					<option value="215">2:15 AM</option>
					<option value="230">2:30 AM</option>
					<option value="245">2:45 AM</option>
					<option value="300">3:00 AM</option>
					<option value="315">3:15 AM</option>
					<option value="330">3:30 AM</option>
					<option value="345">3:45 AM</option>
					<option value="400">4:00 AM</option>
					<option value="415">4:15 AM</option>
					<option value="430">4:30 AM</option>
					<option value="445">4:45 AM</option>
					<option value="500">5:00 AM</option>
					<option value="515">5:15 AM</option>
					<option value="530">5:30 AM</option>
					<option value="545">5:45 AM</option>
					<option value="600">6:00 AM</option>
					<option value="615">6:15 AM</option>
					<option value="630">6:30 AM</option>
					<option value="645">6:45 AM</option>
					<option value="700">7:00 AM</option>
					<option value="715">7:15 AM</option>
					<option value="730">7:30 AM</option>
					<option value="745">7:45 AM</option>
					<option value="800">8:00 AM</option>
					<option value="815">8:15 AM</option>
					<option value="830">8:30 AM</option>
					<option value="845">8:45 AM</option>
					<option value="900">9:00 AM</option>
					<option value="915">9:15 AM</option>
					<option value="930">9:30 AM</option>
					<option value="945">9:45 AM</option>
					<option value="1000">10:00 AM</option>
					<option value="1015">10:15 AM</option>
					<option value="1030">10:30 AM</option>
					<option value="1045">10:45 AM</option>
					<option value="1100">11:00 AM</option>
					<option value="1115">11:15 AM</option>
					<option value="1130">11:30 AM</option>
					<option value="1145">11:45 AM</option>
					<option value="1200">-Noon-</option>
					<option value="1215">12:15 PM</option>
					<option value="1230">12:30 PM</option>
					<option value="1245">12:45 PM</option>
					<option value="1300">1:00 PM</option>
					<option value="1315">1:15 PM</option>
					<option value="1330">1:30 PM</option>
					<option value="1345">1:45 PM</option>
					<option value="1400">2:00 PM</option>
					<option value="1415">2:15 PM</option>
					<option value="1430">2:30 PM</option>
					<option value="1445">2:45 PM</option>
					<option value="1500">3:00 PM</option>
					<option value="1515">3:15 PM</option>
					<option value="1530">3:30 PM</option>
					<option value="1545">3:45 PM</option>
					<option value="1600">4:00 PM</option>
					<option value="1615">4:15 PM</option>
					<option value="1630">4:30 PM</option>
					<option value="1645">4:45 PM</option>
					<option value="1700">5:00 PM</option>
					<option value="1715">5:15 PM</option>
					<option value="1730">5:30 PM</option>
					<option value="1745">5:45 PM</option>
					<option value="1800">6:00 PM</option>
					<option value="1815">6:15 PM</option>
					<option value="1830">6:30 PM</option>
					<option value="1845">6:45 PM</option>
					<option value="1900">7:00 PM</option>
					<option value="1915">7:15 PM</option>
					<option value="1930">7:30 PM</option>
					<option value="1945">7:45 PM</option>
					<option value="2000">8:00 PM</option>
					<option value="2015">8:15 PM</option>
					<option value="2030">8:30 PM</option>
					<option value="2045">8:45 PM</option>
					<option value="2100">9:00 PM</option>
					<option value="2115">9:15 PM</option>
					<option value="2130">9:30 PM</option>
					<option value="2145">9:45 PM</option>
					<option value="2200">10:00 PM</option>
					<option value="2215">10:15 PM</option>
					<option value="2230">10:30 PM</option>
					<option value="2245">10:45 PM</option>
					<option value="2300">11:00 PM</option>
					<option value="2315">11:15 PM</option>
					<option value="2330">11:30 PM</option>
					<option value="2345">11:45 PM</option>				
		 		</select>
		</div>
		<!-- eof time of day !-->

		<!-- minutes !-->
		<div class="col-lg-2" id="pnlMinutes" style="display: none;">
			<strong>Minutes</strong>
			<select class="form-control" id="selMinutes" name="selMinutes">
				<option value="1"<%If Minutes = 1 Then Response.Write(" selected ")%>>1</option>
				<%
					For x = 5 to 480 Step 5 ' 8 hours
						If x mod 60 = 0 Then
							If x = cint(Minutes) Then 
								Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
							else
								Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
							End If
						Else
							If x = cint(Minutes) Then 
								Response.Write("<option value='" & x & "' selected>" & x & "</option>")
							Else
								Response.Write("<option value='" & x & "'>" & x & "</option>")
							End If
						End If
					Next
				%>
			</select>
		</div>
		<!-- eof minutes !-->

		<!-- days !-->
		<div class="col-lg-2" id="pnlDays" style="display: none;">
			<strong>Days</strong>
			<select class="form-control" id="selDays" name="selDAys">
				<option value="1"<%If Days = 1 Then Response.Write(" selected ")%>>1</option>
				<%
					For x = 1 to 10
						If x = cint(Days) Then 
							Response.Write("<option value='" & x & "' selected>" & x & "</option>")
						Else
							Response.Write("<option value='" & x & "'>" & x & "</option>")
						End If
					Next
				%>
			</select>
		</div>
		<!-- eof days !-->

	</div>
	
	<!-- send email alert !-->
	<div class="row row-line">
	
		<!-- send email !-->
		<div class="col-lg-1">
			<p class="email-alert-line" align="right" ><strong>Send an email alert to</strong></p>
		</div>
		<!-- eof send email !-->

		<!-- none from here !-->
		<div class="col-lg-2">
			<select class="form-control email-alert-line email-multi-select" id="selEmailto" name="selEmailto" multiple>
				<option value="0" selected>--- none from here ---</option>
				<option value="<%=Session("UserNo")%>"><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option>
		      	<%'Users dropdown
		      	 
	      	  	SQL = "SELECT UserNo, userFirstName, userLastName FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
	      	  	SQL = SQL & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo")
	      	  	SQL = SQL & " order by  userFirstName, userLastName"

				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.open (Session("ClientCnnString"))
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3 
				Set rs = cnn8.Execute(SQL)
			
				If not rs.EOF Then
					Do
						FullName = rs("userFirstName") & " " & rs("userLastName")
						Response.Write("<option value='" & rs("UserNo") & "'>" & FullName & "</option>")
						rs.movenext
					Loop until rs.eof
				End If
				set rs = Nothing
				cnn8.close
				set cnn8 = Nothing
		      	%>
			</select>
			<strong>Use CTRL and SHIFT to make multiple selections</strong>
		</div>
		<!-- eof none from here !-->
	
		<!-- and !-->
		<div class="col-lg-1">
			<p align="center" class="email-alert-line"><strong>and</strong></p>
		</div>
		<!-- eof and !-->
	
		<!-- separate emails !-->
		<div class="col-lg-3">
			<textarea class="form-control email-alert-line" rows="4" id="txtaAdditionalEmails" name="txtaAdditionalEmails"></textarea>
			<strong>Separate multiple email addresses with a semicolon</strong>
		</div>
		<!-- eof separate emails !-->
	
		<!-- verbiage !-->
		<div class="col-lg-3">
			<strong>Verbiage to include in alert email</strong>
			<textarea class="form-control" rows="4" id="txtaVerbiageEmail" name="txtaVerbiageEmail"></textarea>
			<div class="col-lg-4" id="pnlLog" style="display: none;">
				<input type="checkbox" id="chkLog" name="chkLog" checked><strong> Include log</strong>
			</div>
	 	</div>
		<!-- eof verbiage !-->
	
	</div>
	<!-- eof  send email alert !-->


	<!-- send text alert !-->
	<div class="row row-line">
		
		<!-- send email !-->
		<div class="col-lg-1">
			<p class="email-alert-line" align="right" ><strong>Send a text alert to</strong></p>
		</div>
		<!-- eof send email !-->
		
		<!-- none from here !-->
		<div class="col-lg-2">
			<select class="form-control email-alert-line email-multi-select"  id="selTextto" name="selTextto" multiple>
				<option value="0" selected>--- none from here ---</option>
				<option value="<%=Session("UserNo")%>"><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option>
		      	<%'Users dropdown
		      	 
	      	  	SQL = "SELECT UserNo, userFirstName, userLastName FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
	      	  	SQL = SQL & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo")
	      	  	SQL = SQL & " order by  userFirstName, userLastName"

				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.open (Session("ClientCnnString"))
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3 
				Set rs = cnn8.Execute(SQL)
			
				If not rs.EOF Then
					Do
						FullName = rs("userFirstName") & " " & rs("userLastName")
						Response.Write("<option value='" & rs("UserNo") & "'>" & FullName & "</option>")
						rs.movenext
					Loop until rs.eof
				End If
				set rs = Nothing
				cnn8.close
				set cnn8 = Nothing
		      	%>
			</select>
			<strong>Use CTRL and SHIFT to make multiple selections</strong>
		</div>
		<!-- eof none from here !-->
	
	
	<!-- and !-->
	<div class="col-lg-1">
		<p align="center" class="email-alert-line"><strong>and</strong></p>
	</div>
	<!-- eof and !-->
	
	<!-- separate emails !-->
	<div class="col-lg-3">
		<textarea class="form-control email-alert-line" rows="4" id="txtaAdditionalTexts" name="txtaAdditionalTexts"></textarea>
		<strong>Separate multiple phone numbers with a semicolon</strong>
	</div>
	<!-- eof separate emails !-->
	
	<!-- verbiage !--> 
	<div class="col-lg-3">
		<strong>Message to include in text alert</strong><br>
		<input type="text"  value=""  class="form-control" id="txtAlertTextVerbiage" name="txtAlertTextVerbiage">
 	</div>
	<!-- eof verbiage !-->
	
</div>
<!-- eof  send text alert !-->

<!-- limit / maximum section !-->
<div class="row row-line" id="pnlLimits" style="display: none;">
	<div class="col-lg-10">

		<!-- limit !-->
		<div class="col-lg-3">
	        <strong class="limit-alerts">Limit this alert to sending only once every</strong>
        </div>
        <!-- eof limit !-->
        
        <!-- limit options !-->
        <div class="col-lg-2">
	        <strong>Minutes</strong>
	        <select class="form-control" id="selLimitMinutes" name="selLimitMinutes">
				<%
					For x = 5 to 480 Step 5 ' 8 hours
						If x mod 60 = 0 Then
							Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
						Else
							Response.Write("<option value='" & x & "'>" & x & "</option>")
						End If
					Next
				%>
	        </select>
        </div>
        <!-- eof limit options !-->
        
        <!-- maximum !-->
        <div class="col-lg-3">
	        <strong class="limit-alerts">The maximum # of times to send this alert is</strong>
        </div>
        <!-- eof maximum !-->
        
        <!-- maximum options !-->
        <div class="col-lg-2">
	        <strong>Maximum</strong>
	        <select class="form-control" id="selLimitMaxTimes" name="selLimitMaxTimes">
        	<%
				For x = 1 to 10
					Response.Write("<option value='" & x & "'>" & x & "</option>")
				Next
			%>
	        </select>
        </div>
        <!-- eof maximum options !-->
	</div>
</div>
<!-- eof limit / maximum section !-->


	<!-- reference line !-->
	<div class="row row-line">
		<div class="col-lg-10">
 			
			<div class="row row-line">
				<div class="col-lg-12 alertbutton">
					<p align="right"><a href="<%= BaseURL %>system/alerts/main.asp#System">
    					<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Alert List</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button></p>
		    	</div>
			</div>
		</div>
	</div>
</form>

<!-- parsley form !-->
<script type="text/javascript">
  $('#frmAddAlert').parsley();
  (function($) {
  'use strict';
  var floatingLabel;
  floatingLabel = function(onload) {
    var $input;
    $input = $(this);
    if (onload) {
      $.each($('.bt-flabels__wrapper input'), function(index, value) {
        var $current_input;
        $current_input = $(value);
        if ($current_input.val()) {
          $current_input.closest('.bt-flabels__wrapper').addClass('bt-flabel__float');
        }
      });
    }
    setTimeout((function() {
      if ($input.val()) {
        $input.closest('.bt-flabels__wrapper').addClass('bt-flabel__float');
      } else {
        $input.closest('.bt-flabels__wrapper').removeClass('bt-flabel__float');
      }
    }), 1);
  };
  $('.bt-flabels__wrapper input').keydown(floatingLabel);
  $('.bt-flabels__wrapper input').change(floatingLabel);
  window.addEventListener('load', floatingLabel(true), false);
  $('.js-flabels').parsley().on('form:error', function() {
    $.each(this.fields, function(key, field) {
      if (field.validationResult !== true) {
        field.$element.closest('.bt-flabels__wrapper').addClass('bt-flabels__error');
      }
    });
  });
  $('.js-flabels').parsley().on('field:validated', function() {
    if (this.validationResult === true) {
      this.$element.closest('.bt-flabels__wrapper').removeClass('bt-flabels__error');
    } else {
      this.$element.closest('.bt-flabels__wrapper').addClass('bt-flabels__error');
    }
  });
})(jQuery);

// ---
 </script>
<!-- eof parsley form !-->


<!--#include file="../../inc/footer-main.asp"-->
