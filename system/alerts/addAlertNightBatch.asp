<!--#include file="../../inc/header.asp"-->

<!-- parsley validation !-->
<script src="<%= BaseURL %>js/parsley/parsley.js"></script>

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
		$("#pnlMinutes").hide();
		$("#pnlLog").hide();
		$("#pnlLimits").hide();
		if (document.getElementById('selCond').selectedIndex == '3'){
			$("#pnlMinutes").show();
			$("#pnlLimits").show();
			}
			if (document.getElementById('selCond').selectedIndex == '4'){
			$("#pnlLog").show();
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


<h1 class="page-header"><i class="fa fa-fw fa-exclamation"></i>New Night Batch Alert</h1>


<form method="POST" action="addAlertNightBatch_submit.asp"  onsubmit="return checkform();" name="frmAddAlert" id="frmAddAlert" class="uk-form bt-flabels js-flabels" data-parsley-validate data-parsley-errors-messages-disabled >


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
		<div class="col-lg-3">
			<select class="form-control when-line" name="selCond" id="selCond" onchange="cndChanged();">
				<option value=""> -- none selected -- </option>
				<option value="Start">Night Batch Starts</option>
				<option value="Finish">Night Batch Finishes</option>
				<option value="Running">Night Batch has been running longer than</option>
				<option value="Log">Night Batch Log received</option>
			</select>
		</div>
		<!-- eof when select !-->

		<!-- minutes !-->
		<div class="col-lg-2" id="pnlMinutes" style="display: none;">
			<strong>Minutes</strong>
			<select class="form-control" id="selMinutes" name="selMinutes">
				<option value="1">1</option>
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
		<!-- eof minutes !-->
		
		
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
					<p align="right"><a href="<%= BaseURL %>system/alerts/main.asp#NightBatchAlerts">
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
