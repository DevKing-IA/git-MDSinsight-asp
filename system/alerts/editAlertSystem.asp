<!--#include file="../../inc/header.asp"-->
<% InternalAlertRecNumber = Request.QueryString("a") 
If InternalAlertRecNumber = "" Then Response.Redirect("main.asp")
%>
<body onload="load()">

<script type="text/javascript">
function load(){
if (document.getElementById('selCond').selectedIndex == '2'){
	document.getElementById('pnlMinutes').style.display="block";
	document.getElementById('pnlLimits').style.display="block";
}
if (document.getElementById('selCond').selectedIndex == '3'){
	document.getElementById('pnlLog').style.display="block";
}
}</script>

<script type="text/javascript">
	function cndChanged() {
		$("#pnlTimeOfDay").hide();
		$("#pnlLog").hide();
		$("#pnlLimits").hide();
		$("#pnlMinutes").hide();
		$("#pnlDays").hide();
		if (document.getElementById('selCond').selectedIndex == '0'){
			$("#pnlTimeOfDay").show();
			}
		if (document.getElementById('selCond').selectedIndex == '4'){
			$("#pnlTimeOfDay").show();
			}
		if (document.getElementById('selCond').selectedIndex == '8'){
			$("#pnlTimeOfDay").show();
			}
		if (document.getElementById('selCond').selectedIndex == '1'){
			$("#pnlMinutes").show();
			}
		if (document.getElementById('selCond').selectedIndex == '5'){
			$("#pnlMinutes").show();
			}
		if (document.getElementById('selCond').selectedIndex == '12'){
			$("#pnlTimeOfDay").show();
			}
		if (document.getElementById('selCond').selectedIndex == '13'){
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
	
        if (document.getElementById('selCond').selectedIndex == '2'){

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


<h1 class="page-header"><i class="fa fa-fw fa-exclamation"></i>Edit System Alert</h1>

<%
SQL = "SELECT * FROM SC_Alerts where InternalAlertRecNumber = " & InternalAlertRecNumber 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	AlertName = rs("AlertName")
	Enabled = rs("Enabled")
	Condition = rs("Condition")
	Minutes = rs("NBMinutes")
	Days = rs("NumberOfDays")
	TimeOfDay = rs("TimeOfDay")
	Emailto = rs("EmailToUserNos") 
	AdditionalEmails = rs("AdditionalEmails")
	VerbiageEmail = rs("EmailVerbiage")
	Textto = rs("TextToUserNos")
	AdditionalTexts = rs("AdditionalText")
	TextVerbiage = rs("TextVerbiage") 
	IncludeLog = rs("NBIncludeLog")
	LimitMinutes = rs("NBLimitMiniutes")
	LimitMaxTimes = rs("NBLimitMaxTimes")	
	NotificationType = rs("NotificationType")
	PublicOrPrivate = rs("PublicOrPrivate")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

If IsNull(LimitMinutes) Then LimitMinutes = 60
If IsNull(LimitMaxTimes) Then LimitMaxTimes = 1
If LimitMinutes = "" Then LimitMinutes = 60
If LimitMaxTimes = "" Then LimitMaxTimes = 1


%>


<form method="POST" action="editAlertSystem_submit.asp" name="frmEditAlert" id="frmEditAlert"  onsubmit="return checkform();">


	<input type="hidden" id="txtInternalAlertRecNumber" name="txtInternalAlertRecNumber" value="<%=InternalAlertRecNumber%>"  class="form-control">

	<div class="row row-line">
		<div class="col-lg-2">
			<strong>Alert Name</strong><input type="text" id="txtAlertName" name="txtAlertName" value="<%=AlertName%>"  class="form-control">
		</div>

		<!-- enabled !-->
		<div class="col-lg-2 alert-checkbox">
			<div class="checkbox">
				<label>
					<input type="checkbox" id="chkEnabled" name="chkEnabled" <%If Enabled = vbTrue Then Response.Write(" checked ") %>><strong> Enabled</strong>
				</label>
			</div>
		</div>
		<!-- eof enabled !-->

		<!-- alert or notification !-->
		<div class="col-lg-1 alert-checkbox">
			<div class="checkbox">
				<label>
					<input type="radio" name="optNotificationType" id="optNotificationType" value="Alert" <%If NotificationType = "Alert" Then Response.Write(" checked ") %>><strong> Alert</strong><br>
					<input type="radio" name="optNotificationType" id="optNotificationType" value="Notification" <%If NotificationType = "Notification" Then Response.Write(" checked ") %>><strong> Notification</strong>

				</label>
			</div>
		</div>
		<!-- eof alert or notification !-->

		<!-- public or private !-->
		<div class="col-lg-1 alert-checkbox">
			<div class="checkbox">
				<label>
					<input type="radio" name="optPublicOrPrivate" id="optPublicOrPrivate" value="Private" <%If PublicOrPrivate = "Private" Then Response.Write(" checked ") %>><strong> Private</strong><br>
					<input type="radio" name="optPublicOrPrivate" id="optPublicOrPrivate" value="Public" <%If PublicOrPrivate = "Public" Then Response.Write(" checked ") %>><strong> Public</strong>
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
				<option value="BackendNoStart"<%If Condition = "BackendNoStart" Then Response.Write(" selected ")%>>Backend data import did not start</option>
				<option value="BackendRunTooLong"<%If Condition = "BackendRunTooLong" Then Response.Write(" selected ")%>>Backend data import has been running longer than</option>
				<option value="BackendStarted"<%If Condition = "BackendStarted" Then Response.Write(" selected ")%>>Backend data import started</option>
				<option value="BackendFinished"<%If Condition = "BackendFinished" Then Response.Write(" selected ")%>>Backend data import finished</option>
				<option value="RebuildNotRun"<%If Condition = "RebuildNotRun" Then Response.Write(" selected ")%>>Daily data rebuild did not start</option>
				<option value="RebuildRunTooLong"<%If Condition = "RebuildRunTooLong" Then Response.Write(" selected ")%>>Daily data rebuild has been running longer than</option>
				<option value="RebuildStarted"<%If Condition = "RebuildStarted" Then Response.Write(" selected ")%>>Daily data rebuild started</option>
				<option value="RebuildFinished"<%If Condition = "RebuildFinished" Then Response.Write(" selected ")%>>Daily data rebuild finished</option>
				<option value="DBoardNotRun"<%If Condition = "DBoardNotRun" Then Response.Write(" selected ")%>>Nightly delivery board did not run</option>
				<option value="DBoardSkipped"<%If Condition = "DBoardSkipped" Then Response.Write(" selected ")%>>Nightly delivery board update skipped the update</option>
				<option value="DBoardFinished"<%If Condition = "DBoardFinished" Then Response.Write(" selected ")%>>Nightly delivery board update finished</option>
				<option value="DBoardOnDemandRun"<%If Condition = "DBoardOnDemandRun" Then Response.Write(" selected ")%>>Delivery board update on demand was run</option>
				<option value="AutoCompJSONNotRun"<%If Condition = "AutoCompJSONNotRun" Then Response.Write(" selected ")%>>Auto-complete JSON file not rebuilt</option>
				<option value="HistOldInvoice"<%If Condition = "HistOldInvoice" Then Response.Write(" selected ")%>>Most recent history invoice older than</option>	
				<option value="RouteFileEmpty"<%If Condition = "RouteFileEmpty" Then Response.Write(" selected ")%>>Route file empty</option>
				<% If MUV_Read("prospectingModuleOn") = "Enabled" Then %>
					<option value="ProspectNoNextActivity"<%If Condition = "ProspectNoNextActivity" Then Response.Write(" selected ")%>>Prospect found with no next activity</option>		
				<% End If %>
			</select>
		</div>
		<!-- eof when select !-->


		<!-- time of day !-->
		<div class="col-lg-2" id="pnlTimeOfDay" style="display: none;">
			<strong>Time Of Day</strong>
			<select class="form-control" id="selTimeOfDay" name="selTimeOfDay">			
					<option value="0000"<%If TimeOfDay = "0000" Then Response.Write(" selected ")%>>-Midnight-</option>
					<option value="0015"<%If TimeOfDay = "0015" Then Response.Write(" selected ")%>>12:15 AM</option>
					<option value="0030"<%If TimeOfDay = "0030" Then Response.Write(" selected ")%>>12:30 AM</option>
					<option value="0045"<%If TimeOfDay = "0045" Then Response.Write(" selected ")%>>12:45 AM</option>
					<option value="100"<%If TimeOfDay = "100" Then Response.Write(" selected ")%>>1:00 AM</option>
					<option value="115"<%If TimeOfDay = "115" Then Response.Write(" selected ")%>>1:15 AM</option>
					<option value="130"<%If TimeOfDay = "130" Then Response.Write(" selected ")%>>1:30 AM</option>
					<option value="145"<%If TimeOfDay = "145" Then Response.Write(" selected ")%>>1:45 AM</option>
					<option value="200"<%If TimeOfDay = "200" Then Response.Write(" selected ")%>>2:00 AM</option>
					<option value="215"<%If TimeOfDay = "215" Then Response.Write(" selected ")%>>2:15 AM</option>
					<option value="230"<%If TimeOfDay = "230" Then Response.Write(" selected ")%>>2:30 AM</option>
					<option value="245"<%If TimeOfDay = "245" Then Response.Write(" selected ")%>>2:45 AM</option>
					<option value="300"<%If TimeOfDay = "300" Then Response.Write(" selected ")%>>3:00 AM</option>
					<option value="315"<%If TimeOfDay = "315" Then Response.Write(" selected ")%>>3:15 AM</option>
					<option value="330"<%If TimeOfDay = "330" Then Response.Write(" selected ")%>>3:30 AM</option>
					<option value="345"<%If TimeOfDay = "345" Then Response.Write(" selected ")%>>3:45 AM</option>
					<option value="400"<%If TimeOfDay = "400" Then Response.Write(" selected ")%>>4:00 AM</option>
					<option value="415"<%If TimeOfDay = "415" Then Response.Write(" selected ")%>>4:15 AM</option>
					<option value="430"<%If TimeOfDay = "430" Then Response.Write(" selected ")%>>4:30 AM</option>
					<option value="445"<%If TimeOfDay = "445" Then Response.Write(" selected ")%>>4:45 AM</option>
					<option value="500"<%If TimeOfDay = "500" Then Response.Write(" selected ")%>>5:00 AM</option>
					<option value="515"<%If TimeOfDay = "515" Then Response.Write(" selected ")%>>5:15 AM</option>
					<option value="530"<%If TimeOfDay = "530" Then Response.Write(" selected ")%>>5:30 AM</option>
					<option value="545"<%If TimeOfDay = "545" Then Response.Write(" selected ")%>>5:45 AM</option>
					<option value="600"<%If TimeOfDay = "600" Then Response.Write(" selected ")%>>6:00 AM</option>
					<option value="615"<%If TimeOfDay = "615" Then Response.Write(" selected ")%>>6:15 AM</option>
					<option value="630"<%If TimeOfDay = "630" Then Response.Write(" selected ")%>>6:30 AM</option>
					<option value="645"<%If TimeOfDay = "645" Then Response.Write(" selected ")%>>6:45 AM</option>
					<option value="700"<%If TimeOfDay = "700" Then Response.Write(" selected ")%>>7:00 AM</option>
					<option value="715"<%If TimeOfDay = "715" Then Response.Write(" selected ")%>>7:15 AM</option>
					<option value="730"<%If TimeOfDay = "730" Then Response.Write(" selected ")%>>7:30 AM</option>
					<option value="745"<%If TimeOfDay = "745" Then Response.Write(" selected ")%>>7:45 AM</option>
					<option value="800"<%If TimeOfDay = "800" Then Response.Write(" selected ")%>>8:00 AM</option>
					<option value="815"<%If TimeOfDay = "815" Then Response.Write(" selected ")%>>8:15 AM</option>
					<option value="830"<%If TimeOfDay = "830" Then Response.Write(" selected ")%>>8:30 AM</option>
					<option value="845"<%If TimeOfDay = "845" Then Response.Write(" selected ")%>>8:45 AM</option>
					<option value="900"<%If TimeOfDay = "900" Then Response.Write(" selected ")%>>9:00 AM</option>
					<option value="915"<%If TimeOfDay = "915" Then Response.Write(" selected ")%>>9:15 AM</option>
					<option value="930"<%If TimeOfDay = "930" Then Response.Write(" selected ")%>>9:30 AM</option>
					<option value="945"<%If TimeOfDay = "945" Then Response.Write(" selected ")%>>9:45 AM</option>
					<option value="1000"<%If TimeOfDay = "1000" Then Response.Write(" selected ")%>>10:00 AM</option>
					<option value="1015"<%If TimeOfDay = "1015" Then Response.Write(" selected ")%>>10:15 AM</option>
					<option value="1030"<%If TimeOfDay = "1030" Then Response.Write(" selected ")%>>10:30 AM</option>
					<option value="1045"<%If TimeOfDay = "1045" Then Response.Write(" selected ")%>>10:45 AM</option>
					<option value="1100"<%If TimeOfDay = "1100" Then Response.Write(" selected ")%>>11:00 AM</option>
					<option value="1115"<%If TimeOfDay = "1115" Then Response.Write(" selected ")%>>11:15 AM</option>
					<option value="1130"<%If TimeOfDay = "1130" Then Response.Write(" selected ")%>>11:30 AM</option>
					<option value="1145"<%If TimeOfDay = "1145" Then Response.Write(" selected ")%>>11:45 AM</option>
					<option value="1200"<%If TimeOfDay = "1200" Then Response.Write(" selected ")%>>-Noon-</option>
					<option value="1215"<%If TimeOfDay = "1215" Then Response.Write(" selected ")%>>12:15 PM</option>
					<option value="1230"<%If TimeOfDay = "1230" Then Response.Write(" selected ")%>>12:30 PM</option>
					<option value="1245"<%If TimeOfDay = "1245" Then Response.Write(" selected ")%>>12:45 PM</option>
					<option value="1300"<%If TimeOfDay = "1300" Then Response.Write(" selected ")%>>1:00 PM</option>
					<option value="1315"<%If TimeOfDay = "1315" Then Response.Write(" selected ")%>>1:15 PM</option>
					<option value="1330"<%If TimeOfDay = "1330" Then Response.Write(" selected ")%>>1:30 PM</option>
					<option value="1345"<%If TimeOfDay = "1345" Then Response.Write(" selected ")%>>1:45 PM</option>
					<option value="1400"<%If TimeOfDay = "1400" Then Response.Write(" selected ")%>>2:00 PM</option>
					<option value="1415"<%If TimeOfDay = "1415" Then Response.Write(" selected ")%>>2:15 PM</option>
					<option value="1430"<%If TimeOfDay = "1430" Then Response.Write(" selected ")%>>2:30 PM</option>
					<option value="1445"<%If TimeOfDay = "1445" Then Response.Write(" selected ")%>>2:45 PM</option>
					<option value="1500"<%If TimeOfDay = "1500" Then Response.Write(" selected ")%>>3:00 PM</option>
					<option value="1515"<%If TimeOfDay = "1515" Then Response.Write(" selected ")%>>3:15 PM</option>
					<option value="1530"<%If TimeOfDay = "1530" Then Response.Write(" selected ")%>>3:30 PM</option>
					<option value="1545"<%If TimeOfDay = "1545" Then Response.Write(" selected ")%>>3:45 PM</option>
					<option value="1600"<%If TimeOfDay = "1600" Then Response.Write(" selected ")%>>4:00 PM</option>
					<option value="1615"<%If TimeOfDay = "1615" Then Response.Write(" selected ")%>>4:15 PM</option>
					<option value="1630"<%If TimeOfDay = "1630" Then Response.Write(" selected ")%>>4:30 PM</option>
					<option value="1645"<%If TimeOfDay = "1645" Then Response.Write(" selected ")%>>4:45 PM</option>
					<option value="1700"<%If TimeOfDay = "1700" Then Response.Write(" selected ")%>>5:00 PM</option>
					<option value="1715"<%If TimeOfDay = "1715" Then Response.Write(" selected ")%>>5:15 PM</option>
					<option value="1730"<%If TimeOfDay = "1730" Then Response.Write(" selected ")%>>5:30 PM</option>
					<option value="1745"<%If TimeOfDay = "1745" Then Response.Write(" selected ")%>>5:45 PM</option>
					<option value="1800"<%If TimeOfDay = "1800" Then Response.Write(" selected ")%>>6:00 PM</option>
					<option value="1815"<%If TimeOfDay = "1815" Then Response.Write(" selected ")%>>6:15 PM</option>
					<option value="1830"<%If TimeOfDay = "1830" Then Response.Write(" selected ")%>>6:30 PM</option>
					<option value="1845"<%If TimeOfDay = "1845" Then Response.Write(" selected ")%>>6:45 PM</option>
					<option value="1900"<%If TimeOfDay = "1900" Then Response.Write(" selected ")%>>7:00 PM</option>
					<option value="1915"<%If TimeOfDay = "1915" Then Response.Write(" selected ")%>>7:15 PM</option>
					<option value="1930"<%If TimeOfDay = "1930" Then Response.Write(" selected ")%>>7:30 PM</option>
					<option value="1945"<%If TimeOfDay = "1945" Then Response.Write(" selected ")%>>7:45 PM</option>
					<option value="2000"<%If TimeOfDay = "2000" Then Response.Write(" selected ")%>>8:00 PM</option>
					<option value="2015"<%If TimeOfDay = "2015" Then Response.Write(" selected ")%>>8:15 PM</option>
					<option value="2030"<%If TimeOfDay = "2030" Then Response.Write(" selected ")%>>8:30 PM</option>
					<option value="2045"<%If TimeOfDay = "2045" Then Response.Write(" selected ")%>>8:45 PM</option>
					<option value="2100"<%If TimeOfDay = "2100" Then Response.Write(" selected ")%>>9:00 PM</option>
					<option value="2115"<%If TimeOfDay = "2115" Then Response.Write(" selected ")%>>9:15 PM</option>
					<option value="2130"<%If TimeOfDay = "2130" Then Response.Write(" selected ")%>>9:30 PM</option>
					<option value="2145"<%If TimeOfDay = "2145" Then Response.Write(" selected ")%>>9:45 PM</option>
					<option value="2200"<%If TimeOfDay = "2200" Then Response.Write(" selected ")%>>10:00 PM</option>
					<option value="2215"<%If TimeOfDay = "2215" Then Response.Write(" selected ")%>>10:15 PM</option>
					<option value="2230"<%If TimeOfDay = "2230" Then Response.Write(" selected ")%>>10:30 PM</option>
					<option value="2245"<%If TimeOfDay = "2245" Then Response.Write(" selected ")%>>10:45 PM</option>
					<option value="2300"<%If TimeOfDay = "2300" Then Response.Write(" selected ")%>>11:00 PM</option>
					<option value="2315"<%If TimeOfDay = "2315" Then Response.Write(" selected ")%>>11:15 PM</option>
					<option value="2330"<%If TimeOfDay = "2330" Then Response.Write(" selected ")%>>11:30 PM</option>
					<option value="2345"<%If TimeOfDay = "2345" Then Response.Write(" selected ")%>>11:45 PM</option>	
	
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
			<select class="form-control" id="selDays" name="selDays">
				<% If not IsNumeric(Days) Then Days = 0
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
				<option value="0"<%If EmailTo = "0" Then Response.Write(" selected ")%>>--- none from here ---</option>
				<option value="<%=Session("UserNo")%>"<%If UserInList(Session("UserNo"),Emailto) = True Then Response.Write(" selected ")%>><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option>
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
						If UserInList(rs("UserNo"),Emailto) = True Then
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
			<textarea class="form-control email-alert-line" rows="4" id="txtaAdditionalEmails" name="txtaAdditionalEmails"><%=AdditionalEmails%></textarea>
			<strong>Separate multiple email addresses with a semicolon</strong>
		</div>
		<!-- eof separate emails !-->
	
		<!-- verbiage !-->
		<div class="col-lg-3">
			<strong>Verbiage to include in alert email</strong>
			<textarea class="form-control" rows="4" id="txtaVerbiageEmail" name="txtaVerbiageEmail"><%=VerbiageEmail%></textarea>
			<div class="col-lg-4" id="pnlLog" style="display: none;">
				<input type="checkbox" id="chkLog" name="chkLog" <%If IncludeLog = vbTrue Then Response.Write(" checked ")%>><strong> Include log</strong>
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
				<option value="0" <%If TextTo = "0" Then Response.Write(" selected ")%>>--- none from here ---</option>
				<option value="<%=Session("UserNo")%>"<%If UserInList(Session("UserNo"),Textto) = True Then Response.Write(" selected ")%>><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option>
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
						If UserInList(rs("UserNo"),Textto) = True Then
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
		<textarea class="form-control email-alert-line" rows="4" id="txtaAdditionalTexts" name="txtaAdditionalTexts"><%=AdditionalTexts%></textarea>
		<strong>Separate multiple phone numbers with a semicolon</strong>
	</div>
	<!-- eof separate emails !-->
	
	<!-- verbiage !--> 
	<div class="col-lg-3">
		<strong>Message to include in text alert</strong><br>
		<input type="text" class="form-control" id="txtAlertTextVerbiage" name="txtAlertTextVerbiage" value="<%=TextVerbiage%>">
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
        <div class="col-lg-3">
	        <strong>Minutes</strong>
	        <select class="form-control" id="selLimitMinutes" name="selLimitMinutes">
				<%
					For x = 5 to 480 Step 5 ' 8 hours
						If x mod 60 = 0 Then
							If x = cint(LimitMinutes) Then 
								Response.Write("<option value='" & x & "' selected>" & x & " (" & x/60 & "hours)" & "</option>")
							else
								Response.Write("<option value='" & x & "'>" & x & " (" & x/60 & "hours)" & "</option>")
							End If
						Else
							If x = cint(LimitMinutes) Then 
								Response.Write("<option value='" & x & "' selected>" & x & "</option>")
							Else
								Response.Write("<option value='" & x & "'>" & x & "</option>")
							End If
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
        <div class="col-lg-3">
	        <strong>Maximum</strong>
	        <select class="form-control" id="selLimitMaxTimes" name="selLimitMaxTimes">
	        	<%
					For x = 1 to 10
						If x = cint(LimitMaxTimes) Then
							Response.Write("<option value='" & x & "' selected>" & x & "</option>")
						Else
							Response.Write("<option value='" & x & "'>" & x & "</option>")
						End If
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
<%
Function UserInList(UserToFind,UserList)

	result = False
	
	UserNoList = Split(UserList,",")
	For x = 0 To UBound(UserNoList)
		If cint(UserToFind) = cint(UserNoList(x)) Then
			result = True
			Exit For
		End If
	Next
	
	UserInList = result
	
End Function
%><!--#include file="../../inc/footer-main.asp"-->