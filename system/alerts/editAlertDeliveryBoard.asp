<!--#include file="../../inc/header.asp"-->
<% InternalAlertRecNumber = Request.QueryString("a") 
If InternalAlertRecNumber = "" Then Response.Redirect("main.asp")
%>
<body onload="load()">

<script type="text/javascript">
function load(){
if (document.getElementById('selCond').selectedIndex == '2'){
	document.getElementById('pnlTimeOfDay').style.display="block";
	document.getElementById('pnlLimits').style.display="block";
}
}</script>


<script type="text/javascript">
	function cndChanged() {
		$("#pnlTimeOfDay").hide();
		$("#pnlLimits").hide();
		if (document.getElementById('selCond').selectedIndex == '0'){
			$("#pnlTimeOfDay").show();
			$("#pnlLimits").show();
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


<h1 class="page-header"><i class="fa fa-fw fa-exclamation"></i>Edit Delivery Board Alert</h1>

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
	TimeOfDay = rs("TimeOfDay")
	Emailto = rs("EmailToUserNos") 
	AdditionalEmails = rs("AdditionalEmails")
	VerbiageEmail = rs("EmailVerbiage")
	Textto = rs("TextToUserNos")
	AdditionalTexts = rs("AdditionalText")
	TextVerbiage = rs("TextVerbiage") 
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


<form method="POST" action="editAlertDeliveryBoard_submit.asp" name="frmEditAlert" id="frmEditAlert"  onsubmit="return checkform();">


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
	
		<!-- when select !-->
		<div class="col-lg-3">
			<select class="form-control when-line" name="selCond" id="selCond" onchange="cndChanged();">
				<option value="AM_Overdue"<%If Condition = "AM_Overdue" Then Response.Write(" selected ")%>>An AM Delivery Is Not Delivered By Specified Time</option>
				<option value="Priority_Overdue" <%If Condition = "Priority_Overdue" Then Response.Write(" selected ")%>>A Priority Delivery Is Not Delivered By Specified Time</option>
				<option value="Priority No Delivery" <%If Condition = "Priority No Delivery" Then Response.Write(" selected ")%>>A Priority Delivery Is Marked As No Delivery</option>
				<option value="Delivered"<%If Condition = "Delivered" Then Response.Write(" selected ")%>>An Invoice Is Marked As Delivered</option>
				<option value="No Delivery"<%If Condition = "No Delivery" Then Response.Write(" selected ")%>>An Invoice Is Marked As No Delivery</option>
				<option value="Partial"<%If Condition = "Partial" Then Response.Write(" selected ")%>>A Partial Delivery Is Made</option>
			</select>
		</div>
		<!-- eof when select !-->


		<!-- minutes !-->
		<div class="col-lg-2" id="pnlTimeOfDay" style="display: none;">
			<strong>Time Of Day</strong>
			<select class="form-control" id="selTimeOfDay" name="selTimeOfDay">
				<option value="500"<%If TimeOfDay = "500" Then Response.Write(" selected ")%>>5:00 AM</option>
				<option value="515"<%If TimeOfDay = "515" Then Response.Write(" selected ")%>>5:15 AM</option>
				<option value="503"<%If TimeOfDay = "530" Then Response.Write(" selected ")%>>5:30 AM</option>
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
				<option value="1200"<%If TimeOfDay = "1200" Then Response.Write(" selected ")%>>12:00 PM</option>
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
					<p align="right"><a href="<%= BaseURL %>system/alerts/main.asp#DeliveryAlerts">
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