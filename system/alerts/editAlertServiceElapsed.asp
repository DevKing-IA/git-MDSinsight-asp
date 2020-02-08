<!--#include file="../../inc/header.asp"-->
<% InternalAlertRecNumber = Request.QueryString("a") 
If InternalAlertRecNumber = "" Then Response.Redirect("main.asp")
%>
<body onload="load()">

<script type="text/javascript">
	function cndChanged() {
		$("#pnlminutes").show();
		$("#ticketHold").hide();
		if (document.getElementById('selCond').selectedIndex == '6'){
			$("#pnlminutes").hide();
			}
		if (document.getElementById('selCond').selectedIndex == '7'){
			$("#ticketHold").show();
			}
		if (document.getElementById('selCond').selectedIndex == '8'){
			$("#ticketHold").show();
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


<h1 class="page-header"><i class="fa fa-fw fa-exclamation"></i>Edit Service Ticket Alert</h1>

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
	EmailPrimarySls = rs("EmailPrimarySls")
	EmailSecondarySls = rs("EmailSecondarySls")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

If IsNull(LimitMinutes) Then LimitMinutes = 60
If IsNull(LimitMaxTimes) Then LimitMaxTimes = 1
If LimitMinutes = "" Then LimitMinutes = 60
If LimitMaxTimes = "" Then LimitMaxTimes = 1

If EmailPrimarySls = 0 Then
	PrimarySls = ""
Else
	PrimarySls = "checked"
End If

If EmailSecondarySls = 0 Then
	SecondarySls = ""
Else
	SecondarySls = "checked"
End If
								
%>


<form method="POST" action="editAlertServiceElapsed_submit.asp" name="frmEditAlert" id="frmEditAlert"  onsubmit="return checkform();">


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
				<option value="NotDispatched"<%If Condition = "NotDispatched" Then Response.Write(" selected ")%>>Ticket has not been dispatched for longer than</option>
				<option value="NoACK"<%If Condition = "NoACK" Then Response.Write(" selected ")%>>Ticket dispatched but no acknowledgement has been received longer than</option>
				<option value="NoOnSite"<%If Condition = "NoOnSite" Then Response.Write(" selected ")%>>No tech has been onsite yet and the ticket has been open longer than</option>
				<option value="OpenTooLong"<%If Condition = "OpenTooLong" Then Response.Write(" selected ")%>>Ticket has been open and unresolved longer than</option>
				<option value="RedispatchTooLong"<%If Condition = "RedispatchTooLong" Then Response.Write(" selected ")%>>Ticket has been awaiting redispatch longer than</option>
				<option value="AnyStage"<%If Condition = "AnyStage" Then Response.Write(" selected ")%>>Ticket has been idle in any stage longer than</option>
				<option value="Declined"<%If Condition = "Declined" Then Response.Write(" selected ")%>>The tech has declined the dispatch</option>
				<option value="ARHold"<%If Condition = "ARHold" Then Response.Write(" selected ")%>>When a ticket has been on A/R Hold longer than</option>
				<option value="GPHold"<%If Condition = "GPHold" Then Response.Write(" selected ")%>>When a ticket has been on GP Hold longer than</option>
				
			</select>
		</div>
		<!-- eof when select !-->

		
		<!-- minutes !-->
		<div class="col-lg-2" id="pnlminutes" style="display: none;">
			<strong>Minutes</strong>
			<select class="form-control" id="selMinutes" name="selMinutes">
				<option value="1"<%If Minutes = 1 Then Response.Write(" selected ")%>>1</option>
				<%
					For x = 5 to 7200 Step 5 ' 5 days
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

		<div class="row row-line" id="ticketHold" style="display: none;">
		<div class="col-lg-2">
			<br><br>
			<input type="checkbox" id="chkPrimarySalesperson" name="chkPrimarySalesperson" <%=PrimarySls%>>&nbsp;&nbsp;<strong>Email Primary Salesperson</strong><br>
			<input type="checkbox" id="chkSecondrySalesperson" name="chkSecondrySalesperson" <%=SecondarySls%>>&nbsp;&nbsp;<strong>Email Secondry Salesperson</strong>
	 	</div>
		</div>
		
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
					<p align="right"><a href="<%= BaseURL %>system/alerts/main.asp#ServiceElapsed">
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