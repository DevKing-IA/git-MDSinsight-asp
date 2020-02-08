<!--#include file="../../inc/header.asp"-->

<!-- function that gets the value of the tab when it is clicked and then
updates the value of a hidden form field so when the page posts, it returns
back to the tab that was previously opened -->

<script type="text/javascript">
	$(function () {
		$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
		var target = $(e.target).attr("href");
		$('input[name="txtTab"]').val(target);
		//alert(target);
		});
	})
</script>

<style type="text/css">
		.form-control{
			overflow-x: hidden;
			}
			
		.nav-tabs>li>a{
			background: #f5f5f5;
			border: 1px solid #ccc;
			color: #000;
		}
		
		.nav-tabs>li>a:hover{
			border: 1px solid #ccc;
		}
		
		.nav-tabs>li.active>a, .nav-tabs>li.active>a:focus, .nav-tabs>li.active>a:hover{
			color: #000;
			border: 1px solid #ccc;
		}

		.post-labels{
	 		padding-top: 5px;
	 	}
	 	
	 	.row-margin{
		 	margin-bottom: 20px;
		 	margin-top: 20px;
	 	}
	 	
	 	h3{
		 	margin-top: 0px;
	 	}
	 	
	 	.table-size .category{
		 	width: 35%;
		 	font-weight: normal;
	 	}
	 	
	 	.table-size .group-name{
		 	width: 40%
	 	}
	 	
	 	.table-size .sort-order{
		 	width: 10%;
	 	}
	 	
	 	.table-size .display{
		 	width: 15%;
	 	}
		
	</style>


<%
UseSimpleEmailFormat = Request.Form("chkUseSimpleEmailFormat")
If Request.Form("chkUseSimpleEmailFormat")="on" then UseSimpleEmailFormat = vbTrue Else UseSimpleEmailFormat = vbFalse
OpenViaWeb_SendToWebUser = Request.Form("chkOpenViaWeb_SendToWebUser")
If Request.Form("chkOpenViaWeb_SendToWebUser")="on" then OpenViaWeb_SendToWebUser = vbTrue Else OpenViaWeb_SendToWebUser = vbFalse
OpenViaWeb_SendToCust = Request.Form("chkOpenViaWeb_SendToCust")
If Request.Form("chkOpenViaWeb_SendToCust")="on" then OpenViaWeb_SendToCust = vbTrue Else OpenViaWeb_SendToCust = vbFalse
Open_ViaInsight_SendToCust = Request.Form("chkOpen_ViaInsight_SendToCust")
If Request.Form("chkOpen_ViaInsight_SendToCust")="on" then Open_ViaInsight_SendToCust = vbTrue Else Open_ViaInsight_SendToCust = vbFalse
Open_ViaFServ_SendToCust = Request.Form("chkOpen_ViaFServ_SendToCust")
If Request.Form("chkOpen_ViaFServ_SendToCust")="on" then Open_ViaFServ_SendToCust = vbTrue Else Open_ViaFServ_SendToCust = vbFalse
Open_ViaMDS_SendToCust = Request.Form("chkOpen_ViaMDS_SendToCust")
If Request.Form("chkOpen_ViaMDS_SendToCust")="on" then Open_ViaMDS_SendToCust = vbTrue Else Open_ViaMDS_SendToCust = vbFalse
Close_ViaInsight_SendToCust = Request.Form("chkClose_ViaInsight_SendToCust")
If Request.Form("chkClose_ViaInsight_SendToCust")="on" then Close_ViaInsight_SendToCust = vbTrue Else Close_ViaInsight_SendToCust = vbFalse
Close_ViaFServ_SendToCust = Request.Form("chkClose_ViaFServ_SendToCust")
If Request.Form("chkClose_ViaFServ_SendToCust")="on" then Close_ViaFServ_SendToCust = vbTrue Else Close_ViaFServ_SendToCust = vbFalse
Close_ViaMDS_SendToCust = Request.Form("chkClose_ViaMDS_SendToCust")
If Request.Form("chkClose_ViaMDS_SendToCust")="on" then Close_ViaMDS_SendToCust = vbTrue Else Close_ViaMDS_SendToCust = vbFalse
Cancel_ViaInsight_SendToCust = Request.Form("chkCancel_ViaInsight_SendToCust")
If Request.Form("chkCancel_ViaInsight_SendToCust")="on" then Cancel_ViaInsight_SendToCust = vbTrue Else Cancel_ViaInsight_SendToCust = vbFalse
Cancel_ViaFServ_SendToCust = Request.Form("chkCancel_ViaFServ_SendToCust")
If Request.Form("chkCancel_ViaFServ_SendToCust")="on" then Cancel_ViaFServ_SendToCust = vbTrue Else Cancel_ViaFServ_SendToCust = vbFalse
Cancel_ViaMDS_SendToCust = Request.Form("chkCancel_ViaMDS_SendToCust")
If Request.Form("chkCancel_ViaMDS_SendToCust")="on" then Cancel_ViaMDS_SendToCust = vbTrue Else Cancel_ViaMDS_SendToCust = vbFalse
RealtimeAlertsOn = Request.Form("chkRealtimeAlertsOn")
If Request.Form("chkRealtimeAlertsOn") = "on" then RealtimeAlertsOn = vbTrue Else RealtimeAlertsOn = vbFalse
SendAlertToServiceManagers = Request.Form("chkSendAlertToServiceManagers")
If Request.Form("chkSendAlertToServiceManagers") = "on" then SendAlertToServiceManagers = vbTrue Else SendAlertToServiceManagers = vbFalse
SendAlertHours = Request.Form("selSendAlertHours")
SendAlertsSkipDispatched = Request.Form("chkSendAlertsSkipDispatched")
If Request.Form("chkSendAlertsSkipDispatched") = "on" then SendAlertsSkipDispatched = vbTrue Else SendAlertsSkipDispatched = vbFalse
EscalationAlertsOn = Request.Form("chkEscalationAlertsOn")
If Request.Form("chkEscalationAlertsOn") = "on" then EscalationAlertsOn = vbTrue Else EscalationAlertsOn = vbFalse
EscalationAlertHours = Request.Form("selEscalationAlertHours")
EscalationAlertsSkipDispatched = Request.Form("chkEscalationAlertsSkipDispatched")
If Request.Form("chkEscalationAlertsSkipDispatched") = "on" then EscalationAlertsSkipDispatched = vbTrue Else EscalationAlertsSkipDispatched = vbFalse
AlertsDuringBizHoursOnly = Request.Form("chkAlertsDuringBizHoursOnly")
If Request.Form("chkAlertsDuringBizHoursOnly") = "on" then AlertsDuringBizHoursOnly = vbTrue Else AlertsDuringBizHoursOnly = vbFalse
HoldAlertsOn = Request.Form("chkHoldAlertsOn")
If Request.Form("chkHoldAlertsOn") = "on" then HoldAlertsOn = vbTrue Else HoldAlertsOn = vbFalse

CompletedPMCallEmailOn = Request.Form("chkPMCallEmail")
If Request.Form("chkPMCallEmail") = "on" then CompletedPMCallEmailOn = vbTrue Else CompletedPMCallEmailOn = vbFalse

DoNotSendClientCompletedPMCall = Request.Form("chkDoNotSendClientCompletedPMCall")
If Request.Form("chkDoNotSendClientCompletedPMCall") = "on" then DoNotSendClientCompletedPMCall = vbTrue Else DoNotSendClientCompletedPMCall = vbFalse

PMCallPDFIncludeServiceNotes = Request.Form("chkPMCallPDFIncludeServiceNotes")
If Request.Form("chkPMCallPDFIncludeServiceNotes") = "on" then PMCallPDFIncludeServiceNotes = vbTrue Else PMCallPDFIncludeServiceNotes = vbFalse

SendHoldAlertToFinanceManagers = Request.Form("chkSendHoldAlertToFinanceManagers")
If Request.Form("chkSendHoldAlertToFinanceManagers") = "on" then SendHoldAlertToFinanceManagers = vbTrue Else SendHoldAlertToFinanceManagers = vbFalse
SendHoldAlertHours = Request.Form("selSendHoldAlertHours")
HoldEscalationAlertsOn = Request.Form("chkHoldEscalationAlertsOn")
If Request.Form("chkHoldEscalationAlertsOn") = "on" then HoldEscalationAlertsOn = vbTrue Else HoldEscalationAlertsOn = vbFalse
HoldEscalationAlertHours = Request.Form("selHoldEscalationAlertHours")
SendAlertToAdditionalEmails = Request.Form("txtSendAlertToAdditionalEmails")
SendAlertToAdditionalEmails = Trim(SendAlertToAdditionalEmails)
SendAlertToAdditionalEmails = Replace(SendAlertToAdditionalEmails," ","")

EscalationAlertToEmails = Trim(Request.Form("txtEscalationAlertToEmails"))
EscalationAlertToEmails = Trim(EscalationAlertToEmails)
EscalationAlertToEmails = Replace(EscalationAlertToEmails," ","")


SendHoldAlertToAdditionalEmails = Request.Form("txtSendHoldAlertToAdditionalEmails")
SendHoldAlertToAdditionalEmails = Trim(SendHoldAlertToAdditionalEmails)
SendHoldAlertToAdditionalEmails = Replace(SendHoldAlertToAdditionalEmails," ","")

HoldEscalationAlertToEmails = Trim(Request.Form("txtHoldEscalationAlertToEmails"))
HoldEscalationAlertToEmails = Trim(HoldEscalationAlertToEmails)
HoldEscalationAlertToEmails = Replace(HoldEscalationAlertToEmails," ","")

SendCompletedPMCallTo = Request.Form("txtPMCallEmailsTo")
SendCompletedPMCallTo = Trim(SendCompletedPMCallTo)
SendCompletedPMCallTo = Replace(SendCompletedPMCallTo," ","")


If Request.ServerVariables("REQUEST_METHOD") = "POST" Then 


	Dummy = HandleAuditTrail() ' The code for audit trail is huge, so put it at the bottom
	
	SendAlertToAdditionalEmails = Trim(SendAlertToAdditionalEmails)
	SendAlertToAdditionalEmails = Replace(SendAlertToAdditionalEmails," ","")
	SendAlertToAdditionalEmails= Replace(SendAlertToAdditionalEmails,vbCRLF,"")
	SendAlertToAdditionalEmails= Replace(SendAlertToAdditionalEmails,vbTab,"")
	If Trim(SendAlertToAdditionalEmails) <> "" Then
		If Right(SendAlertToAdditionalEmails,1)<>";" Then SendAlertToAdditionalEmails = SendAlertToAdditionalEmails & ";"
	End If
	EscalationAlertToEmails = Trim(EscalationAlertToEmails)
	EscalationAlertToEmails = Replace(EscalationAlertToEmails," ","")
	EscalationAlertToEmails = Replace(EscalationAlertToEmails ,vbCRLF,"")
	EscalationAlertToEmails = Replace(EscalationAlertToEmails ,vbTAB,"")
	If Trim(EscalationAlertToEmails) <> "" Then
		If Right(EscalationAlertToEmails,1)<>";" Then EscalationAlertToEmails = EscalationAlertToEmails & ";"
	End If
	SendHoldAlertToAdditionalEmails = Trim(SendHoldAlertToAdditionalEmails)
	SendHoldAlertToAdditionalEmails = Replace(SendHoldAlertToAdditionalEmails," ","")
	SendHoldAlertToAdditionalEmails= Replace(SendHoldAlertToAdditionalEmails,vbCRLF,"")
	SendHoldAlertToAdditionalEmails= Replace(SendHoldAlertToAdditionalEmails,vbTab,"")
	If Trim(SendHoldAlertToAdditionalEmails) <> "" Then
		If Right(SendHoldAlertToAdditionalEmails,1)<>";" Then SendHoldAlertToAdditionalEmails = SendHoldAlertToAdditionalEmails & ";"
	End If
	SendCompletedPMCallTo = Trim(SendCompletedPMCallTo)
	SendCompletedPMCallTo = Replace(SendCompletedPMCallTo," ","")
	SendCompletedPMCallTo = Replace(SendCompletedPMCallTo,vbCRLF,"")
	SendCompletedPMCallTo = Replace(SendCompletedPMCallTo,vbTab,"")
	If Trim(SendCompletedPMCallTo) <> "" Then
		If Right(SendCompletedPMCallTo,1)<>";" Then SendCompletedPMCallTo = SendCompletedPMCallTo & ";"
	End If
	HoldEscalationAlertToEmails = Trim(HoldEscalationAlertToEmails)
	HoldEscalationAlertToEmails = Replace(HoldEscalationAlertToEmails," ","")
	HoldEscalationAlertToEmails = Replace(HoldEscalationAlertToEmails ,vbCRLF,"")
	HoldEscalationAlertToEmails = Replace(HoldEscalationAlertToEmails ,vbTAB,"")
	If Trim(HoldEscalationAlertToEmails) <> "" Then
		If Right(HoldEscalationAlertToEmails,1)<>";" Then HoldEscalationAlertToEmails = HoldEscalationAlertToEmails & ";"
	End If

	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_EmailService SET OpenViaWeb_SendToWebUser =" & OpenViaWeb_SendToWebUser &", "
	SQL = SQL & "Internal_Simple= " & UseSimpleEmailFormat & ", "
	SQL = SQL & "OpenViaWeb_SendToCust = " & OpenViaWeb_SendToCust & ", "
	SQL = SQL & "Open_ViaInsight_SendToCust =" & Open_ViaInsight_SendToCust & ", "
	SQL = SQL & "Open_ViaFServ_SendToCust =" & Open_ViaFServ_SendToCust & ", "
	SQL = SQL & "Open_ViaMDS_SendToCust = "& Open_ViaMDS_SendToCust& ", "
	SQL = SQL & "Close_ViaInsight_SendToCust = "& Close_ViaInsight_SendToCust & ", "
	SQL = SQL & "Close_ViaFServ_SendToCust = " & Close_ViaFServ_SendToCust & ", "
	SQL = SQL & "Close_ViaMDS_SendToCust = " & Close_ViaMDS_SendToCust & ","
	SQL = SQL & "Cancel_ViaInsight_SendToCust = " & Cancel_ViaInsight_SendToCust & ","
	SQL = SQL & "Cancel_ViaFServ_SendToCust = " & Cancel_ViaFServ_SendToCust & ","
	SQL = SQL & "Cancel_ViaMDS_SendToCust = " & Cancel_ViaMDS_SendToCust & ","
	SQL = SQL & "RealtimeAlertsOn = " & RealtimeAlertsOn & ","
	SQL = SQL & "AlertsDuringBizHoursOnly = " & AlertsDuringBizHoursOnly & ","
	SQL = SQL & "SendAlertToServiceManagers = " & SendAlertToServiceManagers & ","
	SQL = SQL & "SendAlertToAdditionalEmails = '" & SendAlertToAdditionalEmails & "',"
	SQL = SQL & "SendAlertHours = " & SendAlertHours & ","
	SQL = SQL & "SendAlertsSkipDispatched = " & SendAlertsSkipDispatched & ", "
	SQL = SQL & "EscalationAlertsOn = " & EscalationAlertsOn & ", "
	SQL = SQL & "EscalationAlertToEmails = '" & EscalationAlertToEmails & "', "
	SQL = SQL & "EscalationAlertHours = " & EscalationAlertHours & ", "
	SQL = SQL & "EscalationAlertsSkipDispatched = " & EscalationAlertsSkipDispatched & ","
	SQL = SQL & "HoldAlertsOn = " & HoldAlertsOn & ","
	SQL = SQL & "CompletedPMCallEmailOn = " & CompletedPMCallEmailOn & ","
	SQL = SQL & "DoNotSendClientCompletedPMCall = " & DoNotSendClientCompletedPMCall & ","
	SQL = SQL & "PMCallPDFIncludeServiceNotes = " & PMCallPDFIncludeServiceNotes & ","	
	SQL = SQL & "SendHoldAlertToFinanceManagers = " & SendHoldAlertToFinanceManagers & ","
	SQL = SQL & "SendHoldAlertToAdditionalEmails = '" & SendHoldAlertToAdditionalEmails & "',"	
	SQL = SQL & "SendCompletedPMCallTo = '" & SendCompletedPMCallTo & "',"	
	SQL = SQL & "SendHoldAlertHours = " & SendHoldAlertHours & ","
	SQL = SQL & "HoldEscalationAlertsOn = " & HoldEscalationAlertsOn & ", "
	SQL = SQL & "HoldEscalationAlertToEmails = '" & HoldEscalationAlertToEmails & "', "
	SQL = SQL & "HoldEscalationAlertHours = " & HoldEscalationAlertHours 

	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	
	'Response.Write(SQL)
	
	Set rs = cnn8.Execute(SQL)

	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	ActiveTab = Request.Form("txtTab")
	
	Response.Redirect ("service_tickets.asp" & ActiveTab)
	
End If

SQL = "SELECT * FROM Settings_EmailService "

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)

If not rs.EOF Then
	UseSimpleEmailFormat = rs("Internal_Simple")
	OpenViaWeb_SendToWebUser= rs("OpenViaWeb_SendToWebUser")
	OpenViaWeb_SendToCust= rs("OpenViaWeb_SendToCust")
	Open_ViaInsight_SendToCust = rs("Open_ViaInsight_SendToCust")
	Open_ViaFServ_SendToCust= rs("Open_ViaFServ_SendToCust")
	Open_ViaMDS_SendToCust= rs("Open_ViaMDS_SendToCust")
	Close_ViaInsight_SendToCust= rs("Close_ViaInsight_SendToCust")
	Close_ViaFServ_SendToCust= rs("Close_ViaFServ_SendToCust")
	Close_ViaMDS_SendToCust= rs("Close_ViaMDS_SendToCust")
	Cancel_ViaInsight_SendToCust = rs("Cancel_ViaInsight_SendToCust")
	Cancel_ViaFServ_SendToCust = rs("Cancel_ViaFServ_SendToCust")
	Cancel_ViaMDS_SendToCust = rs("Cancel_ViaMDS_SendToCust")
	RealtimeAlertsOn = rs("RealtimeAlertsOn")
	AlertsDuringBizHoursOnly = rs("AlertsDuringBizHoursOnly")
	SendAlertToServiceManagers = rs("SendAlertToServiceManagers")
	SendAlertToAdditionalEmails = rs("SendAlertToAdditionalEmails")
	SendAlertHours = rs("SendAlertHours")
	SendAlertsSkipDispatched = rs("SendAlertsSkipDispatched")
	EscalationAlertsOn = rs("EscalationAlertsOn")
	EscalationAlertToEmails = rs("EscalationAlertToEmails")
	EscalationAlertHours = rs("EscalationAlertHours")
	EscalationAlertsSkipDispatched = rs("EscalationAlertsSkipDispatched")
	HoldAlertsOn = rs("HoldAlertsOn")
	CompletedPMCallEmailOn = rs("CompletedPMCallEmailOn")
	DoNotSendClientCompletedPMCall= rs("DoNotSendClientCompletedPMCall")
	PMCallPDFIncludeServiceNotes = rs("PMCallPDFIncludeServiceNotes")
	DoNotSendClientCompletedPMCall= rs("DoNotSendClientCompletedPMCall")
	SendHoldAlertToFinanceManagers = rs("SendHoldAlertToFinanceManagers")
	SendHoldAlertToAdditionalEmails = rs("SendHoldAlertToAdditionalEmails")
	SendCompletedPMCallTo = rs("SendCompletedPMCallTo")
	SendHoldAlertHours = rs("SendHoldAlertHours")
	HoldEscalationAlertsOn = rs("HoldEscalationAlertsOn")
	HoldEscalationAlertToEmails = rs("HoldEscalationAlertToEmails")
	HoldEscalationAlertHours = rs("HoldEscalationAlertHours")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing

%>


<h1 class="page-header"><i class="fa fa-envelope-o"></i> Service Ticket Email Settings</h1>

<form method="post" action="service_tickets.asp" name="frmServiceTickets" id="frmServiceTickets">

<input type="hidden" name="txtTab" id="txtTab" value="">

<!-- tabs start here !-->
<div class="row row-margin">
	<div class="col-lg-12">
	
		<!-- tabs navigation !-->
		<ul class="nav nav-tabs" role="tablist">
			    <li role="presentation" <% If ActiveTab = "" OR ActiveTab="open" Then Response.write("class='active'") %>><a href="#open" aria-controls="manage" role="tab" data-toggle="tab">Open</a></li>
   			    <li role="presentation" <% If ActiveTab = "close" Then Response.write("class='active'") %>><a href="#close" aria-controls="tab3" role="tab" data-toggle="tab">Close</a></li>
			    <li role="presentation" <% If ActiveTab = "cancel" Then Response.write("class='active'") %>><a href="#cancel" aria-controls="tab3" role="tab" data-toggle="tab">Cancel</a></li>
		        <li role="presentation" <% If ActiveTab = "realtime" Then Response.write("class='active'") %>><a href="#realtime" aria-controls="tab3" role="tab" data-toggle="tab">Realtime Alerts</a></li>
		</ul>
			<!-- eof tabs navigation !-->
			
	<!-- tabs content !-->
	<div class="tab-content">


		
	<!-- Open Service Ticket Tab !-->
	<div role="tabpanel" class="tab-pane fade in active" id="open"> 
      
	          <!-- main row !-->
	          <div class="row row-margin">
	         	
	         		<% If MUV_READ("BackendSystem") = "Metroplex" Then %>
			         	<!-- left col !-->
			         	<div class="col-lg-3">
					         	<strong><u>Email Hierarchy</u></strong><br>
					         	1. If via web then login email<br>
					         	2. Service email<br>
					         	3. Order email<br>
					         	4. General email<br>
					         	5. Notify Rep<br>
					         	6. Notify IT Dept<br>
			         	</div>
			         	<!-- eof left col -->
		         	<% End If %>
		         	
		         	<!-- right col !-->	 
		         		<div class="col-lg-6">

			         	<!-- row with data !-->
			         	<div class="row row-margin row-data">
					   	   	<div class="col-lg-8">Use "Simple Format" for Internal Emails</div>
				       	  	<div class="col-lg-2">
								<%
								Response.Write("<input type='checkbox' class='check' id='chkUseSimpleEmailFormat' name='chkUseSimpleEmailFormat'")
								If UseSimpleEmailFormat = vbTrue Then Response.Write(" checked ")
								Response.Write(">")
								%>
				       	  	</div>
						</div>
				        <!-- eof row with data !-->
				         	

			         	<!-- row with data !-->
			         	<div class="row row-margin row-data">
					   	   	<div class="col-lg-8">When a service ticket is opened via <b>Website</b>, send an email to email address of the logged in user.</div>
				       	  	<div class="col-lg-2">
								<%
								Response.Write("<input type='checkbox' class='check' id='chkOpenViaWeb_SendToWebUser' name='chkOpenViaWeb_SendToWebUser'")
								If OpenViaWeb_SendToWebUser = vbTrue Then Response.Write(" checked ")
								Response.Write(">")
								%>
				       	  	</div>
						</div>
				        <!-- eof row with data !-->
				         	
				         	
			         	<!-- row with data !-->
			         	<% If MUV_READ("BackendSystem") = "Metroplex" Then %>
				         	<div class="row row-margin row-data">
				         	   	<div class="col-lg-8">When a service ticket is opened via <b>Website</b>, send an email to the customer as defined by the hierarchy.</div>
					         	<div class="col-lg-2">
									<%
									Response.Write("<input type='checkbox' class='check' id='chkOpenViaWeb_SendToCust' name='chkOpenViaWeb_SendToCust'")
									If OpenViaWeb_SendToCust = vbTrue Then Response.Write(" checked ")
									Response.Write(">")
									%>
								</div>
			         		</div>
		         		<% Else %>
		         			<input type="hidden" name="chkOpenViaWeb_SendToCust" id="chkOpenViaWeb_SendToCust">
		         		<% End If %>
			         	<!-- eof row with data !-->
				         	
			         	<!-- row with data !-->
		         		<div class="row row-margin row-data">
		         			<% If MUV_READ("BackendSystem") = "Metroplex" Then %>
				       	   		<div class="col-lg-8">When a service ticket is opened via <b>MDS Insight</b>, send an email to the customer as defined by the hierarchy.</div>
				       	   	<% Else %>
				       	   		<div class="col-lg-8">When a service ticket is opened via <b>MDS Insight</b>, send an email to the customer.</div>
				       	   	<% End If %>
		        		 	<div class="col-lg-2">
								<%
								Response.Write("<input type='checkbox' class='check' id='chkOpen_ViaInsight_SendToCust' name='chkOpen_ViaInsight_SendToCust'")
								If Open_ViaInsight_SendToCust = vbTrue Then Response.Write(" checked ")
								Response.Write(">")
								%>
			    			</div>
		         		</div>
			         	<!-- eof row with data !-->
			         	
			         	<!-- row with data !-->
		         		<div class="row row-margin row-data">
		         		
		         			<% If MUV_READ("BackendSystem") = "Metroplex" Then %>
				       	   		<div class="col-lg-8">When a service ticket is opened via the <b>Field Service Webapp</b>, send an email to the customer as defined by the hierarchy.</div>
				       	   	<% Else %>
				       	   		<div class="col-lg-8">When a service ticket is opened via the <b>Field Service Webapp</b>, send an email to the customer.</div>
				       	   	<% End If %>
		         		
		        		 	<div class="col-lg-2">
			    				<%
								Response.Write("<input type='checkbox' class='check' id='chkOpen_ViaFServ_SendToCust' name='chkOpen_ViaFServ_SendToCust'")
								If Open_ViaFServ_SendToCust = vbTrue Then Response.Write(" checked ")
								Response.Write(">")
								%>
			    			</div>
		         		</div>
			         	<!-- eof row with data !-->
				         	
			         	<!-- row with data !-->
		         		<div class="row row-margin row-data">

		         			<% If MUV_READ("BackendSystem") = "Metroplex" Then %>
				       	   		<div class="col-lg-8">When a service ticket is opened via <b><%= MUV_READ("BackendSystem") %></b>, send an email to the customer as defined by the hierarchy.</div>
				       	   	<% Else %>
				       	   		<div class="col-lg-8">When a service ticket is opened via <b><%= MUV_READ("BackendSystem") %></b>, send an email to the customer.</div>
				       	   	<% End If %>
				       	   	
		        		 	<div class="col-lg-2">
   			    				<%
								Response.Write("<input type='checkbox' class='check' id='chkOpen_ViaMDS_SendToCust' name='chkOpen_ViaMDS_SendToCust'")
								If Open_ViaMDS_SendToCust = vbTrue Then Response.Write(" checked ")
								Response.Write(">")
								%>
							</div>
		         		</div>
			         	<!-- eof row with data !-->
				         	
		         	</div>
		         	<!-- eof right col !-->
	         
          </div>
          <!-- eof main row !-->
          
          
          
          <Br>  	 
          <a href="#" onClick="window.location.reload();"><button type="button" class="btn btn-default">Cancel</button></a> 
            
          <button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>     

        </div>

<!-- OPEN tab !-->


<!--CLOSE Tab !-->
				
<!--  Service Tab !-->
<div role="tabpanel" class="tab-pane fade" id="close"> 

  
      <!-- main row !-->
      <div class="row row-margin">
     	
         	<!-- left col !-->
         	<% If MUV_READ("BackendSystem") = "Metroplex" Then %>
	         	<div class="col-lg-3">
		         	<strong><u>Email Hierarchy</u></strong><br>
		         	1. Service email<br>
		         	2. Order email<br>
		         	3. General email<br>
		         	4. Notify Rep<br>
		         	5. Notify IT Dept<br>
	         	</div>
         	<% End If %>
         	<!-- eof left col
         	
         	<!-- right col !-->	         	
         	<div class="col-lg-6">
	         	
	         	<!-- row with data !-->
         		<div class="row row-margin row-data">

         			<% If MUV_READ("BackendSystem") = "Metroplex" Then %>
		       	   		<div class="col-lg-8">When a service ticket is closed via <b>MDS Insight</b>, send an email to the customer as defined by the hierarchy.</div>
		       	   	<% Else %>
		       	   		<div class="col-lg-8">When a service ticket is closed via <b>MDS Insight</b>, send an email to the customer.</div>
		       	   	<% End If %>
		       	   	
        		 	<div class="col-lg-2">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkClose_ViaInsight_SendToCust' name='chkClose_ViaInsight_SendToCust'")
					If Close_ViaInsight_SendToCust = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
	         	<!-- eof row with data !-->
	         	
	         	<!-- row with data !-->
         		<div class="row row-margin row-data">

         			<% If MUV_READ("BackendSystem") = "Metroplex" Then %>
		       	   		<div class="col-lg-8">When a service ticket is closed via the <b>Field Service Webapp</b>, send an email to the customer as defined by the hierarchy.</div>
		       	   	<% Else %>
		       	   		<div class="col-lg-8">When a service ticket is closed via the <b>Field Service Webapp</b>, send an email to the customer.</div>
		       	   	<% End If %>
		       	   	
        		 	<div class="col-lg-2">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkClose_ViaFServ_SendToCust' name='chkClose_ViaFServ_SendToCust'")
					If Close_ViaFServ_SendToCust = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
	         	<!-- eof row with data !-->
	         	
	          	<!-- row with data 
	          	
	          	THE REASON THIS IS COMMENTED OUT IS BECUASE SERVICE TICKETS DON'T REALLY EVER GET CLOSED ON THE BACKEND SYSTEM
	          	
	          	
	          	
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-8">When a service ticket is closed via <b><%=GetTerm("Backend")%></b>, send an email to the customer as defined by the hierarchy.</div>
        		 	<div class="col-lg-2">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkClose_ViaMDS_SendToCust' name='chkClose_ViaMDS_SendToCust'")
					If Close_ViaMDS_SendToCust = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
	         	<!-- eof row with data !-->

		         	
         	</div>
         	<!-- eof right col !-->
     
  </div>
  <!-- eof main row !-->
  
  <Br>  	 
  <a href="#" onClick="window.location.reload();"><button type="button" class="btn btn-default">Cancel</button></a> 
    
  <button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>        

</div>
<!-- eof CLOSE Tab !-->

        
<!-- CANCEL tab !-->
				
<!--  Service Tab !-->
<div role="tabpanel" class="tab-pane fade" id="cancel"> 

  
      <!-- main row !-->
      <div class="row row-margin">
     	
         	<!-- left col !-->
         	<% If MUV_READ("BackendSystem") = "Metroplex" Then %>
	         	<div class="col-lg-3">
		         	<strong><u>Email Hierarchy</u></strong><br>
		         	1. Service email<br>
		         	2. Order email<br>
		         	3. General email<br>
		         	4. Notify Rep<br>
		         	5. Notify IT Dept<br>	         	
	         	</div>
         	<% End If %>
         	<!-- eof left col
         	
         	<!-- right col !-->	         	
         	<div class="col-lg-6">
	         	
	         	<!-- row with data !-->
         		<div class="row row-margin row-data">

         			<% If MUV_READ("BackendSystem") = "Metroplex" Then %>
		       	   		<div class="col-lg-8">When a service ticket is cancelled via <b>MDS Insight</b>, send an email to the customer as defined by the hierarchy.</div>
		       	   	<% Else %>
		       	   		<div class="col-lg-8">When a service ticket is cancelled via <b>MDS Insight</b>, send an email to the customer.</div>
		       	   	<% End If %>
		       	   	
        		 	<div class="col-lg-2">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkCancel_ViaInsight_SendToCust' name='chkCancel_ViaInsight_SendToCust'")
					If Cancel_ViaInsight_SendToCust = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
	         	<!-- eof row with data !-->
	         	
	         	<!-- row with data !-->
         		<div class="row row-margin row-data">

         			<% If MUV_READ("BackendSystem") = "Metroplex" Then %>
		       	   		<div class="col-lg-8">When a service ticket is cancelled via the <b>Field Service Webapp</b>, send an email to the customer as defined by the hierarchy.</div>
		       	   	<% Else %>
		       	   		<div class="col-lg-8">When a service ticket is cancelled via the <b>Field Service Webapp</b>, send an email to the customer.</div>
		       	   	<% End If %>
		       	   	
        		 	<div class="col-lg-2">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkCancel_ViaFServ_SendToCust' name='chkCancel_ViaFServ_SendToCust'")
					If Cancel_ViaFServ_SendToCust = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
	         	<!-- eof row with data !-->
	         	
		         	
         	</div>
         	<!-- eof right col !-->
     
  </div>
  <!-- eof main row !-->
  
  <Br>  	 
  <a href="#" onClick="window.location.reload();"><button type="button" class="btn btn-default">Cancel</button></a> 
    
  <button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>        

</div>
<!-- eof CANCEL tab !-->





<!-- Realtime Alerts tab !-->
<div role="tabpanel" class="tab-pane fade" id="realtime"> 

  
      <!-- main row !-->
      <div class="row row-margin">
     	
         	<!-- left col !-->
         	<div class="col-lg-3">
	         	<strong>Elapsed Time Threshold</strong><br>
         	</div>
         	<!-- eof left col
         	
         	<!-- right col !-->	         	
         	<div class="col-lg-6">
	         	
	         	
	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Check alerts only during buisness hours (M-F)</div>
        		 	<div class="col-lg-3">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkAlertsDuringBizHoursOnly' name='chkAlertsDuringBizHoursOnly'")
					If AlertsDuringBizHoursOnly  = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->

	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Turn on realtime threshold alerts </div>
        		 	<div class="col-lg-3">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkRealtimeAlertsOn' name='chkRealtimeAlertsOn'")
					If RealtimeAlertsOn = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->
	         	
	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Send alert when a ticket has been open for </div>
	        		 	<div class="col-lg-3">
					         	<select name="selSendAlertHours">
					         	<%
					         	If SendAlertHours = .25 Then Response.Write("<option selected value='.25'>1/4</option>") Else Response.Write("<option value='.25'>1/4</option>")
					         	If SendAlertHours = .5 Then Response.Write("<option selected value='.5'>1/2</option>") Else Response.Write("<option value='.5'>1/2</option>")
					         	If SendAlertHours = .75 Then Response.Write("<option selected value='.75'>3/4</option>") Else Response.Write("<option value='.75'>3/4</option>")
								For x = 1 to 100
					         			If x = cint(SendAlertHours) Then
							               	Response.Write("<option selected>" & x & "</option>")
					         			Else
	   							         	Response.Write("<option>" & x & "</option>")
							         	End If
					         	Next %>
				         	</select>&nbsp;Hours
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->


	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Send alert email to all service managers </div>
        		 	<div class="col-lg-3">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkSendAlertToServiceManagers' name='chkSendAlertToServiceManagers'")
					If SendAlertToServiceManagers = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->

	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Don't send alerts for dispatched tickets </div>
        		 	<div class="col-lg-3">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkSendAlertsSkipDispatched' name='chkSendAlertsSkipDispatched'")
					If SendAlertsSkipDispatched = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->

	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Send alert email to the following email addresses </div>
        		 	<div class="col-lg-5">
						<textarea name="txtSendAlertToAdditionalEmails" id="txtSendAlertToAdditionalEmails" rows="4"  class='form-control'><%= SendAlertToAdditionalEmails %></textarea>
						<strong>Seperate multiple email addresses with a semicolon.</strong>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->
	         	


	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Turn on realtime threshold escalation alerts </div>
        		 	<div class="col-lg-3">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkEscalationAlertsOn' name='chkEscalationAlertsOn'")
					If EscalationAlertsOn = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->

	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Send escalation alert when a ticket has been open for </div>
	        		 	<div class="col-lg-3">
					         	<select name="selEscalationAlertHours">
					         	<%
					         	If EscalationAlertHours = .25 Then Response.Write("<option selected value='.25'>1/4</option>") Else Response.Write("<option value='.25'>1/4</option>")
					         	If EscalationAlertHours = .5 Then Response.Write("<option selected value='.5'>1/2</option>") Else Response.Write("<option value='.5'>1/2</option>")
					         	If EscalationAlertHours = .75 Then Response.Write("<option selected value='.75'>3/4</option>") Else Response.Write("<option value='.75'>3/4</option>")
								For x = 1 to 100
					         			If x = cint(EscalationAlertHours) Then
							               	Response.Write("<option selected>" & x & "</option>")
					         			Else
	   							         	Response.Write("<option>" & x & "</option>")
							         	End If
					         	Next %>
				         	</select>&nbsp;Hours
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->

	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Don't send escalation alerts for dispatched tickets </div>
        		 	<div class="col-lg-3">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkEscalationAlertsSkipDispatched' name='chkEscalationAlertsSkipDispatched'")
					If EscalationAlertsSkipDispatched = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->

	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Send escalation alert email to the following email addresses </div>
        		 	<div class="col-lg-5">
						<textarea name="txtEscalationAlertToEmails" id="txtEscalationAlertToEmails" rows="4"  class="form-control"><%= EscalationAlertToEmails %></textarea>
						<strong>Seperate multiple email addresses with a semicolon.</strong>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->
<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''%>		         	


	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Turn on realtime hold alerts </div>
        		 	<div class="col-lg-3">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkHoldAlertsOn' name='chkHoldAlertsOn'")
					If HoldAlertsOn = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->
	         	
	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Send alert when a ticket has been on hold for </div>
	        		 	<div class="col-lg-3">
					         	<select name="selSendHoldAlertHours">
					         	<%
					         	If SendHoldAlertHours = .25 Then Response.Write("<option selected value='.25'>1/4</option>") Else Response.Write("<option value='.25'>1/4</option>")
					         	If SendHoldAlertHours = .5 Then Response.Write("<option selected value='.5'>1/2</option>") Else Response.Write("<option value='.5'>1/2</option>")
					         	If SendHoldAlertHours = .75 Then Response.Write("<option selected value='.75'>3/4</option>") Else Response.Write("<option value='.75'>3/4</option>")
					         	For x = 1 to 100
					         			If x = cint(SendHoldAlertHours) Then
							               	Response.Write("<option selected>" & x & "</option>")
					         			Else
	   							         	Response.Write("<option>" & x & "</option>")
							         	End If
					         	Next %>
				         	</select>&nbsp;Hours
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->


	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Send hold alert email to all finance managers </div>
        		 	<div class="col-lg-3">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkSendHoldAlertToFinanceManagers' name='chkSendHoldAlertToFinanceManagers'")
					If SendHoldAlertToFinanceManagers = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->


	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Send hold alert email to the following email addresses </div>
        		 	<div class="col-lg-5">
						<textarea name="txtSendHoldAlertToAdditionalEmails" id="txtSendHoldAlertToAdditionalEmails" rows="4"  class='form-control'><%= SendHoldAlertToAdditionalEmails %></textarea>
						<strong>Seperate multiple email addresses with a semicolon.</strong>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->
	         	


	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Turn on realtime hold escalation alerts </div>
        		 	<div class="col-lg-3">
    				<%
					Response.Write("<input type='checkbox' class='check' id='chkHoldEscalationAlertsOn' name='chkHoldEscalationAlertsOn'")
					If HoldEscalationAlertsOn = vbTrue Then Response.Write(" checked ")
					Response.Write(">")
					%>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->

	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Send hold escalation alert when a ticket has been on hold for </div>
	        		 	<div class="col-lg-3">
					         	<select name="selHoldEscalationAlertHours">
					         	<%
					         	If HoldEscalationAlertHours = .25 Then Response.Write("<option selected value='.25'>1/4</option>") Else Response.Write("<option value='.25'>1/4</option>")
					         	If HoldEscalationAlertHours = .5 Then Response.Write("<option selected value='.5'>1/2</option>") Else Response.Write("<option value='.5'>1/2</option>")
					         	If HoldEscalationAlertHours = .75 Then Response.Write("<option selected value='.75'>3/4</option>") Else Response.Write("<option value='.75'>3/4</option>")
								For x = 1 to 100
					         			If x = cint(HoldEscalationAlertHours) Then
							               	Response.Write("<option selected>" & x & "</option>")
					         			Else
	   							         	Response.Write("<option>" & x & "</option>")
							         	End If
					         	Next %>
				         	</select>&nbsp;Hours
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->


	         	<!-- row with data !-->
         		<div class="row row-margin row-data">
		       	   	<div class="col-lg-6">Send hold escalation alert email to the following email addresses </div>
        		 	<div class="col-lg-5">
						<textarea name="txtHoldEscalationAlertToEmails" id="txtHoldEscalationAlertToEmails" rows="4"  class="form-control"><%= HoldEscalationAlertToEmails %></textarea>
						<strong>Seperate multiple email addresses with a semicolon.</strong>
	    			</div>
         		</div>
         		<br>
	         	<!-- eof row with data !-->
		         	
		         	
<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''%>		         	
         	</div>
         	<!-- eof right col !-->
     
  </div>
  <!-- eof main row !-->
  
  
  <Br>  	 
  <a href="#" onClick="window.location.reload();"><button type="button" class="btn btn-default">Cancel</button></a> 
    
  <button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>        

</div>
<!-- eof Realtime Alerts tab !-->



        </div>
        </div>
               </div>
                      
 </form>
         
</div>
</div>
<!-- eof row !-->    
<%
Function HandleAuditTrail()

	'*******************************************
	' This code to write to the audit trail file
	'*******************************************
	SQL = "SELECT * FROM Settings_EmailService"
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	If not rs.EOF Then

		UseSimpleEmailFormat4Compare =  UseSimpleEmailFormat 
		If UseSimpleEmailFormat4Compare <> rs("Internal_Simple") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("Internal_Simple")
			VerbiageForReport = Replace(VerbiageForReport,"1","On")
			VerbiageForReport = Replace(VerbiageForReport,"0","Off")
			If UseSimpleEmailFormat = vbTrue then UseSimpleEmailFormat4Compare = "True" else UseSimpleEmailFormat4Compare = "False"
			CreateAuditLogEntry "Service Ticket Email Settings Change", "Service Ticket Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: When sending an internal email, use simple format was changed from " & VerbiageForReport & " to " & UseSimpleEmailFormat4Compare 
		End If
	
		OpenViaWeb_SendToWebUser4Compare  =  OpenViaWeb_SendToWebUser 
		If OpenViaWeb_SendToWebUser4Compare <> rs("OpenViaWeb_SendToWebUser") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("OpenViaWeb_SendToWebUser")
			VerbiageForReport = Replace(VerbiageForReport,"1","On")
			VerbiageForReport = Replace(VerbiageForReport,"0","Off")
			If OpenViaWeb_SendToWebUser = vbTrue then OpenViaWeb_SendToWebUser4Compare = "True" else OpenViaWeb_SendToWebUser4Compare = "False"
			CreateAuditLogEntry "Service Ticket Email Settings Change", "Service Ticket Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: When a service ticket is opened via Website, send an email to email address of the logged in user from " & VerbiageForReport & " to " & OpenViaWeb_SendToWebUser4Compare
		End If
		OpenViaWeb_SendToCust4Compare  =  OpenViaWeb_SendToCust
		If OpenViaWeb_SendToCust4Compare <> rs("OpenViaWeb_SendToCust") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("OpenViaWeb_SendToCust")
			VerbiageForReport = Replace(VerbiageForReport,"1","On")
			VerbiageForReport = Replace(VerbiageForReport,"0","Off")
			If OpenViaWeb_SendToCust = vbTrue then OpenViaWeb_SendToCust4Compare = "True" else OpenViaWeb_SendToCust4Compare = "False"
			CreateAuditLogEntry "Service Ticket Email Settings Change", "Service Ticket Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: When a service ticket is opened via Website, send an email to the customer as defined by the hierarchy from " & VerbiageForReport & " to " & OpenViaWeb_SendToCust4Compare
		End If
		Open_ViaInsight_SendToCust4Compare  =  Open_ViaInsight_SendToCust
		If Open_ViaInsight_SendToCust4Compare <> rs("Open_ViaInsight_SendToCust") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("Open_ViaInsight_SendToCust")
			VerbiageForReport = Replace(VerbiageForReport,"1","On")
			VerbiageForReport = Replace(VerbiageForReport,"0","Off")
			If Open_ViaInsight_SendToCust = vbTrue then Open_ViaInsight_SendToCust4Compare = "True" else Open_ViaInsight_SendToCust4Compare = "False"
			CreateAuditLogEntry "Service Ticket Email Settings Change", "Service Ticket Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: When a service ticket is opened via MDS Insight, send an email to the customer as defined by the hierarchy from " & VerbiageForReport & " to " & Open_ViaInsight_SendToCust4Compare
		End If
		Open_ViaFServ_SendToCust4Compare  =  Open_ViaFServ_SendToCust
		If Open_ViaFServ_SendToCust4Compare <> rs("Open_ViaFServ_SendToCust") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("Open_ViaFServ_SendToCust")
			VerbiageForReport = Replace(VerbiageForReport,"1","On")
			VerbiageForReport = Replace(VerbiageForReport,"0","Off")
			If Open_ViaFServ_SendToCust = vbTrue then Open_ViaFServ_SendToCust4Compare = "True" else Open_ViaFServ_SendToCust4Compare = "False"
			CreateAuditLogEntry "Service Ticket Email Settings Change", "Service Ticket Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: When a service ticket is opened via the Field Service Webapp, send an email to the customer as defined by the hierarchy from " & VerbiageForReport & " to " & Open_ViaFServ_SendToCust4Compare
		End If
		Open_ViaMDS_SendToCust4Compare  =  Open_ViaMDS_SendToCust
		If Open_ViaMDS_SendToCust4Compare <> rs("Open_ViaMDS_SendToCust") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("Open_ViaMDS_SendToCust")
			VerbiageForReport = Replace(VerbiageForReport,"1","On")
			VerbiageForReport = Replace(VerbiageForReport,"0","Off")
			If Open_ViaMDS_SendToCust = vbTrue then Open_ViaMDS_SendToCust4Compare = "True" else Open_ViaMDS_SendToCust4Compare = "False"
			CreateAuditLogEntry "Service Ticket Email Settings Change", "Service Ticket Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: When a service ticket is opened via " & GetTerm("Backend") & ", send an email to the customer as defined by the hierarchy from " & VerbiageForReport & " to " & Open_ViaMDS_SendToCust4Compare
		End If
		Close_ViaInsight_SendToCust4Compare  =  Close_ViaInsight_SendToCust
		If Close_ViaInsight_SendToCust4Compare <> rs("Close_ViaInsight_SendToCust") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("Close_ViaInsight_SendToCust")
			VerbiageForReport = Replace(VerbiageForReport,"1","On")
			VerbiageForReport = Replace(VerbiageForReport,"0","Off")
			If Close_ViaInsight_SendToCust = vbTrue then Close_ViaInsight_SendToCust4Compare = "True" else Close_ViaInsight_SendToCust4Compare = "False"
			CreateAuditLogEntry "Service Ticket Email Settings Change", "Service Ticket Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: When a service ticket is closed via MDS Insight, send an email to the customer as defined by the hierarchy from " & VerbiageForReport & " to " & Close_ViaInsight_SendToCust4Compare
		End If
		Close_ViaFServ_SendToCust4Compare  =  Close_ViaFServ_SendToCust
		If Close_ViaFServ_SendToCust4Compare <> rs("Close_ViaFServ_SendToCust") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("Close_ViaFServ_SendToCust")
			VerbiageForReport = Replace(VerbiageForReport,"1","On")
			VerbiageForReport = Replace(VerbiageForReport,"0","Off")
			If Close_ViaFServ_SendToCust = vbTrue then Close_ViaFServ_SendToCust4Compare = "True" else Close_ViaFServ_SendToCust4Compare = "False"
			CreateAuditLogEntry "Service Ticket Email Settings Change", "Service Ticket Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: When a service ticket is closed via the Field Service Webapp, send an email to the customer as defined by the hierarchy from " & VerbiageForReport & " to " & Close_ViaFServ_SendToCust4Compare
		End If
		Close_ViaMDS_SendToCust4Compare  =  Close_ViaMDS_SendToCust
		Cancel_ViaInsight_SendToCust4Compare  =  Cancel_ViaInsight_SendToCust
		If Cancel_ViaInsight_SendToCust4Compare <> rs("Cancel_ViaInsight_SendToCust") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("Cancel_ViaInsight_SendToCust")
			VerbiageForReport = Replace(VerbiageForReport,"1","On")
			VerbiageForReport = Replace(VerbiageForReport,"0","Off")
			If Cancel_ViaInsight_SendToCust = vbTrue then Cancel_ViaInsight_SendToCust4Compare = "True" else Cancel_ViaInsight_SendToCust4Compare = "False"
			CreateAuditLogEntry "Service Ticket Email Settings Change", "Service Ticket Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: When a service ticket is cancelled via MDS Insight, send an email to the customer as defined by the hierarchy from " & VerbiageForReport & " to " & Cancel_ViaInsight_SendToCust4Compare
		End If
		Cancel_ViaFServ_SendToCust4Compare  =  Cancel_ViaFServ_SendToCust
		If Cancel_ViaFServ_SendToCust4Compare <> rs("Cancel_ViaFServ_SendToCust") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("Cancel_ViaFServ_SendToCust")
			VerbiageForReport = Replace(VerbiageForReport,"1","On")
			VerbiageForReport = Replace(VerbiageForReport,"0","Off")
			If Cancel_ViaFServ_SendToCust = vbTrue then Cancel_ViaFServ_SendToCust4Compare = "True" else Cancel_ViaFServ_SendToCust4Compare = "False" 
			CreateAuditLogEntry "Service Ticket Email Settings Change", "Service Ticket Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: When a service ticket is cancelled via the Field Service Webapp, send an email to the customer as defined by the hierarchy from " & VerbiageForReport & " to " & Cancel_ViaFServ_SendToCust4Compare
		End If
		RealtimeAlertsOn4Compare  =  RealtimeAlertsOn
		If RealtimeAlertsOn4Compare <> rs("RealtimeAlertsOn") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("RealtimeAlertsOn")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If RealtimeAlertsOn = vbTrue then RealtimeAlertsOn4Compare = "True" else RealtimeAlertsOn4Compare = "False" 
			CreateAuditLogEntry "Service Ticket Realtime Alert Settings Change", "Service Ticket Realtime Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Turn on realtime alerts from " & VerbiageForReport & " to " & RealtimeAlertsOn4Compare  
		End If
		AlertsDuringBizHoursOnly4Compare  =  AlertsDuringBizHoursOnly
		If AlertsDuringBizHoursOnly4Compare <> rs("AlertsDuringBizHoursOnly") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("AlertsDuringBizHoursOnly")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If AlertsDuringBizHoursOnly = vbTrue then AlertsDuringBizHoursOnly4Compare = "True" else AlertsDuringBizHoursOnly4Compare = "False" 
			CreateAuditLogEntry "Service Ticket Realtime Alert Settings Change", "Service Ticket Realtime Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Check only during business hours from " & VerbiageForReport & " to " & AlertsDuringBizHoursOnlyOn4Compare  
		End If
		SendAlertToServiceManagers4Compare  =  SendAlertToServiceManagers
		If SendAlertToServiceManagers4Compare <> rs("SendAlertToServiceManagers") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("SendAlertToServiceManagers")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If SendAlertToServiceManagers = vbTrue then SendAlertToServiceManagers4Compare = "True" else SendAlertToServiceManagers4Compare = "False" 
			CreateAuditLogEntry "Service Ticket Realtime Alert Settings Change", "Service Ticket Realtime Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Send alert email to all service managers from " & VerbiageForReport & " to " & SendAlertToServiceManagers 
		End If
		SendAlertToAdditionalEmails4Compare  =  SendAlertToAdditionalEmails
		If SendAlertToAdditionalEmails4Compare <> rs("SendAlertToAdditionalEmails") Then
			VerbiageForReport = rs("SendAlertToAdditionalEmails")
			CreateAuditLogEntry "Service Ticket Realtime Alert Settings Change", "Service Ticket Realtime Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Send alert email to additional emails from " & VerbiageForReport & " to " & SendAlertToAdditionalEmails4Compare 
		End If
		SendAlertHours4Compare  =  SendAlertHours
		If SendAlertHours4Compare <> rs("SendAlertHours") Then
			VerbiageForReport = rs("SendAlertHours")
			CreateAuditLogEntry "Service Ticket Realtime Alert Settings Change", "Service Ticket Realtime Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Send alert after ticket reaches x hours from " & VerbiageForReport & " to " & SendAlertHours4Compare  
		End If
		SendAlertsSkipDispatched4Compare = SendAlertsSkipDispatched
		If SendAlertsSkipDispatched4Compare <> rs("SendAlertsSkipDispatched") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("SendAlertsSkipDispatched")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If SendAlertsSkipDispatched = vbTrue then SendAlertsSkipDispatched4Compare = "True" else SendAlertsSkipDispatched4Compare = "False" 
			CreateAuditLogEntry "Service Ticket Realtime Alert Settings Change", "Service Ticket Realtime Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Alerts skip dispatched tickets from " & VerbiageForReport & " to " & SendAlertsSkipDispatched4Compare 
		End If
		EscalationAlertsOn4Compare = EscalationAlertsOn
		If EscalationAlertsOn4Compare <> rs("EscalationAlertsOn") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("EscalationAlertsOn")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If EscalationAlertsOn = vbTrue then EscalationAlertsOn4Compare = "True" else EscalationAlertsOn4Compare = "False" 
			CreateAuditLogEntry "Service Ticket Realtime Alert Settings Change", "Service Ticket Realtime Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Turn on realtime escalation alerts from " & VerbiageForReport & " to " & EscalationAlertsOn4Compare 
		End If
		EscalationAlertsSkipDispatched4Compare = EscalationAlertsSkipDispatched
		If EscalationAlertsSkipDispatched4Compare <> rs("EscalationAlertsSkipDispatched") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("EscalationAlertsSkipDispatched")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If EscalationAlertsSkipDispatched = vbTrue then EscalationAlertsSkipDispatched4Compare = "True" else EscalationAlertsSkipDispatched4Compare = "False" 
			CreateAuditLogEntry "Service Ticket Realtime Alert Settings Change", "Service Ticket Realtime Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Escalation alerts skip dispatched tickets from " & VerbiageForReport & " to " & EscalationAlertsSkipDispatched4Compare 
		End If
		EscalationAlertHours4Compare  =  EscalationAlertHours
		If EscalationAlertHours4Compare <> rs("EscalationAlertHours") Then
			VerbiageForReport = rs("EscalationAlertHours")
			CreateAuditLogEntry "Service Ticket Realtime Alert Settings Change", "Service Ticket Realtime Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Send escalation alert after ticket reaches x hours from " & VerbiageForReport & " to " & EscalationAlertHours4Compare  
		End If
		EscalationAlertToEmails4Compare  =  EscalationAlertToEmails
		If EscalationAlertToEmails4Compare <> rs("EscalationAlertToEmails") Then
			VerbiageForReport = rs("EscalationAlertToEmails")
			CreateAuditLogEntry "Service Ticket Realtime Alert Settings Change", "Service Ticket Realtime Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Send escalation alert email addresses from " & VerbiageForReport & " to " & EscalationAlertToEmails4Compare  
		End If
		HoldAlertsOn4Compare  =  HoldAlertsOn
		If HoldAlertsOn4Compare <> rs("HoldAlertsOn") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("HoldAlertsOn")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If HoldAlertsOn = vbTrue then HoldAlertsOn4Compare = "True" else HoldAlertsOn4Compare = "False" 
			CreateAuditLogEntry "Service Ticket Hold Alert Settings Change", "Service Ticket Hold Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Turn on Hold alerts from " & VerbiageForReport & " to " & HoldAlertsOn4Compare  
		End If
		
		CompletedPMCallEmailOn4Compare  =  CompletedPMCallEmailOn
		If CompletedPMCallEmailOn4Compare  <> rs("CompletedPMCallEmailOn") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("CompletedPMCallEmailOn")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If CompletedPMCallEmailOn = vbTrue then CompletedPMCallEmailOn4Compare  = "True" else CompletedPMCallEmailOn4Compare  = "False" 
			CreateAuditLogEntry "PM Call Email Settings Change", "PM Call Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " Completed PMCall triggers email.  from " & VerbiageForReport & " to " & CompletedPMCallEmailOn4Compare  
		End If
		DoNotSendClientCompletedPMCall4Compare  =  DoNotSendClientCompletedPMCall
		If DoNotSendClientCompletedPMCall4Compare  <> rs("DoNotSendClientCompletedPMCall") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("DoNotSendClientCompletedPMCall")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If DoNotSendClientCompletedPMCall= vbTrue then DoNotSendClientCompletedPMCall4Compare  = "True" else DoNotSendClientCompletedPMCall4Compare  = "False" 
			CreateAuditLogEntry "PM Call Email Settings Change", "PM Call Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " Do not send completed PM Call email to client.  from " & VerbiageForReport & " to " & DoNotSendClientCompletedPMCall4Compare  
		End If
		PMCallPDFIncludeServiceNotes4Compare  =  PMCallPDFIncludeServiceNotes
		If PMCallPDFIncludeServiceNotes4Compare  <> rs("PMCallPDFIncludeServiceNotes") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("PMCallPDFIncludeServiceNotes")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If PMCallPDFIncludeServiceNotes = vbTrue then PMCallPDFIncludeServiceNotes4Compare  = "True" else PMCallPDFIncludeServiceNotes4Compare  = "False" 
			CreateAuditLogEntry "PM Call Email Settings Change", "PM Call Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed include service notes in pm call email and .pdf  from " & VerbiageForReport & " to " & PMCallPDFIncludeServiceNotes4Compare  
		End If
		SendHoldAlertToFinanceManagers4Compare  =  SendHoldAlertToFinanceManagers
		If SendHoldAlertToFinanceManagers4Compare <> rs("SendHoldAlertToFinanceManagers") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("SendHoldAlertToFinanceManagers")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If SendHoldAlertToFinanceManagers = vbTrue then SendHoldAlertToFinanceManagers4Compare = "True" else SendHoldAlertToFinanceManagers4Compare = "False" 
			CreateAuditLogEntry "Service Ticket Hold Alert Settings Change", "Service Ticket Hold Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Send hold alert email to all finance managers from " & VerbiageForReport & " to " & SendHoldAlertToFinanceManagers 
		End If
		SendHoldAlertToAdditionalEmails4Compare  =  SendHoldAlertToAdditionalEmails
		If SendHoldAlertToAdditionalEmails4Compare <> rs("SendHoldAlertToAdditionalEmails") Then
			VerbiageForReport = rs("SendHoldAlertToAdditionalEmails")
			CreateAuditLogEntry "Service Ticket Hold Alert Settings Change", "Service Ticket Hold Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Send hold alert email to additional emails from " & VerbiageForReport & " to " & SendHoldAlertToAdditionalEmails4Compare 
		End If
		
		SendCompletedPMCallTo4Compare  =  SendCompletedPMCallTo
		If SendCompletedPMCallTo4Compare  <> rs("SendCompletedPMCallTo") Then
			VerbiageForReport = rs("SendCompletedPMCallTo")
			CreateAuditLogEntry "PM Call Complete Email Settings Change", "PM Call Complete Email Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Send PM Call completed emails to: from " & VerbiageForReport & " to " & SendCompletedPMCallTo4Compare  
		End If


		SendHoldAlertHours4Compare  =  SendHoldAlertHours
		If SendHoldAlertHours4Compare <> rs("SendHoldAlertHours") Then
			VerbiageForReport = rs("SendHoldAlertHours")
			CreateAuditLogEntry "Service Ticket Hold Alert Settings Change", "Service Ticket Hold Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Send hold alert after ticket reaches x hours from " & VerbiageForReport & " to " & SendHoldAlertHours4Compare  
		End If
		SendAlertsSkipDispatched4Compare = SendAlertsSkipDispatched
		HoldEscalationAlertsOn4Compare = HoldEscalationAlertsOn
		If HoldEscalationAlertsOn4Compare <> rs("HoldEscalationAlertsOn") Then
			' Just make it say On/Off instead of True/False
			VerbiageForReport = rs("HoldEscalationAlertsOn")
			VerbiageForReport = Replace(VerbiageForReport,-1,"True")
			VerbiageForReport = Replace(VerbiageForReport,0,"False")
			If HoldEscalationAlertsOn = vbTrue then HoldEscalationAlertsOn4Compare = "True" else HoldEscalationAlertsOn4Compare = "False" 
			CreateAuditLogEntry "Service Ticket Hold Alert Settings Change", "Service Ticket Hold Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Turn on Hold escalation alerts from " & VerbiageForReport & " to " & HoldEscalationAlertsOn4Compare 
		End If
		HoldEscalationAlertHours4Compare  =  HoldEscalationAlertHours
		If HoldEscalationAlertHours4Compare <> rs("HoldEscalationAlertHours") Then
			VerbiageForReport = rs("HoldEscalationAlertHours")
			CreateAuditLogEntry "Service Ticket Hold Alert Settings Change", "Service Ticket Hold Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Send hold escalation alert after ticket reaches x hours from " & VerbiageForReport & " to " & HoldEscalationAlertHours4Compare  
		End If
		HoldEscalationAlertToEmails4Compare  =  HoldEscalationAlertToEmails
		If HoldEscalationAlertToEmails4Compare <> rs("HoldEscalationAlertToEmails") Then
			VerbiageForReport = rs("HoldEscalationAlertToEmails")
			CreateAuditLogEntry "Service Ticket Hold Alert Settings Change", "Service Ticket Hold Alert Settings Change", "Major", 1, MUV_Read("DisplayName") & " changed: Send Hold escalation alert email addresses from " & VerbiageForReport & " to " & HoldEscalationAlertToEmails4Compare  
		End If
	End If

	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

	HandleAuditTrail = 0
End Function
%>
</script>

<!--#include file="../../inc/footer-main.asp"-->