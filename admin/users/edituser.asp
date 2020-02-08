<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/mail.asp"-->
<% 
UserNo = Request.QueryString("uno") 
ActiveTab = Request.QueryString("tab")
If UserNo = "" Then Response.Redirect(BaseURL)

If filterChangeModuleOn() Then filterChangeFlag = 1 Else filterChangeFlag = 0
filterChangeFlag = 1

Response.CacheControl = "no-cache, no-store, must-revalidate" 

If Session("QuickLoginEmailSentTo") <> "" Then 
		msg = "An email has been sent to " & Session("QuickLoginEmailSentTo") %>
		<SCRIPT LANGUAGE="JavaScript">
			swal('<%= msg %>');
		</SCRIPT>     
		<% 
		Session("QuickLoginEmailSentTo") = ""
End If

If Session("QuickLoginTextSentTo") <> "" Then 
		msg = "A text has been sent to " & Session("QuickLoginTextSentTo") %>
		<SCRIPT LANGUAGE="JavaScript">
			swal('<%= msg %>');
		</SCRIPT>     
		<%
		Session("QuickLoginTextSentTo") = ""
End If %>


<script type="text/javascript">
	function onTypeChanged() {
	
		$("#pnlRouteNumber").hide();
		$("#pnlRoutes").hide();
		$("#pnlDriverAdditional").hide();
		$("#pnlProspecting").hide();
		$("#pnlInventoryControl").hide();
		$("#pnlOrderAPI").hide();
		$("#pnlEquipment").hide();
		$("#pnlService").hide();
		
		$("#pnlLeftNavigation").show();
		
		if($("#selUserType").val()=="Driver" || $("#selUserType").val()=="Field Service and Driver")
			$("#pnlRouteNumber").show();
			
		if($("#selUserType").val()=="Driver" || $("#selUserType").val()=="Field Service and Driver")
			$("#pnlDriverAdditional").show();
	
		if($("#selUserType").val()=="Driver" || $("#selUserType").val()=="Field Service and Driver")
			$("#pnlRoutes").show();

		if($("#selUserType").val()=="Field Service" && <%=filterChangeFlag%> == 1)
			$("#pnlRoutes").show();

		if("<%=MUV_Read("prospectingModuleOn")%>" == "Enabled")
			$("#pnlProspecting").show();
			
		if("<%=MUV_Read("InventoryControlModuleOn")%>" == "Enabled")
			$("#pnlInventoryControl").show();
			
		if("<%=MUV_Read("OrderAPIModuleOn")%>" == "Enabled")
			$("#pnlOrderAPI").show();
			
		if("<%=MUV_Read("equipmentModuleOn")%>" == "Enabled")
			$("#pnlEquipment").show();
			
		if("<%=MUV_Read("serviceModuleOn")%>" == "Enabled")
			$("#pnlService").show();			
	}
	
	$(function () {
		onTypeChanged();
	});
</script>	

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

          
<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateUserForm()
    {

        if (document.frmeditUser.txtFirstName.value == "") {
            alert("First name can not be blank.");
            return false;
        }
        if (document.frmeditUser.txtLastName.value == "") {
            alert("Last  name can not be blank.");
            return false;
        }
        if (document.frmeditUser.txtDisplayName.value == "") {
            alert("Display name can not be blank.");
            return false;
        }
        if (document.frmeditUser.txtEmail.value == "") {
            alert("Email can not be blank.");
            return false;
        }
        if (document.frmeditUser.txtPassword.value == "") {
            alert("Password can not be blank.");
            return false;
        }

        if (document.frmeditUser.txtPassword.value != document.frmeditUser.txtPassword2.value) {
            alert("Password entries do not match. Please re-enter the passwords.");
            return false;
        }
		

		if (document.getElementById("seluserSalesPersonNumber") && document.getElementById("seluserSalesPersonNumber2"))
		
		{
			var selectedValueSalesPerson1 = document.getElementById("seluserSalesPersonNumber").selectedIndex;
			var selectedValueSalesPerson2 = document.getElementById("seluserSalesPersonNumber2").selectedIndex;
			
	
	        if ((selectedValueSalesPerson1 == selectedValueSalesPerson2) && (selectedValueSalesPerson1 != "") && (selectedValueSalesPerson2!= "")) {
	            swal("Primary and Secondary Sales Person Numbers Must Be Unique.");
	            return false;
	        }
	    }

        return true;
	  }
// -->
</SCRIPT>     

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
	margin-top: 20px;
}

.last-run-inputs {
    border: 1px solid #eee;
    box-shadow: 0px 0px 0px 0px;
    max-width: 260px !important;
}

.cancel-save{
	margin-top:30px;
}

.custom-select{
	width: auto !important;
	display:inline-block;
}

.select-large{
	max-width:30% !important;
}

.row-line{
	margin-top:20px;
}

.box-border{
	margin-left:15px;
	margin-top:20px;
	padding:20px;
	background: #f5f5f5;
}

.schedule-rows td {
  width: 80px;
  height: 30px;
  margin: 3px;
  padding: 5px;
  background-color: #eee;
  cursor: pointer;
  border:1px solid #fff;
}
.schedule-rows td:first-child {
  background-color: transparent;
  text-align: right;
  position: relative;
  top: -12px;
}
.schedule-rows td[data-selected],
.schedule-rows td[data-selecting] {
  background-color: #6a0bc1;
}
.schedule-rows td[data-disabled] {
  opacity: 0.55;
}

.tab-content{
    padding: 15px;
     
 }
 
  .tab-content ul{
	 list-style-type:none;
	 -webkit-margin-before: 0px;
    -webkit-margin-after: 0px;
    -webkit-margin-start: 0px;
    -webkit-margin-end: 0px;
    -webkit-padding-start: 0px;
 }
 
 .tabs-line{
	 margin-top:20px;
	 margin-bottom:20px;
 }
.nav-tabs>li>a{
	border:1px solid #ddd;
	border-bottom:transparent;
	background:#f5f5f5;
	color:#000;
}

.nav-tabs>li>a:hover{
	border:1px solid #ddd;
	border-bottom:transparent;
	background:#fff;
}

.schedule-table{
	/*transform:rotate(90deg);*/
}

.schedule-header{
}

.schedule-rows td {
  width: 180px;
  height: 30px;
  margin: 3px;
  padding: 5px;
  background-color: #3498DB;
  cursor: pointer;
}

.schedule-rows td:first-child {
  background-color: transparent;
  text-align: right;
  position: relative;
  top: -12px;
}

.schedule-rows td[data-selected],
.schedule-rows td[data-selecting] { background-color: #E74C3C; }

.schedule-rows td[data-disabled] { opacity: 0.55; }

</style>
<!-- eof password strength meter !-->


<h1 class="page-header"><i class="fa fa-users"></i>Edit User</h1>

<%
SQL = "SELECT * FROM tblUsers where Userno = " & UserNo 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	userFirstName = rs("userFirstName")
	userLastName = rs("userLastName")
	userEmail = rs("userEmail")
	userPassword = rs("userPassword")
	userLastLogin = rs("userLastLogin")
	userEnabled = rs("userEnabled")
	userDownloadEmail = rs("userDownloadEmail")
	userUpdateCalendar = rs("userUpdateCalendar")
	userDisplayName = rs("userDisplayName")
	userCellNumber = rs("userCellNumber")
	userType = rs("userType")
	LoginLandingPage = rs("LoginLandingPageURL")
	userTruckNumber = rs("userTruckNumber")
	userCanAuthSwaps = rs("userCanAuthSwaps") 
	userReceivePartsRequestEmails = rs("userReceivePartsRequestEmails") 
	userFilterRoutes = rs("userFilterRoutes")
	userCRMAccessType = rs("userCRMAccessType")
	userOrderAPIAccessType = rs("userOrderAPIAccessType")
	userCRMDeleteAccess = rs("userCRMDeleteAccess")
	userProspectingAddEditAccess = rs("userProspectingAddEditAccess")
	userEmailSystemID = rs("userEmailSystemID")
	userEmailSystemPass = rs("userEmailSystemPass")
	userEmailServer = rs("userEmailServer")
	userSystemVMSID = rs("userVMS_ID")
	userForceNextStopSelectionOverride = rs("userForceNextStopSelectionOverride")
	userNextStopNagMessageOverride = rs("userNextStopNagMessageOverride")
	userNextStopNagMinutes = rs("userNextStopNagMinutes")
	userNextStopNagIntervalMinutes = rs("userNextStopNagIntervalMinutes")
	userNextStopNagMessageMaxToSendPerStop = rs("userNextStopNagMessageMaxToSendPerStop")
	userNextStopNagMessageMaxToSendThisDriverPerDay = rs("userNextStopNagMessageMaxToSendThisDriverPerDay")
	userNextStopNagMessageSendMethod = rs("userNextStopNagMessageSendMethod")
	userNoActivityNagMessageOverride = rs("userNoActivityNagMessageOverride")
	userNoActivityNagMinutes = rs("userNoActivityNagMinutes")
	userNoActivityNagIntervalMinutes = rs("userNoActivityNagIntervalMinutes")
	userNoActivityNagMessageMaxToSendPerStop = rs("userNoActivityNagMessageMaxToSendPerStop")
	userNoActivityNagMessageMaxToSendPerDriverPerDay = rs("userNoActivityNagMessageMaxToSendPerDriverPerDay")
	userNoActivityNagMessageSendMethod = rs("userNoActivityNagMessageSendMethod")
	userNoActivityNagTimeOfDay = rs("userNoActivityNagTimeOfDay")
	userSalesPersonNumber = rs("userSalesPersonNumber")
	userSalesPersonNumber2 = rs("userSalesPersonNumber2")
	userNoActivityNagMessageOverride_FS = rs("userNoActivityNagMessageOverride_FS")
	userNoActivityNagMinutes_FS = rs("userNoActivityNagMinutes_FS")
	userNoActivityNagIntervalMinutes_FS = rs("userNoActivityNagIntervalMinutes_FS")
	userNoActivityNagMessageMaxToSendPerStop_FS = rs("userNoActivityNagMessageMaxToSendPerStop_FS")
	userNoActivityNagMessageMaxToSendPerDriverPerDay_FS = rs("userNoActivityNagMessageMaxToSendPerDriverPerDay_FS")
	userNoActivityNagMessageSendMethod_FS = rs("userNoActivityNagMessageSendMethod_FS")
	userNoActivityNagTimeOfDay_FS = rs("userNoActivityNagTimeOfDay_FS")
	userLoginDisableAccessHolidays = rs("userLoginDisableAccessHolidays")
	userInventoryControlAccessType = rs("userInventoryControlAccessType")
	userMobileInventoryControlAccess = rs("userMobileInventoryControlAccess")	
	userCanEditEqpTablesOnFly = rs("userEditEqpOnTheFly")
	userEditCRMOnTheFly = rs("userEditCRMOnTheFly")
	userCreateEquipmentSymptomCodesOnTheFly = rs("userCreateEquipmentSymptomCodesOnTheFly")
	userCreateEquipmentProblemCodesOnTheFly = rs("userCreateEquipmentProblemCodesOnTheFly")
	userCreateEquipmentResolutionCodesOnTheFly = rs("userCreateEquipmentResolutionCodesOnTheFly")
	userCreateNewServiceTicket = rs("userCreateNewServiceTicket")
	userAccessServiceDispatchCenter = rs("userAccessServiceDispatchCenter")
	userAccessServiceActionsModalButton = rs("userAccessServiceActionsModalButton")
	userAccessServiceDispatchButton = rs("userAccessServiceDispatchButton")
	userAccessServiceCloseCancelButton = rs("userAccessServiceCloseCancelButton")
	userLeftNavAPIModule = rs("userLeftNavAPIModule")
	userLeftNavBIModule = rs("userLeftNavBIModule")
	userLeftNavProspectingModule = rs("userLeftNavProspectingModule")
	userLeftNavCustomerServiceModule = rs("userLeftNavCustomerServiceModule")
	userLeftNavEquipmentModule = rs("userLeftNavEquipmentModule")
	userLeftNavInventoryControlModule = rs("userLeftNavInventoryControlModule")
	userLeftNavAccountsReceivableModule = rs("userLeftNavAccountsReceivableModule")
	userLeftNavAccountsPayableModule = rs("userLeftNavAccountsPayableModule")
	userLeftNavServiceModule = rs("userLeftNavServiceModule")
	userLeftNavRoutingModule = rs("userLeftNavRoutingModule")
	userLeftNavQuickbooksModule = rs("userLeftNavQuickbooksModule")
	userLeftNavFiltertraxModule = rs("userLeftNavFiltertraxModule")
	userLeftNavSystem = rs("userLeftNavSystem")
End If
	
set rs = Nothing
cnn8.close
set cnn8 = Nothing


SQL = "SELECT * FROM tblServerInfo where clientKey='"& MUV_READ("ClientID") &"'"
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (InsightCnnString)
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	userQuickLoginURL = rs("QuickLoginURL")
End If
	
If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV.") <> 0 AND Left(ucase(userQuickLoginURL),11) <> "HTTP://DEV." Then 

	'get rid of http://
	userQuickLoginURL = Right(userQuickLoginURL,len(userQuickLoginURL)-7)

	'Strip the URL first part
	For x = 1 to len(userQuickLoginURL)
		If Mid(userQuickLoginURL,x,1)="." Then
			userQuickLoginURL = right(userQuickLoginURL,len(userQuickLoginURL)-(x))		
			Exit For
		End If
	Next 

	userQuickLoginURL = "http://dev." & userQuickLoginURL 

End If

If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV2.") <> 0 AND Left(ucase(userQuickLoginURL),12) <> "HTTP://DEV2." Then 

	'get rid of http://
	userQuickLoginURL = Right(userQuickLoginURL,len(userQuickLoginURL)-7)

	'Strip the URL first part
	For x = 1 to len(userQuickLoginURL)
		If Mid(userQuickLoginURL,x,1)="." Then
			userQuickLoginURL = right(userQuickLoginURL,len(userQuickLoginURL)-(x))		
			Exit For
		End If
	Next 

	userQuickLoginURL = "http://dev2." & userQuickLoginURL 

End If

	
set rs = Nothing
cnn8.close
set cnn8 = Nothing


%>



<!-- row !-->
<div class="row">

	<div class="col-lg-12">
	
		<form method="POST" action="edituser_submit.asp" name="frmeditUser" onsubmit="return validateUserForm();">		    
      
			<div>
		
				<input type="hidden" id="txtUserNo" name="txtUserNo" value="<%= UserNo %>">	
				<input type="hidden" name="txtTab" id="txtTab" value="<%= ActiveTab %>">	

			        <!-- row !-->		
			        <div class="row ">	
		    	
				        <!-- First Name !-->
				    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-2">
				    		First Name<input type="text" id="txtFirstName" name="txtFirstName" value="<%= userFirstName %>"  class="form-control last-run-inputs">
				        </div>
				    	<!-- First Name !-->
		        
				        <!-- Last Name !-->
				    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-2">
				        	Last Name<input type="text" id="txtLastName" name="txtLastName" value="<%= userLastName %>"  class="form-control last-run-inputs">
				        </div>
				    	<!-- Last Name !-->

				    	<!-- Display Name !-->
				    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-2">
				        	Display Name<input type="text" id="txtDisplayName" name="txtDisplayName" value="<%= userDisplayName %>"  class="form-control last-run-inputs">
				        </div>
				    	<!-- Display Name !-->
		    	
				    	<!-- Cell Number !-->
				    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-2">
				        	Cell Number<input type="text" id="txtCellNumber" name="txtCellNumber" value="<%= userCellNumber %>"  class="form-control last-run-inputs">
				        </div>
				    	<!-- Cell Number !-->
				    	
				    	<!-- User Number For Quick Login!-->
				    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-2">
				        	Quick Login URL &nbsp;&nbsp;
				        	<a href="send_quick_login_email.asp?u=<%= userNo %>&c=<%=MUV_Read("ClientID")%>"><i class="fa fa-envelope-o fa-lg" style="vertical-align: middle;"></i></a>&nbsp;&nbsp;
				        	<a href="send_quick_login_text.asp?u=<%= userNo %>&c=<%=MUV_Read("ClientID")%>"><i class="fa fa-mobile fa-2x" style="vertical-align: middle;"></i></a>
				        	<br><label><%= userQuickLoginURL %>?u=<%= UserNo %>&c=<%=MUV_Read("ClientID")%></label>
				        </div>
				    	<!-- User Number For Quick Login!-->
		
					</div>
					<!-- eof row !-->
		
		
			 		<!-- row !-->
					<div class="row tab-row genline-up">
		        
				    	<div class="col-lg-6 col-md-6 col-sm-12 col-xs-12">
				
						    <div class="row">
				        
						        <!-- Email !-->
						    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
						    	    Email<input type="text" id="txtEmail" name="txtEmail" value="<%= userEmail %>" class="form-control last-run-inputs">
						        </div>
						    	<!-- Email !-->
						    	
						    	
														    	
								<% If InStr(MUV_READ("LicenseStatus"),"MDS Insight Programmer's Super Ultimate Elite license") Then %>
									
							        <!-- Password !-->
							    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
							    	    Password<input type="text" id="txtPassword" name="txtPassword" value="<%= userPassword %>"  class="form-control last-run-inputs password1"  required data-toggle="popover" title="Password Strength" data-content="Enter Password...">
					 		    	     <script type="text/javascript" language="javascript" src="<%= baseURL %>js/password/passwordstrength.js"></script>
							        </div>
							    	<!-- Password !-->
	
									<!-- Password !-->
							    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
							    	    Re-enter Password<input type="text" id="txtPassword2" name="txtPassword2" value="<%= userPassword %>"  class="form-control last-run-inputs password2"   >		   
							        </div>
							    	<!-- Password !-->
								
								
								<% Else %>
								
							        <!-- Password !-->
							    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
							    	    Password<input type="password" id="txtPassword" name="txtPassword" value="<%= userPassword %>"  class="form-control last-run-inputs password1"  required data-toggle="popover" title="Password Strength" data-content="Enter Password...">
					 		    	     <script type="text/javascript" language="javascript" src="<%= baseURL %>js/password/passwordstrength.js"></script>
							        </div>
							    	<!-- Password !-->
	
									<!-- Password !-->
							    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
							    	    Re-enter Password<input type="password" id="txtPassword2" name="txtPassword2" value="<%= userPassword %>"  class="form-control last-run-inputs password2"   >		   
							        </div>
							    	<!-- Password !-->
								
								<% End If %>

		        

							</div>
					       
							<p>&nbsp;</p>
							
							<div class="row">
				        
								

							</div>
					       
							<p>&nbsp;</p>
						   
 							   
							<!-- select !-->
			      		  	<div class="row">
			     				   <div class="select-line">
				        
						  		      	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
											User Type
											<select class="form-control " name="selUserType" id="selUserType" onchange="onTypeChanged()">
												<option value="CSR" <% If userType = "CSR" Then Response.Write("selected")%>><%=GetTerm("CSR")%></option>
												<option value="CSR Manager" <% If userType = "CSR Manager" Then Response.Write("selected")%>><%=GetTerm("CSR Manager")%></option>
												<option value="Repair" <% If userType = "Repair" Then Response.Write("selected")%>><%=GetTerm("Repair")%></option>
												<option value="Field Service" <% If userType = "Field Service" Then Response.Write("selected")%>><%=GetTerm("Field Service")%></option>
												<option value="Service Manager" <% If userType = "Service Manager" Then Response.Write("selected")%>><%=GetTerm("Service Manager")%></option>
												<option value="Field Service and Driver" <% If userType = "Field Service and Driver" Then Response.Write("selected")%>><%=GetTerm("Field Service")%> AND <%=GetTerm("Driver")%></option>
												<option value="Driver" <% If userType = "Driver" Then Response.Write("selected")%>><%=GetTerm("Driver")%></option>
												<option value="Route Manager" <% If userType = "Route Manager" Then Response.Write("selected")%>><%=GetTerm("Route Manager")%></option>
												<option value="Finance" <% If userType = "Finance" Then Response.Write("selected")%>><%=GetTerm("Finance")%></option>
												<option value="Finance Manager" <% If userType = "Finance Manager" Then Response.Write("selected")%>><%=GetTerm("Finance Manager")%></option>
												<option value="Admin" <% If userType = "Admin" Then Response.Write("selected")%>><%=GetTerm("Admin")%></option>
												<option value="Inside Sales" <% If userType = "Inside Sales" Then Response.Write("selected")%>><%=GetTerm("Inside Sales")%></option>
												<option value="Inside Sales Manager" <% If userType = "Inside Sales Manager" Then Response.Write("selected")%>><%=GetTerm("Inside Sales Manager")%></option>
												<option value="Outside Sales" <% If userType = "Outside Sales" Then Response.Write("selected")%>><%=GetTerm("Outside Sales")%></option>
												<option value="Outside Sales Manager" <% If userType = "Outside Sales Manager" Then Response.Write("selected")%>><%=GetTerm("Outside Sales Manager")%></option>
												<option value="Telemarketing" <% If userType = "Telemarketing" Then Response.Write("selected")%>><%=GetTerm("Telemarketing")%></option>
											</select>
							        	</div>

										<!-- Route number !-->
					                    <div class="col-xs-6 col-sm-1 col-md-1 col-lg-4" id="pnlRouteNumber" style="display: none;">
					                    Default Route
									      	<select class="form-control" name="selRouteNumber" id="selRouteNumber">
									      	  	<option value="0">-- none --</option>
											  	<% 'Get all Routes
											  	  	SQL9 = "SELECT DISTINCT TruckID, DriverName FROM RT_Truck Order By TruckID"
													Set cnn9 = Server.CreateObject("ADODB.Connection")
													cnn9.open (Session("ClientCnnString"))
													Set rs9 = Server.CreateObject("ADODB.Recordset")
													rs9.CursorLocation = 3 
													Set rs9 = cnn9.Execute(SQL9)
													If not rs9.EOF Then
														Do
															If UserTruckNumber = rs9("TruckID") Then
																Response.Write("<option value='" & rs9("TruckID") & "' selected>" & rs9("TruckID") & " - " & rs9("driverName") & "</option>")
															Else
																Response.Write("<option value='" & rs9("TruckID") & "'>" & rs9("TruckID") & " - " & rs9("driverName") & "</option>")
															End If
															rs9.movenext
														Loop until rs9.eof
													End If
													set rs9 = Nothing
													cnn9.close
													set cnn9 = Nothing
											  	%>
											</select>
					                    </div>
					                    <!-- eof Route number !-->

										<% If MUV_READ("biModuleOn") = "Enabled" Then %>
											<!-- Salesman number !-->
						                    <div class="col-xs-6 col-sm-1 col-md-1 col-lg-4" id="pnluserSalesPersonNumber">
						                    <%= GetTerm("Primary Salesman") %> Number
										      	<select class="form-control" name="seluserSalesPersonNumber" id="seluserSalesPersonNumber">
										      	  	<option value="0">-- none --</option>
												  	<% 'Get all Salesman
												  	  	SQL9 = "SELECT DISTINCT SalesManSequence, Name FROM Salesman Order By SalesmanSequence"
														Set cnn9 = Server.CreateObject("ADODB.Connection")
														cnn9.open (Session("ClientCnnString"))
														Set rs9 = Server.CreateObject("ADODB.Recordset")
														rs9.CursorLocation = 3 
														Set rs9 = cnn9.Execute(SQL9)
														If not rs9.EOF Then
															Do
																If userSalesPersonNumber = rs9("SalesmanSequence") Then
																	Response.Write("<option value='" & rs9("SalesmanSequence") & "' selected>" & rs9("SalesmanSequence") & " - " & rs9("Name") & "</option>")
																Else
																	Response.Write("<option value='" & rs9("SalesmanSequence") & "'>" & rs9("SalesmanSequence") & " - " & rs9("Name") & "</option>")
																End If
																rs9.movenext
															Loop until rs9.eof
														End If
														set rs9 = Nothing
														cnn9.close
														set cnn9 = Nothing
												  	%>
												</select>
						                    </div>
						                    <!-- eof Salesman number !-->

											<!-- Secondary Salesman number !-->
						                    <div class="col-xs-6 col-sm-1 col-md-1 col-lg-4" id="pnluserSalesPersonNumber2">
						                    <%= GetTerm("Secondary Salesman") %> Number 
										      	<select class="form-control" name="seluserSalesPersonNumber2" id="seluserSalesPersonNumber2">
										      	  	<option value="0">-- none --</option>
												  	<% 'Get all Salesman
												  	  	SQL9 = "SELECT DISTINCT SalesManSequence, Name FROM Salesman Order By SalesmanSequence"
														Set cnn9 = Server.CreateObject("ADODB.Connection")
														cnn9.open (Session("ClientCnnString"))
														Set rs9 = Server.CreateObject("ADODB.Recordset")
														rs9.CursorLocation = 3 
														Set rs9 = cnn9.Execute(SQL9)
														If not rs9.EOF Then
															Do
																If userSalesPersonNumber2 = rs9("SalesmanSequence") Then
																	Response.Write("<option value='" & rs9("SalesmanSequence") & "' selected>" & rs9("SalesmanSequence") & " - " & rs9("Name") & "</option>")
																Else
																	Response.Write("<option value='" & rs9("SalesmanSequence") & "'>" & rs9("SalesmanSequence") & " - " & rs9("Name") & "</option>")
																End If
																rs9.movenext
															Loop until rs9.eof
														End If
														set rs9 = Nothing
														cnn9.close
														set cnn9 = Nothing
												  	%>
												</select>
						                    </div>
						                    <!-- eof  Secondary Salesman number !-->
					                    <% End If %>
					                    
									</div>
									   
									<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
											Login Landing Page
											<select class="form-control " name="selLoginLandingPage" id="selLoginLandingPage">
												<option value="" <% If LoginLandingPage = "" Then Response.Write("selected")%>>Main Page (Default)</option>

												<% If MUV_Read("serviceModuleOn") = "Enabled" Then %>
													<option value="service/main.asp" <% If LoginLandingPage = "service/main.asp" Then Response.Write("selected")%>>Service Tickets Screen</option>
												<% End If %>
												
												<% If MUV_Read("biModuleOn") = "Enabled" Then %>
													<option value="bizintel/tools/MCS/MCS_Report1.asp" <% If LoginLandingPage = "bizintel/tools/MCS/MCS_Report1.asp" Then Response.Write("selected")%>>MCS Analysis</option>
													<option value="bizintel/dashboard/dashboard.asp" <% If LoginLandingPage = "bizintel/dashboard/dashboard.asp" Then Response.Write("selected")%>>Biz Intel Dashboard</option>
												<% End If %>
												
												<% If MUV_Read("routingModuleOn") = "Enabled" Then %>
													<option value="routing/deliveryBoard.asp" <% If LoginLandingPage = "routing/deliveryBoard.asp" Then Response.Write("selected")%>>Delivery Board</option>
												<% End If %>
												
												<% If cint(MUV_Read("arModuleOn")) = 1 Then %>
													<option value="accountsreceivable/TicketsOnHold.asp" <% If LoginLandingPage = "accountsreceivable/TicketsOnHold.asp" Then Response.Write("selected")%>>Service Tickets On Hold</option>
												<% End If %>
												
											</select>
							        	</div>

									</div>
									   
										

								</div>
								
							</div>
				        
							<div class="col-lg-6 col-md-6 col-sm-12 col-xs-12 pass-strength">
 		
								 
                            
	                             <!-- enabled line !-->
	                             <div class="row">
	                             
									<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
										<strong>Enabled</strong>
										<% If userEnabled = vbTrue Then %>
											<input type="checkbox" checked id="chkEnabled"  name="chkEnabled">
										<% Else %>
											<input type="checkbox" unchecked id="chkEnabled"  name="chkEnabled">		    
										<%End If%>
									</div>
									<!-- eof enabled line !-->
									
									<%'If advancedDispatchIsOn() Then %>
										<!-- userCanAuthSwaps line !-->
										<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable" id="pnlSwap" style="display: none;">
											<strong>Can authorize equipment swaps</strong>
											<% If userCanAuthSwaps = vbTrue Then %>
												<input type="checkbox" checked="checked" id="chkuserCanAuthSwaps" name="chkuserCanAuthSwaps">
											<% Else %>
												<input type="checkbox" id="chkuserCanAuthSwaps" name="chkuserCanAuthSwaps">		    
											<%End If%>
											<br>
											<strong>User receives parts request emails</strong>
											<% If userReceivePartsRequestEmails = 1 Then %>
												 <input type="checkbox" checked="checked" id="chkuserReceivePartsRequestEmails"  name="chkuserReceivePartsRequestEmails">		    
											<% Else %>
												 <input type="checkbox" id="chkuserReceivePartsRequestEmails"  name="chkuserReceivePartsRequestEmails">		    
											<%End If%>
										</div>
										<!-- eof userCanAuthSwaps line !-->
									<%'End If%>
								
								
									
									


									
			 			      
                              
                              </div>
                                
						</div> 
						<!-- eof row !-->
						
                        
 
<!-- tabs start here -->
 <div class="col-lg-12 tabs-line">
 
  <!-- Nav tabs -->
  <ul class="nav nav-tabs" role="tablist">
    <li role="presentation" class="active"><a href="#general" aria-controls="general" role="tab" data-toggle="tab">General</a></li>
    <li role="presentation" id="pnlLeftNavigation" style="display: none;"><a href="#leftnav" aria-controls="leftnav" role="tab" data-toggle="tab">Left Navigation</a></li>
    <li role="presentation" id="pnlDriverAdditional" style="display: none;"><a href="#routing" aria-controls="routing" role="tab" data-toggle="tab">Routing</a></li>
    <li role="presentation" id="pnlService" style="display: none;"><a href="#service" aria-controls="service" role="tab" data-toggle="tab"><%= GetTerm("Service") %></a></li>
    <li role="presentation" id="pnlProspecting" style="display: none;"><a href="#prospecting" aria-controls="prospecting" role="tab" data-toggle="tab">Prospecting</a></li>
    <li role="presentation" id="pnlEquipment" style="display: none;"><a href="#equipment" aria-controls="equipment" role="tab" data-toggle="tab">Equipment</a></li>
    <li role="presentation" id="pnlInventoryControl" style="display: none;"><a href="#inventorycontrol" aria-controls="inventorycontrol" role="tab" data-toggle="tab">Inventory Control</a></li>
    <li role="presentation" id="pnlOrderAPI" style="display: none;"><a href="#api" aria-controls="api" role="tab" data-toggle="tab">Order API access</a></li>
    <li role="presentation" id="pnlRoutes" style="display: none;"><a href="#fieldservice" aria-controls="fieldservice" role="tab" data-toggle="tab">Field Service</a></li>
    <li role="presentation" id="pnlLoginAccess"><a href="#loginaccess" aria-controls="loginaccess" role="tab" data-toggle="tab">Access Schedule</a></li>
   </ul>

  <!-- Tab panes -->
  <div class="tab-content">
		
		<!-- General tab -->
		<!--#include file="editUserTabs/general.asp"-->
		<!-- eof General tab -->

	    <!-- Left Navigation tab -->
	    <!--#include file="editUserTabs/leftnavigation.asp"-->
	    <!-- eof Left Navigation tab -->
		
		<!-- routing tab -->
		<!--#include file="editUserTabs/routing.asp"-->
		<!-- eof routing tab -->

		<!-- Service tab -->
		<!--#include file="editUserTabs/service.asp"-->
		<!-- eof Service tab -->
		
		<!-- Prospecting tab -->
		<!--#include file="editUserTabs/prospecting.asp"-->
		<!-- eof Prospecting tab -->
		
		<!-- Equipment tab -->
		<!--#include file="editUserTabs/equipment.asp"-->
		<!-- eof Equipment tab -->
 
	    <!-- Inventory Control tab -->
	    <!--#include file="editUserTabs/inventorycontrol.asp"-->
	    <!-- eof Inventory Control tab -->
		
		<!-- API tab -->
		<!--#include file="editUserTabs/api.asp"-->
		<!-- eof API tab -->
		
		<!-- Field Service tab -->
		<!--#include file="editUserTabs/fieldservice.asp"-->
		<!-- eof Field Service tab -->
		
		<!-- Login Access Schedule tab -->
		<!--#include file="editUserTabs/loginaccess.asp"-->
		<!-- eof Login Access Schedule tab -->
    
  </div>
	<!-- eof Tab panes -->
 	
 </div>
<!-- tabs end here -->
						
						<!-- row !-->
						<div class="row ">


							<div class="col-lg-12 alertbutton cancel-save">
	
							    <a href="<%= BaseURL %>admin/users/main.asp#<%= ActiveTab %>">
							    	<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To User List</button>
								</a>
								
								
								<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
 
						    </div>	

						</div>
				    
				</div>
    
			</div>

		</form>

	</div>	

</div>
<!-- eof row !-->   



   
<!--#include file="../../inc/footer-main.asp"-->
