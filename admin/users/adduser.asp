<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<% 
ActiveTab = Request.QueryString("tab")

If filterChangeModuleOn() Then filterChangeFlag = 1 Else filterChangeFlag = 0
filterChangeFlag = 1

%>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

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

<SCRIPT LANGUAGE="JavaScript">
<!--
	function validateFieldsAndUser()
	{
	    if(validateUserForm()) 
	    {
	        checkIfUserExists();
	    }
	}
	
	
	function checkIfUserExists()
	{	
	
		userEmail = $('#txtEmail').val();
		userPassword = $('#txtPassword').val();
		
		$.ajax({
			type:"POST",
			url: "adduser_ajaxfuncs.asp",
			data: "action=checkIfUserExists&userEmail="+encodeURIComponent(userEmail)+"&userPassword="+encodeURIComponent(userPassword),
			async: true,
			success: function(msg){
	            if(msg !== "success"){
				    swal("The user have entered (email address and password combination) already exists.");
           			return false;
	            } 
	            else
	            {
	            	document.frmAddUser.submit();
	            }
			}
		});
	}

    function validateUserForm()
    {

        if (document.frmAddUser.txtFirstName.value == "") {
            swal("First name cannot be blank.");
            return false;
        }
        if (document.frmAddUser.txtLastName.value == "") {
            swal("Last  name cannot be blank.");
            return false;
        }
        if (document.frmAddUser.txtDisplayName.value == "") {
            swal("Display name cannot be blank.");
            return false;
        }
        if (document.frmAddUser.txtEmail.value == "") {
            swal("Email cannot be blank.");
            return false;
        }
        if (document.frmAddUser.txtPassword.value == "") {
            swal("Password cannot be blank.");
            return false;
        }

        if (document.frmAddUser.txtPassword.value != document.frmAddUser.txtPassword2.value) {
            swal("Password entries do not match. Please re-enter the passwords.");
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
	margin-top:20px;
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


<h1 class="page-header"><i class="fa fa-users"></i>Add New User</h1>



<!-- row !-->
<div class="row">
	<div class="col-lg-12">
	
	<form method="POST" action="adduser_submit.asp" name="frmAddUser" id="frmAddUser">		    
      

		<div>
		        <!-- row !-->		
		        <div class="row ">	
		    	
		        <!-- First Name !-->
		    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-2">
		        First Name<input type="text" id="txtFirstName" name="txtFirstName" value=""  class="form-control last-run-inputs">
		        </div>
		    	<!-- First Name !-->
		        
		        <!-- Last Name !-->
		    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-2">
		        Last Name<input type="text" id="txtLastName" name="txtLastName" value=""  class="form-control last-run-inputs">
		        </div>
		    	<!-- Last Name !-->

		    	<!-- Display Name !-->
		    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-2">
		        Display Name<input type="text" id="txtDisplayName" name="txtDisplayName" value=""  class="form-control last-run-inputs">
		        </div>
		    	<!-- Display Name !-->
		    	
		    	<!-- Cell Number !-->
		    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-2">
		        Cell Number<input type="text" id="txtCellNumber" name="txtCellNumber" value=""  class="form-control last-run-inputs">
		        </div>
		    	<!-- Cell Number !-->
		
		</div>
		<!-- eof row !-->
		
 		<!-- row !-->
		<div class="row tab-row genline-up">
		        
		        <div class="col-lg-6 col-md-6 col-sm-12 col-xs-12">
			      <div class="row">
				        
		        <!-- Email !-->
		    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
		        Email<input type="text" id="txtEmail" name="txtEmail" value=""  class="form-control last-run-inputs">
		        </div>
		    	<!-- Email !-->
		        
		        <!-- Password !-->
		    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
		        Password<input type="password" id="txtPassword" name="txtPassword" value=""  class="form-control last-run-inputs password1"  required data-toggle="popover" title="Password Strength" data-content="Enter Password...">
 		         <script type="text/javascript" language="javascript" src="<%= baseURL %>js/password/passwordstrength.js"></script>
		        </div>
		    	<!-- Password !-->

				<!-- Password !-->
		    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
		        Re-enter Password<input type="password" id="txtPassword2" name="txtPassword2" value=""  class="form-control last-run-inputs password2"   >		   
		        </div>
		    	<!-- Password !-->
		        </div>
		        <p>&nbsp;</p>
		        
				<div class="row">
	        
			        

		       </div>
		        
	        
		        
 			        
			        <!-- select !-->
			        <div class="row">
			        <div class="select-line">
				        
				         	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
				        	
								User Type
								<select class="form-control" name="selUserType" id="selUserType" onchange="onTypeChanged()">
								<option value="CSR" selected><%=GetTerm("CSR")%></option>
								<option value="CSR Manager"><%=GetTerm("CSR Manager")%></option>
								<option value="Repair"><%=GetTerm("Repair")%></option>
								<option value="Field Service"><%=GetTerm("Field Service")%></option>
								<option value="Service Manager"><%=GetTerm("Service Manager")%></option>
								<option value="Field Service and Driver"><%=GetTerm("Field Service")%> AND <%=GetTerm("Driver")%></option>
								<option value="Driver"><%=GetTerm("Driver")%></option>
								<option value="Route Manager"><%=GetTerm("Route Manager")%></option>
								<option value="Finance"><%=GetTerm("Finance")%></option>
								<option value="Finance Manager"><%=GetTerm("Finance Manager")%></option>
								<option value="Admin"><%=GetTerm("Admin")%></option>
								<option value="Inside Sales"><%=GetTerm("Inside Sales")%></option>
								<option value="Inside Sales Manager"><%=GetTerm("Inside Sales Manager")%></option>
								<option value="Outside Sales"><%=GetTerm("Outside Sales")%></option>
								<option value="Outside Sales Manager"><%=GetTerm("Outside Sales Manager")%></option>
								<option value="Telemarketing"><%=GetTerm("Telemarketing")%></option>
								
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
										Response.Write("<option value='" & rs9("TruckID") & "'>" & rs9("TruckID") & " - " & rs9("driverName") & "</option>")
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
						      	<% 'Get all SalesPersons
						      	  	SQL9 = "SELECT DISTINCT SalesManSequence, Name FROM Salesman Order By SalesmanSequence"
									Set cnn9 = Server.CreateObject("ADODB.Connection")
									cnn9.open (Session("ClientCnnString"))
									Set rs9 = Server.CreateObject("ADODB.Recordset")
									rs9.CursorLocation = 3 
									Set rs9 = cnn9.Execute(SQL9)
									If not rs9.EOF Then
										Do
											Response.Write("<option value='" & rs9("SalesmanSequence") & "'>" & rs9("SalesmanSequence") & " - " & rs9("Name") & "</option>")
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

						<!-- Salesman number !-->
	                    <div class="col-xs-6 col-sm-1 col-md-1 col-lg-4" id="pnluserSalesPersonNumber2">
	                    <%= GetTerm("Secondary Salesman") %> Number
					      	<select class="form-control" name="seluserSalesPersonNumber2" id="seluserSalesPersonNumber2">
					      	  	<option value="0">-- none --</option>
						      	<% 'Get all SalesPersons
						      	  	SQL9 = "SELECT DISTINCT SalesManSequence, Name FROM Salesman Order By SalesmanSequence"
									Set cnn9 = Server.CreateObject("ADODB.Connection")
									cnn9.open (Session("ClientCnnString"))
									Set rs9 = Server.CreateObject("ADODB.Recordset")
									rs9.CursorLocation = 3 
									Set rs9 = cnn9.Execute(SQL9)
									If not rs9.EOF Then
										Do
											Response.Write("<option value='" & rs9("SalesmanSequence") & "'>" & rs9("SalesmanSequence") & " - " & rs9("Name") & "</option>")
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
					<% End If %>

			        </div>
 			        <!-- eof select !-->
 

 			        <!-- login landing page !-->
		         	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4">
						Login Landing Page
						<select class="form-control" name="selLoginLandingPage" id="selLoginLandingPage">
							<option value="" selected>Main Page (Default)</option>
							
							<% If MUV_Read("serviceModuleOn") = "Enabled" Then %>
								<option value="service/main.asp">Service Tickets Screen</option>
							<% End If %>
							
							<% If MUV_Read("biModuleOn") = "Enabled" Then %>
								<option value="bizintel/tools/MCS/MCS_Report1.asp">MCS Analysis</option>
								<option value="bizintel/dashboard/dashboard.asp" <% If LoginLandingPage = "bizintel/dashboard/dashboard.asp" Then Response.Write("selected")%>>Biz Intel Dashboard</option>
							<% End If %>
							
							<% If MUV_Read("routingModuleOn") = "Enabled" Then %>
								<option value="routing/deliveryBoard.asp">Delivery Board</option>
							<% End If %>
							
							<% If cint(MUV_Read("arModuleOn")) = 1 Then %>
								<option value="accountsreceivable/TicketsOnHold.asp">Service Tickets On Hold</option>
							<% End If %>
						</select>
		        	</div>


		        		        
		  

 		        </div>
		        </div>
		    	
		    	 <div class="col-lg-6 col-md-6 col-sm-12 col-xs-12 pass-strength">
 		
 
	

   <!-- enabled line !-->
   <div class="row">
   
		   	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable">
		  	  <strong>Enabled</strong> <input type="checkbox" id="chkEnabled" name="chkEnabled" checked="checked">
		    </div>
		    <!-- eof enabled line !-->
		    
			<% 'If advancedDispatchIsOn() Then %>

				<div class="col-xs-6 col-sm-1 col-md-1 col-lg-4 enable-disable" id="pnlSwap" style="display: none;">
					<strong>Can authorize equipment swaps</strong> <input type="checkbox" id="chkuserCanAuthSwaps"  name="chkuserCanAuthSwaps">
					<br>
					<strong>User receives parts request emails</strong> <input type="checkbox" id="chkuserReceivePartsRequestEmails"  name="chkuserReceivePartsRequestEmails">
				</div>

			<%'End If %>

		   

			
		</div> 
	</div> 
    
    
                        
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
    <!--#include file="addUserTabs/general.asp"-->
    <!-- eof General tab -->

    <!-- Left Navigation tab -->
    <!--#include file="addUserTabs/leftnavigation.asp"-->
    <!-- eof Left Navigation tab -->
    
    <!-- routing tab -->
    <!--#include file="addUserTabs/routing.asp"-->
    <!-- eof routing tab -->

	<!-- Service tab -->
	<!--#include file="editUserTabs/service.asp"-->
	<!-- eof Service tab -->
    
    <!-- Prospecting tab -->
    <!--#include file="addUserTabs/prospecting.asp"-->
    <!-- eof Prospecting tab -->
    
	<!-- Equipment tab -->
	<!--#include file="addUserTabs/equipment.asp"-->
	<!-- eof Equipment tab -->
 
    <!-- Inventory Control tab -->
    <!--#include file="addUserTabs/inventorycontrol.asp"-->
    <!-- eof Inventory Control tab -->
	
	<!-- API tab -->
	<!--#include file="addUserTabs/api.asp"-->
	<!-- eof API tab -->
	
	<!-- Field Service tab -->
	<!--#include file="addUserTabs/fieldservice.asp"-->
	<!-- eof Field Service tab -->
	
	<!-- Login Access Schedule tab -->
	<!--#include file="addUserTabs/loginaccess.asp"-->
	<!-- eof Login Access Schedule tab -->
  </div>
	<!-- eof Tab panes -->
 	
 </div>
<!-- tabs end here -->


<div class="genline-up">

	<div class="col-lg-12 alertbutton">
	
    <a href="<%= BaseURL %>admin/users/main.asp#<%= ActiveTab %>">
    	<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To User List</button>
	</a>
	
		<input type="hidden" name="txtTab" id="txtTab" value="<%= ActiveTab %>">

		<button type="button" class="btn btn-primary" onClick="validateFieldsAndUser();"><i class="far fa-save"></i> Save</button>
 
    </div>	

</div>
<!-- eof row !-->    

    
</div>

</form>

</div>	


</div>
<!-- eof row !-->    

 
   
  
<!-- eof day schedule JS --> 
<!--#include file="../../inc/footer-main.asp"-->
