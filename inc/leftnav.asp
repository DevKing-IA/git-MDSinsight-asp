
<script language="javascript">


function ModuleNotEnabled(ModuleName)
{
    swal({
        title: ModuleName+ " Not Enabled",
        text: "Please contact support if you would like to activate the " + ModuleName + " Module.",
        confirmButtonColor: "#337ab7",
        confirmButtonText: 'OK'
	});
}


</script>


<nav>
	<ul class="list-unstyled main-menu">
	 
		<li class="text-right"><a href="#" id="nav-close">X</a></li>
		
		<%'***********************************************************************************************************************************************************************
		'Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main Main 
		'***********************************************************************************************************************************************************************%>
         <li><a href="<%= BaseURL %>main/default.asp"><i class="fa fa-fw fa-home" data-toggle="tooltip" title="Main"></i> Main<span class="icon"></span></a></li>

		
			<% IF MUV_READ("FILTERTRAX") <> "1" Then %>
				
				<!--Include your navigation here		
				<% If MUV_READ("SERNO") <> "1999" Then %>
					<%If Instr(ucase(MUV_READ("LICENSESTATUS")),"PROGRAM") <> 0 or (Session("UserNo")= 4 AND Instr(ucase(MUV_READ("SERNO")),"1071") <> 0)Then %>
					<li class="sub-nav"><a href="<%= BaseURL %>bizintel/tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp"> Category Analysis By Period<span class="icon"></span></a></li>
					<% End If %>
					
					<%If Instr(ucase(MUV_READ("LICENSESTATUS")),"PROGRAM") <> 0 Then %>
						<li class="sub-nav"><a href="<%= BaseURL %>directlaunch/kiosks/routing/DeliveryBoardKioskNoPaging.asp?ri=15&pp=<%=GetPassPhrase(MUV_READ("SERNO"))%>&cl=<%=MUV_READ("SERNO")%>">Launch Delivery Kiosk<span class="icon"></span></a></li>
						<li class="sub-nav"><a href="<%= BaseURL %>directlaunch/kiosks/service/FieldServiceKioskNoPaging.asp?ri=15&pp=<%=GetPassPhrase(MUV_READ("SERNO"))%>&cl=<%=MUV_READ("SERNO")%>">Launch Service Kiosk<span class="icon"></span></a></li>
					<% End If %>
				<% End If %>
				-->


				<%'************************************************************************************************************************************************************************
				'API MODULE API MODULE API MODULE API MODULE API MODULE API MODULE API MODULE API MODULE API MODULE API MODULE API MODULE API MODULE API MODULE API MODULE API MODULE 
				'**************************************************************************************************************************************************************************
				If MUV_Read("OrderAPIModuleOn") = "Enabled" AND userViewLeftNavAPIModule(Session("UserNo")) = true Then %>
				    <li><a href="#"><i class="fa fa-fw fa-plug" data-toggle="tooltip" title="API"></i> API</a>
					<ul class="list-unstyled">
		          		<li class="sub-nav"><a href="<%= BaseURL %>api/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
		          		<li class="sub-nav"><a href="<%= BaseURL %>api/menu.asp"> Menu<span class="icon"></span></a></li>
		          		<li class="sub-nav"><a href="<%= BaseURL %>api/reports.asp"> Reports<span class="icon"></span></a></li>		          		
					</ul>
				<% ElseIf MUV_Read("OrderAPIModuleOn") = "Disabled" Then ' tempt setting  %>	
				    <li id="disabled-item"><a href="#" onclick="ModuleNotEnabled('API');"><i class="fa fa-fw fa-asterisk" data-toggle="tooltip" title="Please contact support if you would like to activate the <%=GetTerm("API")%> module"></i> <%= GetTerm("API") %></a></li>				
				<%End If


				'************************************************************************************************************************************************************************
				'BI Module	BI Module	BI Module	BI Module	BI Module	BI Module	BI Module	BI Module	BI Module	BI Module	BI Module	BI Module	BI Module	BI Module	
				'**************************************************************************************************************************************************************************
				If MUV_Read("biModuleOn") = "Enabled" AND userViewLeftNavBIModule(Session("UserNo")) = true Then %>
				    <li><a href="#"><i class="fa fa-fw fa-graduation-cap" data-toggle="tooltip" title="Business Intelligence"></i> Business Intelligence</a>
					<ul class="list-unstyled">
		          		<li class="sub-nav"><a href="<%= BaseURL %>bizintel/dashboard/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
		          		<li class="sub-nav"><a href="<%= BaseURL %>bizintel/menu.asp"> Menu<span class="icon"></span></a></li>
		          		<li class="sub-nav"><a href="<%= BaseURL %>bizintel/reports.asp"> Reports<span class="icon"></span></a></li>		          		
					</ul>
				<% ElseIf MUV_Read("biModuleOn") = "Disabled" Then ' tempt setting  %>	
				    <li id="disabled-item"><a href="#" onclick="ModuleNotEnabled('Business Intelligence');"><i class="fa fa-fw fa-asterisk" data-toggle="tooltip" title="Please contact support if you would like to activate the <%=GetTerm("Business Intelligence")%> module"></i> <%= GetTerm("Business Intelligence") %></a></li>				
				<%End If


				'***********************************************************************************************************************************************************************************
				'Prospecting Module	Prospecting Module	Prospecting Module	Prospecting Module	Prospecting Module	Prospecting Module	Prospecting Module	Prospecting Module	Prospecting Module	
				'***********************************************************************************************************************************************************************************
			    If MUV_Read("prospectingModuleOn") = "Enabled" AND userViewLeftNavProspectingModule(Session("UserNo")) = true Then %>  						       								      
						<li><a href="#"><i class="fa fa-fw fa-asterisk" data-toggle="tooltip" title="<%= GetTerm("Prospecting") %>"></i> <%= GetTerm("Prospecting") %></a>
							<ul class="list-unstyled">
								<li class="sub-nav"><a href="<%= BaseURL %>prospecting/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>prospecting/menu.asp"> Menu<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>prospecting/reports.asp"> Reports<span class="icon"></span></a></li>	
							</ul>
					    </li>
						
				<% ElseIf MUV_Read("prospectingModuleOn")  = "Disabled" Then ' tempt setting %>
					    <li id="disabled-item"><a href="#" onclick="ModuleNotEnabled('Propsecting');"><i class="fa fa-fw fa-asterisk" data-toggle="tooltip" title="Please contact support if you would like to activate the <%=GetTerm("Prospecting")%> module"></i> <%= GetTerm("Prospecting") %></a></li>
				<% End If



				'**************************************************************************************************************************************************************************
				'Customer Service Customer Service Customer Service  Customer Service Customer Service Customer Service Customer Service Customer Service Customer Service Customer Service
				'**************************************************************************************************************************************************************************
			    If MUV_Read("custServiceOn") = 1 AND userViewLeftNavCustomerServiceModule(Session("UserNo")) = true Then %>  						       								      
						<li><a href="#"><i class="fa fa-fw fa-users" data-toggle="tooltip" title="<%= GetTerm("Customer Service") %>"></i> <%= GetTerm("Customer Service") %></a>
							<ul class="list-unstyled">
								<li class="sub-nav"><a href="<%= BaseURL %>customerservice/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>customerservice/menu.asp"> Menu<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>customerservice/reports.asp"> Reports<span class="icon"></span></a></li>	
							</ul>
					    </li>
				<% End If
			


				'**************************************************************************************************************************************************************************
				'Equipment Equipment Equipment Equipment Equipment Equipment Equipment Equipment Equipment Equipment Equipment Equipment Equipment Equipment Equipment Equipment Equipment 
				'**************************************************************************************************************************************************************************
			    If MUV_Read("equipmentModuleOn") = "Enabled" AND userViewLeftNavEquipmentModule(Session("UserNo")) = true Then %>  						       								      
						<li><a href="#"><i class="fa fa-fw fa-coffee" data-toggle="tooltip" title="<%= GetTerm("Equipment") %>"></i> <%= GetTerm("Equipment") %></a>
							<ul class="list-unstyled">
								<li class="sub-nav"><a href="<%= BaseURL %>equipment/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>equipment/menu.asp"> Menu<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>equipment/reports.asp"> Reports<span class="icon"></span></a></li>	
							</ul>
					    </li>
						
				<% ElseIf MUV_Read("equipmentModuleOn")  = "Disabled" Then ' tempt setting %>
					    <li id="disabled-item"><a href="#" onclick="ModuleNotEnabled('Equipment');"><i class="fa fa-fw fa-coffee" data-toggle="tooltip" title="Please contact support if you would like to activate the <%=GetTerm("Equipment")%> module"></i> <%= GetTerm("Equipment") %></a></li>
				<% End If



				'**************************************************************************************************************************************************************************
				'Inventory Control Inventory Control Inventory Control  Inventory Control Inventory Control Inventory Control Inventory Control Inventory Control Inventory Control 
				'**************************************************************************************************************************************************************************
			    If MUV_Read("InventoryControlModuleOn") = "Enabled" AND userViewLeftNavInventoryControlModule(Session("UserNo")) = true Then %> 						       								      
						<li><a href="#"><i class="fas fa-fw fa-forklift" data-toggle="tooltip" title="<%= GetTerm("Inventory Control") %>"></i> <%= GetTerm("Inventory Control") %></a>
							<ul class="list-unstyled">
								<li class="sub-nav"><a href="<%= BaseURL %>inventorycontrol/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>inventorycontrol/menu.asp"> Menu<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>inventorycontrol/reports.asp"> Reports<span class="icon"></span></a></li>	
							</ul>
					    </li>
						
				<% ElseIf MUV_Read("InventoryControlModuleOn")  = "Disabled" Then ' tempt setting %>
					    <li id="disabled-item"><a href="#" onclick="ModuleNotEnabled('Inventory Control');"><i class="fas fa-fw fa-forklift" data-toggle="tooltip" title="Please contact support if you would like to activate the <%=GetTerm("Inventory Control")%> module"></i> <%= GetTerm("Inventory Control") %></a></li>
				<% End If



				'**********************************************************************************************************************************************************************
				'AR Module	AR Module	AR Module	AR Module	AR Module	AR Module	AR Module	AR Module	AR Module	AR Module	AR Module	AR Module	AR Module	AR Module	
				'**********************************************************************************************************************************************************************
			    If cint(MUV_Read("arModuleOn")) = 1 AND userViewLeftNavAccountsReceivableModule(Session("UserNo")) = true Then %>					       								      
						<li><a href="#"><i class="fas fa-fw fa-file-invoice-dollar" data-toggle="tooltip" title="<%= GetTerm("Accounts Receivable") %>"></i> <%= GetTerm("Accounts Receivable") %></a>
							<ul class="list-unstyled">
								<li class="sub-nav"><a href="<%= BaseURL %>accountsreceivable/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>accountsreceivable/menu.asp"> Menu<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>accountsreceivable/reports/main.asp"> Reports<span class="icon"></span></a></li>	
							</ul>
					    </li>
				<% End If
			


				'**********************************************************************************************************************************************************************
				'AP Module	AP Module	AP Module	AP Module	AP Module	AP Module	AP Module	AP Module	AP Module	AP Module	AP Module	AR Module	AP Module	AP Module	
				'**********************************************************************************************************************************************************************
			    If MUV_Read("apModuleOn") = "Enabled" AND userViewLeftNavAccountsPayableModule(Session("UserNo")) = true Then %>  						       								      
						<li><a href="#"><i class="fas fa-fw fa-envelope-open-dollar" data-toggle="tooltip" title="<%= GetTerm("Accounts Payable") %>"></i> <%= GetTerm("Accounts Payable") %></a>
							<ul class="list-unstyled">
								<li class="sub-nav"><a href="<%= BaseURL %>accountspayable/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>accountspayable/menu.asp"> Menu<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>accountspayable/reports.asp"> Reports<span class="icon"></span></a></li>	
							</ul>
					    </li>
				<% End If


				'**********************************************************************************************************************************************************************
				'Service Module Service Module Service Module Service Module Service Module Service Module Service Module Service Module Service Module Service Module Service Module 
				'*********************************************************************************************************************************************************************
			    If MUV_Read("serviceModuleOn") = "Enabled" AND userViewLeftNavServiceModule(Session("UserNo")) = true Then %>  						       								      
						<li><a href="#"><i class="fa fa-fw fa-wrench" data-toggle="tooltip" title="<%= GetTerm("Service") %>"></i> <%= GetTerm("Service") %></a>
							<ul class="list-unstyled">
								<li class="sub-nav"><a href="<%= BaseURL %>service/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>service/menu.asp"> Menu<span class="icon"></span></a></li>
								<!--<li class="sub-nav"><a href="<%= BaseURL %>service/reports.asp"> Reports<span class="icon"></span></a></li>	-->
							</ul>
					    </li>
						
				<% ElseIf MUV_Read("serviceModuleOn")  = "Disabled" Then ' tempt setting %>
					    <li id="disabled-item"><a href="#" onclick="ModuleNotEnabled('Service');"><i class="fa fa-fw fa-wrench" data-toggle="tooltip" title="Please contact support if you would like to activate the <%=GetTerm("Service")%> module"></i> <%= GetTerm("Service") %></a></li>
				<% End If



				'*************************************************************************************************************************************************************
				'Routing Module	Routing Module	Routing Module	Routing Module	Routing Module	Routing Module	Routing Module	Routing Module	Routing Module Routing Module
				'*************************************************************************************************************************************************************
			    If MUV_Read("routingModuleOn") = "Enabled" AND userViewLeftNavRoutingModule(Session("UserNo")) = true Then %>  						       								      
						<li><a href="#"><i class="fa fa-fw fa-truck" data-toggle="tooltip" title="<%= GetTerm("Routing") %>"></i> <%= GetTerm("Routing") %></a>
							<ul class="list-unstyled">
								<li class="sub-nav"><a href="<%= BaseURL %>routing/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>routing/menu.asp"> Menu<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>routing/reports.asp"> Reports<span class="icon"></span></a></li>	
							</ul>
					    </li>
						
				<% ElseIf MUV_Read("routingModuleOn")  = "Disabled" Then ' tempt setting %>
					    <li id="disabled-item"><a href="#" onclick="ModuleNotEnabled('Service');"><i class="fa fa-fw fa-truck" data-toggle="tooltip" title="Please contact support if you would like to activate the <%=GetTerm("Routing")%> module"></i> <%= GetTerm("Routing") %></a></li>
				<% End If 

 

				'***********************************************************************************************************************************************************************
				'System  System  System  System  System  System  System  System  System  System  System  System  System  System  System  System  System  System  System  System  System 
				'**********************************************************************************************************************************************************************
				If userIsAdmin(Session("userNo")) AND userViewLeftNavSystem(Session("UserNo")) = true Then %>
						<li><a href="#"><i class="fa fa-fw fa-desktop" data-toggle="tooltip" title="<%= GetTerm("System") %>"></i> <%= GetTerm("System") %></a>
							<ul class="list-unstyled">
								<li class="sub-nav"><a href="<%= BaseURL %>system/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>system/menu.asp"> Menu<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>system/reports.asp"> Reports<span class="icon"></span></a></li>	
							</ul>
					    </li>
				<% End If


				'***************************************************************************************************************************************************************
				'Quickbooks Module	Quickbooks Module	Quickbooks Module	Quickbooks Module	Quickbooks Module	Quickbooks Module	Quickbooks Module	Quickbooks Module	
				'***************************************************************************************************************************************************************
			    If MUV_Read("quickbooksModuleOn") = "Enabled" AND userViewLeftNavQuickbooksModule(Session("UserNo")) = true Then %>  						       								      
						<li><a href="#"><i class="fa fa-usd" data-toggle="tooltip" title="Quickbooks"></i> Quickbooks</a>
							<ul class="list-unstyled">
								<li class="sub-nav"><a href="<%= BaseURL %>quickbooks/dashboard.asp"> Dashboard<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>quickbooks/menu.asp"> Menu<span class="icon"></span></a></li>
								<li class="sub-nav"><a href="<%= BaseURL %>quickbooks/reports.asp"> Reports<span class="icon"></span></a></li>	
							</ul>
					    </li>
						
				<% ElseIf MUV_Read("quickbooksModuleOn")  = "Disabled" Then ' tempt setting %>
					    <li id="disabled-item"><a href="#" onclick="ModuleNotEnabled('Service');"><i class="fa fa-usd" data-toggle="tooltip" title="Please contact support if you would like to activate the Quickbooks module"></i> Quickbooks</a></li>
				<% End If %>
     

   			<% ElseIf userViewLeftNavFiltertraxModule(Session("UserNo")) = true Then  'Filtertrax flag %>

				<li><a href="<%= BaseURL %>service/main.asp"><i class="fa fa-fw fa-sticky-note" data-toggle="tooltip" title="Service Tickets"></i> Service Tickets<span class="icon"></span></a></li>
				<li><a href="<%= BaseURL %>service/dispatchcenter/main.asp"><i class="fas fa-fw fa-share-square" data-toggle="tooltip" title="Dispatch Center"></i> Dispatch Center<span class="icon"></span></a></li>
				<li><a href="<%= BaseURL %>service/filters/custfilters/main.asp"><i class="fa fa-fw fa-filter" data-toggle="tooltip" title="Manage Customer Filters"></i> Manage Customer Filters<span class="icon"></span></a></li>
				
				<%'*******************************************************************************************************************************************************************
				'Add / Edit Add / Edit Add / Edit Add / Edit Add / Edit Add / Edit Add / Edit Add / Edit Add / Edit Add / Edit Add / Edit Add / Edit Add / Edit Add / Edit Add / Edit 
				'*********************************************************************************************************************************************************************
				%>
				<li><a href="#"><i class="fa fa-fw fa-plus" data-toggle="tooltip" title="Add / Edit"></i> Add / Edit</a>
					<ul class="list-unstyled">
						<li class="sub-nav"><a href="<%= BaseURL %>service/filters/addeditfilters/main.asp"> Add/Edit Filters<span class="icon"></span></a></li>
						<li class="sub-nav"><a href="<%= BaseURL %>accountsreceivable/regions/main.asp"> Add/Edit Regions<span class="icon"></span></a></li>
						<li class="sub-nav"><a href="<%= BaseURL %>accountsreceivable/customermininfo/main.asp"> Add/Edit Customers<span class="icon"></span></a></li>
						<li class="sub-nav"><a href="<%= BaseURL %>accountsreceivable/customermininfo/contactTitles/main.asp"> Add/Edit Contact Titles<span class="icon"></span></a></li>
					</ul>
			    </li>

			<% End If 'Filtertrax flag %>
			     			
		 	<li><a href="<%= BaseURL %>logout.asp"><i class="fa fa-sign-out" data-toggle="tooltip" title="Sign Out"></i> Sign Out<span class="icon"></span></a></li>
	</ul>
      
  
  <!-- mplex logo !-->
  <div class="build-line">
	<% If MUV_READ("FILTERTRAX") <> "1" Then %>
		<a href="<%= BaseURL %>main/default.asp"><img src="<%= BaseURL %>img/general/logo.png" class="mplexlogo"></a>
	<% Else %>
		<a href="<%= BaseURL %>main/default.asp"><img src="<%= BaseURL %>clientfilesV/filtertrax/logo_small.png" class="mplexlogo"></a>
	<% End If%>
  </div>
  <!-- eof mplex logo !-->
  
  <!-- feed 
  <div class="feed-box">
  	<ul class="list-unstyled main-menu">
	 
		 <% 
		 
			 Function showFeed(url)
			 
				Set xmlObj = Server.CreateObject("MSXML2.FreeThreadedDOMDocument")
				xmlObj.async = False
				xmlObj.setProperty "ServerHTTPRequest", True
				xmlObj.Load(url)
				If xmlObj.parseError.errorCode <> 0 Then
					Response.Write "Sorry, newsfeed is unavailable"
				End If
				Set xmlList = xmlObj.getElementsByTagName("item")
				Set xmlObj = Nothing
				feedCounter = 0
				For Each xmlItem In xmlList
					feedCounter = feedCounter + 1
					If feedCounter < 5 Then
						feedTitleDate = xmlItem.childNodes(0).text 
						feedText = xmlItem.childNodes(5).text
						feedText = Replace(xmlItem.childNodes(5).text,"&#8250;","<br><br>&#8250;")
						Response.write("<li><a href='#'><i class='fa fa-fw fa-file-text-o'></i> " & feedTitleDate & " changes <span class='icon'></span></a><ul class='list-unstyled'><li class='sub-nav'>" & feedText & "</li></ul></li>")
					End If
				Next
				Set xmlList = Nothing
			End Function
			
			showFeed("http://help.mdsinsight.com/feed/?post_type=changelog") 
		
		 %>

    </ul>
   </div>
   <!-- eof feed !-->
  
</nav>  

<style type="text/css">
.mplexlogo{
	display:inline-block;
	padding:5px 10px 5px 10px;
	background: #fff;
	border-radius:5px;
}

.feed-box{
	padding:10px;
	color: #6F7D8C !important;
}

.feed-box .main-menu li{
  color: #6F7D8C !important;
  text-decoration: none;
  font-size:12px;	
  display:block;
  margin-left:-5px;
  margin-bottom:0px;
}

.feed-box .main-menu li a{
  margin-top:-10px;
  color: #6F7D8C !important;
}

.feed-box .main-menu li a:hover{
  color: #FFF !important;
  text-decoration: none;
}

.feed-box .main-menu li.sub-nav{
  color: #FFFF66 !important;
  font-size:12px;	
}

#disabled-item a{
	color:#6F7D8C;
}

</style> 
