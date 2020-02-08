<!--#include file="settings.asp"-->
<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<!--#include file="InsightFuncs_BizIntel.asp"-->
<!--#include file="InsightFuncs_Invoicing.asp"-->
<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************

'Sub DoNotShowWebFulfillmentOrder()
'Sub ShowWebFulfillmentOrder()
'Sub GetContentForWebOrderRemarksModal()
'Sub EditWebOrderRemarksFromModal()
'Sub DeleteWebOrderRemarksFromModal()
'Sub GetContentForWebOrderInvoiceDetailModal()

'***************************************************
'End List of all the AJAX functions & subs
'***************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'ALL AJAX MODAL SUBROUTINES AND FUNCTIONS BELOW THIS AREA

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

action = Request("action")

Select Case action
	Case "DoNotShowWebFulfillmentOrder"
		DoNotShowWebFulfillmentOrder()
	Case "ShowWebFulfillmentOrder"
		ShowWebFulfillmentOrder()
	Case "GetContentForWebOrderRemarksModal"
		GetContentForWebOrderRemarksModal()
	Case "EditWebOrderRemarksFromModal"
		EditWebOrderRemarksFromModal()
	Case "DeleteWebOrderRemarksFromModal"
		DeleteWebOrderRemarksFromModal()
	Case "GetContentForWebOrderInvoiceDetailModal"
		GetContentForWebOrderInvoiceDetailModal()
End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub DoNotShowWebFulfillmentOrder()

	intRecID = Request.Form("InternalRecordIdentifier")
	
	Set rsWebFulfillRecordToHide = Server.CreateObject("ADODB.Recordset")
	rsWebFulfillRecordToHide.CursorLocation = 3 

	SQLWebFulfillRecordToHide = "SELECT * FROM IN_WebFulfillment WHERE InternalRecordIdentifier = " & intRecID	
	
	Set cnnWebFulfillRecordToHide = Server.CreateObject("ADODB.Connection")
	cnnWebFulfillRecordToHide.open (Session("ClientCnnString"))
	Set rsWebFulfillRecordToHide = cnnWebFulfillRecordToHide.Execute(SQLWebFulfillRecordToHide)
	
	If NOT rsWebFulfillRecordToHide.EOF Then
		OCSAccessOrderID = rsWebFulfillRecordToHide("OCSAccessOrderID")
		OCSAccessOrderDate = rsWebFulfillRecordToHide("OCSAccessOrderDate")
		CustID = rsWebFulfillRecordToHide("CustID")		
		CustName = GetCustNameByCustNum(CustID)	
	End If
	
	SQLWebFulfillRecordToHide = "UPDATE IN_WebFulfillment SET DontIncludeOnReport = 1 WHERE InternalRecordIdentifier = " & intRecID
	Set rsWebFulfillRecordToHide = cnnWebFulfillRecordToHide.Execute(SQLWebFulfillRecordToHide)
		
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " hid OCS Order #" & OCSAccessOrderID & " placed on " & FormatDateTime(OCSAccessOrderDate,1) & ", for the customer " & CustName & "(" & CustID & ")."	 			
	CreateAuditLogEntry "Web Fulfillment Order Hidden", "Web Fulfillment Order Hidden", "Minor", 1, Description		
	
	Set rsWebFulfillRecordToHide = Nothing
	cnnWebFulfillRecordToHide.Close
	Set cnnWebFulfillRecordToHide = Nothing
	
	Response.write("Success")
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub ShowWebFulfillmentOrder()

	intRecID = Request.Form("InternalRecordIdentifier")
	
	Set rsWebFulfillRecordToHide = Server.CreateObject("ADODB.Recordset")
	rsWebFulfillRecordToHide.CursorLocation = 3 

	SQLWebFulfillRecordToHide = "SELECT * FROM IN_WebFulfillment WHERE InternalRecordIdentifier = " & intRecID	
	
	Set cnnWebFulfillRecordToHide = Server.CreateObject("ADODB.Connection")
	cnnWebFulfillRecordToHide.open (Session("ClientCnnString"))
	Set rsWebFulfillRecordToHide = cnnWebFulfillRecordToHide.Execute(SQLWebFulfillRecordToHide)
	
	If NOT rsWebFulfillRecordToHide.EOF Then
		OCSAccessOrderID = rsWebFulfillRecordToHide("OCSAccessOrderID")
		OCSAccessOrderDate = rsWebFulfillRecordToHide("OCSAccessOrderDate")
		CustID = rsWebFulfillRecordToHide("CustID")		
		CustName = GetCustNameByCustNum(CustID)	
	End If
	
	SQLWebFulfillRecordToHide = "UPDATE IN_WebFulfillment SET DontIncludeOnReport = 0 WHERE InternalRecordIdentifier = " & intRecID
	Set rsWebFulfillRecordToHide = cnnWebFulfillRecordToHide.Execute(SQLWebFulfillRecordToHide)
		
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " un-hid OCS Order #" & OCSAccessOrderID & " placed on " & FormatDateTime(OCSAccessOrderDate,1) & ", for the customer " & CustName & "(" & CustID & ")."	 			
	CreateAuditLogEntry "Web Fulfillment Order Un-Hidden", "Web Fulfillment Order Un-Hidden", "Minor", 1, Description		
	
	Set rsWebFulfillRecordToHide = Nothing
	cnnWebFulfillRecordToHide.Close
	Set cnnWebFulfillRecordToHide = Nothing
	
	Response.write("Success")
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetContentForWebOrderRemarksModal() 

	InternalRecordIdentifier = Request.Form("InternalRecordIdentifier")
	
	'***************************************************************************************
	'Get values for editing an existing web order
	'***************************************************************************************

	SQLwebOrderRemarksEdit = "SELECT * FROM IN_WebFulfillment WHERE InternalRecordIdentifier = " & InternalRecordIdentifier
	
	Set cnnWebOrderRemarksEdit = Server.CreateObject("ADODB.Connection")
	cnnWebOrderRemarksEdit.open (Session("ClientCnnString"))
	Set rsWebOrderRemarksEdit = Server.CreateObject("ADODB.Recordset")
	rsWebOrderRemarksEdit.CursorLocation = 3 
	Set rsWebOrderRemarksEdit = cnnWebOrderRemarksEdit.Execute(SQLwebOrderRemarksEdit)
		
	If not rsWebOrderRemarksEdit.EOF Then
		OCSAccessOrderID = rsWebOrderRemarksEdit("OCSAccessOrderID")	
		OCSAccessOrderDate = rsWebOrderRemarksEdit("OCSAccessOrderDate")
		CustID = rsWebOrderRemarksEdit("CustID") 
		CustName = GetCustNameByCustNum(CustID)
		CustClassCode = rsWebOrderRemarksEdit("CustClassCode")
		Remarks = rsWebOrderRemarksEdit("Remarks")
	End If
	set rsWebOrderRemarksEdit = Nothing
	cnnWebOrderRemarksEdit.close
	set cnnWebOrderRemarksEdit = Nothing
	
	'***************************************************************************************
	
	
%>

	<input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
	<input type="hidden" name="txtCustID" id="txtCustID" value="<%= CustID %>">
	
	<!-- when line !-->
	<div class="row-line">

		<!-- when !-->
		<div class="col-lg-12">
			<p><strong>Customer</strong>: <%= CustName %> (<%= CustID %>)</p>
			<p><strong>Web Order Date</strong>: <%= FormatDateTime(OCSAccessOrderDate,2) %></p>
			<p><strong>Web Order ID</strong>: <%= OCSAccessOrderID %></p>
		</div>
		<!-- eof when !-->
    </div>
    <!-- eof when line !-->


	<!-- email alert line !-->
	<div class="row-line">

		<!-- email alert !-->
		<div class="col-lg-2" style="padding-top:20px">
			<label class="right">Remarks:</label>
		</div>
		<!-- eof email alert !-->

		<!-- multi select !-->
		<div class="col-lg-10" style="padding-top:20px">
			<% If Remarks <> "" Then %>
				<textarea rows="4" cols="50" name="txtWebOrderRemarks" id="txtWebOrderRemarks" class="form-control"><%= Remarks %></textarea>
			<% Else %>
				<textarea rows="4" cols="50" name="txtWebOrderRemarks" id="txtWebOrderRemarks" class="form-control"></textarea>
			<% End If %>
        </div>
		<!-- eof multi select !-->
    </div>
    <!-- eof email alert line !-->

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub EditWebOrderRemarksFromModal() 

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs8 = Server.CreateObject("ADODB.Recordset")
	rs8.CursorLocation = 3 
	
	InternalRecordIdentifier = Request.Form("InternalRecordIdentifier")
	NewWebOrderRemarks = Request.Form("WebOrderRemarks")
	NewWebOrderRemarks = Replace(NewWebOrderRemarks,"'","''")
	
	SQLwebOrderRemarksEdit = "SELECT * FROM IN_WebFulfillment WHERE InternalRecordIdentifier = " & InternalRecordIdentifier
	
	Set cnnWebOrderRemarksEdit = Server.CreateObject("ADODB.Connection")
	cnnWebOrderRemarksEdit.open (Session("ClientCnnString"))
	Set rsWebOrderRemarksEdit = Server.CreateObject("ADODB.Recordset")
	rsWebOrderRemarksEdit.CursorLocation = 3 
	Set rsWebOrderRemarksEdit = cnnWebOrderRemarksEdit.Execute(SQLwebOrderRemarksEdit)
		
	If not rsWebOrderRemarksEdit.EOF Then
		OCSAccessOrderID = rsWebOrderRemarksEdit("OCSAccessOrderID")	
		OCSAccessOrderDate = rsWebOrderRemarksEdit("OCSAccessOrderDate")
		CustID = rsWebOrderRemarksEdit("CustID") 
		CustName = GetCustNameByCustNum(CustID)
		Remarks = rsWebOrderRemarksEdit("Remarks")
	End If
	set rsWebOrderRemarksEdit = Nothing
	cnnWebOrderRemarksEdit.close
	set cnnWebOrderRemarksEdit = Nothing
	
	SQL = "UPDATE IN_WebFulfillment SET Remarks='" & NewWebOrderRemarks & "' WHERE InternalRecordIdentifier = " & InternalRecordIdentifier
	Set rs8 = cnn8.Execute(SQL)
	set rs8 = Nothing
		
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " changed the original remarks from, <em><strong>" & Remarks & "</strong></em>, to <em><strong>" & NewWebOrderRemarks & "</strong></em>,"
	Description = Description  & " for OCS Order #" & OCSAccessOrderID & " placed on " & FormatDateTime(OCSAccessOrderDate,1) & ", for the customer " & CustName & "(" & CustID & ")."	 			
	CreateAuditLogEntry "Web Fulfillment Order Remarks Edited", "Web Fulfillment Order Remarks Edited", "Major", 1, Description	

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub DeleteWebOrderRemarksFromModal() 

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs8 = Server.CreateObject("ADODB.Recordset")
	rs8.CursorLocation = 3 
	
	InternalRecordIdentifier = Request.Form("InternalRecordIdentifier")
	
	SQLWebOrderRemarksDelete = "SELECT * FROM IN_WebFulfillment WHERE InternalRecordIdentifier = " & InternalRecordIdentifier
	
	Set cnnWebOrderRemarksDelete = Server.CreateObject("ADODB.Connection")
	cnnWebOrderRemarksDelete.open (Session("ClientCnnString"))
	Set rsWebOrderRemarksDelete = Server.CreateObject("ADODB.Recordset")
	rsWebOrderRemarksDelete.CursorLocation = 3 
	Set rsWebOrderRemarksDelete = cnnWebOrderRemarksDelete.Execute(SQLWebOrderRemarksDelete)
		
	If not rsWebOrderRemarksDelete.EOF Then
		OCSAccessOrderID = rsWebOrderRemarksDelete("OCSAccessOrderID")	
		OCSAccessOrderDate = rsWebOrderRemarksDelete("OCSAccessOrderDate")
		CustID = rsWebOrderRemarksDelete("CustID") 
		CustName = GetCustNameByCustNum(CustID)
		Remarks = rsWebOrderRemarksDelete("Remarks")
	End If
	set rsWebOrderRemarksDelete = Nothing
	cnnWebOrderRemarksDelete.close
	set cnnWebOrderRemarksDelete = Nothing
	
	SQL = "UPDATE IN_WebFulfillment SET Remarks='' WHERE InternalRecordIdentifier = " & InternalRecordIdentifier
	Set rs8 = cnn8.Execute(SQL)
	set rs8 = Nothing
		
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " deleted the remarks, <em><strong>" & Remarks & "</strong></em>, for OCS Order #" & OCSAccessOrderID & " placed on " & FormatDateTime(OCSAccessOrderDate,1) & ", for the customer " & CustName & "(" & CustID & ")."	 			
	CreateAuditLogEntry "Web Fulfillment Order Remarks Deleted", "Web Fulfillment Order Remarks Deleted", "Major", 1, Description		
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetContentForWebOrderInvoiceDetailModal() 


	InternalRecordIdentifier = Request.Form("InternalRecordIdentifier")
	OCSAccessOrderID = Request.Form("OrderID")
	CustID = Request.Form("CustID")
	
	Set cnnInsight = Server.CreateObject("ADODB.Connection")
	
	'Response.Write("InsightCnnString: " & InsightCnnString & "<br>")
	
	cnnInsight.open (InsightCnnString)
	Set rsInsight = Server.CreateObject("ADODB.Recordset")
	rsInsight.CursorLocation = 3 
		
	SQLInsight = "SELECT * FROM tblServerInfo where clientKey='"& MUV_READ("CLIENTID") &"'"
	Set rsInsight = cnnInsight.Execute(SQLInsight)
	
	If NOT rsInsight.EOF Then
	
		OCSAccessCnnString = "Driver={SQL Server};Server=" & rsInsight.Fields("OCSAccess_dbServer")
		OCSAccessCnnString = OCSAccessCnnString & ";Database=" & rsInsight.Fields("OCSAccess_dbCatalog")
		OCSAccessCnnString = OCSAccessCnnString & ";Uid=" & rsInsight.Fields("OCSAccess_dbLogin")
		OCSAccessCnnString = OCSAccessCnnString & ";Pwd=" & rsInsight.Fields("OCSAccess_dbPassword") & ";"
		
	End If
	
	set rsInsight = Nothing
	cnnInsight.Close
	
	'Response.Write("OCSAccessCnnString: " & OCSAccessCnnString & "<br>")
	
	Set cnnOCSAccess = Server.CreateObject("ADODB.Connection")
	cnnOCSAccess.open (OCSAccessCnnString)
	Set rsOCSAccess = Server.CreateObject("ADODB.Recordset")
	rsOCSAccess.CursorLocation = 3 
		
	SQLOCSAccess = "SELECT * FROM tblOrders WHERE OrderID ='"& OCSAccessOrderID & "'"
	Set rsOCSAccess = cnnOCSAccess.Execute(SQLOCSAccess)
	
	If NOT rsOCSAccess.EOF Then
	
		total 		= rsOCSAccess("merchTotal")
		subtotal	= rsOCSAccess("merchSubtotal")
		tax 		= rsOCSAccess("Tax")
		freight 	= rsOCSAccess("Freight")
		OrderDate 	= rsOCSAccess("OrderDate")
		userPO		= rsOCSAccess("PO")
		Comments	= rsOCSAccess("Comments")

		ShipToName = rsOCSAccess("Name")
		ShipToAddress1 = rsOCSAccess("Address1")
		ShipToAddress2 = rsOCSAccess("Address2")
		ShipToCity = rsOCSAccess("City")
		ShipToState = rsOCSAccess("State")
		ShipToZip = rsOCSAccess("Zip")
		ShipToPhone = rsOCSAccess("Phone")
		ShipToCompany = rsOCSAccess("Company")
		
		BillToName = rsOCSAccess("billName") 
		BillToCompany = rsOCSAccess("billCompany") 
		BillToAddress1 = rsOCSAccess("billAddress1")
		BillToAddress2 = rsOCSAccess("billAddress2")
		BillToCity = rsOCSAccess("billCity")
		BillToState = rsOCSAccess("billState")
		BillToZip = rsOCSAccess("billState")
		BillToPhone = rsOCSAccess("billPhone")	
		
	End If
	
	
	If Total <> "" Then
		Total = formatCurrency(Total,2)
	End If
	
	If Tax <> "" Then
		Tax = formatCurrency(Tax,2)
	End If
	
	If Freight <> "" Then
		Freight = formatCurrency(Freight,2)
	End If
	
	If SubTotal <> "" Then
		SubTotal = formatCurrency(SubTotal,2)
	End If
	

	SQLOCSAccess = "SELECT tblUser.userNo, tblUser.custID, tblUser.userName, tblUser.userPhone,  "
	SQLOCSAccess = SQLOCSAccess & "tblUser.userFax, tblUser.userEmail, tblUser.userAddress1, tblUser.userAddress2, tblUser.userCity, tblUser.userState, tblUser.userState, "
	SQLOCSAccess = SQLOCSAccess & "tblUser.userZip, tblUser.userCompany, tblUser.userPO, tblUser.userDepartment, tblUser.userCostCenter, tblUser.userCCDescription, "
	SQLOCSAccess = SQLOCSAccess & "tblUser.userPassword, tblUser.userReceipt, tblCustomer.Name, tblCustomer.BlanketPONumber, "
	SQLOCSAccess = SQLOCSAccess & "tblCustomer.Address1, tblCustomer.Address2, tblCustomer.City, tblCustomer.State, tblCustomer.Zip, tblCustomer.Terms, "
	SQLOCSAccess = SQLOCSAccess & "tblCustomer.Contact1, tblCustomer.Contact2, tblCustomer.Phone, tblCustomer.Fax, tblCustomer.devDate "
	SQLOCSAccess = SQLOCSAccess & "FROM tblUser INNER JOIN tblCustomer ON tblUser.CustID = tblCustomer.CustID "
	SQLOCSAccess = SQLOCSAccess & "WHERE tblCustomer.CustID = '" & CustID & "'"
	
	Set rsOCSAccess = cnnOCSAccess.Execute(SQLOCSAccess)
	
	If NOT rsOCSAccess.EOF Then
	
		userEmail = rsOCSAccess("userEmail")
		userName = rsOCSAccess("userName")
	
	End If
	
	
	%>
	<div class="container-fluid container-modal">
	
	    <div class="row">
	        <div class="col-xs-12">
	    		<div class="invoice-title">
	    			<h2>Invoice Copy</h2><h3 class="pull-right">Order #<%= OCSAccessOrderID %></h3>
	    		</div>
	    		<hr>
	    		<div class="row">
	    			<div class="col-xs-6">
	    				<address>
	    				<strong>Billed To:</strong><br>
		                <strong><%= BillToCompany %></strong><br>
						<%= BillToName %><br>
						<%= BillToAddress1 %>&nbsp;
						<% If BillToAddress2 <> "" then 
							Response.write(BillToAddress2)
						   End If %>
						<br><%= BillToCity %>,&nbsp;<%= BillToState %>&nbsp;&nbsp;<%= BillToZip %><br>
						<%= BillToPhone %>
	    				</address>
	    			</div>
	    			<div class="col-xs-6 text-right">
	    				<address>
	        			<strong>Shipped To:</strong><br>
		                <strong><%= ShipToCompany %></strong><br>
						<%= ShipToName %><br>
						<%= ShipToAddress1 %>&nbsp;
						<% If ShipToAddress2 <> "" then 
							Response.write(ShipToAddress2)
						   End If %>
						<br><%= ShipToCity %>,&nbsp;<%= ShipToState %>&nbsp;&nbsp;<%= ShipToZip %><br>
						<%= ShipToPhone %>
	    				</address>
	    			</div>
	    		</div>
	    		<div class="row">
	    			<div class="col-xs-6">
	    				<address>
	    					<strong>PO Number:</strong><br>
	    					<%= userPO %><br>
	    					<strong>Comments:</strong><br>
	    					<%= Comments %>
	    				</address>
	    			</div>
	    			<div class="col-xs-6 text-right">
	    				<address>
	    					<strong>Order Date:</strong><br>
	    					<%= OrderDate %><br><br>
	    					<strong>Placed By:</strong><br>
	    					<%= userName %><br>
	    					<%= userEmail %>
	    					
	    				</address>
	    			</div>
	    		</div>
	    	</div>
	    </div>
	    
	    <div class="row">
	    	<div class="col-md-12">
	    		<div class="panel panel-default">
	    			<div class="panel-heading">
	    				<h3 class="panel-title"><strong>Order summary</strong></h3>
	    			</div>
	    			<div class="panel-body">
	    				<div class="table-responsive">
	    					<table class="table table-condensed">
	    						<thead>
	                                <tr>
	        							<td><strong>Item</strong></td>
	        							<td class="text-center"><strong>Price</strong></td>
	        							<td class="text-center"><strong>Quantity</strong></td>
	        							<td class="text-right"><strong>Totals</strong></td>
	                                </tr>
	    						</thead>
	    						<tbody>
	    						
	    						<%
	    										
								SQLOCSAccess ="SELECT tblProducts.prodSKU, tblProducts.prodShortDesc, tblProducts.ProdTaxable, "
								SQLOCSAccess = SQLOCSAccess & "tblProducts.prodOutOfStock, tblProducts.prodOutOfStockMessage, "
								SQLOCSAccess = SQLOCSAccess & "tblProducts.prodIndicators, tblProducts.prodDisplayOnWeb, "
								SQLOCSAccess = SQLOCSAccess & "tblProducts.prodUMdesc, tblOrderDetails.UMQty, tblOrderDetails.UM, tblOrderDetails.Qty, tblOrderDetails.SellPrice "
								SQLOCSAccess = SQLOCSAccess & "FROM tblProducts RIGHT OUTER JOIN tblOrderDetails ON tblProducts.ProdSKU = tblOrderDetails.ProdSKU "
								SQLOCSAccess = SQLOCSAccess & "WHERE tblOrderDetails.OrderID =" & OCSAccessOrderID & " ORDER BY tblOrderDetails.ProdSKU DESC"
								
								Set rsOCSAccess = cnnOCSAccess.Execute(SQLOCSAccess)
								
								If NOT rsOCSAccess.EOF Then
																
									Do While NOT rsOCSAccess.EOF
															
										SpecialOrderItem = False
										
										cProdSKU = rsOCSAccess("prodSKU")
										
										SQLWEBCONTENT = "SELECT * FROM tblProductWebContent WHERE prodSKU ='" & cProdSKU & "'"
										Set rsWebContent = Server.CreateObject("ADODB.Recordset")
										rsWebContent.Open SQLWEBCONTENT, OCSAccessCnnString
								
										If Not rsWebContent.EOF Then
										
											If  rsWebContent("prodShortDesc") <> "" Then
											
												cDesc = rsWebContent("prodShortDesc")
												
											    Set regEx= New RegExp
											
											    With regEx
											     .Pattern = "&#(\d+);" 'Match html unicode escapes
											     .Global = True
											    End With
											
											    Set matches = regEx.Execute(cDesc)
											
											    'Iterate over matches
											    For Each match in matches
											        'For each unicode match, replace the whole match, with the ChrW of the digits.
											
											        cDesc = Replace(cDesc, match.Value, ChrW(match.SubMatches(0)))
											    Next
											    
											    cDesc = Replace(cDesc, "&quot;", Chr(34))
											    cDesc = Replace(cDesc, "&lt;"  , Chr(60))
											    cDesc = Replace(cDesc, "&gt;"  , Chr(62))
											    cDesc = Replace(cDesc, "&amp;" , Chr(38))
											    cDesc = Replace(cDesc, "&nbsp;", Chr(32))
																							
											Else
												SQLWEBCONTENT2 = "SELECT * FROM tblProducts WHERE prodSKU ='" & cProdSKU & "'"
												Set rsWebContent2 = Server.CreateObject("ADODB.Recordset")
												rsWebContent2.Open SQLWEBCONTENT2, OCSAccessCnnString	
												If not rsWebContent2.EOF Then
													cDesc = rsWebContent2("prodShortDesc")
												End If
											End If
										Else
											SQLWEBCONTENT2 = "SELECT * FROM tblProducts WHERE prodSKU ='" & cProdSKU & "'"
											Set rsWebContent2 = Server.CreateObject("ADODB.Recordset")
											rsWebContent2.Open SQLWEBCONTENT2, OCSAccessCnnString	
											If not rsWebContent2.EOF Then
												cDesc = rsWebContent2("prodShortDesc")
											End If
										End If
										
										prodIndicators = rsOCSAccess("prodIndicators")
																
										If Instr(Ucase(prodIndicators),"SI:") <> 0 Then SpecialOrderItem = True
										
										If cDesc = "" or IsNull(cDesc) or IsEmpty(cDesc) Then
											cDesc = "Item No Longer Available or Under New SKU"
										End If
										
										cQty = rsOCSAccess("Qty")
										cProdTaxable = rsOCSAccess("prodTaxable")
										cUnitMeas = rsOCSAccess("UM")
										cUnitMeasQty = rsOCSAccess("UMqty")
						
										IF isNull(rsOCSAccess("SellPrice")) then 
											cPrice = 0
										ELSE
											cPrice = rsOCSAccess("SellPrice") 
										END IF	
										
										clineTotal = cQty * cPrice
										cSubTotal = formatCurrency((cSubTotal + clineTotal),2)
										If cPrice <> "" Then cPrice = formatCurrency(cPrice,2)				
										clineTotal = formatCurrency(clineTotal,2)
									
									%>
		    						
		    							<!-- foreach ($order->lineItems as $line) or some such thing here -->
		    							<tr>
		    								<td><strong><%= cProdSKU %></strong>&nbsp;&nbsp;<%= cDesc %>&nbsp;&nbsp;[<%= cUnitMeasQty %>/<%= cUnitMeas %>]
												<% If SpecialOrderItem = True Then %>
													<br>*Special Order Item*
												<% End If %>
		    								</td>
		    								<td class="text-center"><%= cPrice %></td>
		    								<td class="text-center"><%= cQty %></td>
		    								<td class="text-right"><%= cLineTotal %></td>
		    							</tr>
	
									<%
									rsOCSAccess.MoveNext
									Loop
									
								End If
								
								
								%>
	    							<tr>
	    								<td class="thick-line"></td>
	    								<td class="thick-line"></td>
	    								<td class="thick-line text-center"><strong>Subtotal</strong></td>
	    								<td class="thick-line text-right"><%= SubTotal %></td>
	    							</tr>
	    							<tr>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line text-center"><strong>Shipping</strong></td>
	    								<td class="no-line text-right"><%= freight %></td>
	    							</tr>
	    							<tr>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line text-center"><strong>Tax</strong></td>
	    								<td class="no-line text-right"><%= tax %></td>
	    							</tr>
	    							
	    							<tr>
	    								<td class="no-line"></td>
	    								<td class="no-line"></td>
	    								<td class="no-line text-center"><strong>Total</strong></td>
	    								<td class="no-line text-right"><%= total %></td>
	    							</tr>
	    						</tbody>
	    					</table>
	    				</div>
	    			</div>
	    		</div>
	    	</div>
	    </div>
	</div>
	<%
End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'END ALL AJAX MODAL SUBROUTINES AND FUNCTIONS

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

%>