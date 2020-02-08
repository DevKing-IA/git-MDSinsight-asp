<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<!--#include file="InSightFuncs_Orders.asp"-->
<!--#include file="mail.asp"-->

<%
'***************************************************
'List of all the AJAX functions & subs
'***************************************************
'Sub rePostOrderToBackend()
'Sub returnsUMSForUnmappedProductCode()
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
	Case "rePostOrderToBackend"
		rePostOrderToBackend()
	Case "returnsUMSForUnmappedTaxableProductCode"
		returnsUMSForUnmappedTaxableProductCode()
	Case "returnsUMSForUnmappedNonTaxableProductCode"
		returnsUMSForUnmappedNonTaxableProductCode()
End Select




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub returnsUMSForUnmappedTaxableProductCode() 

	prodSKU = Request.Form("prodSKU")
	UM = Request.Form("UM")
	%>

	Unit of Measure: 
	<select class="C_Country_Modal form-control" id="txtUnmappedTaxableUM" name="txtUnmappedTaxableUM" style="width:50px;"> 
		<% 
		  	SQLProductsTableUM = "SELECT Distinct(prodCasePricing) FROM IC_Product WHERE prodSKU = '" & prodSKU & "'"
		
			Set cnnProductsTableUM = Server.CreateObject("ADODB.Connection")
			cnnProductsTableUM.open (Session("ClientCnnString"))
			Set rsProductsTableUM = Server.CreateObject("ADODB.Recordset")
			rsProductsTableUM.CursorLocation = 3 
			Set rsProductsTableUM = cnnProductsTableUM.Execute(SQLProductsTableUM)
				
			If not rsProductsTableUM.EOF Then
			
				If rsProductsTableUM("prodCasePricing") = "N" Then
					%><option value="N" selected="selected">N</option><%
				ElseIf rsProductsTableUM("prodCasePricing") = "U" Then
					%><option value="U" <% If UM = rsProductsTableUM("prodCasePricing") Then Response.Write("selected='selected'") %>>U</option><%
					%><option value="C" <% If UM = rsProductsTableUM("prodCasePricing") Then Response.Write("selected='selected'") %>>C</option><%
				ElseIf rsProductsTableUM("prodCasePricing") = "C" Then
					%><option value="U" <% If UM = rsProductsTableUM("prodCasePricing") Then Response.Write("selected='selected'") %>>U</option><%
					%><option value="C" <% If UM = rsProductsTableUM("prodCasePricing") Then Response.Write("selected='selected'") %>>C</option><%
				End If 
			Else
			%>
				<option value="U">U</option>
     			<option value="C">C</option>
     			<option value="N">N</option>
			<%
			End If
			
			set rsProductsTableUM = Nothing
			cnnProductsTableUM.close
			set cnnProductsTableUM = Nothing
		%>									
	</select>

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub returnsUMSForUnmappedNonTaxableProductCode() 

	prodSKU = Request.Form("prodSKU")
	%>

	Unit of Measure: 
	<select class="C_Country_Modal form-control" id="txtUnmappedNonTaxableUM" name="txtUnmappedNonTaxableUM" style="width:50px;"> 
		<% 
		  	SQLProductsTableUM = "SELECT Distinct(prodCasePricing) FROM IC_Product WHERE prodSKU = '" & prodSKU & "'"
		
			Set cnnProductsTableUM = Server.CreateObject("ADODB.Connection")
			cnnProductsTableUM.open (Session("ClientCnnString"))
			Set rsProductsTableUM = Server.CreateObject("ADODB.Recordset")
			rsProductsTableUM.CursorLocation = 3 
			Set rsProductsTableUM = cnnProductsTableUM.Execute(SQLProductsTableUM)
				
			If not rsProductsTableUM.EOF Then
			
				If rsProductsTableUM("prodCasePricing") = "N" Then
					%><option value="N">N</option><%
				ElseIf rsProductsTableUM("prodCasePricing") = "U" Then
					%><option value="U">U</option><%
					%><option value="C">C</option><%
				ElseIf rsProductsTableUM("prodCasePricing") = "C" Then
					%><option value="U">U</option><%
					%><option value="C">C</option><%
				End If 
			Else
			%>
				<option value="U">U</option>
     			<option value="C">C</option>
     			<option value="N">N</option>
			<%
			End If
			
			set rsProductsTableUM = Nothing
			cnnProductsTableUM.close
			set cnnProductsTableUM = Nothing
		%>									
	</select>

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub rePostOrderToBackend()
	
	If Request.Form("IntRecID") <> "" Then
		

		IntRecID = Request.Form("IntRecID") 
	
		Set cnnPostOrderToBackend = Server.CreateObject("ADODB.Connection")
		cnnPostOrderToBackend.open (Session("ClientCnnString"))

		Set rsRepost = Server.CreateObject("ADODB.Recordset")
		rsRepost.CursorLocation = 3 
		
		
		SQLrsRepost = "SELECT * FROM API_OR_OrderHeader WHERE InternalRecordIdentifier = " & IntRecID 

		Set rsRepost = cnnPostOrderToBackend.Execute(SQLrsRepost)

		If Not rsRepost.Eof Then

			'Construct xml fields based on record
			xmlData = "<DATASTREAM>"
			xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			
			xmlData = xmlData & "<MODE>" & GetPOSTParams("MODE") & "</MODE>"
			
			xmlData = xmlData & "<RECORD_TYPE>ORDER</RECORD_TYPE>"
		
			xmlData = xmlData & "<RECORD_SUBTYPE>UPSERT</RECORD_SUBTYPE>"

			xmlData = xmlData & "<CLIENT_ID>" & MUV_READ("ClientID") & "</CLIENT_ID>"
			xmlData = xmlData & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
			xmlData = xmlData & "<ORDER>"
				xmlData = xmlData & "<ORDER_HEADER>"
					xmlData = xmlData & "<ORDER_ID>" & rsRepost("OrderID") & "</ORDER_ID>"
					xmlData = xmlData & "<ORDER_DATE>" & FormatDateTime(rsRepost("OrderDate"),2) & "</ORDER_DATE>"		
					xmlData = xmlData & "<DELIVERY_DATE>" & FormatDateTime(rsRepost("RequestedDeliveryDate"),2) & "</DELIVERY_DATE>"			
					xmlData = xmlData & "<CUST_ID>" & rsRepost("CustId") & "</CUST_ID>"		
					
					xmlData = xmlData & "<NUM_DETAIL_LINES>" & Number_Of_Lines(IntRecID) & "</NUM_DETAIL_LINES>"	
			
					xmlData = xmlData & "<BILL_COMPANY_NAME>" & rsRepost("BillToCompany") & "</BILL_COMPANY_NAME>"		
					xmlData = xmlData & "<BILL_ADDR1>" & rsRepost("BillToAddressLine1") & "</BILL_ADDR1>"		
					xmlData = xmlData & "<BILL_ADDR2>" & rsRepost("BillToAddressLine2") & "</BILL_ADDR2>"		
					xmlData = xmlData & "<BILL_CITY>" & rsRepost("BillToCity") & "</BILL_CITY>"		
					xmlData = xmlData & "<BILL_STATE>" & rsRepost("BillToState") & "</BILL_STATE>"
					xmlData = xmlData & "<BILL_ZIP>" & rsRepost("BillToZip") & "</BILL_ZIP>"		
					xmlData = xmlData & "<BILL_PHONE>" & rsRepost("BillToPhone") & "</BILL_PHONE>"
					xmlData = xmlData & "<BILL_ATTN>" & rsRepost("BillToAttention") & "</BILL_ATTN>"
					xmlData = xmlData & "<BILL_EMAIL>" & rsRepost("BillToEmail") & "</BILL_EMAIL>"
			
					xmlData = xmlData & "<SHIP_COMPANY_NAME>" & rsRepost("ShipToCompany") & "</SHIP_COMPANY_NAME>"		
					xmlData = xmlData & "<SHIP_ADDR1>" & rsRepost("ShipToAddressLine1") & "</SHIP_ADDR1>"		
					xmlData = xmlData & "<SHIP_ADDR2>" & rsRepost("ShipToAddressLine2") & "</SHIP_ADDR2>"		
					xmlData = xmlData & "<SHIP_CITY>" & rsRepost("ShipToCity") & "</SHIP_CITY>"		
					xmlData = xmlData & "<SHIP_STATE>" & rsRepost("ShipToState") & "</SHIP_STATE>"
					xmlData = xmlData & "<SHIP_ZIP>" & rsRepost("ShipToZip") & "</SHIP_ZIP>"		
					xmlData = xmlData & "<SHIP_PHONE>" & rsRepost("ShipToPhone") & "</SHIP_PHONE>"
					xmlData = xmlData & "<SHIP_ATTN>" & rsRepost("ShipToAttention") & "</SHIP_ATTN>"
					xmlData = xmlData & "<SHIP_EMAIL>" & rsRepost("ShipToEmail") & "</SHIP_EMAIL>"

					xmlData = xmlData & "<ROUTE>" & rsRepost("Route") & "</ROUTE>"	
					xmlData = xmlData & "<SALESPER1>" & rsRepost("SalesPerson1") & "</SALESPER1>"	
					xmlData = xmlData & "<DEPT>" & rsRepost("Department") & "</DEPT>"	
					xmlData = xmlData & "<CUST_PO_NUM>" & rsRepost("CustomerPONumber") & "</CUST_PO_NUM>"	
					xmlData = xmlData & "<COST_CENTER>" & rsRepost("CostCenter") & "</COST_CENTER>"
					xmlData = xmlData & "<APPROVED_BY>" & rsRepost("ApprovedBy") & "</APPROVED_BY>"	
					xmlData = xmlData & "<SUB_TOTAL>" & FormatCurrency(Round(rsRepost("OrderSubTotal"),2),2) & "</SUB_TOTAL>"
					xmlData = xmlData & "<PLACED_BY>" & rsRepost("OrderPlaceByName") & "</PLACED_BY>"	
					xmlData = xmlData & "<SHIPPING_CHARGE>" & FormatCurrency(Round(rsRepost("ShippingCharge"),2),2) & "</SHIPPING_CHARGE>"	
					xmlData = xmlData & "<TOTAL_TAX>" & FormatCurrency(Round(rsRepost("Tax"),2),2) & "</TOTAL_TAX>"	
					
					xmlData = xmlData & "<DEPOSIT_CHARGE>" & FormatCurrency(Round(rsRepost("DepositCharge"),2),2) & "</DEPOSIT_CHARGE>"	
					xmlData = xmlData & "<FUEL_SURCHARGE>" & FormatCurrency(Round(rsRepost("FuelSurcharge"),2),2) & "</FUEL_SURCHARGE>"	
					xmlData = xmlData & "<COUPON_CHARGE>" & FormatCurrency(Round(rsRepost("CouponCharge"),2),2) & "</COUPON_CHARGE>"	


					xmlData = xmlData & "<GRAND_TOTAL>" & FormatCurrency(Round(rsRepost("GrandTotal"),2),2) & "</GRAND_TOTAL>"	
					xmlData = xmlData & "<TOTAL_COST>" & FormatCurrency(Round(rsRepost("TotalCost"),2),2) & "</TOTAL_COST>"	
					xmlData = xmlData & "<TERMS>" & rsRepost("Terms") & "</TERMS>"	
					xmlData = xmlData & "<DRIVER_NOTES>" & rsRepost("DriverNotes") & "</DRIVER_NOTES>"			
					xmlData = xmlData & "<WH_NOTES>" & rsRepost("WarehouseNotes") & "</WH_NOTES>"			
					
				xmlData = xmlData & "</ORDER_HEADER>"
				
				
				' Open a recordset and get the details


				Set rsRepostDetails = Server.CreateObject("ADODB.Recordset")
				rsRepostDetails.CursorLocation = 3 
			
		
				SQLrsRepostDetails = "SELECT * FROM API_OR_OrderDetail WHERE OrderHeaderRecID = " & IntRecID & " ORDER BY OrderDetailID"

				Set rsRepostDetails = cnnPostOrderToBackend.Execute(SQLrsRepostDetails)


				xmlData = xmlData & "<ORDER_DETAILS>"
		
	
				Do While NOT rsRepostDetails.EOF
				
					xmlData = xmlData & "<DETAIL_LINE>" 
					xmlData = xmlData & "<DETAIL_NUM>" & rsRepostDetails("OrderDetailID") & "</DETAIL_NUM>"
					xmlData = xmlData & "<PROD_ID>" & rsRepostDetails("prodSKU") & "</PROD_ID>"
					xmlData = xmlData & "<DESCRIPT>" & rsRepostDetails("prodDescription") & "</DESCRIPT>"
					xmlData = xmlData & "<QTY_ORD>" & rsRepostDetails("QtyOrd") & "</QTY_ORD>"
					xmlData = xmlData & "<UOM>" & rsRepostDetails("prodUM") & "</UOM>"
					xmlData = xmlData & "<PROD_COST>" & FormatCurrency(Round(rsRepostDetails("Cost"),2),2) & "</PROD_COST>"
					xmlData = xmlData & "<SELL_PRICE>" & FormatCurrency(Round(rsRepostDetails("SellPrice"),2),2) & "</SELL_PRICE>"
					xmlData = xmlData & "<LINE_EXTENSION>" & FormatCurrency(Round(rsRepostDetails("QtyOrd") * cDbl(rsRepostDetails("SellPrice")),2),2) & "</LINE_EXTENSION>"
					If Not IsNull(rsRepostDetails("DeportAmount"))  Then
						xmlData = xmlData & "<DEPOSIT_AMT>" & FormatCurrency(Round(rsRepostDetails("DeportAmount"),2),2) & "</DEPOSIT_AMT>"
					Else
					 xmlData = xmlData & "<DEPOSIT_AMT>" & FormatCurrency(0,2) & "</DEPOSIT_AMT>"
					End If
					xmlData = xmlData & "<TAXABLE_FLAG>" & rsRepostDetails("Taxable") & "</TAXABLE_FLAG>"
					xmlData = xmlData & "<TAXABLE_PERCENT>" & rsRepostDetails("TaxPercent") & "</TAXABLE_PERCENT>"
					xmlData = xmlData & "<DROP_SHIP>" & rsRepostDetails("DropShip") & "</DROP_SHIP>"
					xmlData = xmlData & "</DETAIL_LINE>" 
					
					rsRepostDetails.MoveNext	
				Loop		
	
			
		xmlData = xmlData & "</ORDER_DETAILS>"
		
	xmlData = xmlData & "</ORDER>"		
	xmlData = xmlData & "</DATASTREAM>"
	
	Set rs = Nothing
	cnnPostOrderToBackend.Close
	Set cnnPostOrderToBackend = Nothing

	
	xmlDataForDisp = Replace(xmlData,"<","[")
	xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
	xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
	xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
	xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")
	
	Response.write("xmlDataForDisp: " & xmlDataForDisp)
	Response.end
	
	
	Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
	httpRequest.Open "POST", "http://98.6.75.158:3291/ocsmds/ocsapi", False
'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	httpRequest.SetRequestHeader "Content-Type", "text/xml"
	
	xmlData = Replace(xmlData,"&","&amp;")
	httpRequest.Send xmlData

	data = xmlData

	Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

	If (Err.Number <> 0 ) Then
		emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br>"
		emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
		emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		emailBody = emailBody & "Posted to http://apidev.mdsinsight.com/apiIn/receive_order_xml.asp<br>"
		emailBody = emailBody & "POSTED DATA:" & data & "<br>"
		emailBody = emailBody & "SERNO: 1071d<br>"
		SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com","1071d" & " POST Error",emailBody, "Order API", "Order API"
	
		Description = emailBody 
		CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,"TEST","1071d","1071d","Order API"
response.end	
	End If

	
	If httpRequest.status = 200 THEN 
	
		If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
	
			Description ="success! httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br>"
			Description = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			Description = Description & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			Description = Description & "Posted to http://apidev.mdsinsight.com/apiIn/receive_order_xml.asp<br>"
			Description = Description & "POSTED DATA:" & data & "<br>"
			Description = Description & "SERNO: 1071d<br>"
			
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,"TEST","1071d","1071d","Order API"
			
			'****************************
			'TEMPORARY CODE TO BE REMOVED
			'****************************
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br>"
			emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			Description = Description & "Posted to http://apidev.mdsinsight.com/apiIn/receive_order_xml.asp<br>"
			Description = Description & "POSTED DATA:" & data & "<br>"
			Description = Description & "SERNO: 1071d<br>"
			SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com","1071d" & " GOOD POST",emailBody, "Order API", "Order API"
			'********************************
			'END TEMPORARY CODE TO BE REMOVED
			'********************************
response.end	
		Else
			'FAILURE
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			Description = Description & "Posted to http://apidev.mdsinsight.com/apiIn/receive_order_xml.asp<br>"
			Description = Description & "POSTED DATA:" & data & "<br>"
			Description = Description & "SERNO: 1071d<br>"
			SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com","1071d" & " POST Error",emailBody, "Order API", "Order API"
		
			Description = emailBody 
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,"TEST","1071d","1071d","Order API"
			
response.end
		End If
	Else
	
			'FAILURE
			emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			Description = Description & "Posted to http://apidev.mdsinsight.com/apiIn/receive_order_xml.asp<br>"
			Description = Description & "POSTED DATA:" & data & "<br>"
			Description = Description & "SERNO: 1071d<br>"
			SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com","1071d" & " POST Error",emailBody, "Order API", "Order API"
		
			Description = emailBody 
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,"TEST","1071d","1071d","Order API"
response.end			
	End If
	


End If


		End If

End Sub


Function Number_Of_Lines (passedOrderHeaderRecID)

	resultNumber_Of_Lines = ""
	
	Set cnnNumber_Of_Lines = Server.CreateObject("ADODB.Connection")
	cnnNumber_Of_Lines.open (Session("ClientCnnString"))
	
	Set rsNumber_Of_Lines = Server.CreateObject("ADODB.Recordset")
	rsNumber_Of_Lines.CursorLocation = 3 

	SQLNumber_Of_Lines = "SELECT COUNT(*) as Expr1 FROM API_OR_OrderDetail WHERE OrderHeaderRecID = " & passedOrderHeaderRecID

	Set rsNumber_Of_Lines = cnnNumber_Of_Lines.Execute(SQLNumber_Of_Lines)
	
	If Not rsNumber_Of_Lines.EOF Then resultNumber_Of_Lines = rsNumber_Of_Lines("Expr1")
	
	Set rsNumber_Of_Lines = Nothing
	cnnNumber_Of_Lines.Close
	Set cnnNumber_Of_Lines = Nothing

	Number_Of_Lines = resultNumber_Of_Lines

End Function

%>