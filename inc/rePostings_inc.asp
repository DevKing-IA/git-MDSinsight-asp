<!--#include file="mail.asp"-->
<%
Sub rePostOrderToBackend1(passedIntRecID_or_OrderID,postType)

If UCASE(TRIM(postType)) = "DELETE" Then 

	If passedIntRecID_or_OrderID <> "" Then
	
		OrderID = passedIntRecID_or_OrderID
	
		'Construct xml fields 
		xmlData = "<DATASTREAM>"
		xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
		xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTORDERMODE") & "</MODE>"
		xmlData = xmlData & "<RECORD_TYPE>ORDER</RECORD_TYPE>"
		xmlData = xmlData & "<RECORD_SUBTYPE>DELETE</RECORD_SUBTYPE>"
		xmlData = xmlData & "<SERNO>" & SERNO & "</SERNO>"
		xmlData = xmlData & "<ORDER_ID>" & OrderID & "</ORDER_ID>"
		xmlData = xmlData & "</DATASTREAM>"

		xmlDataForDisp = Replace(xmlData,"<","[")
		xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
		xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")

		Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
		httpRequest.Open "POST", GetAPIRepostURL(), False
	'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		httpRequest.SetRequestHeader "Content-Type", "text/xml"
		
		xmlData = Replace(xmlData,"&","&amp;")
		xmlData = Replace(xmlData,chr(34),"")
		
		httpRequest.Send xmlData
	
		data = xmlData

		'Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

		If (Err.Number <> 0 ) Then
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>DELETE"& "<br><br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & SERNO & "<br>"
			SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Order Delete",emailBody, "Order API", "Order API"
		
			Description = emailBody 
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,MODE,"1071d","1071d","Order API"
		End If

		If httpRequest.status = 200 THEN 
		
			If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
		
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost Order Delete",emailBody, "Order API", "Order API"
				
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
				
			Else
				'FAILURE
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Order Delete",emailBody, "Order API", "Order API"
			
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
				
			End If
			
		Else
		
				'FAILURE
				emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Order Delete",emailBody, "Order API", "Order API"
			
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
	
		End If

	End If

End If

If UCASE(TRIM(postType)) = "UPSERT" Then 

	If passedIntRecID_or_OrderID <> "" Then

		IntRecID = passedIntRecID_or_OrderID
	
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
			
			xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTORDERMODE") & "</MODE>"
			
			xmlData = xmlData & "<RECORD_TYPE>ORDER</RECORD_TYPE>"
		
			xmlData = xmlData & "<RECORD_SUBTYPE>UPSERT</RECORD_SUBTYPE>"

			'xmlData = xmlData & "<CLIENT_ID>" & GetPOSTParams("SERNO") & "</CLIENT_ID>"
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
			xmlData = xmlData & "<SHIPPING_COST>" & FormatCurrency(Round(rsRepost("ShippingCost"),2),2) & "</SHIPPING_COST>"				
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
	
	
	
			Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
			httpRequest.Open "POST", GetAPIRepostURL(), False
		'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			httpRequest.SetRequestHeader "Content-Type", "text/xml"
			
			xmlData = Replace(xmlData,"&","&amp;")
			xmlData = Replace(xmlData,chr(34),"")			
			httpRequest.Send xmlData
		
			data = xmlData
		
			Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

			If (Err.Number <> 0 ) Then
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
				emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Order Upsert",emailBody, "Order API", "Order API"
			
				Description = emailBody 
				CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTORDERMODE"),"1071d","1071d","Order API"
			End If

			If httpRequest.status = 200 THEN 
			
				If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
			
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost Order Upsert",emailBody, "Order API", "Order API"
					
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
					
				Else
					'FAILURE
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Order Upsert",emailBody, "Order API", "Order API"
				
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
					
				End If
				
			Else
			
					'FAILURE
					emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Order Upsert",emailBody, "Order API", "Order API"
				
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTORDERMODE"),"rePostings.asp")
		
			End If

		End If
	
	End If
	
End If
	
End Sub

Sub rePostInvoiceToBackend (passedIntRecID_or_InvoiceID,postType)

If UCASE(TRIM(postType)) = "DELETE" Then 

	If passedIntRecID_or_InvoiceID <> "" Then
	
		InvoiceID = passedIntRecID_or_InvoiceID
	
		'Construct xml fields 
		xmlData = "<DATASTREAM>"
		xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
		xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTINVOICEMODE") & "</MODE>"
		xmlData = xmlData & "<RECORD_TYPE>INVOICE</RECORD_TYPE>"
		xmlData = xmlData & "<RECORD_SUBTYPE>DELETE</RECORD_SUBTYPE>"
		xmlData = xmlData & "<SERNO>" & SERNO & "</SERNO>"
		xmlData = xmlData & "<INVOICE_ID>" & InvoiceID & "</INVOICE_ID>"
		xmlData = xmlData & "</DATASTREAM>"

		xmlDataForDisp = Replace(xmlData,"<","[")
		xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
		xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")

		Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
		httpRequest.Open "POST", GetAPIRepostInvoicesURL(), False
	'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		httpRequest.SetRequestHeader "Content-Type", "text/xml"
		
		xmlData = Replace(xmlData,"&","&amp;")
		xmlData = Replace(xmlData,chr(34),"")		
		httpRequest.Send xmlData
	
		data = xmlData

		Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

		If (Err.Number <> 0 ) Then
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>DELETE"& "<br><br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetAPIRepostInvoicesURL() & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & SERNO & "<br>"
			SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Invoice Delete",emailBody, "Invoice API", "Invoice API"
		
			Description = emailBody 
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTINVOICEMODE"),"1071d","1071d","Order API"
		End If

		If httpRequest.status = 200 THEN 
		
			If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
		
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostInvoicesURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost Invoice Delete",emailBody, "Invoice API", "Invoice API"
				
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTINVOICEMODE"),"rePostings.asp")
				
			Else
				'FAILURE
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostInvoicesURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Invoice Delete",emailBody, "Invoice API", "Invoice API"
			
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTINVOICEMODE"),"rePostings.asp")
				
			End If
			
		Else
		
				'FAILURE
				emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostInvoicesURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Invoice Delete",emailBody, "Invoice API", "Invoice API"
			
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTINVOICEMODE"),"rePostings.asp")
	
		End If

	End If

End If

If UCASE(TRIM(postType)) = "UPSERT" Then 

	If passedIntRecID_or_InvoiceID <> "" Then

		IntRecID = passedIntRecID_or_InvoiceID
	
		Set cnnPostInvoiceToBackend = Server.CreateObject("ADODB.Connection")
		cnnPostInvoiceToBackend.open (Session("ClientCnnString"))

		Set rsRepost = Server.CreateObject("ADODB.Recordset")
		rsRepost.CursorLocation = 3 
		
		SQLrsRepost = "SELECT * FROM API_IN_InvoiceHeader WHERE InternalRecordIdentifier = " & IntRecID 

		Set rsRepost = cnnPostInvoiceToBackend.Execute(SQLrsRepost)

		If Not rsRepost.Eof Then

			'Construct xml fields based on record
			xmlData = "<DATASTREAM>"
			xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTINVOICEMODE") & "</MODE>"
			xmlData = xmlData & "<RECORD_TYPE>INVOICE</RECORD_TYPE>"
			xmlData = xmlData & "<RECORD_SUBTYPE>UPSERT</RECORD_SUBTYPE>"
			xmlData = xmlData & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
			xmlData = xmlData & "<INVOICE>"
			xmlData = xmlData & "<INVOICE_HEADER>"
			xmlData = xmlData & "<INVOICE_ID>" & rsRepost("InvoiceID") & "</INVOICE_ID>"
			xmlData = xmlData & "<ORDER_ID>" & rsRepost("InvoiceID") & "</ORDER_ID>"
			xmlData = xmlData & "<INVOICE_DATE>" & FormatDateTime(rsRepost("InvoiceDate"),2) & "</INVOICE_DATE>"		
			xmlData = xmlData & "<DELIVERED_DATE>" & FormatDateTime(rsRepost("DeliveredDate"),2) & "</DELIVERED_DATE>"		
			xmlData = xmlData & "<DELIVERED_TIME>" & FormatDateTime(rsRepost("DeliveredTime"),2) & "</DELIVERED_TIME>"	
			xmlData = xmlData & "<CUST_ID>" & rsRepost("CustId") & "</CUST_ID>"		
			xmlData = xmlData & "<NUM_DETAIL_LINES>" & Number_Of_Lines_Invoice(IntRecID) & "</NUM_DETAIL_LINES>"	
			xmlData = xmlData & "<SUB_TOTAL>" & FormatCurrency(Round(rsRepost("InvoiceSubTotal"),2),2) & "</SUB_TOTAL>"
			xmlData = xmlData & "<SHIPPING_CHARGE>" & FormatCurrency(Round(rsRepost("ShippingCharge"),2),2) & "</SHIPPING_CHARGE>"	
			xmlData = xmlData & "<SHIPPING_COST>" & FormatCurrency(Round(rsRepost("ShippingCost"),2),2) & "</SHIPPING_COST>"				
			xmlData = xmlData & "<TOTAL_TAX>" & FormatCurrency(Round(rsRepost("Tax"),2),2) & "</TOTAL_TAX>"	
			xmlData = xmlData & "<DEPOSIT_CHARGE>" & FormatCurrency(Round(rsRepost("DepositCharge"),2),2) & "</DEPOSIT_CHARGE>"	
			xmlData = xmlData & "<FUEL_SURCHARGE>" & FormatCurrency(Round(rsRepost("FuelSurcharge"),2),2) & "</FUEL_SURCHARGE>"	
			xmlData = xmlData & "<COUPON_CHARGE>" & FormatCurrency(Round(rsRepost("CouponCharge"),2),2) & "</COUPON_CHARGE>"	
			xmlData = xmlData & "<GRAND_TOTAL>" & FormatCurrency(Round(rsRepost("GrandTotal"),2),2) & "</GRAND_TOTAL>"	
			xmlData = xmlData & "<TOTAL_COST>" & FormatCurrency(Round(rsRepost("TotalCost"),2),2) & "</TOTAL_COST>"	
			xmlData = xmlData & "<INVOICE_STATUS>" & GetInvoiceReportStatus() & "</INVOICE_STATUS>"	
			xmlData = xmlData & "</INVOICE_HEADER>"
				
				
			' Open a recordset and get the details
			Set rsRepostDetails = Server.CreateObject("ADODB.Recordset")
			rsRepostDetails.CursorLocation = 3 
		
			SQLrsRepostDetails = "SELECT * FROM API_IN_InvoiceDetail WHERE InvoiceHeaderRecID = " & IntRecID & " ORDER BY InvoiceDetailID"

			Set rsRepostDetails = cnnPostInvoiceToBackend.Execute(SQLrsRepostDetails)


			xmlData = xmlData & "<INVOICE_DETAILS>"
		
	
			Do While NOT rsRepostDetails.EOF
			
				xmlData = xmlData & "<DETAIL_LINE>" 
				xmlData = xmlData & "<DETAIL_NUM>" & rsRepostDetails("InvoiceDetailID") & "</DETAIL_NUM>"
				xmlData = xmlData & "<PROD_ID>" & rsRepostDetails("prodSKU") & "</PROD_ID>"
				xmlData = xmlData & "<DESCRIPT>" & rsRepostDetails("prodDescription") & "</DESCRIPT>"
				xmlData = xmlData & "<QTY_ORD>" & rsRepostDetails("QtyOrd") & "</QTY_ORD>"
				xmlData = xmlData & "<QTY_SHIPPED>" & rsRepostDetails("QtyShipped") & "</QTY_SHIPPED>"
				xmlData = xmlData & "<QTY_BACKORD>" & rsRepostDetails("QtyBackOrd") & "</QTY_BACKORD>"
				xmlData = xmlData & "<QTY_CANCELLED>" & rsRepostDetails("QtyCancelled") & "</QTY_CANCELLED>"				
				xmlData = xmlData & "<UOM>" & rsRepostDetails("prodUM") & "</UOM>"
				xmlData = xmlData & "<PROD_COST>" & FormatCurrency(Round(rsRepostDetails("Cost"),2),2) & "</PROD_COST>"
				xmlData = xmlData & "<SELL_PRICE>" & FormatCurrency(Round(rsRepostDetails("SellPrice"),2),2) & "</SELL_PRICE>"
				xmlData = xmlData & "<LINE_EXTENSION>" & FormatCurrency(Round(rsRepostDetails("LineExtension"),2),2) & "</LINE_EXTENSION>"
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
	
			
			xmlData = xmlData & "</INVOICE_DETAILS>"
				
			xmlData = xmlData & "</INVOICE>"		
			xmlData = xmlData & "</DATASTREAM>"
			
			Set rs = Nothing
			cnnPostInvoiceToBackend.Close
			Set cnnPostInvoiceToBackend = Nothing
			
			xmlDataForDisp = Replace(xmlData,"<","[")
			xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
			xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
			xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
			xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")
	
	
			Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
			httpRequest.Open "POST", GetAPIRepostInvoicesURL(), False
		'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			httpRequest.SetRequestHeader "Content-Type", "text/xml"
			
			xmlData = Replace(xmlData,"&","&amp;")
			xmlData = Replace(xmlData,chr(34),"")			
			httpRequest.Send xmlData
		
			data = xmlData
		
			Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

			If (Err.Number <> 0 ) Then
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
				emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Invoice Upsert",emailBody, "Invoice API", "Invoice API"
			
				Description = emailBody 
				CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTINVOICEMODE"),"1071d","1071d","Order API"
			End If

			If httpRequest.status = 200 THEN 
			
				If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
			
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost Invoice Upsert",emailBody, "Invoice API", "Invoice API"
					
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTINVOICEMODE"),"rePostings.asp")
					
				Else
					'FAILURE
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Invoice Upsert",emailBody, "Invoice API", "Invoice API"
				
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTINVOICEMODE"),"rePostings.asp")
					
				End If
				
			Else
			
					'FAILURE
					emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVOICE and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Invoice Upsert",emailBody, "Invoice API", "Invoice API"
				
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTINVOICEMODE"),"rePostings.asp")
		
			End If

		End If
	
	End If
	
End If
	
End Sub

Sub rePostRAToBackend(passedIntRecID_or_RAID,postType)

If UCASE(TRIM(postType)) = "DELETE" Then 

	If passedIntRecID_or_RAID <> "" Then
	
		RAID = passedIntRecID_or_RAID
	
		'Construct xml fields 
		xmlData = "<DATASTREAM>"
		xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
		xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTRAMODE") & "</MODE>"
		xmlData = xmlData & "<RECORD_TYPE>RETURN_AUTHORIZATION</RECORD_TYPE>"
		xmlData = xmlData & "<RECORD_SUBTYPE>DELETE</RECORD_SUBTYPE>"
		xmlData = xmlData & "<SERNO>" & SERNO & "</SERNO>"
		xmlData = xmlData & "<RAID>" & RAID & "</RAID>"
		xmlData = xmlData & "</DATASTREAM>"

		xmlDataForDisp = Replace(xmlData,"<","[")
		xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
		xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")

		Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
		httpRequest.Open "POST", GetAPIRepostRAURL(), False
	'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		httpRequest.SetRequestHeader "Content-Type", "text/xml"
		
		xmlData = Replace(xmlData,"&","&amp;")
		xmlData = Replace(xmlData,chr(34),"")		
		httpRequest.Send xmlData
	
		data = xmlData

		Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

		If (Err.Number <> 0 ) Then
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN AUTHORIZATION and <RECORD_SUBTYPE>DELETE"& "<br><br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetAPIRepostRAURL() & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & SERNO & "<br>"
			SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error RA Delete",emailBody, "RA API", "RA API"
		
			Description = emailBody 
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTRAMODE"),"1071d","1071d","Order API"
		End If

		If httpRequest.status = 200 THEN 
		
			If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
		
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN AUTHORIZATION and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostRAURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost RA Delete",emailBody, "RA API", "RA API"
				
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTRAMODE"),"rePostings.asp")
				
			Else
				'FAILURE
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN AUTHORIZATION and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostRAURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error RA Delete",emailBody, "RA API", "RA API"
			
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTRAMODE"),"rePostings.asp")
				
			End If
			
		Else
		
				'FAILURE
				emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN AUTHORIZATION and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostRAURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error RA Delete",emailBody, "RA API", "RA API"
			
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTRAMODE"),"rePostings.asp")
	
		End If

	End If

End If

If UCASE(TRIM(postType)) = "UPSERT" Then 

	If passedIntRecID_or_RAID <> "" Then

		IntRecID = passedIntRecID_or_RAID
	
		Set cnnPostRAToBackend = Server.CreateObject("ADODB.Connection")
		cnnPostRAToBackend.open (Session("ClientCnnString"))

		Set rsRepost = Server.CreateObject("ADODB.Recordset")
		rsRepost.CursorLocation = 3 
		
		SQLrsRepost = "SELECT * FROM API_OR_RAHeader WHERE InternalRecordIdentifier = " & IntRecID 

		Set rsRepost = cnnPostRAToBackend.Execute(SQLrsRepost)

		If Not rsRepost.Eof Then

			'Construct xml fields based on record
			xmlData = "<DATASTREAM>"
			xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			
			xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTRAMODE") & "</MODE>"
			
			xmlData = xmlData & "<RECORD_TYPE>RETURN_AUTHORIZATION</RECORD_TYPE>"
		
			xmlData = xmlData & "<RECORD_SUBTYPE>UPSERT</RECORD_SUBTYPE>"

			xmlData = xmlData & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
			
			xmlData = xmlData & "<RETURN_AUTHORIZATION>"
			xmlData = xmlData & "<RA_HEADER>"
			xmlData = xmlData & "<RA_ID>" & rsRepost("RAID") & "</RA_ID>"
			xmlData = xmlData & "<ORDER_ID>" & rsRepost("OrderID") & "</ORDER_ID>"
			xmlData = xmlData & "<RA_DATE>" & FormatDateTime(rsRepost("RDDate"),2) & "</RA_DATE>"		
			xmlData = xmlData & "<PICKUP_DATE>" & FormatDateTime(rsRepost("PickupDate"),2) & "</PICKUP_DATE>"			
			xmlData = xmlData & "<CUST_ID>" & rsRepost("CustId") & "</CUST_ID>"		
			xmlData = xmlData & "<NUM_DETAIL_LINES>" & Number_Of_Lines_RA(IntRecID) & "</NUM_DETAIL_LINES>"	
			xmlData = xmlData & "<SUB_TOTAL>" & FormatCurrency(Round(rsRepost("SubTotal"),2),2) & "</SUB_TOTAL>"
			xmlData = xmlData & "<PLACED_BY>" & rsRepost("PlaceByName") & "</PLACED_BY>"	
			xmlData = xmlData & "<SHIPPING_CHARGE>" & FormatCurrency(Round(rsRepost("ShippingCharge"),2),2) & "</SHIPPING_CHARGE>"	
			xmlData = xmlData & "<SHIPPING_COST>" & FormatCurrency(Round(rsRepost("ShippingCost"),2),2) & "</SHIPPING_COST>"				
			xmlData = xmlData & "<TOTAL_TAX>" & FormatCurrency(Round(rsRepost("Tax"),2),2) & "</TOTAL_TAX>"	
			xmlData = xmlData & "<DEPOSIT_CHARGE>" & FormatCurrency(Round(rsRepost("DepositCharge"),2),2) & "</DEPOSIT_CHARGE>"	
			xmlData = xmlData & "<FUEL_SURCHARGE>" & FormatCurrency(Round(rsRepost("FuelSurcharge"),2),2) & "</FUEL_SURCHARGE>"	
			xmlData = xmlData & "<COUPON_CHARGE>" & FormatCurrency(Round(rsRepost("CouponCharge"),2),2) & "</COUPON_CHARGE>"	
			xmlData = xmlData & "<GRAND_TOTAL>" & FormatCurrency(Round(rsRepost("GrandTotal"),2),2) & "</GRAND_TOTAL>"	
			xmlData = xmlData & "<TOTAL_COST>" & FormatCurrency(Round(rsRepost("TotalCost"),2),2) & "</TOTAL_COST>"
			xmlData = xmlData & "<DRIVER_NOTES>" & rsRepost("DriverNotes") & "</DRIVER_NOTES>"			
			xmlData = xmlData & "<WH_NOTES>" & rsRepost("WarehouseNotes") & "</WH_NOTES>"			
			xmlData = xmlData & "<RA_NOTES>" & rsRepost("RA_Notes") & "</RA_NOTES>"
			xmlData = xmlData & "</RA_HEADER>"
			
				
			' Open a recordset and get the details


			Set rsRepostDetails = Server.CreateObject("ADODB.Recordset")
			rsRepostDetails.CursorLocation = 3 
		
	
			SQLrsRepostDetails = "SELECT * FROM API_OR_RADetail WHERE RAHeaderRecID = " & IntRecID & " ORDER BY RADetailID"

			Set rsRepostDetails = cnnPostRAToBackend.Execute(SQLrsRepostDetails)


			xmlData = xmlData & "<RA_DETAILS>"
		
	
			Do While NOT rsRepostDetails.EOF
			
				xmlData = xmlData & "<DETAIL_LINE>" 
				xmlData = xmlData & "<DETAIL_NUM>" & rsRepostDetails("RADetailID") & "</DETAIL_NUM>"
				xmlData = xmlData & "<PROD_ID>" & rsRepostDetails("prodSKU") & "</PROD_ID>"
				xmlData = xmlData & "<DESCRIPT>" & rsRepostDetails("prodDescription") & "</DESCRIPT>"
				xmlData = xmlData & "<QTY_RET>" & rsRepostDetails("QtyRet") & "</QTY_RET>"
				xmlData = xmlData & "<UOM>" & rsRepostDetails("prodUM") & "</UOM>"
				xmlData = xmlData & "<PICKUP_REQUIRED>" & rsRepostDetails("PickupRequired") & "</PICKUP_REQUIRED>"
				xmlData = xmlData & "<RA_DETAIL_NOTE>" & rsRepostDetails("RADetailNote") & "</RA_DETAIL_NOTE>"				
				xmlData = xmlData & "</DETAIL_LINE>" 
				
				rsRepostDetails.MoveNext	
			Loop		
	
			
			xmlData = xmlData & "</RA_DETAILS>"
				
			xmlData = xmlData & "</RETURN_AUTHORIZATION>"		
			xmlData = xmlData & "</DATASTREAM>"
			
			Set rs = Nothing
			cnnPostRAToBackend.Close
			Set cnnPostRAToBackend = Nothing
		
			
			xmlDataForDisp = Replace(xmlData,"<","[")
			xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
			xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
			xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
			xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")
	
	
	
			Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
			httpRequest.Open "POST", GetAPIRepostRAURL(), False
		'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			httpRequest.SetRequestHeader "Content-Type", "text/xml"
			
			xmlData = Replace(xmlData,"&","&amp;")
			xmlData = Replace(xmlData,chr(34),"")			
			httpRequest.Send xmlData
		
			data = xmlData
		
			Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

			If (Err.Number <> 0 ) Then
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN_AUTHORIZATION and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
				emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostRAURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error RA Upsert",emailBody, "Order API", "Order API"
			
				Description = emailBody 
				CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTRAMODE"),"1071d","1071d","Order API"
			End If

			If httpRequest.status = 200 THEN 
			
				If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
			
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN_AUTHORIZATION and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostRAURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost RA Upsert",emailBody, "Order API", "Order API"
					
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTRAMODE"),"rePostings.asp")
					
				Else
					'FAILURE
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN_AUTHORIZATION and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostRAURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error RA Upsert",emailBody, "Order API", "Order API"
				
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTRAMODE"),"rePostings.asp")
					
				End If
				
			Else
			
					'FAILURE
					emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>RETURN_AUTHORIZATION and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostRAURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error RA Upsert",emailBody, "Order API", "Order API"
				
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTRAMODE"),"rePostings.asp")
		
			End If

		End If
	
	End If
	
End If
	
End Sub

Sub rePostCMToBackend(passedIntRecID_or_CMID,postType)

If UCASE(TRIM(postType)) = "DELETE" Then 

	If passedIntRecID_or_CMID <> "" Then
	
		CMID = passedIntRecID_or_CMID
	
		'Construct xml fields 
		xmlData = "<DATASTREAM>"
		xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
		xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTCMMODE") & "</MODE>"
		xmlData = xmlData & "<RECORD_TYPE>CREDIT_MEMO</RECORD_TYPE>"
		xmlData = xmlData & "<RECORD_SUBTYPE>DELETE</RECORD_SUBTYPE>"
		xmlData = xmlData & "<SERNO>" & SERNO & "</SERNO>"
		xmlData = xmlData & "<CMID>" & CMID & "</CMID>"
		xmlData = xmlData & "</DATASTREAM>"

		xmlDataForDisp = Replace(xmlData,"<","[")
		xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
		xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")

		Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
		httpRequest.Open "POST", GetAPIRepostCMURL(), False
	'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		httpRequest.SetRequestHeader "Content-Type", "text/xml"
		
		xmlData = Replace(xmlData,"&","&amp;")
		xmlData = Replace(xmlData,chr(34),"")		
		httpRequest.Send xmlData
	
		data = xmlData

		Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

		If (Err.Number <> 0 ) Then
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>CREDIT MEMO and <RECORD_SUBTYPE>DELETE"& "<br><br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetAPIRepostCMURL() & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & SERNO & "<br>"
			SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CM  Delete",emailBody, "CM API", "CM API"
		
			Description = emailBody 
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTCMMODE"),"1071d","1071d","Order API"
		End If

		If httpRequest.status = 200 THEN 
		
			If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
		
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>CREDIT MEMO and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostCMURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost CM Delete",emailBody, "CM API", "CM API"
				
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTCMMODE"),"rePostings.asp")
				
			Else
				'FAILURE
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>CREDIT MEMO and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostCMURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CM Delete",emailBody, "CM API", "CM API"
			
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTCMMODE"),"rePostings.asp")
				
			End If
			
		Else
		
				'FAILURE
				emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>CREDIT MEMO and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostCMURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CM Delete",emailBody, "CM API", "CM API"
			
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTCMMODE"),"rePostings.asp")
	
		End If

	End If

End If

If UCASE(TRIM(postType)) = "UPSERT" Then 

	If passedIntRecID_or_CMID <> "" Then

		IntRecID = passedIntRecID_or_CMID
	
		Set cnnPostCMToBackend = Server.CreateObject("ADODB.Connection")
		cnnPostCMToBackend.open (Session("ClientCnnString"))

		Set rsRepost = Server.CreateObject("ADODB.Recordset")
		rsRepost.CursorLocation = 3 
		
		SQLrsRepost = "SELECT * FROM API_IN_CMHeader WHERE InternalRecordIdentifier = " & IntRecID 

		Set rsRepost = cnnPostCMToBackend.Execute(SQLrsRepost)

		If Not rsRepost.Eof Then

			'Construct xml fields based on record
			xmlData = "<DATASTREAM>"
			xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			
			xmlData = xmlData & "<MODE>" & GetPOSTParams("CMAPIRepostMode") & "</MODE>"
			
			xmlData = xmlData & "<RECORD_TYPE>CREDIT_MEMO</RECORD_TYPE>"
		
			xmlData = xmlData & "<RECORD_SUBTYPE>UPSERT</RECORD_SUBTYPE>"

			xmlData = xmlData & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
			
			xmlData = xmlData & "<CREDIT_MEMO>"
			xmlData = xmlData & "<CM_HEADER>"
			xmlData = xmlData & "<CM_ID>" & rsRepost("CMID") & "</CM_ID>"
			'xmlData = xmlData & "<RA_ID>" & rsRepost("RAID") & "</RA_ID>"
			xmlData = xmlData & "<CM_DATE>" & FormatDateTime(rsRepost("CMDate"),2) & "</CM_DATE>"		
			xmlData = xmlData & "<PICKUP_DATE>" & FormatDateTime(rsRepost("PickupDate"),2) & "</PICKUP_DATE>"		
			'xmlData = xmlData & "<PICKUP_TIME>" & rsRepost("PickupTime") & "</PICKUP_TIME>"
			xmlData = xmlData & "<CUST_ID>" & rsRepost("CustId") & "</CUST_ID>"		
			xmlData = xmlData & "<NUM_DETAIL_LINES>" & Number_Of_Lines_CM(IntRecID) & "</NUM_DETAIL_LINES>"	
			xmlData = xmlData & "<SUB_TOTAL>" & FormatCurrency(Round(rsRepost("CMSubTotal"),2),2) & "</SUB_TOTAL>"
			xmlData = xmlData & "<SHIPPING_CHARGE>" & FormatCurrency(Round(rsRepost("ShippingCharge"),2),2) & "</SHIPPING_CHARGE>"	
			xmlData = xmlData & "<SHIPPING_COST>" & FormatCurrency(Round(rsRepost("ShippingCost"),2),2) & "</SHIPPING_COST>"				
			xmlData = xmlData & "<TOTAL_TAX>" & FormatCurrency(Round(rsRepost("Tax"),2),2) & "</TOTAL_TAX>"	
			xmlData = xmlData & "<DEPOSIT_CHARGE>" & FormatCurrency(Round(rsRepost("DepositCharge"),2),2) & "</DEPOSIT_CHARGE>"	
			xmlData = xmlData & "<FUEL_SURCHARGE>" & FormatCurrency(Round(rsRepost("FuelSurcharge"),2),2) & "</FUEL_SURCHARGE>"	
			xmlData = xmlData & "<COUPON_CHARGE>" & FormatCurrency(Round(rsRepost("CouponCharge"),2),2) & "</COUPON_CHARGE>"	
			xmlData = xmlData & "<GRAND_TOTAL>" & FormatCurrency(Round(rsRepost("GrandTotal"),2),2) & "</GRAND_TOTAL>"	
			xmlData = xmlData & "<TOTAL_COST>" & FormatCurrency(Round(rsRepost("TotalCost"),2),2) & "</TOTAL_COST>"
			xmlData = xmlData & "</CM_HEADER>"
			
				
			' Open a recordset and get the details


			Set rsRepostDetails = Server.CreateObject("ADODB.Recordset")
			rsRepostDetails.CursorLocation = 3 
		
	
			SQLrsRepostDetails = "SELECT * FROM API_IN_CMDetail WHERE CMHeaderRecID = " & IntRecID & " ORDER BY CMDetailID"

			Set rsRepostDetails = cnnPostCMToBackend.Execute(SQLrsRepostDetails)


			xmlData = xmlData & "<CM_DETAILS>"
		
	
			Do While NOT rsRepostDetails.EOF
			
				xmlData = xmlData & "<DETAIL_LINE>" 
				xmlData = xmlData & "<DETAIL_NUM>" & rsRepostDetails("CMDetailID") & "</DETAIL_NUM>"
				xmlData = xmlData & "<PROD_ID>" & rsRepostDetails("prodSKU") & "</PROD_ID>"
				xmlData = xmlData & "<DESCRIPT>" & rsRepostDetails("prodDescription") & "</DESCRIPT>"
				xmlData = xmlData & "<QTY_RET>" & rsRepostDetails("QtyReturned") & "</QTY_RET>"
				xmlData = xmlData & "<QTY_PICKED_UP>" & rsRepostDetails("QtyPickedUp") & "</QTY_PICKED_UP>"
				xmlData = xmlData & "<UOM>" & rsRepostDetails("prodUM") & "</UOM>"
				xmlData = xmlData & "<PROD_COST>" & FormatCurrency(Round(rsRepostDetails("Cost"),2),2) & "</PROD_COST>"
				xmlData = xmlData & "<SELL_PRICE>" & FormatCurrency(Round(rsRepostDetails("SellPrice"),2),2) & "</SELL_PRICE>"
				xmlData = xmlData & "<LINE_EXTENSION>" & FormatCurrency(Round(rsRepostDetails("LineExtension"),2),2) & "</LINE_EXTENSION>"
				If Not IsNull(rsRepostDetails("DeportAmount"))  Then
					xmlData = xmlData & "<DEPOSIT_AMT>" & FormatCurrency(Round(rsRepostDetails("DeportAmount"),2),2) & "</DEPOSIT_AMT>"
				Else
				 xmlData = xmlData & "<DEPOSIT_AMT>" & FormatCurrency(0,2) & "</DEPOSIT_AMT>"
				End If
				xmlData = xmlData & "<TAXABLE_FLAG>" & rsRepostDetails("Taxable") & "</TAXABLE_FLAG>"
				xmlData = xmlData & "<TAXABLE_PERCENT>" & rsRepostDetails("TaxPercent") & "</TAXABLE_PERCENT>"
				xmlData = xmlData & "</DETAIL_LINE>" 
				
				rsRepostDetails.MoveNext	
			Loop		
	
			
			xmlData = xmlData & "</CM_DETAILS>"
				
			xmlData = xmlData & "</CREDIT_MEMO>"		
			xmlData = xmlData & "</DATASTREAM>"
			
			Set rs = Nothing
			cnnPostCMToBackend.Close
			Set cnnPostCMToBackend = Nothing
		
			
			xmlDataForDisp = Replace(xmlData,"<","[")
			xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
			xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
			xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
			xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")
	
	
	
			Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
			httpRequest.Open "POST", GetAPIRepostCMURL(), False
		'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			httpRequest.SetRequestHeader "Content-Type", "text/xml"
			
			xmlData = Replace(xmlData,"&","&amp;")
			xmlData = Replace(xmlData,chr(34),"")			
			httpRequest.Send xmlData
		
			data = xmlData
		
			Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

			If (Err.Number <> 0 ) Then
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>CREDIT MEMO and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
				emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostCMURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CM  Upsert",emailBody, "Order API", "Order API"
			
				Description = emailBody 
				CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTCMMODE"),"1071d","1071d","Order API"
			End If

			If httpRequest.status = 200 THEN 
			
				If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
			
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>CREDIT MEMO and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostCMURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost CM Upsert",emailBody, "Order API", "Order API"
					
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTCMMODE"),"rePostings.asp")
					
				Else
					'FAILURE
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>CREDIT MEMO and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostCMURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CM Upsert",emailBody, "Order API", "Order API"
				
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTCMMODE"),"rePostings.asp")
					
				End If
				
			Else
			
					'FAILURE
					emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>CREDIT MEMO and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostCMURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error CM Upsert",emailBody, "Order API", "Order API"
				
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTCMMODE"),"rePostings.asp")
		
			End If

		End If
	
	End If
	
End If
	
End Sub

Sub rePostSummaryInvoiceToBackend (passedIntRecID_or_SumInvID,postType)

If UCASE(TRIM(postType)) = "DELETE" Then 

	If passedIntRecID_or_SumInvID <> "" Then

		SumInvID = passedIntRecID_or_SumInvID
	
		'Construct xml fields 
		xmlData = "<DATASTREAM>"
		xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
		xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTSUMINVMODE") & "</MODE>"
		xmlData = xmlData & "<RECORD_TYPE>SUMMARY_INVOICE</RECORD_TYPE>"
		xmlData = xmlData & "<RECORD_SUBTYPE>DELETE</RECORD_SUBTYPE>"
		xmlData = xmlData & "<SERNO>" & SERNO & "</SERNO>"
		xmlData = xmlData & "<SUM_INVOICE_ID>" & SumInvID & "</SUM_INVOICE_ID>"
		xmlData = xmlData & "</DATASTREAM>"

		xmlDataForDisp = Replace(xmlData,"<","[")
		xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
		xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
		xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")

		Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
		httpRequest.Open "POST", GetAPIRepostSumInvURL(), False
	'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		httpRequest.SetRequestHeader "Content-Type", "text/xml"
		
		xmlData = Replace(xmlData,"&","&amp;")
		xmlData = Replace(xmlData,chr(34),"")		
		httpRequest.Send xmlData
	
		data = xmlData

		Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

		If (Err.Number <> 0 ) Then
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY INVOICE and <RECORD_SUBTYPE>DELETE"& "<br><br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & SERNO & "<br>"
			SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Summary Inv Delete",emailBody, "CM API", "CM API"
		
			Description = emailBody 
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTSUMINVMODE"),"1071d","1071d","Order API"
		End If

		If httpRequest.status = 200 THEN 
		
			If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
		
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY INVOICE and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost Summary Inv Delete",emailBody, "CM API", "CM API"
				
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTSUMINVMODE"),"rePostings.asp")
				
			Else
				'FAILURE
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY INVOICE and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Summary Inv Delete",emailBody, "CM API", "CM API"
			
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTSUMINVMODE"),"rePostings.asp")
				
			End If
			
		Else
		
				'FAILURE
				emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY INVOICE and <RECORD_SUBTYPE>DELETE"& "<br><br>"
				emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Summary Inv Delete",emailBody, "CM API", "CM API"
			
				Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTSUMINVMODE"),"rePostings.asp")
	
		End If

	End If

End If

If UCASE(TRIM(postType)) = "UPSERT" Then 

	If passedIntRecID_or_SumInvID <> "" Then

		IntRecID = passedIntRecID_or_SumInvID
	
		Set cnnPostSumInvToBackend = Server.CreateObject("ADODB.Connection")
		cnnPostSumInvToBackend.open (Session("ClientCnnString"))

		Set rsRepost = Server.CreateObject("ADODB.Recordset")
		rsRepost.CursorLocation = 3 
		
		SQLrsRepost = "SELECT * FROM API_IN_SummaryInvoiceHeader WHERE InternalRecordIdentifier = " & IntRecID 

		Set rsRepost = cnnPostSumInvToBackend.Execute(SQLrsRepost)

		If Not rsRepost.Eof Then

			'Construct xml fields based on record
			xmlData = "<DATASTREAM>"
			xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			
			xmlData = xmlData & "<MODE>" & GetPOSTParams("REPOSTSUMINVMODE") & "</MODE>"
			
			xmlData = xmlData & "<RECORD_TYPE>SUMMARY_INVOICE</RECORD_TYPE>"
		
			xmlData = xmlData & "<RECORD_SUBTYPE>UPSERT</RECORD_SUBTYPE>"

			xmlData = xmlData & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
			
			xmlData = xmlData & "<INVOICE>"
			xmlData = xmlData & "<INVOICE_HEADER>"
			xmlData = xmlData & "<SUM_INVOICE_ID>" & rsRepost("SumInvID") & "</SUM_INVOICE_ID>"
			xmlData = xmlData & "<SUM_INVOICE_DATE>" & FormatDateTime(rsRepost("SumInvDate"),2) & "</SUM_INVOICE_DATE>"
			xmlData = xmlData & "<SUM_INVOICE_AGE_DATE>" & FormatDateTime(rsRepost("SumInvAgingDate"),2) & "</SUM_INVOICE_AGE_DATE>"
			xmlData = xmlData & "<CUST_ID>" & rsRepost("CustId") & "</CUST_ID>"		
			xmlData = xmlData & "<NUM_INVOICES>" & Number_Of_Lines_SumInv(IntRecID) & "</NUM_INVOICES>"			
			xmlData = xmlData & "<SUB_TOTAL>" & FormatCurrency(Round(rsRepost("Sub_Total"),2),2) & "</SUB_TOTAL>"
			xmlData = xmlData & "<SHIPPING_CHARGE>" & FormatCurrency(Round(rsRepost("Shipping_Charge"),2),2) & "</SHIPPING_CHARGE>"	
			xmlData = xmlData & "<SHIPPING_COST>" & FormatCurrency(Round(rsRepost("Shipping_Cost"),2),2) & "</SHIPPING_COST>"				
			xmlData = xmlData & "<TOTAL_TAX>" & FormatCurrency(Round(rsRepost("Total_Tax"),2),2) & "</TOTAL_TAX>"	
			xmlData = xmlData & "<DEPOSIT_CHARGE>" & FormatCurrency(Round(rsRepost("DepositCharge"),2),2) & "</DEPOSIT_CHARGE>"	
			xmlData = xmlData & "<FUEL_SURCHARGE>" & FormatCurrency(Round(rsRepost("FuelSurcharge"),2),2) & "</FUEL_SURCHARGE>"	
			xmlData = xmlData & "<COUPON_CHARGE>" & FormatCurrency(Round(rsRepost("CouponCharge"),2),2) & "</COUPON_CHARGE>"	
			xmlData = xmlData & "<GRAND_TOTAL>" & FormatCurrency(Round(rsRepost("Grand_Total"),2),2) & "</GRAND_TOTAL>"	
			xmlData = xmlData & "<TOTAL_COST>" & FormatCurrency(Round(rsRepost("Total_Cost"),2),2) & "</TOTAL_COST>"
			xmlData = xmlData & "</INVOICE_HEADER>"
			
				
			' Open a recordset and get the details


			Set rsRepostDetails = Server.CreateObject("ADODB.Recordset")
			rsRepostDetails.CursorLocation = 3 
		
	
			SQLrsRepostDetails = "SELECT * FROM API_IN_SummaryInvoiceDetail WHERE SumInvHeaderRecID = " & IntRecID & " ORDER BY Detail_Number"

			Set rsRepostDetails = cnnPostSumInvToBackend.Execute(SQLrsRepostDetails)


			xmlData = xmlData & "<INVOICE_DETAILS>"
		
	
			Do While NOT rsRepostDetails.EOF
			
				xmlData = xmlData & "<INVOICE_DETAIL_LINE>"
				
				xmlData = xmlData & "<DETAIL_NUM>" & rsRepostDetails("Detail_Number") & "</DETAIL_NUM>"
				xmlData = xmlData & "<INVOICE_ID>" & rsRepostDetails("InvoiceID") & "</INVOICE_ID>"
				xmlData = xmlData & "<CUST_ID>" & rsRepostDetails("CustID") & "</CUST_ID>"
				xmlData = xmlData & "</INVOICE_DETAIL_LINE>"
				
				rsRepostDetails.MoveNext	
			Loop		
	
			
			xmlData = xmlData & "</INVOICE_DETAILS>"
			xmlData = xmlData & "</INVOICE>"
			xmlData = xmlData & "</DATASTREAM>"
			
			Set rs = Nothing
			cnnPostSumInvToBackend.Close
			Set cnnPostSumInvToBackend = Nothing
		
			
			xmlDataForDisp = Replace(xmlData,"<","[")
			xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
			xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
			xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
			xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")
	
	
	
			Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
			httpRequest.Open "POST", GetAPIRepostSumInvURL(), False
		'	httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			httpRequest.SetRequestHeader "Content-Type", "text/xml"
			
			xmlData = Replace(xmlData,"&","&amp;")
			xmlData = Replace(xmlData,chr(34),"")			
			httpRequest.Send xmlData
		
			data = xmlData
		
			Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

			If (Err.Number <> 0 ) Then
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY INVOICE and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
				emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
				emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
				emailBody = emailBody & "SERNO: " & SERNO & "<br>"
				SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Summary Inv  Upsert",emailBody, "Invoice API", "Invoice API"
			
				Description = emailBody 
				CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("REPOSTSUMINVMODE"),"1071d","1071d","Order API"
			End If

			If httpRequest.status = 200 THEN 
			
				If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
			
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY INVOICE and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com", SERNO & " Good RePost Summary Inv Upsert",emailBody, "Invoice API", "Invoice API"
					
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTSUMINVMODE"),"rePostings.asp")
					
				Else
					'FAILURE
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY INVOICE and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Summary Inv Upsert",emailBody, "Invoice API", "Invoice API"
				
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody ,GetPOSTParams("REPOSTSUMINVMODE"),"rePostings.asp")
					
				End If
				
			Else
			
					'FAILURE
					emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SUMMARY INVOICE and <RECORD_SUBTYPE>UPSERT"& "<br><br>"
					emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetAPIRepostSumInvURL() & "<br><br>"
					emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
					emailBody = emailBody & "SERNO: " & SERNO & "<br>"
					SendMail "mailsender@" & maildomain ,"rsmith@ocsaccess.com",SERNO & " RePost Error Summary Inv Upsert",emailBody, "Invoice API", "Invoice API"
				
					Call CreateINSIGHTAuditLogEntry(Identity ,emailBody,GetPOSTParams("REPOSTSUMINVMODE"),"rePostings.asp")
		
			End If

		End If
	
	End If
	
End If
	
End Sub

%>