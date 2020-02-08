<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/settings.asp"-->
<!--#include file="../../inc/mail.asp"-->
<!--#include file="../../inc/InsightFuncs_Orders.asp"-->
<!--#include file="../../inc/InsightFuncs_API.asp"-->
<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
OKtoPost = True

If InternalRecordIdentifier = "" Then OKtoPost = False

If OKtoPost = True Then
	'Lookup the Order
	Set cnnAPIOrderHeader = Server.CreateObject("ADODB.Connection")
	cnnAPIOrderHeader.open (Session("ClientCnnString"))
	Set rsAPIOrderHeader  = Server.CreateObject("ADODB.Recordset")
	
	SQLHeader = "SELECT * FROM API_OR_OrderHeader WHERE InternalRecordIdentifier = " & InternalRecordIdentifier 
	
	Set rsAPIOrderHeader = cnnAPIOrderHeader.Execute(SQLHeader)
	
	If rsAPIOrderHeader.EOF Then
		OKtoPost = False
	Else
		'Grab the fields needed for the XML

		APIKey = rsAPIOrderHeader("APIKey") 
		OrderID = rsAPIOrderHeader("OrderID") 
		OrderDate = rsAPIOrderHeader("OrderDate") 
		OrderThread = rsAPIOrderHeader("OrderThread") 
		BaseOrderID = rsAPIOrderHeader("BaseOrderID")
		RequestedDeliveryDate = rsAPIOrderHeader("RequestedDeliveryDate")
		CustID = rsAPIOrderHeader("CustID") 
		BillToCompany = rsAPIOrderHeader("BillToCompany") 
		BillToAttention = rsAPIOrderHeader("BillToAttention") 
		BillToAddressLine1 = rsAPIOrderHeader("BillToAddressLine1")
		BillToAddressLine2 = rsAPIOrderHeader("BillToAddressLine2")
		BillToCity = rsAPIOrderHeader("BillToCity")
		BillToState = rsAPIOrderHeader("BillToState")
		BillToZip = rsAPIOrderHeader("BillToZip") 
		BillToPhone = rsAPIOrderHeader("BillToPhone")
		BillToEmail = rsAPIOrderHeader("BillToEmail")
		ShipToCompany = rsAPIOrderHeader("ShipToCompany")
		ShipToAttention = rsAPIOrderHeader("ShipToAttention")
		ShipToAddressLine1 = rsAPIOrderHeader("ShipToAddressLine1") 
		ShipToAddressLine2 = rsAPIOrderHeader("ShipToAddressLine2") 
		ShipToCity = rsAPIOrderHeader("ShipToCity")
		ShipToState = rsAPIOrderHeader("ShipToState") 
		ShipToZip = rsAPIOrderHeader("ShipToZip") 
		ShipToPhone = rsAPIOrderHeader("ShipToPhone") 
		ShipToEmail = rsAPIOrderHeader("ShipToEmail") 
		Route = rsAPIOrderHeader("Route") 
		SalesPerson1 = rsAPIOrderHeader("SalesPerson1")
		Department = rsAPIOrderHeader("Department") 
		CustomerPONumber = rsAPIOrderHeader("CustomerPONumber") 
		CostCenter = rsAPIOrderHeader("CostCenter") 
		ApprovedBy = rsAPIOrderHeader("ApprovedBy") 
		OrderPlaceByName = rsAPIOrderHeader("OrderPlaceByName") 
		OrderSubTotal = rsAPIOrderHeader("OrderSubTotal") 
		ShippingCharge = rsAPIOrderHeader("ShippingCharge")
		Tax = rsAPIOrderHeader("Tax") 
		DepositCharge = rsAPIOrderHeader("DepositCharge") 
		FuelSurcharge = rsAPIOrderHeader("FuelSurcharge") 
		CouponCharge = rsAPIOrderHeader("CouponCharge") 
		GrandTotal = rsAPIOrderHeader("GrandTotal") 
		TotalCost = rsAPIOrderHeader("TotalCost")
		Terms = rsAPIOrderHeader("Terms")
		WarehouseNotes = rsAPIOrderHeader("WarehouseNotes") 
		DriverNotes = rsAPIOrderHeader("DriverNotes") 
		Voided = rsAPIOrderHeader("Voided")
		VoidedDateTime	 = rsAPIOrderHeader("VoidedDateTime")
	
	End If
	
	Set rsAPIOrderHeader = Nothing
	cnnAPIOrderHeader.Close
	Set cnnAPIOrderHeader = Nothing

End If

If OKtoPost = True Then
		'Begin to construct the XML from the header info
	
		xmlData = "<DATASTREAM>"
		xmlData = "<DATASTREAM xmlns=''>"
		xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>" ' Uses our API key
		sURL = Request.ServerVariables("SERVER_NAME")
		If Instr(ucase(sURL),"DEV.") <> 0 Then MODE = "TEST" Else MODE = "LIVE"
		xmlData = xmlData & "<MODE>" & MODE & "</MODE>"
		xmlData = xmlData & "<RECORD_TYPE>ORDER</RECORD_TYPE>"
		xmlData = xmlData & "<RECORD_SUBTYPE>UPSERT</RECORD_SUBTYPE>"
		xmlData = xmlData & "<CLIENT_ID>" & MUV_READ("SERNO") & "</CLIENT_ID>"
		xmlData = xmlData & "<SERNO>" & MUV_READ("SERNO") & "</SERNO>"
	
		xmlData = xmlData & "<ORDER>"
		xmlData = xmlData & "<ORDER_HEADER>"
		xmlData = xmlData & "<ORDER_ID>" & OrderID & "</ORDER_ID>"
		xmlData = xmlData & "<BASE_ORDER_ID>" & BaseOrderID & "</BASE_ORDER_ID>"
		xmlData = xmlData & "<ORDER_DATE>" & OrderDate & "</ORDER_DATE>"
		xmlData = xmlData & "<DELIVERY_DATE>" & RequestedDeliveryDate & "</DELIVERY_DATE>"
		xmlData = xmlData & "<CUST_ID>" & CustID & "</CUST_ID>"
		xmlData = xmlData & "<NUM_DETAIL_LINES>" & NumberOfAPIOrderLines(OrderID,GetAPIOrderHighestThread(OrderID)) & "</NUM_DETAIL_LINES>"
		xmlData = xmlData & "<BILL_COMPANY_NAME>" & BillToCompany & "</BILL_COMPANY_NAME>"
		xmlData = xmlData & "<BILL_ADDR1>" & BillToAddressLine1 & "</BILL_ADDR1>"
		xmlData = xmlData & "<BILL_ADDR2>" & BillToAddressLine2 & "</BILL_ADDR2>"
		xmlData = xmlData & "<BILL_CITY>" & BillToCity & "</BILL_CITY>"
		xmlData = xmlData & "<BILL_STATE>" & BillToState & "</BILL_STATE>"
		xmlData = xmlData & "<BILL_ZIP>" & BillToZip & "</BILL_ZIP>"
		xmlData = xmlData & "<BILL_PHONE>" & BillToPhone & "</BILL_PHONE>"
		xmlData = xmlData & "<BILL_ATTN>" & BillToAttention  & "</BILL_ATTN>"
		xmlData = xmlData & "<BILL_EMAIL>" & BillToEmail & "</BILL_EMAIL>"
		xmlData = xmlData & "<SHIP_COMPANY_NAME>" & ShipToCompany  & "</SHIP_COMPANY_NAME>"
		xmlData = xmlData & "<SHIP_ADDR1>" & ShipToAddressLine1  & "</SHIP_ADDR1>"
		xmlData = xmlData & "<SHIP_ADDR2>" & ShipToAddressLine2  & "</SHIP_ADDR2>"
		xmlData = xmlData & "<SHIP_CITY>" & ShipToCity & "</SHIP_CITY>"
		xmlData = xmlData & "<SHIP_STATE>" & ShipToState & "</SHIP_STATE>"
		xmlData = xmlData & "<SHIP_ZIP>" & ShipToZip & "</SHIP_ZIP>"
		xmlData = xmlData & "<SHIP_PHONE>" & ShipToPhone  & "</SHIP_PHONE>"
		xmlData = xmlData & "<SHIP_ATTN>" & ShipToAttention  & "</SHIP_ATTN>"
		xmlData = xmlData & "<SHIP_EMAIL>" & ShipToEmail  & "</SHIP_EMAIL>"
		xmlData = xmlData & "<ROUTE>" &  Route & "</ROUTE>"
		xmlData = xmlData & "<SALESPER1>" & SalesPerson1  & "</SALESPER1>"
		xmlData = xmlData & "<DEPT>" & Department & "</DEPT>"
		xmlData = xmlData & "<CUST_PO_NUM>" & CustomerPONumber  & "</CUST_PO_NUM>"
		xmlData = xmlData & "<COST_CENTER>" & CostCenter & "</COST_CENTER>"
		xmlData = xmlData & "<APPROVED_BY>" & ApprovedBy  & "</APPROVED_BY>"
		xmlData = xmlData & "<SUB_TOTAL>" & OrderSubTotal  & "</SUB_TOTAL>"
		xmlData = xmlData & "<PLACED_BY>" & OrderPlaceByName  & "</PLACED_BY>"
		xmlData = xmlData & "<SHIPPING_CHARGE>" & ShippingCharge  & "</SHIPPING_CHARGE>"
		xmlData = xmlData & "<TOTAL_TAX>" & Tax  & "</TOTAL_TAX>"
		xmlData = xmlData & "<DEPOSIT_CHARGE>" & DepositCharge & "</DEPOSIT_CHARGE>"
		xmlData = xmlData & "<FUEL_SURCHARGE>" & FuelSurcharge & "</FUEL_SURCHARGE>"
		xmlData = xmlData & "<COUPON_CHARGE>" & CouponCharge & "</COUPON_CHARGE>"
		xmlData = xmlData & "<GRAND_TOTAL>" & GrandTotal  & "</GRAND_TOTAL>"
		xmlData = xmlData & "<TOTAL_COST>" & TotalCost  & "</TOTAL_COST>"
		xmlData = xmlData & "<TERMS>" & Terms  & "</TERMS>"
		xmlData = xmlData & "<DRIVER_NOTES>" & DriverNotes  & "</DRIVER_NOTES>"
		xmlData = xmlData & "<WH_NOTES>" & WarehouseNotes  & "</WH_NOTES>"
		xmlData = xmlData & "</ORDER_HEADER>"
End If

If OKtoPost = True Then

	'Now Do the line items
	
	Set cnnAPIOrderDetails = Server.CreateObject("ADODB.Connection")
	cnnAPIOrderDetails.open (Session("ClientCnnString"))
	Set rsAPIOrderDetails  = Server.CreateObject("ADODB.Recordset")
	
	SQLDetails = "SELECT * FROM API_OR_OrderDetail WHERE OrderHeaderRecID = " & InternalRecordIdentifier & " ORDER BY OrderDetailID"
	
	Set rsAPIOrderDetails = cnnAPIOrderDetails.Execute(SQLDetails)
	
	If rsAPIOrderDetails.EOF Then
		OKtoPost = False
	Else
		'Grab the fields needed for the XML
		xmlData = xmlData & "<ORDER_DETAILS>"
		
		Do
		
			xmlData = xmlData & "<DETAIL_LINE>"
			xmlData = xmlData & "<DETAIL_NUM>" & rsAPIOrderDetails("OrderDetailID") & "</DETAIL_NUM>"
			xmlData = xmlData & "<PROD_ID>" & rsAPIOrderDetails("prodSKU") & "</PROD_ID>"
			xmlData = xmlData & "<DESCRIPT>" & rsAPIOrderDetails("prodDescription") & "</DESCRIPT>"
			xmlData = xmlData & "<QTY_ORD>" & rsAPIOrderDetails("QtyOrd") & "</QTY_ORD>"
			xmlData = xmlData & "<UOM>" & rsAPIOrderDetails("prodUM") & "</UOM>"
			xmlData = xmlData & "<PROD_COST>" & rsAPIOrderDetails("Cost") & "</PROD_COST>"
			xmlData = xmlData & "<SELL_PRICE>" & rsAPIOrderDetails("SellPrice") & "</SELL_PRICE>"
			xmlData = xmlData & "<LINE_EXTENSION>" & rsAPIOrderDetails("LineExtension") & "</LINE_EXTENSION>"
			xmlData = xmlData & "<DEPOSIT_AMT>" & rsAPIOrderDetails("DeportAmount") & "</DEPOSIT_AMT>"
			xmlData = xmlData & "<TAXABLE_FLAG>" & rsAPIOrderDetails("Taxable") & "</TAXABLE_FLAG>"
			xmlData = xmlData & "<TAXABLE_PERCENT>" & rsAPIOrderDetails("TaxPercent") & "</TAXABLE_PERCENT>"
			xmlData = xmlData & "<DROP_SHIP>" & rsAPIOrderDetails("DropShip") & "</DROP_SHIP>"
			xmlData = xmlData & "</DETAIL_LINE>"

		
			rsAPIOrderDetails.MoveNext
		
		Loop While Not rsAPIOrderDetails.EOF
		
		xmlData = xmlData & "</ORDER_DETAILS>"
		xmlData = xmlData & "</ORDER>"
		xmlData = xmlData & "</DATASTREAM>"

	
	End If

	Set rsAPIOrderDetails = Nothing
	cnnAPIOrderDetails.Close
	Set cnnAPIOrderDetails = Nothing

End If

xmlDataForDisp = xmlData 
xmlDataForDisp = Replace(xmlDataForDisp,"     ","")
xmlDataForDisp = Replace(xmlDataForDisp,"    <","<")	
xmlDataForDisp = Replace(xmlDataForDisp,"   <","<")	
xmlDataForDisp = Replace(xmlDataForDisp,"  <","<")	
xmlDataForDisp = Replace(xmlDataForDisp," <","<")	
xmlDataForDisp = Replace(xmlDataForDisp,"<","[")
xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")
'Response.Write(GetAPIRepostURL()& "<br>")
'Response.Write(OKtoPost& "<br>")
'Response.Write(xmlDataForDisp)
'response.end


If OKtoPost= True Then

	xmlData = Replace(xmlData,"&","&amp;")
	
	'Got all the data, go ahead & post it
	Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
	httpRequest.Open "POST", GetAPIRepostURL(), False
	httpRequest.SetRequestHeader "Content-Type", "text/xml"
	httpRequest.Send xmlData

	data = xmlData


	If (Err.Number <> 0 ) Then
		emailbody="httpRequest.status returned " & httpRequest.status & " when Reposting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br>"
		emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
		emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br>"
		emailBody = emailBody & "POSTED DATA:" & data & "<br>"
		emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
	End If

	If httpRequest.status = 200 THEN 
	
		If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
	
			emailbody="httpRequest.status returned " & httpRequest.status & " when Reposting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br>"
			emailBody = emailBody & "POSTED DATA:" & data & "<br>"
			emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
			
		Else
			emailbody="httpRequest.status returned " & httpRequest.status & " when Reposting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br>"
			emailBody = emailBody & "POSTED DATA:" & data & "<br>"
			emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
				

		End If
	Else
	
			'FAILURE
			emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>ORDER and <RECORD_SUBTYPE>UPSERT"& "<br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetAPIRepostURL() & "<br>"
			emailBody = emailBody & "POSTED DATA:" & data & "<br>"
			emailBody = emailBody & "SERNO: " & MUV_READ("SERNO") & "<br>"
	
			
	End If

End If

Response.Redirect("main.asp")

'response.end
''Write audit trail for dispatch
''*******************************
'Description = GetUserDisplayNameByUserNo(UserToDispatch) & " was dispatched to service ticket number " & ServiceTicketNumber & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
'CreateAuditLogEntry "Service Ticket System","Dispatched","Minor",0,Description 
'


%>

 
