<% @ Language = VBScript %>
<%
Response.Buffer = True
Response.Expires = 0
Response.Clear
%>
<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_API.asp"-->

<%
Dim PageNo, LineCount


dummy=MUV_Write("ClientID","") 'Need this here

ClientKey = Request.QueryString("c")

SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open InsightCnnString
'Response.Write("InsightCnnString:" & InsightCnnString & "<br>")

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3
'Response.Write("SQL:" & SQL& "<br>")

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - API Daily Activity Summary By Partner Report<%
	Response.End
Else
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & Recordset.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & Recordset.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	Recordset.close
	Connection.close	
End If	



'This is here so we only open it once for the whole page
Set cnn_Settings_Global = Server.CreateObject("ADODB.Connection")
cnn_Settings_Global.open (Session("ClientCnnString"))
Set rs_Settings_Global = Server.CreateObject("ADODB.Recordset")
rs_Settings_Global.CursorLocation = 3 
SQL_Settings_Global = "SELECT * FROM Settings_Global"
Set rs_Settings_Global = cnn_Settings_Global.Execute(SQL_Settings_Global)
If not rs_Settings_Global.EOF Then
	APIDailyActivityReportOnOff = rs_Settings_Global("APIDailyActivityReportOnOff")
	APIDailyActivityReportUserNos = rs_Settings_Global("APIDailyActivityReportUserNos")
	APIDailyActivityReportAdditionalEmails = rs_Settings_Global("APIDailyActivityReportAdditionalEmails")
	APIDailyActivityReportEmailSubject = rs_Settings_Global("APIDailyActivityReportEmailSubject")
	Order_OffSetFromToday = rs_Settings_Global("OrderAPIOffsetDays")
	Invoice_OffSetFromToday = rs_Settings_Global("InvoiceAPIOffsetDays")
	RA_OffSetFromToday = rs_Settings_Global("RAAPIOffsetDays")
	CM_OffSetFromToday = rs_Settings_Global("CMAPIOffsetDays")
	SumInv_OffSetFromToday = rs_Settings_Global("SumInvAPIOffsetDays")
Else
	APIDailyActivityReportOnOff = vbFalse
End If
Set rs_Settings_Global = Nothing
cnn_Settings_Global.Close
Set cnn_Settings_Global = Nothing

Order_currentDay = day(date()) - Order_OffSetFromToday 
Order_currentMonth = month(date())
Order_currentYear = year(date())

Invoice_currentDay = day(date()) - Invoice_OffSetFromToday
Invoice_currentMonth = month(date())
Invoice_currentYear = year(date())

RA_currentDay = day(date()) - RA_OffSetFromToday
RA_currentMonth = month(date())
RA_currentYear = year(date())

CM_currentDay = day(date()) - CM_OffSetFromToday
CM_currentMonth = month(date())
CM_currentYear = year(date())

SumInv_currentDay = day(date()) - SumInv_OffSetFromToday
SumInv_currentMonth = month(date())
SumInv_currentYear = year(date())

%>
<!DOCTYPE html>
<!--[if lt IE 7 ]> <html class="no-js ie6 oldie" lang="en"> <![endif]-->
<!--[if IE 7 ]>    <html class="no-js ie7 oldie" lang="en"> <![endif]-->
<!--[if IE 8 ]>    <html class="no-js ie8 oldie" lang="en"> <![endif]-->
<!--[if IE 9 ]>    <html class="no-js ie9" lang="en"> <![endif]-->
<!--[if (gte IE 9)|!(IE)]><![endif]--><!-->
<html class="no-js" lang="en">
<!--<![endif]-->
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta name="description" content="">
    <meta name="author" content="">

    <title>API Daily Activity Summary By Partner Report</title>

<%
    
Response.Write("<script src='https://use.fontawesome.com/3382135cdc.js'></script>")


Response.Write("<style type='text/css'>")
Response.Write("mark {")
Response.Write("    background-color: yellow;")
Response.Write("    color: black;")
Response.Write("}")
Response.Write("</style>")

Response.Write("<style type='text/css'>")
	
Response.Write("	body{font-family: arial, helvetica, sans-serif;}")
	
Response.Write("	div.table-title {")
Response.Write("	  display: block;")
Response.Write("	  margin: auto;")
Response.Write("	  max-width: 800px;")
Response.Write("	  padding:5px;")
Response.Write("	  width: 100%;")
Response.Write("	}")

	
Response.Write("	div.table-data {")
Response.Write("	  display: block;")
Response.Write("	  margin: auto;")
Response.Write("	  max-width: 1200px;")
Response.Write("	  padding:5px;")
Response.Write("	  width: 100%;")
Response.Write("	}")
	
Response.Write("	p, h1, h2 {")
Response.Write("	  display: block;")
Response.Write("	  margin: auto;")
Response.Write("	  max-width: 800px;")
Response.Write("	  padding:5px;")
Response.Write("	  width: 100%;")
Response.Write("	}")

Response.Write("	h1 {")
Response.Write("		color: #193048;")
Response.Write("	    font-size: 30px;")
Response.Write("	    font-weight: 400;")
Response.Write("	    font-style: normal;")
Response.Write("	    font-family: arial, helvetica, sans-serif;")
Response.Write("	    text-transform: uppercase;")
Response.Write("	    text-align:center;")
Response.Write("	}")
	
Response.Write("	.generated {")
Response.Write("		color: #3e94ec;")
Response.Write("	    font-size: 20px;")
Response.Write("	    font-weight: 400;")
Response.Write("	    font-style: normal;")
Response.Write("	    font-family: arial, helvetica, sans-serif;")
Response.Write("	    text-transform: uppercase;")
Response.Write("	    text-align:center;")
Response.Write("	}")
	
Response.Write("	h2 {")
Response.Write("		color: #3e94ec;")
Response.Write("	    font-size: 20px;")
Response.Write("	    font-weight: 400;")
Response.Write("	    font-style: normal;")
Response.Write("	    font-family: arial, helvetica, sans-serif;")
Response.Write("	    text-transform: uppercase;")
Response.Write("	    text-align:center;")
Response.Write("	}")
	
Response.Write("	hr{")
Response.Write("	   /* margin-top: 40px;")
Response.Write("	    margin-bottom: 40px;*/")
Response.Write("	}")
	
Response.Write("	.table-title h3 {")
Response.Write("	   color: #193048;")
Response.Write("	   font-size: 22px;")
Response.Write("	   font-weight: 400;")
Response.Write("	   font-style:normal;")
Response.Write("	   font-family: arial, helvetica, sans-serif;")
Response.Write("	   text-transform:uppercase;")
Response.Write("	   font-weight:bold;")
Response.Write("	}")
	
	
Response.Write("	/*** Table Styles **/")

Response.Write("	.table-fill {")
Response.Write("	  background: white;")
Response.Write("	  border-collapse: collapse;")
Response.Write("	  margin: auto;")
Response.Write("	  max-width: 800px;")
Response.Write(" 	  padding:5px;")
Response.Write("	  width: 100%;")
Response.Write("	  font-family: arial, helvetica, sans-serif;")
Response.Write("	}")
	 
Response.Write("	th {")
Response.Write("	   color:#483D8B;")
Response.Write("	  /*font-size:23px;*/")
Response.Write("	  font-size: 18px;")
Response.Write("	  font-weight: 100;")
Response.Write("	  padding:13px !important;")
Response.Write("	  text-align:left;")
Response.Write("	  vertical-align:middle;")
Response.Write("	  border: 1px solid #C1C3D1;")
Response.Write("	  width: 12.5% !important;")
Response.Write("  	}")

Response.Write("	tr {")
Response.Write("	  color:#666B85;")
Response.Write("	  font-size:16px;")
Response.Write("	  font-weight:normal;")
Response.Write("	}")
	 	 
Response.Write("	tr:nth-child(odd) td {")
Response.Write("	  background:#EBEBEB;")
Response.Write("	}")
	 	
	 
Response.Write("	td {")
Response.Write("	  background:#FFFFFF;")
Response.Write("	  padding:9px 13px 8px 20px !important;")
Response.Write("	  text-align:left;")
Response.Write("	  vertical-align:middle;")
Response.Write("	  font-weight:300;")
Response.Write("	  font-size:18px;")
Response.Write("	  border: 1px solid #C1C3D1;")
Response.Write("	}")
	
Response.Write("	/* custom table */")
	
	 
	
Response.Write("	.custom-table th{")
Response.Write("		padding:5px;")
Response.Write("	}")
	
Response.Write("	.custom-table td{")
Response.Write("		padding:5px;")
Response.Write("	}")
	
Response.Write("	#leftcol{")
Response.Write("		width:65%;")
Response.Write("	}")
	
Response.Write("	#rightcol{")
Response.Write("		width:35%;")
Response.Write("	}")
	
Response.Write("	#table-fill-short{")
Response.Write("		max-width: 500px;")
Response.Write("	}")
Response.Write("	/* eof custom table */")
	
Response.Write("	.cust-logo{")
Response.Write("		position: absolute;")
Response.Write("		margin-left: -280px;")
Response.Write("	}")


Response.Write("	</style>")
     
Response.Write("</head>")



Response.Write("<body bgcolor='#FFFFFF' text='#000000' link='#000080' topmargin='0' leftmargin='0' rightmargin='0' bottommargin='0' marginwidth='0' marginheight='0'>")
	 

Response.Write("<div class='table-title'>")

PageNo = 0
Call PageHeader 

Response.Write("<br>")
Response.Write("</div>")

Response.Write("<div class='table-data'>")


SQLDailyAPIPartnersLoop = "SELECT DISTINCT(partnerAPIKey) FROM IC_PARTNERS	"	

Set cnnDailyAPIPartnersLoop = Server.CreateObject("ADODB.Connection")
cnnDailyAPIPartnersLoop.open(Session("ClientCnnString"))
Set rsDailyAPIPartnersLoop = Server.CreateObject("ADODB.Recordset")
rsDailyAPIPartnersLoop.CursorLocation = 3 
Set rsDailyAPIPartnersLoop = cnnDailyAPIPartnersLoop.Execute(SQLDailyAPIPartnersLoop)

If NOT rsDailyAPIPartnersLoop.EOF Then

	rowCount = 1

	Do While Not rsDailyAPIPartnersLoop.EOF
	
		currentPartnerAPIKey = rsDailyAPIPartnersLoop("partnerAPIKey")
	

			Response.Write("<hr>")
			Response.Write("<h2>Partner: " & GetPartnerNameByAPIKey(currentPartnerAPIKey) & "</h2>")
			Response.Write("<hr>")
			Response.Write("<div><span style='background:yellow; display:inline-block; padding:3px;'>Yellow highlights on orders and invoices indicates a gap in sequence</span></div>")
				    	
		   	Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Orders with an order date of " & FormatDateTime(date() - Order_OffSetFromToday)  & "</h4><br>")
		    
		    
			Response.Write("<table style='margin-left:50px;width:1000px;'>")	
		        Response.Write("<thead>")
		            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
			            Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right'>Count</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right'>Order ID</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Subtotal</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Shipping</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Tax</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Fuel</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Deposit</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Coupon</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Grand Total</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Orig Date</th>")
   		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Adjstd Date</th>")
		            Response.Write("</tr>")
		        Response.Write("</thead>")
		        Response.Write("<tbody>")

				
				SQLDailyAPIOrders = "SELECT * FROM API_OR_OrderHeader "
				SQLDailyAPIOrders = SQLDailyAPIOrders & " WHERE (DAY(OrderDate) = " & Order_CurrentDay & " AND MONTH(OrderDate) =  " & Order_CurrentMonth & " AND YEAR(OrderDate) =  " & Order_CurrentYear & " AND "
				SQLDailyAPIOrders = SQLDailyAPIOrders & " (Voided = 0) AND (APIKey = '" & currentPartnerAPIKey & "')) "
				'SQLDailyAPIOrders = SQLDailyAPIOrders & " AND OrderID Not In (SELECT OrderID FROM API_IN_InvoiceHeader) "
				SQLDailyAPIOrders = SQLDailyAPIOrders & "  Order By BaseOrderID"
'	Response.Write(SQLDailyAPIOrders )
				
				Set cnnDailyAPIOrders = Server.CreateObject("ADODB.Connection")
				cnnDailyAPIOrders.open(Session("ClientCnnString"))
				Set rsDailyAPIOrders = Server.CreateObject("ADODB.Recordset")
				rsDailyAPIOrders.CursorLocation = 3 
				Set rsDailyAPIOrders = cnnDailyAPIOrders.Execute(SQLDailyAPIOrders)

				rowCount = 1
				DailyCount = 0
				DailySubtotal = 0
				DailyTaxTot = 0 
				DailyShipTot = 0 
				DailyFuelTot  = 0
				DailyDepositTot   = 0
				DailyCouponTot  = 0
				DailyGranTot  = 0
				LastCMID = 0 
				Subtotal = 0 
				TaxTot = 0
				ShipTot = 0
				FuelTot = 0
				DepositTot = 0
				CouponTot = 0
				GranTot = 0
				
				If NOT rsDailyAPIOrders.EOF Then
					
					Do While Not rsDailyAPIOrders.EOF
						
						OrderID = rsDailyAPIOrders("OrderID")
						DailyCount = DailyCount + 1
						
						Subtotal = rsDailyAPIOrders("OrderSubTotal")
						DailySubtotal = DailySubtotal + cdbl(Subtotal)
						
						TaxTot = rsDailyAPIOrders("Tax")
						DailyTaxTot = DailyTaxTot + cdbl(TaxTot)
						
						ShipTot = rsDailyAPIOrders("ShippingCharge")
						DailyShipTot = DailyShipTot + cdbl(ShipTot)
						
						FuelTot = rsDailyAPIOrders("FuelSurcharge")
						DailyFuelTot  = DailyFuelTot  + cdbl(FuelTot)
						
						DepositTot = rsDailyAPIOrders("DepositCharge")
						DailyDepositTot   = DailyDepositTot + cdbl(DepositTot)
						
						CouponTot = rsDailyAPIOrders("CouponCharge")
						DailyCouponTot  = DailyCouponTot + cdbl(CouponTot)
						
						GranTot = rsDailyAPIOrders("GrandTotal")
						DailyGranTot  = DailyGranTot + cdbl(GranTot)
						
						If Subtotal <> "" Then
							Subtotal = formatCurrency(Subtotal ,2)
						End If
						If TaxTot <> "" Then
							TaxTot = formatCurrency(TaxTot ,2)
						End If
						If ShipTot <> "" Then
							ShipTot = formatCurrency(ShipTot ,2)
						End If
						If FuelTot <> "" Then
							FuelTot = formatCurrency(FuelTot ,2)
						End If
						If DepositTot <> "" Then
							DepositTot = formatCurrency(DepositTot ,2)
						End If
						If CouponTot <> "" Then
							CouponTot = formatCurrency(CouponTot ,2)
						End If
						If GranTot <> "" Then
							GranTot = formatCurrency(GranTot ,2)
						End If
						

						Response.Write("<tr>")
						Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & DailyCount & "</td>")
							If LastOrderID <> 0 Then
								If LastOrderID + 1 <> CDbl(rsDailyAPIOrders("BaseOrderID")) Then
									 Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><span style='background:yellow; display:inline-block; padding:3px; color: #000;'>" & OrderID & "</span></td>")
								Else
									Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" &  OrderID & "</td>")
								End If
							Else 
								Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" &  OrderID & "</td>")
							End If
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & Subtotal & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & ShipTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & TaxTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & FuelTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & DepositTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & CouponTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & GranTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & rsDailyAPIOrders("Orig_OrderDate")& "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & rsDailyAPIOrders("OrderDate")& "</td>")
							
			            Response.Write("</tr>")

						rowCount = rowCount + 1
						
						LineCount = LineCount + 1
						
						LastOrderID = CDbl(rsDailyAPIOrders("BaseOrderID"))
						
						rsDailyAPIOrders.MoveNext
						
						If LineCount >= 21 Then 
					        Response.Write("</tbody>")
						    Response.Write("</table>")

							Call PageHeader

							Call SubHeader("Orders","Body")
							
							Response.Write("<table style='margin-left:50px;width:1000px;'>")	
					        Response.Write("<thead>")
				            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
				            Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right'>Count</th>")
			                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right'>Order ID</th>")
			                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Subtotal</th>")
			                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Shipping</th>")
			                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Tax</th>")
			                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Fuel</th>")
			                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Deposit</th>")
			                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Coupon</th>")
			                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Grand Total</th>")
			                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Orig Date</th>")
	   		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Adjstd Date</th>")
							Response.Write("</tr>")
							Response.Write("</thead>")
					        Response.Write("<tbody>")

						End if
						
					Loop
				Else
					Response.Write("<tr ><td colspan='8'>No Order API Data</td></tr>")
				End If

				If DailySubtotal <> "" Then
					DailySubtotal = formatCurrency(DailySubtotal ,2)
				End If
				If DailyTaxTot <> "" Then
					DailyTaxTot = formatCurrency(DailyTaxTot ,2)
				End If
				If DailyShipTot <> "" Then
					DailyShipTot = formatCurrency(DailyShipTot ,2)
				End If
				If DailyFuelTot  <> "" Then
					DailyFuelTot  = formatCurrency(DailyFuelTot  ,2)
				End If
				If DailyDepositTot   <> "" Then
					DailyDepositTot   = formatCurrency(DailyDepositTot   ,2)
				End If
				If DailyCouponTot  <> "" Then
					DailyCouponTot  = formatCurrency(DailyCouponTot  ,2)
				End If
				If DailyGranTot  <> "" Then
					DailyGranTot  = formatCurrency(DailyGranTot  ,2)
				End If
				
				Response.Write("<tr style='border-top: 2px solid #ddd;'>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'><strong>Count:&nbsp;&nbsp;" & DailyCount & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'>&nbsp;&nbsp;</td>")
	                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailySubtotal & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyShipTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyTaxTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyFuelTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyDepositTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyCouponTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyGranTot & "</strong></td>")
	            Response.Write("</tr>")
		        Response.Write("</tbody>")
		    Response.Write("</table>")

		    
		    
		    '*********************************
			'		    I N V O I C E S 
		    '*********************************
		    If LineCount + 5 >= 21 Then
				Call PageHeader
				Call SubHeader("Invoices","Top")
			End If

		    LineCount = LineCount + 5
		    	    
        	Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Invoices with an invoice date of " & FormatDateTime(Date() - Invoice_OffSetFromToday) & "</h4>")
						
			
			Response.Write("<table style='margin-left:50px;width:1000px;'>")	
		        Response.Write("<thead>")
		            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' >Invoice ID</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Subtotal</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Shipping</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Tax</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Fuel</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Deposit</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Coupon</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Grand Total</th>")
		            Response.Write("</tr>")
		        Response.Write("</thead>")
		        Response.Write("<tbody>")
		        


				SQLDailyAPIInvoices = "SELECT * FROM API_IN_InvoiceHeader "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " WHERE (DAY(InvoiceDate) = " & Invoice_CurrentDay & " AND MONTH(InvoiceDate) =  " & Invoice_CurrentMonth & " AND YEAR(InvoiceDate) =  "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & Invoice_CurrentYear & " AND (Voided = 0)  AND (APIKey = '" & currentPartnerAPIKey & "')) "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & "ORDER BY InvoiceID"

				Set cnnDailyAPIInvoices = Server.CreateObject("ADODB.Connection")
				cnnDailyAPIInvoices.open(Session("ClientCnnString"))
				Set rsDailyAPIInvoices = Server.CreateObject("ADODB.Recordset")
				rsDailyAPIInvoices.CursorLocation = 3 
				'Response.Write(SQLDailyAPIInvoices)
				Set rsDailyAPIInvoices = cnnDailyAPIInvoices.Execute(SQLDailyAPIInvoices)

				rowCount = 1
				DailyCount = 0
				DailySubtotal = 0
				DailyTaxTot = 0 
				DailyShipTot = 0 
				DailyFuelTot  = 0
				DailyDepositTot   = 0
				DailyCouponTot  = 0
				DailyGranTot  = 0
				LastCMID = 0 
				Subtotal = 0 
				TaxTot = 0
				ShipTot = 0
				FuelTot = 0
				DepositTot = 0
				CouponTot = 0
				GranTot = 0
				
				If NOT rsDailyAPIInvoices.EOF Then

					Do While Not rsDailyAPIInvoices.EOF
						
						InvoiceID = rsDailyAPIInvoices("InvoiceID")
						DailyCount = DailyCount + 1
						
						Subtotal = rsDailyAPIInvoices("InvoiceSubTotal")
						DailySubtotal = DailySubtotal + cdbl(Subtotal)

						TaxTot = rsDailyAPIInvoices("Tax")
						DailyTaxTot = DailyTaxTot + cdbl(TaxTot)
						
						ShipTot = rsDailyAPIInvoices("ShippingCharge")
						DailyShipTot = DailyShipTot + cdbl(ShipTot)

						FuelTot = rsDailyAPIInvoices("FuelSurcharge")
						DailyFuelTot  = DailyFuelTot  + cdbl(FuelTot)
						
						DepositTot = rsDailyAPIInvoices("DepositCharge")
						DailyDepositTot   = DailyDepositTot + cdbl(DepositTot)
						
						CouponTot = rsDailyAPIInvoices("CouponCharge")
						DailyCouponTot  = DailyCouponTot + cdbl(CouponTot)
						
						GranTot = rsDailyAPIInvoices("GrandTotal")
						DailyGranTot  = DailyGranTot + cdbl(GranTot)

						
						If Subtotal <> "" Then
							Subtotal = formatCurrency(Subtotal,2)
						End If
						If TaxTot <> "" Then
							TaxTot = formatCurrency(TaxTot,2)
						End If
						If ShipTot <> "" Then
							ShipTot = formatCurrency(ShipTot,2)
						End If
						If FuelTot <> "" Then
							FuelTot = formatCurrency(FuelTot,2)
						End If
						If DepositTot <> "" Then
							DepositTot = formatCurrency(DepositTot,2)
						End If
						If CouponTot <> "" Then
							CouponTot = formatCurrency(CouponTot,2)
						End If
						If GranTot <> "" Then
							GranTot = formatCurrency(GranTot,2)
						End If
						
						
						Response.Write("<tr>")
						If IsNumeric(LastInvoiceID) Then
							If LastInvoiceID <> 0 Then
								If IsNumeric(rsDailyAPIInvoices("InvoiceID")) Then
									If LastInvoiceID + 1 <> CDbl(rsDailyAPIInvoices("InvoiceID")) Then
										 Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><span style='background:yellow; display:inline-block; padding:3px; color: #000;'>" & InvoiceID & "</span></td>")
									Else
										Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & InvoiceID & "</td>")
									End If
								Else
									Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & InvoiceID & "</td>")
								End IF
							Else
								Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & InvoiceID & "</td>")
							End If 
						Else
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & InvoiceID & "</td>")
						End If 

			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & Subtotal & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & ShipTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & TaxTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & FuelTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & DepositTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & CouponTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & GranTot & "</td>")
			            Response.Write("</tr>")


						rowCount = rowCount + 1
						
						LineCount = LineCount + 1
						
						If IsNumeric(rsDailyAPIInvoices("InvoiceID")) Then LastInvoiceID = CDbl(rsDailyAPIInvoices("InvoiceID")) Else LastInvoiceID = rsDailyAPIInvoices("InvoiceID")
						
						
						rsDailyAPIInvoices.MoveNext
						
						If LineCount >= 21 Then 
					        Response.Write("</tbody>")
						    Response.Write("</table>")

							Call PageHeader

							Call SubHeader("Invoices","Body")
							
							Response.Write("<table style='margin-left:50px;width:1000px;'>")	
						        Response.Write("<thead>")
						            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' >Invoice ID</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Subtotal</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Shipping</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Tax</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Fuel</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Deposit</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Coupon</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Grand Total</th>")
						            Response.Write("</tr>")
						        Response.Write("</thead>")
						        Response.Write("<tbody>")

						End if

					Loop
				Else
					response.Write("<tr><td colspan='8'>No Invoice API Data</td></tr>")
				End If
				
				If DailySubtotal <> "" Then
					DailySubtotal = formatCurrency(DailySubtotal ,2)
				End If
				If DailyTaxTot <> "" Then
					DailyTaxTot = formatCurrency(DailyTaxTot ,2)
				End If
				If DailyShipTot <> "" Then
					DailyShipTot = formatCurrency(DailyShipTot ,2)
				End If
				If DailyFuelTot  <> "" Then
					DailyFuelTot  = formatCurrency(DailyFuelTot  ,2)
				End If
				If DailyDepositTot   <> "" Then
					DailyDepositTot   = formatCurrency(DailyDepositTot   ,2)
				End If
				If DailyCouponTot  <> "" Then
					DailyCouponTot  = formatCurrency(DailyCouponTot  ,2)
				End If
				If DailyGranTot  <> "" Then
					DailyGranTot  = formatCurrency(DailyGranTot  ,2)
				End If


				Response.Write("<tr style='border-top: 2px solid #ddd;'>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'><strong>Count:&nbsp;&nbsp;" & DailyCount & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailySubtotal & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyShipTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyTaxTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyFuelTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyDepositTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyCouponTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyGranTot & "</strong></td>")
	            Response.Write("</tr>")
		        Response.Write("</tbody>")
		    Response.Write("</table>")


		    '****************************************
			' R E T U R N   A U T H O R I Z A T I O N S
		    '****************************************
		    If LineCount + 5 >= 21 Then
				Call PageHeader
				Call SubHeader("RAs","Top")
			End If
			
			LineCount = LineCount + 5
			
        	Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Return Authorizations with a return authorization date of " & FormatDateTime(Date() - RA_OffSetFromToday) & "</h4>")
	    
		    
			Response.Write("<table style='margin-left:50px;width:1000px;'>")	
		        Response.Write("<thead>")
		            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' >RA ID</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Subtotal</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Shipping</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Tax</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Fuel</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Deposit</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Coupon</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Grand Total</th>")
		            Response.Write("</tr>")
		        Response.Write("</thead>")
		        Response.Write("<tbody>")
		        
				
				SQLDailyAPIReturnAuths = "SELECT * FROM API_OR_RAHeader "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " WHERE (DAY(RDDate) = " & RA_CurrentDay & " AND MONTH(RDDate) =  " & RA_CurrentMonth & " AND YEAR(RDDate) =  " & RA_CurrentYear & " AND "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " (Voided = 0) AND (APIKey = '" & currentPartnerAPIKey & "')) "
				'SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " AND RAID Not In (SELECT RAID FROM API_IN_CMHeader) "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & "  Order By RAID"
			

				Set cnnDailyAPIReturnAuths = Server.CreateObject("ADODB.Connection")
				cnnDailyAPIReturnAuths.open(Session("ClientCnnString"))
				Set rsDailyAPIReturnAuths = Server.CreateObject("ADODB.Recordset")
				rsDailyAPIReturnAuths.CursorLocation = 3 
				Set rsDailyAPIReturnAuths = cnnDailyAPIReturnAuths.Execute(SQLDailyAPIReturnAuths)

				rowCount = 1
				DailyCount = 0
				DailySubtotal = 0
				DailyTaxTot = 0 
				DailyShipTot = 0 
				DailyFuelTot  = 0
				DailyDepositTot   = 0
				DailyCouponTot  = 0
				DailyGranTot  = 0
				LastCMID = 0 
				Subtotal = 0 
				TaxTot = 0
				ShipTot = 0
				FuelTot = 0
				DepositTot = 0
				CouponTot = 0
				GranTot = 0
			
				If NOT rsDailyAPIReturnAuths.EOF Then
				
				
					Do While Not rsDailyAPIReturnAuths.EOF
						
						RAID = rsDailyAPIReturnAuths("RAID")
						DailyCount = DailyCount + 1
						
						Subtotal = rsDailyAPIReturnAuths("SubTotal")
						DailySubtotal = DailySubtotal + cdbl(Subtotal)
						
						TaxTot = rsDailyAPIReturnAuths("Tax")
						DailyTaxTot = DailyTaxTot + cdbl(TaxTot)
						
						ShipTot = rsDailyAPIReturnAuths("ShippingCharge")
						DailyShipTot = DailyShipTot + cdbl(ShipTot)
						
						FuelTot = rsDailyAPIReturnAuths("FuelSurcharge")
						DailyFuelTot  = DailyFuelTot  + cdbl(FuelTot)
						
						DepositTot = rsDailyAPIReturnAuths("DepositCharge")
						DailyDepositTot   = DailyDepositTot + cdbl(DepositTot)
						
						CouponTot = rsDailyAPIReturnAuths("CouponCharge")
						DailyCouponTot  = DailyCouponTot + cdbl(CouponTot)
						
						GranTot = rsDailyAPIReturnAuths("GrandTotal")
						DailyGranTot  = DailyGranTot + cdbl(GranTot)
						
						If Subtotal <> "" Then
							Subtotal = formatCurrency(Subtotal,2)
						End If
						If TaxTot <> "" Then
							TaxTot = formatCurrency(TaxTot,2)
						End If
						If ShipTot <> "" Then
							ShipTot = formatCurrency(ShipTot,2)
						End If
						If FuelTot <> "" Then
							FuelTot = formatCurrency(FuelTot,2)
						End If
						If DepositTot <> "" Then
							DepositTot = formatCurrency(DepositTot,2)
						End If
						If CouponTot <> "" Then
							CouponTot = formatCurrency(CouponTot,2)
						End If
						If GranTot <> "" Then
							GranTot = formatCurrency(GranTot,2)
						End If
						

						Response.Write("<tr>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & RAID & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & Subtotal & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & ShipTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & TaxTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & FuelTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & DepositTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & CouponTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & GranTot & "</td>")
			            Response.Write("</tr>")


						rowCount = rowCount + 1
						
						LineCount = LineCount + 1
						
						'LastRAID = CDbl(rsDailyAPIReturnAuths("RAID"))
						
						rsDailyAPIReturnAuths.MoveNext
						
						If LineCount >= 21 Then 
					        Response.Write("</tbody>")
						    Response.Write("</table>")

							Call PageHeader

							Call SubHeader("RAs","Body")
							
							Response.Write("<table style='margin-left:50px;width:1000px;'>")	
						        Response.Write("<thead>")
						            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' >RA ID</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Subtotal</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Shipping</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Tax</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Fuel</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Deposit</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Coupon</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Grand Total</th>")
						            Response.Write("</tr>")
						        Response.Write("</thead>")
						        Response.Write("<tbody>")

						End if


					Loop
				Else
					Response.Write("<tr><td colspan='8'>No Return Authorization API Data</td></tr>")
				End If

				If DailySubtotal <> "" Then
					DailySubtotal = formatCurrency(DailySubtotal ,2)
				End If
				If DailyTaxTot <> "" Then
					DailyTaxTot = formatCurrency(DailyTaxTot ,2)
				End If
				If DailyShipTot <> "" Then
					DailyShipTot = formatCurrency(DailyShipTot ,2)
				End If
				If DailyFuelTot  <> "" Then
					DailyFuelTot  = formatCurrency(DailyFuelTot  ,2)
				End If
				If DailyDepositTot   <> "" Then
					DailyDepositTot   = formatCurrency(DailyDepositTot   ,2)
				End If
				If DailyCouponTot  <> "" Then
					DailyCouponTot  = formatCurrency(DailyCouponTot  ,2)
				End If
				If DailyGranTot  <> "" Then
					DailyGranTot  = formatCurrency(DailyGranTot  ,2)
				End If

				
				Response.Write("<tr style='border-top: 2px solid #ddd;'>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'><strong>Count:&nbsp;&nbsp;" & DailyCount & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailySubtotal & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyShipTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyTaxTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyFuelTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyDepositTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyCouponTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyGranTot & "</strong></td>")
	            Response.Write("</tr>")
		        Response.Write("</tbody>")
		    Response.Write("</table>")



		    
		    '*********************************
			'		C R E D I T  M E M O S 
		    '*********************************
		    If LineCount + 5 >= 21 Then
				Call PageHeader
				Call SubHeader("CMs","Top")
			End If
		    
			LineCount = LineCount + 5
			
        	Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Credit Memos with a credit memo date of " & FormatDateTime(Date() - CM_OffSetFromToday) & "</h4>")
		        
	    
			Response.Write("<table style='margin-left:50px;width:1000px;'>")	
		        Response.Write("<thead>")
		            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' >CM ID</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Subtotal</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Shipping</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Tax</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Fuel</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Deposit</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Coupon</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Grand Total</th>")
		            Response.Write("</tr>")
		        Response.Write("</thead>")
		        Response.Write("<tbody>")

		        				
				SQLDailyAPICreditMemos= "SELECT * FROM API_IN_CMHeader "
				SQLDailyAPICreditMemos= SQLDailyAPICreditMemos& " WHERE (DAY(CMDate) = " & CM_CurrentDay & " AND MONTH(CMDate) =  " & CM_CurrentMonth & " AND YEAR(CMDate) =  " & CM_CurrentYear & " AND "
				SQLDailyAPICreditMemos= SQLDailyAPICreditMemos& " (Voided = 0) AND (APIKey = '" & currentPartnerAPIKey & "')) "
				'SQLDailyAPICreditMemos= SQLDailyAPICreditMemos& " AND RAID Not In (SELECT RAID FROM API_IN_CMHeader) "
				SQLDailyAPICreditMemos= SQLDailyAPICreditMemos& "  Order By CMID"

				Set cnnDailyAPICreditMemos = Server.CreateObject("ADODB.Connection")
				cnnDailyAPICreditMemos.open(Session("ClientCnnString"))
				Set rsDailyAPICreditMemos = Server.CreateObject("ADODB.Recordset")
				rsDailyAPICreditMemos.CursorLocation = 3 
				Set rsDailyAPICreditMemos = cnnDailyAPICreditMemos.Execute(SQLDailyAPICreditMemos)
				
				rowCount = 1
				DailyCount = 0
				DailySubtotal = 0
				DailyTaxTot = 0 
				DailyShipTot = 0 
				DailyFuelTot  = 0
				DailyDepositTot   = 0
				DailyCouponTot  = 0
				DailyGranTot  = 0
				LastCMID = 0 
				Subtotal = 0 
				TaxTot = 0
				ShipTot = 0
				FuelTot = 0
				DepositTot = 0
				CouponTot = 0
				GranTot = 0
				
				If NOT rsDailyAPICreditMemos.EOF Then

					Do While Not rsDailyAPICreditMemos.EOF
						
						CMID = rsDailyAPICreditMemos("CMID")
						DailyCount = DailyCount + 1
						
						Subtotal = rsDailyAPICreditMemos("CMSubTotal")
						DailySubtotal = DailySubtotal + cdbl(Subtotal)
						
						TaxTot = rsDailyAPICreditMemos("Tax")
						DailyTaxTot = DailyTaxTot + cdbl(TaxTot)
						
						ShipTot = rsDailyAPICreditMemos("ShippingCharge")
						DailyShipTot = DailyShipTot + cdbl(ShipTot)
						
						FuelTot = rsDailyAPICreditMemos("FuelSurcharge")
						DailyFuelTot  = DailyFuelTot  + cdbl(FuelTot)
						
						DepositTot = rsDailyAPICreditMemos("DepositCharge")
						DailyDepositTot   = DailyDepositTot + cdbl(DepositTot)
						
						CouponTot = rsDailyAPICreditMemos("CouponCharge")
						DailyCouponTot  = DailyCouponTot + cdbl(CouponTot)
						
						GranTot = rsDailyAPICreditMemos("GrandTotal")
						DailyGranTot  = DailyGranTot + cdbl(GranTot)
						
						If Subtotal <> "" Then
							Subtotal = formatCurrency(Subtotal,2)
						End If
						If TaxTot <> "" Then
							TaxTot = formatCurrency(TaxTot,2)
						End If
						If ShipTot <> "" Then
							ShipTot = formatCurrency(ShipTot,2)
						End If
						If FuelTot <> "" Then
							FuelTot = formatCurrency(FuelTot,2)
						End If
						If DepositTot <> "" Then
							DepositTot = formatCurrency(DepositTot,2)
						End If
						If CouponTot <> "" Then
							CouponTot = formatCurrency(CouponTot,2)
						End If
						If GranTot <> "" Then
							GranTot = formatCurrency(GranTot,2)
						End If
						

						Response.Write("<tr>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & CMID & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & Subtotal & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & ShipTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & TaxTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & FuelTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & DepositTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & CouponTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & GranTot & "</td>")
			            Response.Write("</tr>")


						rowCount = rowCount + 1
						
						LineCount = LineCount + 1
						
						rsDailyAPICreditMemos.MoveNext
						
						If LineCount >= 21 Then 
					        Response.Write("</tbody>")
						    Response.Write("</table>")

							Call PageHeader

							Call SubHeader("CMs","Body")
							
							Response.Write("<table style='margin-left:50px;width:1000px;'>")	
						        Response.Write("<thead>")
						            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' >CM ID</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Subtotal</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Shipping</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Tax</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Fuel</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Deposit</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Coupon</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Grand Total</th>")
						            Response.Write("</tr>")
						        Response.Write("</thead>")
						        Response.Write("<tbody>")

						End if
						
					Loop
				Else
					Response.Write("<tr><td colspan='8'>No Credit Memo API Data</td></tr>")
				End If

				If DailySubtotal <> "" Then
					DailySubtotal = formatCurrency(DailySubtotal ,2)
				End If
				If DailyTaxTot <> "" Then
					DailyTaxTot = formatCurrency(DailyTaxTot ,2)
				End If
				If DailyShipTot <> "" Then
					DailyShipTot = formatCurrency(DailyShipTot ,2)
				End If
				If DailyFuelTot  <> "" Then
					DailyFuelTot  = formatCurrency(DailyFuelTot  ,2)
				End If
				If DailyDepositTot   <> "" Then
					DailyDepositTot   = formatCurrency(DailyDepositTot   ,2)
				End If
				If DailyCouponTot  <> "" Then
					DailyCouponTot  = formatCurrency(DailyCouponTot  ,2)
				End If
				If DailyGranTot  <> "" Then
					DailyGranTot  = formatCurrency(DailyGranTot  ,2)
				End If

			    
				Response.Write("<tr style='border-top: 2px solid #ddd;'>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'><strong>Count:&nbsp;&nbsp;" &  DailyCount & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailySubtotal & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyShipTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyTaxTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyFuelTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyDepositTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyCouponTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyGranTot & "</strong></td>")
	            Response.Write("</tr>")
		        Response.Write("</tbody>")
		    Response.Write("</table>")

		    
		    
			'*************************************
			'		S U M M A R Y  I N V O I C E S
		    '*************************************
		    If LineCount + 5 >= 21 Then
				Call PageHeader
				Call SubHeader("SumInv","Top")
			End If
	        
			Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Summary Invoices with a summary invoice date of " & FormatDateTime(Date() - SumInv_OffSetFromToday) & "</h4>")			
			

			LineCount = LineCount + 5
			
		    
			Response.Write("<table style='margin-left:50px;width:1000px;'>")	
		        Response.Write("<thead>")
		            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' >Sum Inv ID</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Subtotal</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Shipping</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Tax</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Fuel</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Deposit</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Coupon</th>")
		                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Grand Total</th>")
		            Response.Write("</tr>")
		        Response.Write("</thead>")
		        Response.Write("<tbody>")
		        

				
				SQLDailyAPISummInv= "SELECT * FROM API_IN_SummaryInvoiceHeader "
				SQLDailyAPISummInv= SQLDailyAPISummInv & " WHERE (DAY(SumInvDate) = " & SumInv_CurrentDay & " AND MONTH(SumInvDate) =  " & SumInv_CurrentMonth & " AND YEAR(SumInvDate) =  " & SumInv_CurrentYear & " AND "
				SQLDailyAPISummInv= SQLDailyAPISummInv & " (Voided = 0) AND (APIKey = '" & currentPartnerAPIKey & "')) "
				'SQLDailyAPISummInv= SQLDailyAPISummInv & " AND RAID Not In (SELECT RAID FROM API_IN_CMHeader) "
				SQLDailyAPISummInv= SQLDailyAPISummInv & "  Order By SumInvID"

				Set cnnDailyAPISummInv = Server.CreateObject("ADODB.Connection")
				cnnDailyAPISummInv.open(Session("ClientCnnString"))
				Set rsDailyAPISummInv = Server.CreateObject("ADODB.Recordset")
				rsDailyAPISummInv.CursorLocation = 3 
				Set rsDailyAPISummInv = cnnDailyAPISummInv.Execute(SQLDailyAPISummInv)
				
				rowCount = 1
				DailyCount = 0
				DailySubtotal = 0
				DailyTaxTot = 0 
				DailyShipTot = 0 
				DailyFuelTot  = 0
				DailyDepositTot   = 0
				DailyCouponTot  = 0
				DailyGranTot  = 0
				LastCMID = 0 
				Subtotal = 0 
				TaxTot = 0
				ShipTot = 0
				FuelTot = 0
				DepositTot = 0
				CouponTot = 0
				GranTot = 0
				
				If NOT rsDailyAPISummInv.EOF Then
				
					Do While Not rsDailyAPISummInv.EOF
						
						CMID = rsDailyAPISummInv("SumInvID")
						DailyCount = DailyCount + 1
						
						Subtotal = rsDailyAPISummInv("Sub_Total")
						DailySubtotal = DailySubtotal + cdbl(Subtotal)
						
						TaxTot = rsDailyAPISummInv("Total_Tax")
						DailyTaxTot = DailyTaxTot + cdbl(TaxTot)
						
						ShipTot = rsDailyAPISummInv("Shipping_Charge")
						DailyShipTot = DailyShipTot + cdbl(ShipTot)
						
						FuelTot = rsDailyAPISummInv("FuelSurcharge")
						DailyFuelTot  = DailyFuelTot  + cdbl(FuelTot)
						
						DepositTot = rsDailyAPISummInv("DepositCharge")
						DailyDepositTot   = DailyDepositTot + cdbl(DepositTot)
						
						CouponTot = rsDailyAPISummInv("CouponCharge")
						DailyCouponTot  = DailyCouponTot + cdbl(CouponTot)
						
						GranTot = rsDailyAPISummInv("Grand_Total")
						DailyGranTot  = DailyGranTot + cdbl(GranTot)
						
						If Subtotal <> "" Then
							Subtotal = formatCurrency(Subtotal,2)
						End If
						If TaxTot <> "" Then
							TaxTot = formatCurrency(TaxTot,2)
						End If
						If ShipTot <> "" Then
							ShipTot = formatCurrency(ShipTot,2)
						End If
						If FuelTot <> "" Then
							FuelTot = formatCurrency(FuelTot,2)
						End If
						If DepositTot <> "" Then
							DepositTot = formatCurrency(DepositTot,2)
						End If
						If CouponTot <> "" Then
							CouponTot = formatCurrency(CouponTot,2)
						End If
						If GranTot <> "" Then
							GranTot = formatCurrency(GranTot,2)
						End If
						
						Response.Write("<tr>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & CMID & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & Subtotal & "</td>")
			                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & ShipTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & TaxTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & FuelTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & DepositTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & CouponTot & "</td>")
							Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'>" & GranTot & "</td>")
			            Response.Write("</tr>")

						rowCount = rowCount + 1
						
						LineCount = LineCount + 1
						
						rsDailyAPISummInv.MoveNext
						
						If LineCount >= 21 Then 
					        Response.Write("</tbody>")
						    Response.Write("</table>")

							Call PageHeader

							Call SubHeader("SumInv","Body")
							
							Response.Write("<table style='margin-left:50px;width:1000px;'>")	
						        Response.Write("<thead>")
						            Response.Write("<tr style='border-bottom: 2px solid #ddd;'>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;'  align='right' >Sum Inv ID</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Subtotal</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Shipping</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Tax</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Fuel</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Deposit</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right' >Coupon</th>")
						                Response.Write("<th style='padding-top: 8px; text-align: right;' align='right'  >Grand Total</th>")
						            Response.Write("</tr>")
						        Response.Write("</thead>")
						        Response.Write("<tbody>")

						End if
						
					Loop
				Else
					Response.Write("<tr><td colspan='8'>No Summary Invoice API Data</td></tr>")
				End If

				If DailySubtotal <> "" Then
					DailySubtotal = formatCurrency(DailySubtotal ,2)
				End If
				If DailyTaxTot <> "" Then
					DailyTaxTot = formatCurrency(DailyTaxTot ,2)
				End If
				If DailyShipTot <> "" Then
					DailyShipTot = formatCurrency(DailyShipTot ,2)
				End If
				If DailyFuelTot  <> "" Then
					DailyFuelTot  = formatCurrency(DailyFuelTot  ,2)
				End If
				If DailyDepositTot   <> "" Then
					DailyDepositTot   = formatCurrency(DailyDepositTot   ,2)
				End If
				If DailyCouponTot  <> "" Then
					DailyCouponTot  = formatCurrency(DailyCouponTot  ,2)
				End If
				If DailyGranTot  <> "" Then
					DailyGranTot  = formatCurrency(DailyGranTot  ,2)
				End If

			    	Response.Write("<tr style='border-top: 2px solid #ddd;'>")
	                Response.Write("<td style='padding-top: 8px;text-align: right;' align='right'><strong>Count:&nbsp;&nbsp;" & DailyCount & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailySubtotal & "</strong></td>")
	                Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyShipTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyTaxTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyFuelTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyDepositTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyCouponTot & "</strong></td>")
					Response.Write("<td style='padding-top: 8px; text-align: right;' align='right'><strong>" & DailyGranTot & "</strong></td>")
	            Response.Write("</tr>")
		        Response.Write("</tbody>")
		    Response.Write("</table>")

	rsDailyAPIPartnersLoop.MoveNext
Loop
End If


Sub PageHeader


	LineCount = 0	
 	PageNo = PageNo + 1

	If PageNo > 1 Then Response.Write("<div style='page-break-before: always'>")

 	Response.Write("<div style='width:100%;'>")

 	Response.Write("<img src='/clientfiles/" & ClientKey & "/logos/logo.png' style='float:left; margin-top:30px;'><center><h1 >DAILY API ACTIVITY DETAIL <Br>BY PARTNER"  & "</h1><h2 class='generated' >Generated " & WeekDayName(WeekDay(DateValue(Now()))) & "&nbsp;" &  Now() & "</h2></center>")

 	Response.Write("</div><BR><BR>")

 	If PageNo > 1 Then Response.Write("</div>") 	
End Sub

Sub SubHeader(passedSection, passedTopOrBody)

	If passedSection="Orders" Then
		Response.Write("<hr>")
		Response.Write("<h2>Partner: " & GetPartnerNameByAPIKey(currentPartnerAPIKey) & "</h2>")
		Response.Write("<hr>")
		Response.Write("<div><span style='background:yellow; display:inline-block; padding:3px;'>Yellow highlights on orders and invoices indicates a gap in sequence</span></div>")
				    	
	   	Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Orders with an order date of " & FormatDateTime(date() - Order_OffSetFromToday)  & " (continued)</h4><br>")
	End If
	
	If passedSection="Invoices" Then
		Response.Write("<hr>")
		Response.Write("<h2>Partner: " & GetPartnerNameByAPIKey(currentPartnerAPIKey) & "</h2>")
		Response.Write("<hr>")
		Response.Write("<div><span style='background:yellow; display:inline-block; padding:3px;'>Yellow highlights on orders and invoices indicates a gap in sequence</span></div>")
				    	
	    If passedTopOrBody = "Body" Then Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Invocies with an order date of " & FormatDateTime(date() - Invoice_OffSetFromToday)  & " (continued)</h4><br>")
	End If

	If passedSection="RAs" Then
		Response.Write("<hr>")
		Response.Write("<h2>Partner: " & GetPartnerNameByAPIKey(currentPartnerAPIKey) & "</h2>")
		Response.Write("<hr>")
				    	
	   	If passedTopOrBody = "Body" Then Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Return Authorizations with an order date of " & FormatDateTime(date() - RA_OffSetFromToday)  & " (continued)</h4><br>")
	End If

	If passedSection="CMs" Then
		Response.Write("<hr>")
		Response.Write("<h2>Partner: " & GetPartnerNameByAPIKey(currentPartnerAPIKey) & "</h2>")
		Response.Write("<hr>")
				    	
	   	If passedTopOrBody = "Body" Then Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Credit Memos with an order date of " & FormatDateTime(date() - CM_OffSetFromToday)  & " (continued)</h4><br>")
	End If

	If passedSection="SumInv" Then
		Response.Write("<hr>")
		Response.Write("<h2>Partner: " & GetPartnerNameByAPIKey(currentPartnerAPIKey) & "</h2>")
		Response.Write("<hr>")
				    	
	   	If passedTopOrBody = "Body" Then Response.Write("<h4 style='color: #3c763d; margin-top: 40px; font-size:23px;'>Summary Invoices with an order date of " & FormatDateTime(date() - SumInv_OffSetFromToday)  & " (continued)</h4><br>")
	End If

End Sub

%></div></body></html>