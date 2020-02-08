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


    <!-- icons and notification styles !-->
    <script src="https://use.fontawesome.com/3382135cdc.js"></script>

	<style type="text/css">
	
	body{font-family: arial, helvetica, sans-serif;}
	
	div.table-title {
	  display: block;
	  margin: auto;
	  max-width: 800px;
	  padding:5px;
	  width: 100%;
	}

	
	div.table-data {
	  display: block;
	  margin: auto;
	  max-width: 1200px;
	  padding:5px;
	  width: 100%;
	}
	
	p, h1, h2 {
	  display: block;
	  margin: auto;
	  max-width: 800px;
	  padding:5px;
	  width: 100%;
	}

	h1 {
		color: #193048;
	    font-size: 30px;
	    font-weight: 400;
	    font-style: normal;
	    font-family: arial, helvetica, sans-serif;
	    text-transform: uppercase;
	    text-align:center;
	}
	
	.generated {
		color: #3e94ec;
	    font-size: 20px;
	    font-weight: 400;
	    font-style: normal;
	    font-family: arial, helvetica, sans-serif;
	    text-transform: uppercase;
	    text-align:center;
	}
	
	h2 {
		color: #3e94ec;
	    font-size: 20px;
	    font-weight: 400;
	    font-style: normal;
	    font-family: arial, helvetica, sans-serif;
	    text-transform: uppercase;
	    text-align:center;
	}
	
	hr{
	   /* margin-top: 40px;
	    margin-bottom: 40px;*/
	}
	
	.table-title h3 {
	   color: #193048;
	   font-size: 22px;
	   font-weight: 400;
	   font-style:normal;
	   font-family: arial, helvetica, sans-serif;
	   text-transform:uppercase;
	   font-weight:bold;
	}
	
	
	/*** Table Styles **/

	.table-fill {
	  background: white;
	  border-collapse: collapse;
	  margin: auto;
	  max-width: 800px;
 	  padding:5px;
	  width: 100%;
	  font-family: arial, helvetica, sans-serif;
	}
	 
	th {
	  color:#483D8B;
	  /*font-size:23px;*/
	  font-size: 18px;
	  font-weight: 100;
	  padding:13px !important;
	  text-align:left;
	  vertical-align:middle;
	  border: 1px solid #C1C3D1;
	  width: 12.5% !important;
   	}

	tr {
	  color:#666B85;
	  font-size:16px;
	  font-weight:normal;
	}
	 	 
	tr:nth-child(odd) td {
	  background:#EBEBEB;
	}
	 	
	 
	td {
	  background:#FFFFFF;
	  padding:20px 13px 14px 20px !important;
	  text-align:left;
	  vertical-align:middle;
	  font-weight:300;
	  font-size:18px;
	  border: 1px solid #C1C3D1;
	}
	
	/* custom table */
	
	 
	
	.custom-table th{
		padding:5px;
	}
	
	.custom-table td{
		padding:5px;
	}
	
	#leftcol{
		width:65%;
	}
	
	#rightcol{
		width:35%;
	}
	
	#table-fill-short{
		max-width: 500px;
	}
	/* eof custom table */
	

 
 

	</style>
     
</head>




<body bgcolor="#FFFFFF" text="#000000" link="#000080" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">


	 
  	<div class='table-title'>
<% Call PageHeader %>
<br>
 </div>
 

<div class="line">
<div class="table-data">

<%
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
	
%>
	<hr>
	<h2>Partner: <%= GetPartnerNameByAPIKey(currentPartnerAPIKey) %></h2>
	<hr>
	
        	<h4 style="color: #3c763d; margin-top: 40px; font-size:23px;">Orders with an order date of <%=FormatDateTime(date() - Order_OffSetFromToday)%></h4><br>
			<!-- HTML -->
		    
		    
		    <% 'Details Go Here %>
		    
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" ># Orders</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Subtotal</th>
		                <th style="padding-top: 8px; text-align: right;" align="right"  >Shipping</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Tax</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Fuel</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Deposit</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Coupon</th>
		                <th style="padding-top: 8px; text-align: right;" align="right"  >Grand Total</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%		
				
				SQLDailyAPIOrders = "SELECT COUNT(*) AS NumOrders, SUM(OrderSubTotal) AS Subtotal, SUM(Tax) AS TaxTot, SUM(ShippingCharge) AS ShipTot, "
				SQLDailyAPIOrders = SQLDailyAPIOrders & " SUM(DepositCharge) AS DepositTot, SUM(FuelSurcharge) AS FuelTot, SUM(CouponCharge) AS CouponTot, SUM(GrandTotal) AS GranTot"
				SQLDailyAPIOrders = SQLDailyAPIOrders & " FROM            API_OR_OrderHeader"
				SQLDailyAPIOrders = SQLDailyAPIOrders & " WHERE        (InternalRecordIdentifier IN"
				SQLDailyAPIOrders = SQLDailyAPIOrders & " (SELECT        MAX(InternalRecordIdentifier) AS Expr1"
				SQLDailyAPIOrders = SQLDailyAPIOrders & " FROM            API_OR_OrderHeader AS API_OR_OrderHeader_1"
				SQLDailyAPIOrders = SQLDailyAPIOrders & " WHERE        (DAY(OrderDate) = " & Order_CurrentDay & " AND MONTH(OrderDate) =  " & Order_CurrentMonth & " AND YEAR(OrderDate) =  " & Order_CurrentYear & " AND (Voided = 0) AND (APIKey = '" & currentPartnerAPIKey & "')) "
				'SQLDailyAPIOrders = SQLDailyAPIOrders & " AND OrderID Not In (SELECT OrderID FROM API_IN_InvoiceHeader) "
				SQLDailyAPIOrders = SQLDailyAPIOrders & " GROUP BY OrderID)) "	
	

				
				Set cnnDailyAPIOrders = Server.CreateObject("ADODB.Connection")
				cnnDailyAPIOrders.open(Session("ClientCnnString"))
				Set rsDailyAPIOrders = Server.CreateObject("ADODB.Recordset")
				rsDailyAPIOrders.CursorLocation = 3 
				Set rsDailyAPIOrders = cnnDailyAPIOrders.Execute(SQLDailyAPIOrders)
				
				If NOT rsDailyAPIOrders.EOF and rsDailyAPIOrders("NumOrders") <> 0 Then
				
					rowCount = 1
				
					Do While Not rsDailyAPIOrders.EOF
						
						NumOrders = rsDailyAPIOrders("NumOrders")
						Subtotal = rsDailyAPIOrders("Subtotal")					
						TaxTot = rsDailyAPIOrders("TaxTot")
						ShipTot = rsDailyAPIOrders("ShipTot")
						FuelTot = rsDailyAPIOrders("FuelTot")
						DepositTot = rsDailyAPIOrders("DepositTot")
						CouponTot = rsDailyAPIOrders("CouponTot")
						GranTot = rsDailyAPIOrders("GranTot")
						
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
						
						%>
						<tr>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= NumOrders %></td>
			                <td style="padding-top: 8px; text-align: right;" align="right"><%= Subtotal %></td>
			                <td style="padding-top: 8px; text-align: right;" align="right"><%= ShipTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= TaxTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= FuelTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= DepositTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= CouponTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= GranTot %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsDailyAPIOrders.MoveNext
					Loop
				Else
					%><tr><td colspan="8">No Order API Data</td></tr><%
				End If
				%>
		        </tbody>
		    </table>
		    
		    
        	<h4 style="color: #3c763d; margin-top: 40px; font-size:23px;">Invoices with an invoice date of <%=FormatDateTime(Date() - Invoice_OffSetFromToday)%></h4>
			<!-- HTML -->
			
			
			
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" ># Invoices</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Subtotal</th>
		                <th style="padding-top: 8px; text-align: right;" align="right"  >Shipping</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Tax</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Fuel</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Deposit</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Coupon</th>
		                <th style="padding-top: 8px; text-align: right;" align="right"  >Grand Total</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%

				SQLDailyAPIInvoices = "SELECT COUNT(*) AS NumInv, SUM(InvoiceSubTotal) AS Subtotal, SUM(Tax) AS TaxTot, SUM(ShippingCharge) AS ShipTot, "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " SUM(DepositCharge) AS DepositTot, SUM(FuelSurcharge) AS FuelTot, SUM(CouponCharge) AS CouponTot, SUM(GrandTotal) AS GranTot "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " FROM            API_IN_InvoiceHeader "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " WHERE        (InternalRecordIdentifier IN "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " (SELECT        MAX(InternalRecordIdentifier) AS Expr1 "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " FROM            API_IN_InvoiceHeader AS API_IN_InvoiceHeader_1 "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " WHERE        (DAY(InvoiceDate) = " & Invoice_CurrentDay & " AND MONTH(InvoiceDate) =  " & Invoice_CurrentMonth & " AND YEAR(InvoiceDate) =  " & Invoice_CurrentYear & " AND (Voided = 0)  AND (APIKey = '" & currentPartnerAPIKey & "')) "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " GROUP BY InvoiceID))	"	

				Set cnnDailyAPIInvoices = Server.CreateObject("ADODB.Connection")
				cnnDailyAPIInvoices.open(Session("ClientCnnString"))
				Set rsDailyAPIInvoices = Server.CreateObject("ADODB.Recordset")
				rsDailyAPIInvoices.CursorLocation = 3 
				Set rsDailyAPIInvoices = cnnDailyAPIInvoices.Execute(SQLDailyAPIInvoices)
				
				If NOT rsDailyAPIInvoices.EOF and rsDailyAPIInvoices("NumInv") <> 0 Then
				
					rowCount = 1
				
					Do While Not rsDailyAPIInvoices.EOF
						
						NumOrders = rsDailyAPIInvoices("NumInv")
						Subtotal = rsDailyAPIInvoices("Subtotal")						
						TaxTot = rsDailyAPIInvoices("TaxTot")
						ShipTot = rsDailyAPIInvoices("ShipTot")
						FuelTot = rsDailyAPIInvoices("FuelTot")
						DepositTot = rsDailyAPIInvoices("DepositTot")
						CouponTot = rsDailyAPIInvoices("CouponTot")
						GranTot = rsDailyAPIInvoices("GranTot")
						
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
						
						%>
						<tr>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= NumOrders %></td>
			                <td style="padding-top: 8px; text-align: right;" align="right"><%= Subtotal %></td>
			                <td style="padding-top: 8px; text-align: right;" align="right"><%= ShipTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= TaxTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= FuelTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= DepositTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= CouponTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= GranTot %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsDailyAPIInvoices.MoveNext
					Loop
				Else
					%><tr><td colspan="8">No Invoice API Data</td></tr><%
				End If
				%>
		        </tbody>
		    </table>


        	<h4 style="color: #3c763d; margin-top: 40px; font-size:23px;">Return Authorizations with a return authorization date of <%=FormatDateTime(date() - RA_OffSetFromToday)%></h4><br>
			<!-- HTML -->
		    
		    
		    <% 'Details Go Here %>
		    
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" ># RAs</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Subtotal</th>
		                <th style="padding-top: 8px; text-align: right;" align="right"  >Shipping</th>
		                <th style="padding-top: 8px; text-align: right;" align="right">Tax</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Fuel</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Deposit</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Coupon</th>
		                <th style="padding-top: 8px; text-align: right;" align="right"  >Grand Total</th>
		            </tr>
		        </thead>
		        <tbody>
		        
		        
				<%
				
				SQLDailyAPIReturnAuths = " SELECT  COUNT(*) AS NumRA, SUM(SubTotal) AS Subtotal, SUM(Tax) AS TaxTot, SUM(ShippingCharge) AS shiptot, "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " SUM(DepositCharge) AS DepositTot, SUM(FuelSurcharge) AS FuelTot, SUM(CouponCharge) AS CouponTot, "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " SUM(GrandTotal) AS GranTot "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " FROM            API_OR_RAHeader "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " WHERE        (InternalRecordIdentifier IN "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " (SELECT        MAX(InternalRecordIdentifier) AS Expr1 "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " FROM            API_OR_RAHeader AS API_OR_RAHeader_1 "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " WHERE        (DAY(RDDate) = " & RA_CurrentDay & " AND MONTH(RDDate) =  " & RA_CurrentMonth & " AND YEAR(RDDate) =  " & RA_CurrentYear & " AND (Voided = 0) AND (APIKey = '" & currentPartnerAPIKey & "')) "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " GROUP BY RAID))	"				

				Set cnnDailyAPIReturnAuths = Server.CreateObject("ADODB.Connection")
				cnnDailyAPIReturnAuths.open(Session("ClientCnnString"))
				Set rsDailyAPIReturnAuths = Server.CreateObject("ADODB.Recordset")
				rsDailyAPIReturnAuths.CursorLocation = 3 
				Set rsDailyAPIReturnAuths = cnnDailyAPIReturnAuths.Execute(SQLDailyAPIReturnAuths)
				
				If NOT rsDailyAPIReturnAuths.EOF and rsDailyAPIReturnAuths("NumRA") <> 0 Then
				
					rowCount = 1
				
					Do While Not rsDailyAPIReturnAuths.EOF
						
						NumRA = rsDailyAPIReturnAuths("NumRA")
						Subtotal = rsDailyAPIReturnAuths("Subtotal")					
						TaxTot = rsDailyAPIReturnAuths("TaxTot")
						ShipTot = rsDailyAPIReturnAuths("ShipTot")
						FuelTot = rsDailyAPIReturnAuths("FuelTot")
						DepositTot = rsDailyAPIReturnAuths("DepositTot")
						CouponTot = rsDailyAPIReturnAuths("CouponTot")
						GranTot = rsDailyAPIReturnAuths("GranTot")
						
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
						
						%>
						<tr>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= NumRA %></td>
			                <td style="padding-top: 8px; text-align: right;" align="right"><%= Subtotal %></td>
			                <td style="padding-top: 8px; text-align: right;" align="right"><%= ShipTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= TaxTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= FuelTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= DepositTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= CouponTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= GranTot %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsDailyAPIReturnAuths.MoveNext
					Loop
				Else
					%><tr><td colspan="8">No Return Authorization API Data</td></tr><%
				End If
				%>
		        </tbody>
		    </table>

        	<h4 style="color: #3c763d; margin-top: 40px; font-size:23px;">Credit Memos with a credit memo date of <%=FormatDateTime(date() - CM_OffSetFromToday)%></h4><br>
			<!-- HTML -->
		        
		    
		    <% 'Details Go Here %>
		    
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" ># CMs</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Subtotal</th>
		                <th style="padding-top: 8px; text-align: right;" align="right"  >Shipping</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Tax</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Fuel</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Deposit</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Coupon</th>
		                <th style="padding-top: 8px; text-align: right;" align="right"  >Grand Total</th>
		            </tr>
		        </thead>
		        <tbody>

		        
				<%
				SQLDailyAPICreditMemos = "SELECT COUNT(*) AS NumInv, SUM(CMSubTotal) AS Subtotal, SUM(Tax) AS TaxTot, SUM(ShippingCharge) AS ShipTot, "
				SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " SUM(DepositCharge) AS DepositTot, SUM(FuelSurcharge) AS FuelTot, SUM(CouponCharge) AS CouponTot, SUM(GrandTotal) AS GranTot "
				SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " FROM            API_IN_CMHeader "
				SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " WHERE        (InternalRecordIdentifier IN "
				SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " (SELECT        MAX(InternalRecordIdentifier) AS Expr1 "
				SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " FROM            API_IN_CMHeader AS API_IN_CMHeader_1 "
				SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " WHERE        (DAY(CMDate) = " & CM_CurrentDay & " AND MONTH(CMDate) =  " & CM_CurrentMonth & " AND YEAR(CMDate) =  " & CM_CurrentYear & " AND (Voided = 0)  AND (APIKey = '" & currentPartnerAPIKey & "')) "
				SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " GROUP BY CMID))	"	

				Set cnnDailyAPICreditMemos = Server.CreateObject("ADODB.Connection")
				cnnDailyAPICreditMemos.open(Session("ClientCnnString"))
				Set rsDailyAPICreditMemos = Server.CreateObject("ADODB.Recordset")
				rsDailyAPICreditMemos.CursorLocation = 3 
				Set rsDailyAPICreditMemos = cnnDailyAPICreditMemos.Execute(SQLDailyAPICreditMemos)
				
				If NOT rsDailyAPICreditMemos.EOF and rsDailyAPICreditMemos("NumInv") <> 0 Then
				
					rowCount = 1
				
					Do While Not rsDailyAPICreditMemos.EOF
						
						NumInv = rsDailyAPICreditMemos("NumInv")
						Subtotal = rsDailyAPICreditMemos("Subtotal")					
						TaxTot = rsDailyAPICreditMemos("TaxTot")
						ShipTot = rsDailyAPICreditMemos("ShipTot")
						FuelTot = rsDailyAPICreditMemos("FuelTot")
						DepositTot = rsDailyAPICreditMemos("DepositTot")
						CouponTot = rsDailyAPICreditMemos("CouponTot")
						GranTot = rsDailyAPICreditMemos("GranTot")
						
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
						
						%>
						<tr>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= NumInv %></td>
			                <td style="padding-top: 8px; text-align: right;" align="right"><%= Subtotal %></td>
			                <td style="padding-top: 8px; text-align: right;" align="right"><%= ShipTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= TaxTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= FuelTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= DepositTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= CouponTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= GranTot %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsDailyAPICreditMemos.MoveNext
					Loop
				Else
					%><tr><td colspan="8">No Credit Memo API Data</td></tr><%
				End If
				%>
		        </tbody>
		    </table>
		    
		    
        	<h4 style="color: #3c763d; margin-top: 40px; font-size:23px;">Summary Invoices with a summary invoice date of <%=FormatDateTime(date() - SumInv_OffSetFromToday)%></h4><br>
			<!-- HTML -->
		        
		    <% 'Details Go Here %>
		    
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" ># Sum Inv</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Subtotal</th>
		                <th style="padding-top: 8px; text-align: right;" align="right"  >Shipping</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Tax</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Fuel</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Deposit</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" >Coupon</th>
		                <th style="padding-top: 8px; text-align: right;" align="right"  >Grand Total</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%
				SQLDailyAPISummInv = "SELECT COUNT(*) AS NumInv, SUM(Sub_Total) AS Subtotal, SUM(Total_Tax) AS TaxTot, SUM(Shipping_Charge) AS ShipTot, "
				SQLDailyAPISummInv = SQLDailyAPISummInv & " SUM(DepositCharge) AS DepositTot, SUM(FuelSurcharge) AS FuelTot, SUM(CouponCharge) AS CouponTot, SUM(Grand_Total) AS GranTot "
				SQLDailyAPISummInv = SQLDailyAPISummInv & " FROM            API_IN_SummaryInvoiceHeader "
				SQLDailyAPISummInv = SQLDailyAPISummInv & " WHERE        (InternalRecordIdentifier IN "
				SQLDailyAPISummInv = SQLDailyAPISummInv & " (SELECT        MAX(InternalRecordIdentifier) AS Expr1 "
				SQLDailyAPISummInv = SQLDailyAPISummInv & " FROM            API_IN_SummaryInvoiceHeader AS API_IN_SummaryInvoiceHeader_1 "
				SQLDailyAPISummInv = SQLDailyAPISummInv & " WHERE        (DAY(SumInvDate) = " & SumInv_CurrentDay & " AND MONTH(SumInvDate) =  " & SumInv_CurrentMonth & " AND YEAR(SumInvDate) =  " & SumInv_CurrentYear & " AND (Voided = 0)  AND (APIKey = '" & currentPartnerAPIKey & "')) "
				SQLDailyAPISummInv = SQLDailyAPISummInv & " GROUP BY SumInvID))	"	

				Set cnnDailyAPISummInv = Server.CreateObject("ADODB.Connection")
				cnnDailyAPISummInv.open(Session("ClientCnnString"))
				Set rsDailyAPISummInv = Server.CreateObject("ADODB.Recordset")
				rsDailyAPISummInv.CursorLocation = 3 
				Set rsDailyAPISummInv = cnnDailyAPISummInv.Execute(SQLDailyAPISummInv)
				
				If NOT rsDailyAPISummInv.EOF and rsDailyAPISummInv("NumInv") <> 0 Then
				
					rowCount = 1
				
					Do While Not rsDailyAPISummInv.EOF
						
						NumInv = rsDailyAPISummInv("NumInv")
						Subtotal = rsDailyAPISummInv("Subtotal")					
						TaxTot = rsDailyAPISummInv("TaxTot")
						ShipTot = rsDailyAPISummInv("ShipTot")
						FuelTot = rsDailyAPISummInv("FuelTot")
						DepositTot = rsDailyAPISummInv("DepositTot")
						CouponTot = rsDailyAPISummInv("CouponTot")
						GranTot = rsDailyAPISummInv("GranTot")
						
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
						
						%>
						<tr>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= NumInv %></td>
			                <td style="padding-top: 8px; text-align: right;" align="right"><%= Subtotal %></td>
			                <td style="padding-top: 8px; text-align: right;" align="right"><%= ShipTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= TaxTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= FuelTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= DepositTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= CouponTot %></td>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= GranTot %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsDailyAPISummInv.MoveNext
					Loop
				Else
					%><tr><td colspan="8">No Summary Invoice API Data</td></tr><%
				End If
				%>
		        </tbody>
		    </table>
<%

	rsDailyAPIPartnersLoop.MoveNext
Loop
End If



Sub PageHeader

 	Response.Write("<div style='width:100%;'><img src='" & BaseURL & "clientfiles/" & MUV_Read("ClientID") & "/logos/logo.png' style='float:left; margin-top:30px;' ><center ><h1>DAILY API ACTIVITY SUMMARY <br>BY PARTNER" & "</h1><h2 class='generated'>Generated " & WeekDayName(WeekDay(DateValue(Now()))) & "&nbsp;" &  Now() & "</h2></center></div>")
 	Response.Write("<BR><BR>")	
End Sub


%>   
</div> </div>        
</body>

</html>


 