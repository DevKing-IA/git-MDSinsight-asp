<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_API.asp"-->
<!-- Styles -->
<style>
h2 {
    margin-top: 40px;
    margin-bottom: 10px;
}
    .text-success {
    color: #3c763d;
    margin-top:40px;
}
</style>

<!-- Resources -->

<%
OffSetFromToday = 2
currentDay = day(date()) - OffSetFromToday 
currentMonth = month(date())
currentYear = year(date())

%>


<h1 class="page-header"><i class="fa fa-fw fa-calendar"></i> Daily API Activity Summary By Partner for <%= FormatDateTime(DateAdd("d",-OffSetFromToday,Date()),2) %></h1>

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
	<h2><i class="fa fa-handshake-o" aria-hidden="true"></i> Partner: <%= GetPartnerNameByAPIKey(currentPartnerAPIKey) %></h2>
	<hr>
	
        	<h4 style="color: #3c763d; margin-top: 40px;">Orders </h4><br>
			<!-- HTML -->
		    
		    
		    <% 'Details Go Here %>
		    
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" width="5%"># Orders</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="8%">Subtotal</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%" >Shipping</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Tax</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Fuel</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Deposit</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Coupon</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="8%" >Grand Total</th>
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
				SQLDailyAPIOrders = SQLDailyAPIOrders & " WHERE        (DAY(OrderDate) = " & CurrentDay & " AND MONTH(OrderDate) =  " & CurrentMonth & " AND YEAR(OrderDate) =  " & CurrentYear & " AND (Voided = 0) AND (APIKey = '" & currentPartnerAPIKey & "')) "
				SQLDailyAPIOrders = SQLDailyAPIOrders & " AND OrderID Not In (SELECT OrderID FROM API_IN_InvoiceHeader) "
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
		    
		    
        	<h4 style="color: #3c763d; margin-top: 40px;">Invoices </h4>
			<!-- HTML -->
			
			
			
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" width="5%"># Invoices</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="8%">Subtotal</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%" >Shipping</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Tax</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Fuel</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Deposit</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Coupon</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="8%" >Grand Total</th>
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
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " WHERE        (DAY(InvoiceDate) = " & CurrentDay & " AND MONTH(InvoiceDate) =  " & CurrentMonth & " AND YEAR(InvoiceDate) =  " & CurrentYear & " AND (Voided = 0)  AND (APIKey = '" & currentPartnerAPIKey & "')) "
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


        	<h4 style="color: #3c763d; margin-top: 40px;">Return Autorizations </h4><br>
			<!-- HTML -->
		    
		    
		    <% 'Details Go Here %>
		    
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" width="5%"># RAs</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="8%">Subtotal</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%" >Shipping</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Tax</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Fuel</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Deposit</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Coupon</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="8%" >Grand Total</th>
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
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " WHERE        (DAY(RDDate) = " & CurrentDay & " AND MONTH(RDDate) =  " & CurrentMonth & " AND YEAR(RDDate) =  " & CurrentYear & " AND (Voided = 0) AND (APIKey = '" & currentPartnerAPIKey & "')) "
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

        	<h4 style="color: #3c763d; margin-top: 40px;">Credit Memos </h4><br>
			<!-- HTML -->
		        
		    
		    <% 'Details Go Here %>
		    
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" width="5%"># CMs</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="8%">Subtotal</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%" >Shipping</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Tax</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Fuel</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Deposit</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Coupon</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="8%" >Grand Total</th>
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
				SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " WHERE        (DAY(CMDate) = " & CurrentDay & " AND MONTH(CMDate) =  " & CurrentMonth & " AND YEAR(CMDate) =  " & CurrentYear & " AND (Voided = 0)  AND (APIKey = '" & currentPartnerAPIKey & "')) "
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
		    
		    
        	<h4 style="color: #3c763d; margin-top: 40px;">Summary Invoices</h4><br>
			<!-- HTML -->
		        
		    <% 'Details Go Here %>
		    
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" width="5%"># Sum Inv</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="8%">Subtotal</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%" >Shipping</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Tax</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Fuel</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Deposit</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="10%">Coupon</th>
		                <th style="padding-top: 8px; text-align: right;" align="right" width="8%" >Grand Total</th>
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
				SQLDailyAPISummInv = SQLDailyAPISummInv & " WHERE        (DAY(SumInvDate) = " & CurrentDay & " AND MONTH(SumInvDate) =  " & CurrentMonth & " AND YEAR(SumInvDate) =  " & CurrentYear & " AND (Voided = 0)  AND (APIKey = '" & currentPartnerAPIKey & "')) "
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
%>		    

	<hr>

<!--#include file="../../inc/footer-main.asp"-->


 