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

mark {
    background-color: yellow;
    color: black;
}

</style>

<!-- Resources -->

<%
OffSetFromToday = 2
currentDay = day(date()) - OffSetFromToday 
currentMonth = month(date())
currentYear = year(date())

%>


<h1 class="page-header"><i class="fa fa-fw fa-calendar"></i> Daily API Activity Detail By Partner for <%= FormatDateTime(DateAdd("d",-OffSetFromToday,Date()),2) %></h1>

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
		                <th style="padding-top: 8px; text-align: right;"  align="right" width="5%">Order ID</th>
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
				
				SQLDailyAPIOrders = "SELECT * FROM API_OR_OrderHeader "
				SQLDailyAPIOrders = SQLDailyAPIOrders & " WHERE (DAY(OrderDate) = " & CurrentDay & " AND MONTH(OrderDate) =  " & CurrentMonth & " AND YEAR(OrderDate) =  " & CurrentYear & " AND "
				SQLDailyAPIOrders = SQLDailyAPIOrders & " (Voided = 0) AND (APIKey = '" & currentPartnerAPIKey & "')) "
				SQLDailyAPIOrders = SQLDailyAPIOrders & " AND OrderID Not In (SELECT OrderID FROM API_IN_InvoiceHeader) "
				SQLDailyAPIOrders = SQLDailyAPIOrders & "  Order By BaseOrderID"
	
				
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
						
						%>
						<tr>
							<% If LastOrderID <> 0 Then
								If LastOrderID + 1 <> CDbl(rsDailyAPIOrders("BaseOrderID")) Then %>
									 <td style="padding-top: 8px; text-align: right;" align="right"><mark><%= OrderID %></mark></td>
								<% Else %>
									<td style="padding-top: 8px; text-align: right;" align="right"><%= OrderID %></td>
								<% End If %>
							<% Else %>
								<td style="padding-top: 8px; text-align: right;" align="right"><%= OrderID %></td>
							<% End If %>
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
						
						LastOrderID = CDbl(rsDailyAPIOrders("BaseOrderID"))
						
						rsDailyAPIOrders.MoveNext
					Loop
				Else
					%><tr ><td colspan="8">No Order API Data</td></tr><%
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
			    %>
				
				<tr style="border-top: 2px solid #ddd;">
	                <td style="padding-top: 8px;text-align: right;" align="right"><strong>Count:&nbsp;&nbsp;<%= DailyCount %></strong></td>
	                <td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailySubtotal %></strong></td>
	                <td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyShipTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyTaxTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyFuelTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyDepositTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyCouponTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyGranTot %></strong></td>
	            </tr>
		        </tbody>
		    </table>

		    
		    
		    <%'*********************************
			'		    I N V O I C E S 
		    '*********************************%>
		    	    
        	<h4 style="color: #3c763d; margin-top: 40px;">Invoices </h4>
			<!-- HTML -->
			
			
			
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" width="5%">Invoice ID</th>
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

				SQLDailyAPIInvoices = "SELECT * FROM API_IN_InvoiceHeader "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & " WHERE (DAY(InvoiceDate) = " & CurrentDay & " AND MONTH(InvoiceDate) =  " & CurrentMonth & " AND YEAR(InvoiceDate) =  "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & CurrentYear & " AND (Voided = 0)  AND (APIKey = '" & currentPartnerAPIKey & "')) "
				SQLDailyAPIInvoices = SQLDailyAPIInvoices & "ORDER BY InvoiceID"

				Set cnnDailyAPIInvoices = Server.CreateObject("ADODB.Connection")
				cnnDailyAPIInvoices.open(Session("ClientCnnString"))
				Set rsDailyAPIInvoices = Server.CreateObject("ADODB.Recordset")
				rsDailyAPIInvoices.CursorLocation = 3 
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
						
						%>
						<tr>
							<% If LastInvoiceID <> 0 Then
								If LastInvoiceID + 1 <> CDbl(rsDailyAPIInvoices("InvoiceID")) Then %>
									 <td style="padding-top: 8px; text-align: right;" align="right"><mark><%= InvoiceID %></mark></td>
								<% Else %>
									<td style="padding-top: 8px; text-align: right;" align="right"><%= InvoiceID %></td>
								<% End If %>
							<% Else %>
								<td style="padding-top: 8px; text-align: right;" align="right"><%= InvoiceID %></td>
							<% End If %>
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
						
						LastInvoiceID = CDbl(rsDailyAPIInvoices("InvoiceID"))
						
						rsDailyAPIInvoices.MoveNext
					Loop
				Else
					%><tr><td colspan="8">No Invoice API Data</td></tr><%
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
			    %>

				<tr style="border-top: 2px solid #ddd;">
	                <td style="padding-top: 8px;text-align: right;" align="right"><strong>Count:&nbsp;&nbsp;<%= DailyCount %></strong></td>
	                <td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailySubtotal %></strong></td>
	                <td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyShipTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyTaxTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyFuelTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyDepositTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyCouponTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyGranTot %></strong></td>
	            </tr>
		        </tbody>
		    </table>



		    
		    
		    <%'****************************************
			' R E T U R N   A U T H O R I Z A T I O N S
		    '****************************************%>

        	<h4 style="color: #3c763d; margin-top: 40px;">Return Autorizations </h4><br>
			<!-- HTML -->
		    
		    
		    <% 'Details Go Here %>
		    
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" width="5%">RA ID</th>
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
				
				SQLDailyAPIReturnAuths = "SELECT * FROM API_OR_RAHeader "
				SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " WHERE (DAY(RDDate) = " & CurrentDay & " AND MONTH(RDDate) =  " & CurrentMonth & " AND YEAR(RDDate) =  " & CurrentYear & " AND "
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
						
						%>
						<tr>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= RAID %></td>
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
						
						'LastRAID = CDbl(rsDailyAPIReturnAuths("RAID"))
						
						rsDailyAPIReturnAuths.MoveNext
					Loop
				Else
					%><tr><td colspan="8">No Return Authorization API Data</td></tr><%
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
			    %>

				
				<tr style="border-top: 2px solid #ddd;">
	                <td style="padding-top: 8px;text-align: right;" align="right"><strong>Count:&nbsp;&nbsp;<%= DailyCount %></strong></td>
	                <td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailySubtotal %></strong></td>
	                <td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyShipTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyTaxTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyFuelTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyDepositTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyCouponTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyGranTot %></strong></td>
	            </tr>
		        </tbody>
		    </table>



		    
		    <%'*********************************
			'		C R E D I T  M E M O S 
		    '*********************************%>

        	<h4 style="color: #3c763d; margin-top: 40px;">Credit Memos </h4><br>
			<!-- HTML -->
		        
		    
		    <% 'Details Go Here %>
		    
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" width="5%">CM ID</th>
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
				
				SQLDailyAPICreditMemos= "SELECT * FROM API_IN_CMHeader "
				SQLDailyAPICreditMemos= SQLDailyAPICreditMemos& " WHERE (DAY(CMDate) = " & CurrentDay & " AND MONTH(CMDate) =  " & CurrentMonth & " AND YEAR(CMDate) =  " & CurrentYear & " AND "
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
						
						Subtotal = rsDailyAPICreditMemos("SubTotal")
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
						
						%>
						<tr>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= CMID %></td>
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
			    %>
			    
				<tr style="border-top: 2px solid #ddd;">
	                <td style="padding-top: 8px;text-align: right;" align="right"><strong>Count:&nbsp;&nbsp;<%= DailyCount %></strong></td>
	                <td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailySubtotal %></strong></td>
	                <td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyShipTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyTaxTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyFuelTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyDepositTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyCouponTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyGranTot %></strong></td>
	            </tr>
		        </tbody>
		    </table>

		    
		    
		    <%'*********************************
			'		C R E D I T  M E M O S 
		    '*********************************%>

        	<h4 style="color: #3c763d; margin-top: 40px;">Summary Invoices</h4><br>
			<!-- HTML -->
		        
		    <% 'Details Go Here %>
		    
			<table style="margin-left:50px;width:1000px;">	
		        <thead>
		            <tr style="border-bottom: 2px solid #ddd;">
		                <th style="padding-top: 8px; text-align: right;"  align="right" width="5%">Sum Inv ID</th>
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
				
				SQLDailyAPISummInv= "SELECT * FROM API_IN_SummaryInvoiceHeader "
				SQLDailyAPISummInv= SQLDailyAPISummInv & " WHERE (DAY(SumInvDate) = " & CurrentDay & " AND MONTH(SumInvDate) =  " & CurrentMonth & " AND YEAR(SumInvDate) =  " & CurrentYear & " AND "
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
						
						%>
						<tr>
							<td style="padding-top: 8px; text-align: right;" align="right"><%= CMID %></td>
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
			    %>
			    	<tr style="border-top: 2px solid #ddd;">
	                <td style="padding-top: 8px;text-align: right;" align="right"><strong>Count:&nbsp;&nbsp;<%= DailyCount %></strong></td>
	                <td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailySubtotal %></strong></td>
	                <td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyShipTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyTaxTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyFuelTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyDepositTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyCouponTot %></strong></td>
					<td style="padding-top: 8px; text-align: right;" align="right"><strong><%= DailyGranTot %></strong></td>
	            </tr>
		        </tbody>
		    </table>

<%

	rsDailyAPIPartnersLoop.MoveNext
Loop
End If
%>		    

	<hr>

<!--#include file="../../inc/footer-main.asp"-->


 