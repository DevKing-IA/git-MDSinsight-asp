<div style="width: 100%; overflow-y: scroll; margin: 0 auto">
		
 	<%

 	'Get all Referral Codes
 	
	TotProductSalesReferral = 0
	Tot3PAvgReferral = 0
	TotDollarDiff =0
	TotalNegDiff = 0
	Tot12PAvgReferral = 0
	TotDollarDiff12 = 0
	TotalNegDiff12 = 0

	GrandTotalLCPSales = 0 : GrandTotal3PAvgSales = 0 : GrantTotal12PAvgSales = 0 : GrandTotalSPLYSales = 0
	GrandTotalLCPADS = 0 : GrandTotal3PPADS = 0 : GrandTotal12PPADS = 0 : GrandTotalSPLYADS = 0 : GrandTotalCPADS = 0 : GrandTotalCPLYADS = 0
	GrandTotalLCPvs3PAvgADS = 0 : GrandTotalLCPvs12PAvg = 0 : GrandTotalLCPvsSPLY = 0 : GrandTotalCPvsvs3PAvgADS = 0 : GrandTotalCPLYvsvs3PAvgADS = 0

	LeftOverReferral = ""

	SQL = "SELECT ReferralDesc2 "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	SQL = SQL & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
	SQL = SQL & " GROUP BY ReferralDesc2 "
	SQL = SQL & " EXCEPT "
	SQL = SQL & " SELECT ReferralDesc2 "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY ReferralDesc2 "

'Response.Write(SQL)

	Set rs = cnn8.Execute(SQL)
	If Not rs.EOF Then
		Do While Not rs.EOF
				LeftOverReferral = LeftOverReferral & rs("ReferralDesc2") & ","
			rs.MoveNext
		Loop
	End IF

	If Right(LeftOverReferral,1)="," Then LeftOverReferral = Left(LeftOverReferral,len(LeftOverReferral)-1)
	
'Response.Write(SQL&"<br>")	
	Set rs = cnn8.Execute(SQL)

	SQL = "SELECT * FROM BI_DashBoard WHERE Segment='Referral'"


	Set rs = cnn8.Execute(SQL)

	If not rs.EOF Then

		%>
			
		<table id="tableSuperSumReferral" class="display compact" style="width:100%;">
	
		<thead>
		
			<tr id="referralTopHeaderRow">	
				<th id="referralTopHeaderRow1" rowspan="2" colspan="1" style="width:10% !important; border-right:2px solid #555 !important;border-bottom:2px solid #555 !important;">Referral Code</th>
				<th id="referralTopHeaderRow2" colspan="4" class="td-align1 cust-type-color" style="width:24% !important; border-right:2px solid #555 !important;border-bottom:2px solid #555 !important;">Sales</th>
				<th id="referralTopHeaderRow4" colspan="6" class="td-align1 gen-info-header" style="width:36% !important; border-right:2px solid #555 !important;border-bottom:2px solid #555 !important;">Average Daily Sales</th>
				<th id="referralTopHeaderRow3" colspan="5" class="td-align1 referral-color" style="width:30% !important; border-right:2px solid #555 !important;border-bottom:2px solid #555 !important;">Average Daily Sales Variances</th>
			</tr>
		
		  	<tr id="referralBottomHeaderRow">
				<th id="referralBottomHeaderRow1" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">Period&nbsp;<%=LCP_Display_Month%><br>Sales</th>
				<th id="referralBottomHeaderRow2" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">3 Prior Periods<br>Average Sales</th>
				<th id="referralBottomHeaderRow3" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">12 Prior Periods<br>Average Sales</th>
				<th id="referralBottomHeaderRow4" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">Same Period<br>Last Year</th>
				
				<th id="referralBottomHeaderRow5" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">Period&nbsp;<%=LCP_Display_Month%><br></th>
				<th id="referralBottomHeaderRow6" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">3 Prior Periods<br>Average</th>
				<th id="referralBottomHeaderRow7" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">12 Prior Periods<br>Average</th>
				<th id="referralBottomHeaderRow8" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">Same Period<br>Last Year</th>
				<th id="referralBottomHeaderRow9" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">Current Period<br>(So Far)</th>
				<%
				PeriodDisplayVar = ""
				If LCP_Display_Month <> 12 Then
					PeriodDisplayVar = LCP_Display_Month + 1
				Else
					PeriodDisplayVar = 1
				End If %>
				<th id="referralBottomHeaderRow10" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">Current<br>Last Year</th>
				
				<th id="referralBottomHeaderRow11" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">Period&nbsp;<%=LCP_Display_Month%> vs<br>3 Prior Periods Average</th>
				<th id="referralBottomHeaderRow12" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">Period&nbsp;<%=LCP_Display_Month%> vs<br>12 Prior Periods Average</th>
				<th id="referralBottomHeaderRow13" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">Period&nbsp;<%=LCP_Display_Month%> vs<br>Same Period Last Year</th>
				<th id="referralBottomHeaderRow14" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">Current vs 3<br>Prior Periods Average</th>
				<th id="referralBottomHeaderRow15" class="dollar-amount-header" style="border-right: 2px solid #555 !important;">Current vs<br>Current<br>Last Year</th>
			</tr>
		</thead>
		
		<tbody>
		
			<%
	
			ChartElementNumber = 1
			ChartDataReferral = ""
			ChartRemainder = 100
			NextPeriodProj = 0

									
			Do
			


				Response.Write("<tr>")
				
				If rs("SalesLCP") + rs("Sales3PPAvg")  < 1 Then
					Response.Write("<td align='left' class='smaller-detail-line'>" & rs("ReferralCodeNumber") & " - " & rs("ReferralCodeDesc2") & "</td>")													
				Else
					Response.Write("<td align='left' class='smaller-detail-line'><a href='dashboard_segment_secondarysalesman.asp?p=" & rs("ReferralCodeDesc2") & "' target='_blank'>"& rs("ReferralCodeNumber") & " - " & rs("ReferralCodeDesc2") & "</a></td>")									
				End IF
				
					Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(rs("SalesLCP") ,0,-2,0) & "</td>")
				
				Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(rs("Sales3PPAvg") ,0,-2,0) & "</td>")
				
				Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(rs("Sales12PPAvg"),0,-2,0) & "</td>")

				Response.Write("<td  class='smaller-detail-line' style='border-right: 2px solid #555 !important;'>" & FormatCurrency(rs("SalesSPLY"),0,-2,0) & "</td>")
				
				Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(rs("ADSLCP"),0,-2,0) & "</td>")
				Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(rs("ADS3PPAvg"),0,-2,0) & "</td>")
				Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(rs("ADS12PPAvg"),0,-2,0) & "</td>")
				Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(rs("ADSSPLY"),0,-2,0) & "</td>")

				Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(rs("ADSCP"),0,-2,0) & "</td>")

				Response.Write("<td  class='smaller-detail-line' style='border-right: 2px solid #555 !important;'>" & FormatCurrency(rs("ADSCPLY"),0,-2,0) & "</td>")	

				If rs("VARLCPv3ppAvg") > 0 Then
					Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(rs("VARLCPv3ppAvg"),0,-2,0) & "</td>")
				Else
					Response.Write("<td  class='smaller-detail-line negative'>" & FormatCurrency(rs("VARLCPv3ppAvg"),0,-2,0) & "</td>")
				End If
			
				If rs("VARLCPv12PPAvg") > 0 Then
					Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(rs("VARLCPv12PPAvg") ,0,-2,0) & "</td>")
				Else
					Response.Write("<td  class='smaller-detail-line negative'>" & FormatCurrency(rs("VARLCPv12PPAvg") ,0,-2,0) & "</td>")
				End If

				If rs("VARLCPvSPLY")  > 0 Then
					Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(rs("VARLCPvSPLY"),0,-2,0) & "</td>")
				Else
					Response.Write("<td  class='smaller-detail-line negative'>" & FormatCurrency(rs("VARLCPvSPLY"),0,-2,0) & "</td>")
				End If

				If rs("VARCPv3PPAvg") > 0 Then
					Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(rs("VARCPv3PPAvg"),0,-2,0) & "</td>")
				Else
					Response.Write("<td  class='smaller-detail-line negative'>" & FormatCurrency(rs("VARCPv3PPAvg"),0,-2,0) & "</td>")
				End If

				If rs("VARCPvCPLY") > 0 Then
					Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(rs("VARCPvCPLY"),0,-2,0) & "</td>")
				Else
					Response.Write("<td  class='smaller-detail-line negative'>" & FormatCurrency(rs("VARCPvCPLY"),0,-2,0) & "</td>")
				End If
				
				Response.Write("</tr>")

					
				'Sales
				GrandTotalLCPSales = GrandTotalLCPSales + rs("SalesLCP")
				GrandTotal3PAvgSales = GrandTotal3PAvgSales + rs("Sales3PPAvg")
				GrantTotal12PAvgSales = GrantTotal12PAvgSales + rs("Sales12PPAvg")
				GrandTotalSPLYSales = GrandTotalSPLYSales + rs("SalesSPLY")
				
				'ADS
				GrandTotalLCPADS = GrandTotalLCPADS + rs("ADSLCP")
				GrandTotal3PPADS = GrandTotal3PPADS + rs("ADS3PPAvg")
				GrandTotal12PPADS = GrandTotal12PPADS + rs("ADS12PPAvg")
				GrandTotalSPLYADS = GrandTotalSPLYADS + rs("ADSSPLY")
				GrandTotalCPADS = GrandTotalCPADS + rs("ADSCP")
				GrandTotalCPLYADS = GrandTotalCPLYADS + rs("ADSCPLY")
				
				'ADS Variance
				GrandTotalLCPvs3PAvgADS = GrandTotalLCPADS - GrandTotal3PPADS 
				GrandTotalLCPvs12PAvg = GrandTotalLCPADS - GrandTotal12PPADS 
				GrandTotalLCPvsSPLY = GrandTotalLCPADS - GrandTotalSPLYADS 
				GrandTotalCPvsvs3PAvgADS = GrandTotalCPADS - GrandTotal3PPADS 
				GrandTotalCPLYvsvs3PAvgADS = GrandTotalCPLYADS - GrandTotal3PPADS 
					
				rs.movenext
			Loop until rs.eof
		End If



		'***********
		'***********
		' LEFT OVERS
		'***********
		'***********
		If LeftOverSLs2  <> "" Then
		
				SQL2 = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales ELSE 0 END) AS SPLYTotalSales ,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
				SQL2 = SQL2 & " FROM CustCatPeriodSales "
				SQL2 = SQL2 & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
				SQL2 = SQL2 & " AND ReferralDesc2 = '" & Sls22Find & "' "

	
			Set rs = cnn8.Execute(SQL2)
			If Not rs.EOF Then
				Do While Not rs.EOF
				
					SPLYProductsSales = rs2("SPLYTotalSales")
					SPLYTotalRentals = rs2("SPLYTotalRentals")
				
					If SPLYProductsSales <> 0 Then 
								
						Response.Write("<tr>")
						
						Response.Write("<td align='left' class='smaller-detail-line'>" & rs("ReferralCode") & " - " & rs("ReferralDesc2") & "</td>")									
						
						Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
						Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
						Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
		
		
						Response.Write("<td  class='smaller-detail-line' style='border-right: 2px solid #555 !important;'>" & FormatCurrency(SPLYProductsSales,0,-2,0) & "</td>")

						Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")
						Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")
						Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")						
						Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")
						Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")
						Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(0.00,0,-2,0) & "</td>")

					
						If SPLYProductsSales * -1 > 0 Then
							Response.Write("<td  class='smaller-detail-line'>" & FormatCurrency(SPLYProductsSales * -1,0,-2,0) & "</td>")				
						Else
							Response.Write("<td  class='smaller-detail-line negative'>" & FormatCurrency(SPLYProductsSales * -1,0,-2,0) & "</td>")							
						End If
						
						Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
						Response.Write("<td class='smaller-detail-line' style='border-right: 2px solid #555 !important;'>" & FormatCurrency(0,0,-2,0) & "</td>")
						Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
						Response.Write("<td class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
				
						Response.Write("</tr>")
					
						GrandTotalSPLYSales = GrandTotalSPLYSales + SPLYProductsSales 			
						GrandTotalLCPvsSPLY = GrandTotalLCPSales - GrandTotalSPLYSales 
	
					End If
								
					rs.MoveNext
				Loop
			End IF
		End If
		'***************
		'***************
		' END LEFT OVERS
		'***************
		'***************



	%>
	
		
		</tbody>

		<tfoot>
		  	<tr>	
				<td align="center"><strong>Totals</strong></td>
				<% ' Sales %>
				<td ><strong><%= FormatCurrency(GrandTotalLCPSales,0,-2,0) %></strong></td>
				<td ><strong><%= FormatCurrency(GrandTotal3PAvgSales,0,-2,0) %></strong></td>
				<td ><strong><%= FormatCurrency(GrantTotal12PAvgSales,0,-2,0) %></strong></td>
				<td ><strong><%= FormatCurrency(GrandTotalSPLYSales,0,-2,0) %></strong></td>
				
				<% ' ADS %>
				<td ><strong><%= FormatCurrency(GrandTotalLCPADS,0,-2,0) %></strong></td>			
				<td ><strong><%= FormatCurrency(GrandTotal3PPADS,0,-2,0) %></strong></td>			
				<td ><strong><%= FormatCurrency(GrandTotal12PPADS,0,-2,0) %></strong></td>	
				<td ><strong><%= FormatCurrency(GrandTotalSPLYADS,0,-2,0) %></strong></td>			
				<td ><strong><%= FormatCurrency(GrandTotalCPADS,0,-2,0) %></strong></td>			
				<td ><strong><%= FormatCurrency(GrandTotalCPLYADS,0,-2,0) %></strong></td>	
				
				<% ' ADS VAriance %>
				<% If GrandTotalLCPvs3PAvgADS < 0 Then %>
					<td  class="negative" ><strong><%= FormatCurrency(GrandTotalLCPvs3PAvgADS,0,-2,0) %></strong></td>
				<% Else %>
					<td ><strong><%= FormatCurrency(GrandTotalLCPvs3PAvgADS,0,-2,0) %></strong></td>
				<% End If %>
				
				<% If GrandTotalLCPvs12PAvg < 0 Then %>
					<td  class="negative"><strong><%= FormatCurrency(GrandTotalLCPvs12PAvg,0,-2,0) %></strong></td>
				<% Else %>
					<td ><strong><%= FormatCurrency(GrandTotalLCPvs12PAvg,0,-2,0) %></strong></td>
				<% End If %>
				
				<% If GrandTotalLCPvsSPLY < 0 Then %>
					<td  class="negative"><strong><%= FormatCurrency(GrandTotalLCPvsSPLY,0,-2,0) %></strong></td>
				<% Else %>
					<td ><strong><%= FormatCurrency(GrandTotalLCPvsSPLY,0,-2,0) %></strong></td>			
				<% End IF %>

				<% If GrandTotalCPvsvs3PAvgADS < 0 Then %>
					<td  class="negative"><strong><%= FormatCurrency(GrandTotalCPvsvs3PAvgADS,0,-2,0) %></strong></td>
				<% Else %>
					<td ><strong><%= FormatCurrency(GrandTotalCPvsvs3PAvgADS,0,-2,0) %></strong></td>			
				<% End IF %>

				<% If GrandTotalCPLYvsvs3PAvgADS < 0 Then %>
					<td  class="negative"><strong><%= FormatCurrency(GrandTotalCPLYvsvs3PAvgADS,0,-2,0) %></strong></td>
				<% Else %>
					<td ><strong><%= FormatCurrency(GrandTotalCPLYvsvs3PAvgADS,0,-2,0) %></strong></td>			
				<% End IF %>
				
			</tr>
		</tfoot>
	
	
	</table><br>
		
</div>	


