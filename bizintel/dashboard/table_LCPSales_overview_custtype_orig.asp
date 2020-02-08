<div style="width: 100%; overflow-y: scroll; margin: 0 auto">
	
<% 


 	'Get all Custmoer Types
 	
	TotSalesTyp = 0
	Tot3PAvgTyp = 0
	TotDollarDiff = 0
	TotalNegDiff = 0
	Tot12PAvgTyp = 0
	TotDollarDiff12 = 0
	TotalNegDiff12 = 0

	GrandTotalLCPSales = 0
	GrandTotal3PAvgSales = 0
	GrantTotal12PAvgSales = 0
	GrandTotalSPLYSales = 0
	GrandTotalLCPvs3PAvg = 0
	GrandTotalLCPvs12PAvg = 0
	GrandTotalLCPvsSPLY = 0


	LeftOverTyp = ""

	SQL = "SELECT CustType "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData  "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
	SQL = SQL & " GROUP BY AR_Customer.CustType"
	SQL = SQL & " EXCEPT "
	SQL = SQL & " SELECT CustType "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData  "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY AR_Customer.CustType"

	Set rs = cnn8.Execute(SQL)
	If Not rs.EOF Then
		Do While Not rs.EOF
				LeftOverTyp = LeftOverTyp & rs("CustType") & ","
			rs.MoveNext
		Loop
	End IF

	If Right(LeftOverTyp ,1)="," Then LeftOverTyp = Left(LeftOverTyp ,len(LeftOverTyp )-1)
	
'Response.Write(SQL&"<br>")	
	Set rs = cnn8.Execute(SQL)

	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3) As Tot3PPAvg"
	SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
	SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg"
	SQL = SQL & ",SUM(PriorPeriod1Sales) As TotPP1Sales"
	SQL = SQL & ",SUM(PriorPeriod2Sales) As TotPP2Sales"
	SQL = SQL & ",SUM(PriorPeriod3Sales) As TotPP3Sales"
	SQL = SQL & ",CustType  "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData  "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	If LimitSelection = 1 Then
		SQL = SQL & " AND TotalSales < [3PriorPeriodsAeverage] "
		SQL = SQL & " AND [3PriorPeriodsAeverage] - TotalSales > " & FilterSalesDollars 
		SQL = SQL & " AND (CASE WHEN [3PriorPeriodsAeverage] <> 0 THEN (((TotalSales - [3PriorPeriodsAeverage] ) / [3PriorPeriodsAeverage]) * 100) * -1 END) >= " & FilterPercentage 
	End If
	SQL = SQL & " GROUP BY AR_Customer.CustType"
	SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"


	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then

		%>		
		<br>
		<table id="tableSuperSumCustType" class="display compact" style="width:100%;">
		
			<thead>
				<tr>	
					<th colspan="1" class="td-align1" width='16%' style="border-right: 2px solid #555 !important;">&nbsp;</th>
					<th colspan="4" class="td-align1 cust-type-color" width='48%' style="border-right: 2px solid #555 !important;">Sales</th>
					<th colspan="3" class="td-align1 referral-color" width='36%' style="border-right: 2px solid #555 !important;">Variance</th>
				</tr>
			  	<tr>
					<th class="td-align1 dollar-amount-header width='16%'" style="border-right: 2px solid #555 !important;">Customer Type</th>
					<th class="td-align1 dollar-amount-header width='12%'" style="border-right: 2px solid #555 !important;">Period&nbsp;<%=LCP_Display_Month%><br>Sales</th>
					<th class="td-align1 dollar-amount-header width='12%'" style="border-right: 2px solid #555 !important;">3 Prior Periods<br>Average Sales</th>
					<th class="td-align1 dollar-amount-header width='12%'" style="border-right: 2px solid #555 !important;">12 Prior Periods<br>Average Sales</th>
					<th class="td-align1 dollar-amount-header width='12%'" style="border-right: 2px solid #555 !important;">Same Period<br>Last Year</th>
					<th class="td-align1 dollar-amount-header width='12%'" style="border-right: 2px solid #555 !important;">Period&nbsp;<%=LCP_Display_Month%> vs<br>3 Prior Periods Average</th>
					<th class="td-align1 dollar-amount-header width='12%'" style="border-right: 2px solid #555 !important;">Period&nbsp;<%=LCP_Display_Month%> vs<br>12 Prior Periods Average</th>
					<th class="td-align1 dollar-amount-header width='12%'" style="border-right: 2px solid #555 !important;">Period&nbsp;<%=LCP_Display_Month%> vs<br>Same Period Last Year</th>
				
				</tr>
			</thead>
			
			<tbody>
				<%													
				
				ChartElementNumber = 1
				ChartDataCustType = ""
				ChartRemainder = 100
				NextPeriodProj = 0
				
				Do
				
					Response.Write("<tr>")
				    Response.Write("<td align='left' class='smaller-detail-line'><a href='dashboard_segment_customertype.asp?t=" & rs("CustType") & "' target='_blank'>"& rs("CustType") & " - " & GetCustTypeByCode(rs("CustType")) & "</a></td>")									
													    
					Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line'>" & FormatCurrency(rs("TotSales"),0,-2,0) & "</td>")
					
					Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line'>" & FormatCurrency(rs("Tot3PPAvg"),0,-2,0) & "</td>")
					
					Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line'>" & FormatCurrency(rs("Tot12PPAvg"),0,-2,0) & "</td>")
					
					
					Type2Find = rs("CustType")
					
					'Now get all the SPLY Numbers
					SQL2 = "SELECT SUM(TotalSales) AS SPLYTotalSales "
					SQL2 = SQL2 & " FROM CustCatPeriodSales "
					SQL2 = SQL2 & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales.CustNum"
					SQL2 = SQL2 & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
					SQL2 = SQL2 & " AND AR_Customer.CustType= '" & Type2Find & "' "
					If LimitSelection = 1 Then
						SQL2 = SQL2 & " AND TotalSales < [3PriorPeriodsAeverage] "
						SQL2 = SQL2 & " AND [3PriorPeriodsAeverage] - TotalSales > " & FilterSalesDollars 
						SQL2 = SQL2 & " AND (CASE WHEN [3PriorPeriodsAeverage] <> 0 THEN (((TotalSales - [3PriorPeriodsAeverage] ) / [3PriorPeriodsAeverage]) * 100) * -1 END) >= " & FilterPercentage 
					End If

		
					Set rs2 = cnn8.Execute(SQL2)
					If not rs2.EOF Then SPLYTotalSales  = rs2("SPLYTotalSales") Else SPLYTotalSales= 0

				
					Response.Write("<td align='" & ColumnAlign & "'width='12%' class='smaller-detail-line'>" & FormatCurrency(SPLYTotalSales,0,-2,0) & "</td>")
					
					DollarDiffSPLY = rs("TotSales") - SPLYTotalSales

					
					DollarDiff = rs("TotSales") - rs("Tot3PPAvg")
					
					If DollarDiff > 0 Then
						Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line'>" & FormatCurrency(DollarDiff,0,-2,0) & "</td>")
					Else
						Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line negative'>" & FormatCurrency(DollarDiff,0,-2,0) & "</td>")
					End If
					
				
					DollarDiff12 = rs("TotSales") - rs("Tot12PPAvg")
					
					If DollarDiff12 > 0 Then
						Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line'>" & FormatCurrency(DollarDiff12,0,-2,0) & "</td>")
					Else
						Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line negative'>" & FormatCurrency(DollarDiff12,0,-2,0) & "</td>")
					End If
					
					If DollarDiffSPLY > 0 Then
						Response.Write("<td align='" & ColumnAlign & "'width='12%' class='smaller-detail-line'>" & FormatCurrency(DollarDiffSPLY,0,-2,0) & "</td>")
					Else
						Response.Write("<td align='" & ColumnAlign & "'width='12%' class='smaller-detail-line negative'>" & FormatCurrency(DollarDiffSPLY,0,-2,0) & "</td>")
					End If
					
					
					
					Response.Write("</tr>")

					GrandTotalLCPSales = GrandTotalLCPSales + rs("TotSales")
					GrandTotal3PAvgSales = GrandTotal3PAvgSales + rs("Tot3PPAvg")
					GrantTotal12PAvgSales = GrantTotal12PAvgSales + rs("Tot12PPAvg")
					GrandTotalSPLYSales = GrandTotalSPLYSales + SPLYTotalSales
					
					GrandTotalLCPvs3PAvg = GrandTotalLCPSales - GrandTotal3PAvgSales 
					
					GrandTotalLCPvs12PAvg = GrandTotalLCPSales - GrantTotal12PAvgSales 
					
					GrandTotalLCPvsSPLY = GrandTotalLCPSales - GrandTotalSPLYSales 
					
					rs.movenext
				Loop until rs.eof
				
			End If

			
					
							        		


		'***********
		'***********
		' LEFT OVERS
		'***********
		'***********
		If LeftOverTyp <> "" Then
	      	'This part is a little crazy but now we have to do the left over customer type
			'Now get all the SPLY Numbers for the leftovers
			SQL2 = "SELECT SUM(TotalSales) AS SPLYTotalSales, CustType "
			SQL2 = SQL2 & " FROM CustCatPeriodSales "
			SQL2 = SQL2 & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales.CustNum"
			SQL2 = SQL2 & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
			SQL2 = SQL2 & " AND AR_Customer.CustType IN (" & LeftOverTyp & ") GROUP BY AR_Customer.CustType "
	
			Set rs = cnn8.Execute(SQL2)
			If Not rs.EOF Then
				Do While Not rs.EOF
				
					If rs("SPLYTotalSales") <> 0 Then 
								
						Response.Write("<tr>")
					    Response.Write("<td align='left' class='smaller-detail-line'><a href='dashboard_segment_customertype.asp?t=" & rs("CustType") & "' target='_blank'>"& rs("CustType") & " - " & GetCustTypeByCode(rs("CustType")) & "</a></td>")									
						
						Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
						Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
						Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
		
						SPLYTotalSales  = rs("SPLYTotalSales") 
		
						If SPLYTotalSales > 0 Then
							Response.Write("<td align='" & ColumnAlign & "'width='12%' class='smaller-detail-line'>" & FormatCurrency(SPLYTotalSales,0,-2,0) & "</td>")
						Else
							Response.Write("<td align='" & ColumnAlign & "'width='12%' class='smaller-detail-line negative'>" & FormatCurrency(SPLYTotalSales,0,-2,0) & "</td>")
						End If
						
											
						Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
						Response.Write("<td align='" & ColumnAlign & "'class='smaller-detail-line'>" & FormatCurrency(0,0,-2,0) & "</td>")
					
						If SPLYTotalSales * -1 > 0 Then
							Response.Write("<td align='" & ColumnAlign & "'width='12%' class='smaller-detail-line'>" & FormatCurrency(SPLYTotalSales * -1,0,-2,0) & "</td>")				
						Else
							Response.Write("<td align='" & ColumnAlign & "'width='12%' class='smaller-detail-line negative'>" & FormatCurrency(SPLYTotalSales * -1,0,-2,0) & "</td>")
						End If
						
						Response.Write("</tr>")
					
						GrandTotalSPLYSales = GrandTotalSPLYSales + SPLYTotalSales 			
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
				<td align="center" width='16%'><strong>Totals</strong></td>
				<td align="<%= ColumnAlign %>" width='12%'><strong><%= FormatCurrency(GrandTotalLCPSales,0,-2,0) %></strong></td>
				<td align="<%= ColumnAlign %>" width='12%'><strong><%= FormatCurrency(GrandTotal3PAvgSales,0,-2,0) %></strong></td>
				<td align="<%= ColumnAlign %>" width='12%'><strong><%= FormatCurrency(GrantTotal12PAvgSales,0,-2,0) %></strong></td>
				<td align="<%= ColumnAlign %>" width='12%'><strong><%= FormatCurrency(GrandTotalSPLYSales,0,-2,0) %></strong></td>
				<% If GrandTotalLCPvs3PAvg < 0 Then %>
					<td align="<%= ColumnAlign %>" class="negative" width='12%'><strong><%= FormatCurrency(GrandTotalLCPvs3PAvg,0,-2,0) %></strong></td>
				<% Else %>
					<td align="<%= ColumnAlign %>" width='12%'><strong><%= FormatCurrency(GrandTotalLCPvs3PAvg,0,-2,0) %></strong></td>
				<% End If %>
				<% If GrandTotalLCPvs12PAvg < 0 Then %>
					<td align="<%= ColumnAlign %>" class="negative" width='12%'><strong><%= FormatCurrency(GrandTotalLCPvs12PAvg,0,-2,0) %></strong></td>
				<% Else %>
					<td align="<%= ColumnAlign %>" width='12%'><strong><%= FormatCurrency(GrandTotalLCPvs12PAvg,0,-2,0) %></strong></td>
				<% End If %>
				<% If GrandTotalLCPvsSPLY < 0 Then %>
					<td align="<%= ColumnAlign %>" class="negative" width='12%'><strong><%= FormatCurrency(GrandTotalLCPvsSPLY,0,-2,0) %></strong></td>
				<% Else %>
					<td align="<%= ColumnAlign %>" width='12%'><strong><%= FormatCurrency(GrandTotalLCPvsSPLY,0,-2,0) %></strong></td>			
				<% End IF %>
			</tr>
		</tfoot>
	
	</table><br>	
</div>	

