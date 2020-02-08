<!--<div class='table-responsive' style="border:1px #ddd solid;">-->
<%

	PeriodBeingEvaluated = GetLastClosedPeriodAndYear()
	PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum()
	
	WorkDaysIn3PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -3), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1
	WorkDaysIn12PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -12), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1)) + 1 
	WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated )) 
	WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated +1)) + 1 
	WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1),Date()) + 1
	WorkDaysInProjectionBasis =  (NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -2), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1) + WorkDaysInLastClosedPeriod
	
	LCP_Display_MonthArr = Split(GetLastClosedPeriodAndYear(),"-")
	LCP_Display_Month = Trim(LCP_Display_MonthArr(0))
	LCP_Display_Year = Trim(LCP_Display_MonthArr(1))
	
	LCP_Display_Var = "P" & LCP_Display_Month & "/" & LCP_Display_Year
	
	SQL = "SELECT BeginDate As SPLYBeginDate, EndDate As SPLYEndDate FROM BillingPeriodHistory WHERE Period = " & LCP_Display_Month & " AND Year = " & LCP_Display_Year - 1
'Response.Write(SQL)
	Set rs = cnn8.Execute(SQL)
	SPLYBeginDate = rs("SPLYBeginDate")
	SPLYEndDate = rs("SPLYEndDate")
	WorkDaysInSPLYPeriodBasis =  NumberOfWorkDays(SPLYBeginDate, SPLYEndDate)

	SQL = "SELECT BeginDate As CPLYBeginDate, EndDate As CPLYEndDate FROM BillingPeriodHistory WHERE Period = " & LCP_Display_Month+1 & " AND Year = " & LCP_Display_Year - 1
'Response.Write(SQL)
	Set rs = cnn8.Execute(SQL)
	CPLYBeginDate = rs("CPLYBeginDate")
	CPLYEndDate = rs("CPLYEndDate")
	WorkDaysInCurrentPLY =  NumberOfWorkDays(CPLYBeginDate , CPLYEndDate )

Select Case MUV_READ("LOHVAR")
	
	Case "Secondary"

		SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS CPLYTotalSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS CPLYTotalRentals "
		SQL = SQL & " FROM CustCatPeriodSales "
		SQL = SQL & " WHERE Period = " & LCP_Display_Month+1 & " AND PeriodYear = " & LCP_Display_Year - 1
		SQL = SQL & " AND SecondarySalesman = '" & FilterSlsmn2 & "'"
		Set rs = cnn8.Execute(SQL)
		
	Case "Primary"
	
		SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS CPLYTotalSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS CPLYTotalRentals "
		SQL = SQL & " FROM CustCatPeriodSales "
		SQL = SQL & " WHERE Period = " & LCP_Display_Month+1 & " AND PeriodYear = " & LCP_Display_Year - 1
		SQL = SQL & " AND PrimarySalesman = '" & FilterSlsmn1 & "'"
		Set rs = cnn8.Execute(SQL)

	Case "CustType"
	
		SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS CPLYTotalSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS CPLYTotalRentals "
		SQL = SQL & " FROM CustCatPeriodSales "
		SQL = SQL & " WHERE Period = " & LCP_Display_Month+1 & " AND PeriodYear = " & LCP_Display_Year - 1
		SQL = SQL & " AND CustType = '" & CustType & "'"
		Set rs = cnn8.Execute(SQL)
	
End Select
	
	TotSalesCPLY = rs("CPLYTotalSales")
	TotRentalsCPLY = rs("CPLYTotalRentals")

	'Current Period Last Year
	'See if it has a decimal > 5
	If WorkDaysInCurrentPLY  - Int(WorkDaysInCurrentPLY) < .5 Then
		WorkDaysInCurrentPLY = Int(WorkDaysInCurrentPLY)
	Else
		WorkDaysInCurrentPLY = Int(WorkDaysInCurrentPLY) + .5
	End If

	Tot_CPLYPADS = TotSalesCPLY / WorkDaysInCurrentPLY 

	TotSalesHeader = 0
	Tot3PAvgHeader = 0
	TotDollarDiff = 0
	TotalNegDiff = 0
	TotDollarDiff12 = 0
	TotalNegDiff12 = 0
	Tot12PAvgHeader = 0
	SPLYTotalSales  = 0

Select Case MUV_READ("LOHVAR")
	
	Case "Secondary"
	
		'Now get all the current info that we need
		SQL = "SELECT SUM(TotalSales) AS TotSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS TotRentals "
		SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg, SUM(CASE WHEN Category = 0 THEN [3PriorPeriodsAeverage] END) AS Tot3PPAvgRentals "
		SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
		SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg "
		SQL = SQL & ",SUM( CASE WHEN Category = 0 THEN( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
		SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12 END) As Tot12PPAvgRentals "		
		SQL = SQL & " FROM CustCatPeriodSales_ReportData "
		SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
		SQL = SQL & " AND SecondarySalesman = '" & FilterSlsmn2 & "'"

	Case "Primary"
	
		'Now get all the current info that we need
		SQL = "SELECT SUM(TotalSales) AS TotSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS TotRentals "
		SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg, SUM(CASE WHEN Category = 0 THEN [3PriorPeriodsAeverage] END) AS Tot3PPAvgRentals "
		SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
		SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg "
		SQL = SQL & ",SUM( CASE WHEN Category = 0 THEN( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
		SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12 END) As Tot12PPAvgRentals "		
		SQL = SQL & " FROM CustCatPeriodSales_ReportData "
		'SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum "
		SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
		SQL = SQL & " AND PrimarySalesman = '" & FilterSlsmn1 & "'"

	Case "CustType"
	
		'Now get all the current info that we need
		SQL = "SELECT SUM(TotalSales) AS TotSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS TotRentals "
		SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg, SUM(CASE WHEN Category = 0 THEN [3PriorPeriodsAeverage] END) AS Tot3PPAvgRentals "
		SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
		SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg "
		SQL = SQL & ",SUM( CASE WHEN Category = 0 THEN( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
		SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12 END) As Tot12PPAvgRentals "		
		SQL = SQL & " FROM CustCatPeriodSales_ReportData "
		'SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum "
		SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
		SQL = SQL & " AND CustType = '" & CustType & "'"
		
		
End Select		

'Response.write("<br><br><br>"&SQL&"<br>")

	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		'Need to get totals first

		If rs("TotSales") - rs("Tot3PPAvg") < 0 Then TotalNegDiff = (rs("TotSales") - rs("Tot3PPAvg"))

		TotDollarDiff = ( rs("TotSales") - rs("Tot3PPAvg")) 
		TotSalesHeader = rs("TotSales")
		TotRentalsHeader = rs("TotRentals")
		Tot3PAvgHeader = rs("Tot3PPAvg")
		Tot3PAvgHeaderRentals = rs("Tot3PPAvgRentals")
		Tot12PAvgHeaderRentals = rs("Tot12PPAvgRentals")
		
		If rs("TotSales") - rs("Tot12PPAvg") < 0 Then TotalNegDiff12 = (rs("TotSales") - rs("Tot12PPAvg"))
		
		TotDollarDiff12 = ( rs("TotSales") - rs("Tot12PPAvg")) 
		Tot12PAvgHeader = rs("Tot12PPAvg")

		If Not IsNumeric(TotDollarDiff12) Then TotDollarDiff12 = 0		
		If Not IsNumeric(TotDollarDiff) Then TotDollarDiff = 0
		If Not IsNumeric(Tot12PAvgHeader) Then Tot12PAvgHeader = 0	
		If Not IsNumeric(TotSalesHeader) Then TotSalesHeader = 0
		If Not IsNumeric(Tot3PAvgHeader) Then Tot3PAvgHeader = 0
		
	End If

	Tot_CurrentADS = 0

	'Current Period
	WD_CurrentSoFar = WorkDaysSoFar 
	WD_CurrentPeriod = WorkDaysInCurrentPeriod  
	'See if it has a decimal > 5
	If WD_SPLYPADS - Int(WD_SPLYPADS ) < .5 Then
		WD_SPLYPADS = Int(WD_P12PADS )
	Else
		WD_SPLYPADS = Int(WD_SPLYPADS ) + .5
	End If

	
	Select Case MUV_READ("LOHVAR")
		
		Case "Secondary"
	
			'Now get all the SPLY Numbers
			SQL = "SELECT SUM(TotalSales) AS SPLYTotalSales,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
			SQL = SQL & " FROM CustCatPeriodSales "
			SQL = SQL & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
			SQL = SQL & " AND SecondarySalesman = '" & FilterSlsmn2 & "'"
		
			Tot_CurrentADS = (GetCurrentPeriod_PostedTotalSls2(FilterSlsmn2) + GetCurrentPeriod_UnPostedTotalSls2(FilterSlsmn2)) / WD_CurrentSoFar

		Case "Primary"
	
			'Now get all the SPLY Numbers
			SQL = "SELECT SUM(TotalSales) AS SPLYTotalSales,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
			SQL = SQL & " FROM CustCatPeriodSales "
			SQL = SQL & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
			SQL = SQL & " AND PrimarySalesman = '" & FilterSlsmn1 & "'"
		
			Tot_CurrentADS = (GetCurrentPeriod_PostedTotalSls2(FilterSlsmn1) + GetCurrentPeriod_UnPostedTotalSls2(FilterSlsmn1)) / WD_CurrentSoFar

		Case "CustType"
	
			'Now get all the SPLY Numbers
			SQL = "SELECT SUM(TotalSales) AS SPLYTotalSales,SUM(CASE WHEN Category = 0 THEN TotalSales ELSE 0 END) AS SPLYTotalRentals "
			SQL = SQL & " FROM CustCatPeriodSales "
			SQL = SQL & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
			SQL = SQL & " AND CustType = '" & CustType & "'"
		
			Tot_CurrentADS = (GetCurrentPeriod_PostedTotalSls2(FilterSlsmn1) + GetCurrentPeriod_UnPostedTotalSls2(FilterSlsmn1)) / WD_CurrentSoFar

	End Select

'Response.write("<br><br><br>"&SQL&"<br>")

	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		
		SPLYTotalSales = rs("SPLYTotalSales")
		SPLYTotalRentals = rs("SPLYTotalRentals")
		TotDollarDiffSPLY = ( TotSalesHeader  - rs("SPLYTotalSales")) 
			
	End If


	WD_P3PADS = WorkDaysIn3PeriodBasis / 3
	'See if it has a decimal > 5
	If WD_P3PADS - Int(WD_P3PADS) < .5 Then
		WD_P3PADS = Int(WD_P3PADS)
	Else
		WD_P3PADS = Int(WD_P3PADS) + .5
	End If
	Tot_P3PADS = Tot3PAvgHeader / WD_P3PADS


	WD_P12PADS = WorkDaysIn12PeriodBasis / 12
	'See if it has a decimal > 5
	If WD_P12PADS - Int(WD_P12PADS ) < .5 Then
		WD_P12PADS = Int(WD_P12PADS )
	Else
		WD_P12PADS = Int(WD_P12PADS ) + .5
	End If
	Tot_P12PADS = Tot12PAvgHeader / WD_P12PADS
	
	
	WD_SPLYPADS = WorkDaysInSPLYPeriodBasis 
	'See if it has a decimal > 5
	If WD_SPLYPADS - Int(WD_SPLYPADS ) < .5 Then
		WD_SPLYPADS = Int(WD_P12PADS )
	Else
		WD_SPLYPADS = Int(WD_SPLYPADS ) + .5
	End If
	Tot_SPLYPADS = SPLYTotalSales  / WD_SPLYPADS
	

%>


<div class='table-responsive' style="width:50%;">
	<table class='table table-condensed table-top2'>
		<tbody>
			<tr>


				<!----- BOX 1 ----->
				<td width="27%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
						<table class="table table-striped table-condensed table-hover">
							<tbody>
								<tr>
									<td align="right" width="40%"><strong>&nbsp;</strong></td>
									<td align="center"width="10%"><strong>Total Sales</strong></td>
									<td align="center"width="10%"><strong>Variance</strong></td>
									<td align="center"width="10%"><strong>ADS</strong></td>
									<td align="center"width="10%"><strong>Days</strong></td>
									<td align="center"width="10%"><strong>Product Sales</strong></td>
									<td align="center"width="10%"><strong>Rentals</strong></td>
								</tr>
								
								
								<tr>
									<!-- Title --><td align="right" width="15%"><strong>Period&nbsp;<%=LCP_Display_Month%></strong></td>
									<!-- Total Sales --><td align="center"><%= FormatCurrency(TotSalesHeader,0,0) %></td>
									<!-- Variance -->
									<td align="center"><strong>&nbsp;</strong></td>
									<!-- ADS --><td align="center"><%= FormatCurrency(((Round(TotSalesHeader,0)-Round(TotRentalsHeader,0))/WorkDaysInLastClosedPeriod),0,0) %></td>
									<!-- Days --><td align="center"><%= WorkDaysInLastClosedPeriod%></td>
									<!-- Product Sales --><td align="center"><%= FormatCurrency(Round(TotSalesHeader,0)-Round(TotRentalsHeader,0),0,0) %></td>
									<!-- Rentals --><td align="center"><%= FormatCurrency(TotRentalsHeader,0,0) %></td>
								</tr>

								<tr>
									<!-- Title --><td align="right" width="15%"><strong>3 Prior Periods Avg</strong></td>
									<!-- Total Sales --><td align="center"><%= FormatCurrency(Tot3PAvgHeader,0,0) %></td>
									<!-- Variance -->
									<% If TotDollarDiff < 0 Then %>
										<td align="center" class="negative"><%= FormatCurrency(TotDollarDiff,0,0) %></td>
									<% Else %>
										<td align="center"><%= FormatCurrency(TotDollarDiff,0,0) %></td>
									<% End If %>	
									<!-- ADS --><td align="center"><%= FormatCurrency(((Round(Tot3PAvgHeader,0)-Round(Tot3PAvgHeaderRentals,0))/WD_P3PADS),0,0) %></td>
									<!-- Days --><td align="center"><%= WD_P3PADS %></td>
									<!-- Product Sales --><td align="center"><%= FormatCurrency(Round(Tot3PAvgHeader,0)-Round(Tot3PAvgHeaderRentals,0),0,0) %></td>
									<!-- Rentals --><td align="center"><%= FormatCurrency(Tot3PAvgHeaderRentals,0,0) %></td>
								</tr>

								<tr>
									<!-- Title --><td align="right" width="15%"><strong>12 Prior Periods Avg</strong></td>
									<!-- Total Sales --><td align="center"><%= FormatCurrency(Tot12PAvgHeader,0,0) %></td>
									<!-- Variance -->
									<% If TotDollarDiff12 < 0 Then %>
										<td align="center" class="negative"><%= FormatCurrency(TotDollarDiff12,0,0) %></td>
									<% Else %>
										<td  align="center"><%= FormatCurrency(TotDollarDiff12,0,0) %></td>
									<% End If %>
									<!-- ADS --><td align="center"><%= FormatCurrency((Round(Tot12PAvgHeader,0)-Round(Tot12PAvgHeaderRentals,0))/WD_P12PADS,0,0) %></td>
									<!-- Days --><td align="center"><%= WD_P12PADS %></td>
									<!-- Product Sales --><td align="center"><%= FormatCurrency(Round(Tot12PAvgHeader,0)-Round(Tot12PAvgHeaderRentals,0),0,0) %></td>
									<!-- Rentals --><td align="center"><%= FormatCurrency(Tot12PAvgHeaderRentals,0,0) %></td>
								</tr>
						

								<tr>
									<!-- Title --><td align="right" width="15%"><strong>Same Period Last Year</strong></td>
									<!-- Total Sales --><td align="center"><%= FormatCurrency(SPLYTotalSales,0,0) %></td>
									<!-- Variance -->
									<% If TotDollarDiffSPLY< 0 Then %>
										<td align="center" class="negative"><%= FormatCurrency(TotDollarDiffSPLY,0,0) %></td>
									<% Else %>
										<td align="center"><%= FormatCurrency(TotDollarDiffSPLY,0,0) %></td>
									<% End If %>
									<!-- ADS --><td align="center"><%= FormatCurrency((Round(SPLYTotalSales,0)-Round(SPLYTotalRentals,0))/WD_SPLYPADS,0,0) %></td>
									<!-- Days --><td align="center"><%= WD_SPLYPADS %></td>
									<!-- Product Sales --><td align="center"><%= FormatCurrency((Round(SPLYTotalSales,0)-Round(SPLYTotalRentals,0)),0,0) %></td>
									<!-- Rentals --><td align="center"><%= FormatCurrency(SPLYTotalRentals,0,0) %></td>
								</tr>
								

								<tr>
									<%
									Select Case MUV_READ("LOHVAR")
	
										Case "Secondary"

											RentalsHolder = GetCurrentPeriod_PostedRentalsSls2(FilterSlsmn2) + GetCurrentPeriod_UnPostedRentalsSls2(FilterSlsmn2)
											ProdSalesHolder = (GetCurrentPeriod_PostedTotalSls2(FilterSlsmn2) + GetCurrentPeriod_UnPostedTotalSls2(FilterSlsmn2)) - RentalsHolder 

										Case "Primary"

											RentalsHolder = GetCurrentPeriod_PostedRentalsSls(FilterSlsmn1) + GetCurrentPeriod_UnPostedRentalsSls(FilterSlsmn1)
											ProdSalesHolder = (GetCurrentPeriod_PostedTotalSls(FilterSlsmn1) + GetCurrentPeriod_UnPostedTotalSls(FilterSlsmn1)) - RentalsHolder 

										Case "CustType"

											RentalsHolder = GetCurrentPeriod_PostedRentalsCustType(CustType) + UnPostedRentalsCustType(CustType)
											ProdSalesHolder = (GetCurrentPeriod_PostedTotalCustType(CustType) + GetCurrentPeriod_UnPostedTotalCustType(CustType)) - RentalsHolder 
											
									End Select
									
									tmp3PADS = (Round(Tot3PAvgHeader,0)-Round(Tot3PAvgHeaderRentals,0))/WD_P3PADS
									tmpCurADS = ProdSalesHolder/WorkDaysSoFar
									
									%>
									<!-- Title --><td align="right" width="15%"><strong>Current (So Far)</strong></td>
									<!-- Total Sales --><td align="center"><%= FormatCurrency(ProdSalesHolder + RentalsHolder ,0,0) %></td>
									<!-- Variance -->
									<td align="center"><strong>&nbsp;</strong></td>
									<% If tmpCurADS > tmp3PADS then %>
										<!-- ADS --><td align="center" class="blue"><%= FormatCurrency(ProdSalesHolder/WorkDaysSoFar,0,0) %></td>
									<% Else %>
										<!-- ADS --><td align="center" class="red"><%= FormatCurrency(ProdSalesHolder/WorkDaysSoFar,0,0) %></td>
									<% End If %>
									<!-- Days --><td align="center"><%= WorkDaysSoFar %> of <%=WorkDaysInCurrentPeriod %></td>
									<!-- Product Sales --><td align="center"><%= FormatCurrency((ProdSalesHolder),0,0) %></td>
									<!-- Rentals --><td align="center"><%= FormatCurrency(RentalsHolder,0,0) %></td>
								</tr>
								

								<tr>
									<% 
									PeriodDisplayVar = ""
									If LCP_Display_Month <> 12 Then
										PeriodDisplayVar = LCP_Display_Month + 1
									Else
										PeriodDisplayVar = 1
									End If
									%>
									<!-- Title --><td align="right" width="15%"><strong>Current Period (P<%=PeriodDisplayVar%>) Last Year</strong></td>
									<!-- Total Sales --><td align="center"><%= FormatCurrency(TotSalesCPLY+TotRentalsCPLY,0,0) %></td>
									<% If TotSalesHeader-(TotSalesCPLY+TotRentalsCPLY) < 0 Then %>
										<td align="center" class="negative"><%= FormatCurrency(TotSalesHeader-(TotSalesCPLY+TotRentalsCPLY),0,0) %></td>
									<% Else %>
										<td align="center"><%= FormatCurrency(TotSalesHeader-(TotSalesCPLY+TotRentalsCPLY),0,0) %></td>
									<% End If %>
									<!-- ADS --><td align="center"><%= FormatCurrency(TotSalesCPLY/WorkDaysInCurrentPLY ,0,0) %></td>
									<!-- Days --><td align="center"><%= WorkDaysInCurrentPLY %></td>
									<!-- Product Sales --><td align="center"><%= FormatCurrency((TotSalesCPLY),0,0) %></td>
									<!-- Rentals --><td align="center"><%= FormatCurrency(TotRentalsCPLY,0,0) %></td>
								</tr>
								
								
							</tbody>
						</table>
					</div>
				</td>
				<!----- END BOX 1 ----->

		
			</tr>
		</tbody>
	</table>
</div>
