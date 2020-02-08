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

	SQL = "SELECT SUM(TotalSales) AS CPLYTotalSales  "
	SQL = SQL & " FROM CustCatPeriodSales "
	SQL = SQL & " WHERE Period = " & LCP_Display_Month+1 & " AND PeriodYear = " & LCP_Display_Year - 1
	SQL = SQL & " AND SecondarySalesman = '" & FilterSlsmn2 & "'"
	Set rs = cnn8.Execute(SQL)
	
	TotSalesCPLY = rs("CPLYTotalSales")

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
		SQL = "SELECT SUM(CASE WHEN Category <> 0 THEN TotalSales END) AS TotSales,SUM(CASE WHEN Category = 0 THEN TotalSales END) AS TotRentals "
		SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg"
		SQL = SQL & ",SUM(( PriorPeriod1Sales+ PriorPeriod2Sales+ PriorPeriod3Sales+ PriorPeriod4Sales+ PriorPeriod5Sales+ PriorPeriod6Sales+ "
		SQL = SQL & "PriorPeriod7Sales+ PriorPeriod8Sales+ PriorPeriod9Sales+ PriorPeriod10Sales+ PriorPeriod11Sales+ PriorPeriod12Sales )/12) As Tot12PPAvg"
		SQL = SQL & " FROM CustCatPeriodSales_ReportData "
		'SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum "
		SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
		SQL = SQL & " AND SecondarySalesman = '" & FilterSlsmn2 & "'"
		
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
		SQL = "SELECT SUM(TotalSales) AS SPLYTotalSales  "
		SQL = SQL & " FROM CustCatPeriodSales "
		'SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales.CustNum "
		SQL = SQL & " WHERE Period = " & LCP_Display_Month & " AND PeriodYear = " & LCP_Display_Year - 1
		SQL = SQL & " AND SecondarySalesman = '" & FilterSlsmn2 & "'"
	
		Tot_CurrentADS = (GetCurrentPeriod_PostedTotalSls2(FilterSlsmn2) + GetCurrentPeriod_UnPostedTotalSls2(FilterSlsmn2)) / WD_CurrentSoFar
	

	End Select

'Response.write("<br><br><br>"&SQL&"<br>")

	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
		
		SPLYTotalSales  = rs("SPLYTotalSales")
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


<div class='table-responsive'>
	<table class='table table-condensed table-top2'>
		<tbody>
			<tr>


				<!----- BOX 1 ----->
				<td width="27%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
						<table class="table table-striped table-condensed table-hover">
							<tbody>
								<tr>
									<td colspan="2" align="center"><strong>Sales</strong></td>
								</tr>
								<tr>
									<td width="50%" align="left"><strong>Period&nbsp;<%=LCP_Display_Month%> Product Sales</strong></td>	
									<td width="50%" align="right"><%= FormatCurrency(TotSalesHeader,0,0) %></td>
								</tr>
								<tr>
									<td width="50%" align="left"><strong>Period&nbsp;<%=LCP_Display_Month%> Rentals</strong></td>	
									<td width="50%" align="right"><%= FormatCurrency(TotRentalsHeader,0,0) %></td>
								</tr>
								<tr>
									<td width="50%" align="left"><strong>3 Prior Periods Average</strong></td>	
									<td width="50%" align="right"><%= FormatCurrency(Tot3PAvgHeader,0,0) %></td>
								</tr>
								<tr>
									<td width="50%" align="left"><strong>12 Prior Periods Average</strong></td>	
									<td width="50%" align="right"><%= FormatCurrency(Tot12PAvgHeader,0,0) %></td>
								</tr>
								<tr>
									<td width="50%" align="left"><strong>Same Period Last Year</strong></td>	
									<td width="50%" align="right"><%= FormatCurrency(SPLYTotalSales,0,0) %></td>
								</tr>
						</tbody>
						</table>
					</div>
				</td>
				<!----- END BOX 1 ----->

				<!----- BOX 1.5 ----->
				<td width="10%">&nbsp;</td>
				<!----- END BOX 1.5 ----->

				<!----- BOX 2 ----->
				<td width="28%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
						<table class="table table-striped table-condensed table-hover">
							<tbody>
								<tr>
									<td colspan="2" align="center"><strong>Variances</strong></td>
								</tr>
								<tr>
									<td width="60%" align="left">&nbsp;</td>	
									<td width="40%" align="right">&nbsp;</td>
								</tr>
								<tr>
									<td width="60%" align="left"><strong>Period&nbsp;<%=LCP_Display_Month%> vs 3 Prior Periods Average</strong></td>	
									<% If TotDollarDiff < 0 Then %>
										<td width="40%" align="right" class="negative"><%= FormatCurrency(TotDollarDiff,0,0) %></td>
									<% Else %>
										<td width="40%" align="right"><%= FormatCurrency(TotDollarDiff,0,0) %></td>
									<% End If %>										
								</tr>
								<tr>
									<td width="60%" align="left"><strong>Period&nbsp;<%=LCP_Display_Month%> vs 12 Prior Periods Average</strong></td>	
									<% If TotDollarDiff12 < 0 Then %>
										<td width="40%" align="right" class="negative"><%= FormatCurrency(TotDollarDiff12,0,0) %></td>
									<% Else %>
										<td width="40%" align="right"><%= FormatCurrency(TotDollarDiff12,0,0) %></td>
									<% End If %>
								</tr>
								<tr>
									<td width="60%" align="left"><strong>Period&nbsp;<%=LCP_Display_Month%> vs Same Period Last Year</strong></td>	
									<% If TotDollarDiffSPLY< 0 Then %>
										<td width="40%" align="right" class="negative"><%= FormatCurrency(TotDollarDiffSPLY,0,0) %></td>
									<% Else %>
										<td width="40%" align="right"><%= FormatCurrency(TotDollarDiffSPLY,0,0) %></td>
									<% End If %>
								</tr>
							</tbody>
						</table>
					</div>
				</td>
				<!----- END BOX 2 ----->


				<!----- BOX 2.5 ----->
				<td width="10%">&nbsp;</td>
				<!----- END BOX 2.5 ----->
		
				<!----- BOX 2 ----->
				<td width="25%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
						<table class="table table-striped table-condensed table-hover">
							<tbody>
								<tr>
									<td colspan="2" align="center"><strong>Average Daily Sales</strong></td>
									<td align="right"><strong>Days</strong></td>
								</tr>
								<tr>
									<td width="40%" align="left"><strong>Period&nbsp;<%=LCP_Display_Month%></strong></td>	
									<td width="40%" align="right"><%= FormatCurrency((TotSalesHeader/WorkDaysInLastClosedPeriod),0,0) %></td>
									<td width="20%" align="right"><%= WorkDaysInLastClosedPeriod%></td>
								</tr>
								<tr>
									<td width="40%" align="left"><strong>3 Prior Periods Average</strong></td>	
									<td width="40%" align="right"><%= FormatCurrency(Tot_P3PADS,0,0) %></td>
									<td width="20%" align="right"><%= WD_P3PADS %></td>
								</tr>
								<tr>
									<td width="40%" align="left"><strong>12 Prior Periods Average</strong></td>	
									<td width="40%" align="right"><%= FormatCurrency(Tot_P12PADS,0,0) %></td>
									<td width="20%" align="right"><%= WD_P12PADS %></td>
								</tr>
								<tr>
									<td width="40%" align="left"><strong>Same Period Last Year</strong></td>	
									<td width="40%" align="right"><%= FormatCurrency(Tot_SPLYPADS,0,0) %></td>
									<td width="20%" align="right"><%= WD_SPLYPADS %></td>
								</tr>
								<tr>
									<td width="40%" align="left"><strong>Period <%=LCP_Display_Month+1%> So Far</strong></td>	
									<td width="40%" align="right"><%= FormatCurrency(Tot_CurrentADS,0,0) %></td>
									<td width="20%" align="right"><%= WorkDaysSoFar %> of <%=WorkDaysInCurrentPeriod %></td>
								</tr>
								<tr>
									<td width="40%" align="left"><strong>Period <%=LCP_Display_Month+1%> Last Year</strong></td>	
									<td width="40%" align="right"><%= FormatCurrency(Tot_CPLYPADS,0,0) %></td>
									<td width="20%" align="right"><%=WorkDaysInCurrentPLY %></td>
								</tr>
							</tbody>
						</table>
					</div>
				</td>
				<!----- END BOX 2 ----->
		
			</tr>
		</tbody>
	</table>
</div>
