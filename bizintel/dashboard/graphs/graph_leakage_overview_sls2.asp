<%
amChartDataSls2 = ""

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

%>
<div id="chartdivSls2" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<% 
 	'Get all Secondary Salesman

	TotSalesSls2 = 0
	Tot3PAvgSls2 = 0
	TotDollarDiff =0
	TotalNegDiff = 0		

	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg"
	SQL = SQL & " ,SUM(TotalSales+PriorPeriod1Sales+PriorPeriod2Sales) / 3 AS ProjectionBasis "
	SQL = SQL & ",SecondarySalesman "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	'SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY SecondarySalesman "
	SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"


	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		'Need to get totals first
		Do
			If rs("TotSales") - rs("Tot3PPAvg") < 0 Then TotalNegDiff = TotalNegDiff + (rs("TotSales") - rs("Tot3PPAvg"))
	
			TotDollarDiff = TotDollarDiff + ( rs("TotSales") - rs("Tot3PPAvg")) 
			TotSalesSls2 = TotSalesSls2 + rs("TotSales")
			Tot3PAvgSls2 = Tot3PAvgSls2 + rs("Tot3PPAvg")

			rs.MoveNext
			
		Loop While Not rs.Eof

		rs.MoveFirst
		
		ChartElementNumber = 1
		ChartDataSls2 = ""
		ChartRemainder = 100
		amChartDataSls2 = ""
		RemainderDollarDiff = 0
		
		Do
		
			DollarDiff = rs("TotSales") - rs("Tot3PPAvg")
			
			If TotalNegDiff <> 0 Then ContributionPercent = (DollarDiff / TotalNegDiff ) * 100 Else ContributionPercent = 0 * 100

			'Now handle the part for the chart (Hah! "The part for the chart")
			If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
				ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
				'am Charts
				If Instr(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman"))," ") <> 0 Then 
					amChartDataSls2  = amChartDataSls2  & "{'secondary': '" & Left(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman")),Instr(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman"))," ")+1)  & "',"
				Else
					amChartDataSls2  = amChartDataSls2  & "{'secondary': '" & GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman")) & "',"										
				End If
				amChartDataSls2  = amChartDataSls2  &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
				amChartDataSls2  = amChartDataSls2  &  "'contribDollars': " & Round(DollarDiff ,0) & "}," 
			Else
				RemainderDollarDiff = RemainderDollarDiff + DollarDiff	
				
			End If
			
			ChartElementNumber = ChartElementNumber + 1
			
			rs.movenext
		Loop until rs.eof
		
		'am Charts
		amChartDataSls2  = amChartDataSls2  & "{'secondary': 'Other',"
		amChartDataSls2  = amChartDataSls2  &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
		amChartDataSls2  = amChartDataSls2  &  "'contribDollars': " & Round((RemainderDollarDiff * -1) ,0) & "}" 

	End If	
			
%>