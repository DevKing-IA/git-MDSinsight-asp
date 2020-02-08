<%
amChartDataSls1 = ""

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

%>
<div id="chartdivSls1" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<% 
 	'Get all Salesmen 1

	TotSalesSls1 = 0
	Tot3PAvgSls1 = 0
	TotDollarDiff =0
	TotalNegDiff = 0		

	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg"
	SQL = SQL & " ,SUM(TotalSales+PriorPeriod1Sales+PriorPeriod2Sales) / 3 AS ProjectionBasis "
	SQL = SQL & ",Salesman "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY AR_Customer.Salesman"
	SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"

'Response.Write(SQL)

	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		'Need to get totals first
		Do
			If rs("TotSales") - rs("Tot3PPAvg") < 0 Then TotalNegDiff = TotalNegDiff + (rs("TotSales") - rs("Tot3PPAvg"))
	
			TotDollarDiff = TotDollarDiff + ( rs("TotSales") - rs("Tot3PPAvg")) 
			TotSalesSls1 = TotSalesSls1 + rs("TotSales")
			Tot3PAvgSls1 = Tot3PAvgSls1 + rs("Tot3PPAvg")

			rs.MoveNext
			
		Loop While Not rs.Eof

		rs.MoveFirst
		
		ChartElementNumber = 1
		ChartDataSls1 = ""
		ChartRemainder = 100
		amChartDataSls1 = ""
		RemainderDollarDiff = 0
		
		Do
		
			DollarDiff = rs("TotSales") - rs("Tot3PPAvg")
			
			If TotalNegDiff <> 0 Then ContributionPercent = (DollarDiff / TotalNegDiff ) * 100 Else ContributionPercent = 0
			

			'Now handle the part for the chart (Hah! "The part for the chart")
			If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
				ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
				'am Charts
				If Instr(GetSalesmanNameBySlsmnSequence(rs("Salesman"))," ") <> 0 Then 
					amChartDataSls1  = amChartDataSls1  & "{'primary': '" & Left(GetSalesmanNameBySlsmnSequence(rs("Salesman")),Instr(GetSalesmanNameBySlsmnSequence(rs("Salesman"))," ")+1)  & "',"
				Else
					amChartDataSls1  = amChartDataSls1  & "{'primary': '" & GetSalesmanNameBySlsmnSequence(rs("Salesman")) & "',"										
				End If
				amChartDataSls1  = amChartDataSls1  &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
				amChartDataSls1  = amChartDataSls1  &  "'contribDollars': " & Round(DollarDiff ,0) & "}," 
			Else
				RemainderDollarDiff = RemainderDollarDiff + DollarDiff	
				
			End If
			
			ChartElementNumber = ChartElementNumber + 1
			
			rs.movenext
		Loop until rs.eof
		
		'am Charts
		amChartDataSls1  = amChartDataSls1  & "{'primary': 'Other',"
		amChartDataSls1  = amChartDataSls1  &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
		amChartDataSls1  = amChartDataSls1  &  "'contribDollars': " & Round((RemainderDollarDiff * -1) ,0) & "}" 

	End If	
			
%>