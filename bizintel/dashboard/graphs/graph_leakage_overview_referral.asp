<%
amChartDataReferral = ""

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

%>
<div id="chartdivRef" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<% 
 	'Get all Referral Codes

	TotSalesRef = 0
	Tot3PAvgRef = 0
	TotDollarDiff =0
	TotalNegDiff = 0
	

	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg"
	SQL = SQL & " ,Referal.Description2 As ReferralDesc2"
	SQL = SQL & " ,SUM(TotalSales+PriorPeriod1Sales+PriorPeriod2Sales) / 3 AS ProjectionBasis "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " INNER JOIN Referal ON Referal.ReferalCode = AR_Customer.ReferalCode"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY Referal.Description2"
	SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"


	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		'Need to get totals first
		Do
			If rs("TotSales") - rs("Tot3PPAvg") < 0 Then TotalNegDiff = TotalNegDiff + (rs("TotSales") - rs("Tot3PPAvg"))
	
			TotDollarDiff = TotDollarDiff + ( rs("TotSales") - rs("Tot3PPAvg")) 
			TotSalesRef = TotSalesRef + rs("TotSales")
			Tot3PAvgRef = Tot3PAvgRef + rs("Tot3PPAvg")

			rs.MoveNext
			
		Loop While Not rs.Eof

		rs.MoveFirst
		
		ChartElementNumber = 1
		ChartDataReferral = ""
		ChartRemainder = 100
		amChartDataReferral = ""
		RemainderDollarDiff = 0
		
		Do
		
			DollarDiff = rs("TotSales") - rs("Tot3PPAvg")

			If TotalNegDiff <> 0 Then ContributionPercent = (DollarDiff / TotalNegDiff ) * 100 Else ContributionPercent = 0 * 100
			
			'Now handle the part for the chart (Hah! "The part for the chart")
			If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
				ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
				'am Charts
				amChartDataReferral = amChartDataReferral & "{'referral': '" & rs("ReferralDesc2") & "',"
				amChartDataReferral = amChartDataReferral &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
				amChartDataReferral = amChartDataReferral &  "'contribDollars': " & Round(DollarDiff ,0) & "}," 
				
			Else
				RemainderDollarDiff = RemainderDollarDiff + DollarDiff	
			End If
			
			ChartElementNumber = ChartElementNumber + 1
			rs.movenext
			
		Loop until rs.eof
		
		'am Charts
		amChartDataReferral = amChartDataReferral & "{'referral': 'Other',"
		amChartDataReferral = amChartDataReferral &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
		amChartDataReferral = amChartDataReferral &  "'contribDollars': " & Round((RemainderDollarDiff * -1) ,0) & "}" 
		
	End If	
			
%>