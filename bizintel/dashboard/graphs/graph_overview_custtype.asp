<%
amChartDataCustType = ""

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

%>
<div id="chartdivCustType" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<% 
 	'Get all Referral Codes

	TotSalesTyp = 0
	Tot3PAvgTyp = 0
	TotDollarDiff =0
	TotalNegDiff = 0


	SQL = "SELECT SUM(SalesLCP) AS TotSales ,SUM(Sales3PPAvg) As Tot3PPAvg ,CustomerTypeNumber AS CustType "
	SQL = SQL & " FROM BI_Dashboard "
	SQL = SQL & " WHERE Segment = 'CUSTOMERTYPE' "
	SQL = SQL & " GROUP BY CustomerTypeNumber "
	SQL = SQL & " ORDER BY SUM(SalesLCP)- SUM(Sales3PPAvg)"


'	SQL = "SELECT SUM(TotalSales) AS TotSales "
'	SQL = SQL & ",SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3) As Tot3PPAvg"
'	SQL = SQL & ",SUM(PriorPeriod1Sales) As TotPP1Sales"
'	SQL = SQL & ",SUM(PriorPeriod2Sales) As TotPP2Sales"
'	SQL = SQL & ",SUM(PriorPeriod3Sales) As TotPP3Sales"
'	SQL = SQL & " ,SUM(TotalSales+PriorPeriod1Sales+PriorPeriod2Sales) / 3 AS ProjectionBasis "
'	SQL = SQL & ",AR_Customer.CustType  "
'	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
'	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
'	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
'	SQL = SQL & " GROUP BY AR_Customer.CustType"
'	SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"

response.write(sql)

	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		'Need to get totals first
		Do
			If rs("TotSales") - rs("Tot3PPAvg") < 0 Then TotalNegDiff = TotalNegDiff + (rs("TotSales") - rs("Tot3PPAvg"))
	
			TotDollarDiff = TotDollarDiff + ( rs("TotSales") - rs("Tot3PPAvg")) 
			TotSalesTyp = TotSalesTyp + rs("TotSales")
			Tot3PAvgTyp = Tot3PAvgTyp + rs("Tot3PPAvg")

			rs.MoveNext
			
		Loop While Not rs.Eof

		rs.MoveFirst
		
		ChartElementNumber = 1
		ChartDataType = ""
		ChartRemainder = 100
		amChartDataType = ""
		RemainderDollarDiff = 0
		
		Do
		
			DollarDiff = rs("TotSales") - rs("Tot3PPAvg")

			
			If TotalNegDiff <> 0 Then ContributionPercent = (DollarDiff / TotalNegDiff ) * 100 Else ContributionPercent = 0 * 100

			'Now handle the part for the chart (Hah! "The part for the chart")
			If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
				ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
				'am Charts
				amChartDataCustType  = amChartDataCustType  & "{'custtype': '" & GetCustTypeByCode(rs("CustType")) & "',"
				amChartDataCustType  = amChartDataCustType  &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
				amChartDataCustType  = amChartDataCustType  &  "'contribDollars': " & Round(DollarDiff ,0) & "}," 
			Else
				RemainderDollarDiff = RemainderDollarDiff + DollarDiff	

			End If
			
			ChartElementNumber = ChartElementNumber + 1
			
			rs.movenext
		Loop until rs.eof
		
		'am Charts
		amChartDataCustType  = amChartDataCustType  & "{'custtype': 'Other',"
		amChartDataCustType  = amChartDataCustType  &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
		amChartDataCustType  = amChartDataCustType  &  "'contribDollars': " & Round((RemainderDollarDiff * -1) ,0) & "}" 

	End If	

		
%>