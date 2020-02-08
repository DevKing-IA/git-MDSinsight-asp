<%
amChartDataCustTypeSALES = ""

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

%>
<div id="chartdivCustTypeSALES" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<% 
 	'Get all Customer Types
 	
	GrandTotalAllSalesTyp = 0

	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",AR_Customer.CustType  "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY AR_Customer.CustType ORDER BY TotSales DESC"

	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		'Need to get totals first
		Do

			GrandTotalAllSalesTyp = GrandTotalAllSalesTyp + rs("TotSales")

			rs.MoveNext
			
		Loop While Not rs.Eof

		rs.MoveFirst
		
		ChartElementNumber = 1
		ChartDataType = ""
		ChartRemainder = 100
		amChartDataTypeSALES = ""
		RemainderDollarDiff = 0
		
		Do
		
			ContributionPercent = (rs("TotSales") / GrandTotalAllSalesTyp ) * 100

			'Now handle the part for the chart (Hah! "The part for the chart")
			If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
				ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
				'am Charts
				amChartDataCustTypeSALES  = amChartDataCustTypeSALES  & "{'custtype': '" & GetCustTypeByCode(rs("CustType")) & "',"
				amChartDataCustTypeSALES  = amChartDataCustTypeSALES  &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
				amChartDataCustTypeSALES  = amChartDataCustTypeSALES  &  "'contribDollars': " & Round(rs("TotSales"),0) & "}," 
			Else
				RemainderDollarDiff = RemainderDollarDiff +  + rs("TotSales")		
			End If
			
			ChartElementNumber = ChartElementNumber + 1
			
			rs.movenext
		Loop until rs.eof
		
		'am Charts
		amChartDataCustTypeSALES  = amChartDataCustTypeSALES  & "{'custtype': 'Other',"
		amChartDataCustTypeSALES  = amChartDataCustTypeSALES  &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
		amChartDataCustTypeSALES  = amChartDataCustTypeSALES  &  "'contribDollars': " & Round((RemainderDollarDiff) ,0) & "}" 

	End If	

		
%>