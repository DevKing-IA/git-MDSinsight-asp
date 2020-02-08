<%
amChartDataSls2SALES = ""

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

%>
<div id="chartdivSls2SALES" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<% 
 	'Get all Secondary Salesman

	GrandTotalAllSalesSls2 = 0		

	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",SecondarySalesman "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	'SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY SecondarySalesman ORDER BY TotSales DESC"


	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		'Need to get totals first
		Do

			GrandTotalAllSalesSls2 = GrandTotalAllSalesSls2 + rs("TotSales")
			
			rs.MoveNext
			
		Loop While Not rs.Eof

		rs.MoveFirst
		
		ChartElementNumber = 1
		ChartDataSls2 = ""
		ChartRemainder = 100
		amChartDataSls2SALES = ""
		RemainderDollarDiff = 0
		
		Do
		
			ContributionPercent = (rs("TotSales") / GrandTotalAllSalesSls2 ) * 100
			
			
			'Now handle the part for the chart (Hah! "The part for the chart")
			If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
				ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
				'am Charts
				If Instr(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman"))," ") <> 0 Then 
					amChartDataSls2SALES  = amChartDataSls2SALES  & "{'secondary': '" & Left(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman")),Instr(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman"))," ")+1)  & "',"
				Else
					amChartDataSls2SALES  = amChartDataSls2SALES  & "{'secondary': '" & GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman")) & "',"										
				End If
				amChartDataSls2SALES  = amChartDataSls2SALES  &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
				amChartDataSls2SALES  = amChartDataSls2SALES  &  "'contribDollars': " & Round(rs("TotSales"),0) & "}," 
			Else
				RemainderDollarDiff = RemainderDollarDiff +  + rs("TotSales")		
				
			End If
			
			ChartElementNumber = ChartElementNumber + 1
			
			rs.movenext
		Loop until rs.eof
		
		'am Charts
		amChartDataSls2SALES  = amChartDataSls2SALES  & "{'secondary': 'Other',"
		amChartDataSls2SALES  = amChartDataSls2SALES  &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
		amChartDataSls2SALES  = amChartDataSls2SALES  &  "'contribDollars': " & Round((RemainderDollarDiff) ,0) & "}" 

	End If	
			
%>