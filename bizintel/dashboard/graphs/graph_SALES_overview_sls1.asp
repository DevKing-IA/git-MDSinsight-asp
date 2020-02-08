<%
amChartDataSls1SALES = ""

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

%>
<div id="chartdivSls1SALES" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<% 
 	'Get all Salesmen 1

	GrandTotalAllSls1 = 0	

	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",Salesman "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY AR_Customer.Salesman ORDER BY TotSales DESC"

	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		'Need to get totals first
		Do

			GrandTotalAllSls1 = GrandTotalAllSls1 + rs("TotSales")
			
			rs.MoveNext
			
		Loop While Not rs.Eof

		rs.MoveFirst
		
		ChartElementNumber = 1
		ChartDataSls1 = ""
		ChartRemainder = 100
		amChartDataSls1SALES = ""
		RemainderDollarDiff = 0
		
		Do
		
			ContributionPercent = (rs("TotSales") / GrandTotalAllSls1) * 100
			
			
			'Now handle the part for the chart (Hah! "The part for the chart")
			If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
				ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
				'am Charts
				If Instr(GetSalesmanNameBySlsmnSequence(rs("Salesman"))," ") <> 0 Then 
					amChartDataSls1SALES  = amChartDataSls1SALES  & "{'primary': '" & Left(GetSalesmanNameBySlsmnSequence(rs("Salesman")),Instr(GetSalesmanNameBySlsmnSequence(rs("Salesman"))," ")+1)  & "',"
				Else
					amChartDataSls1SALES  = amChartDataSls1SALES  & "{'primary': '" & GetSalesmanNameBySlsmnSequence(rs("Salesman")) & "',"										
				End If
				amChartDataSls1SALES  = amChartDataSls1SALES  &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
				amChartDataSls1SALES  = amChartDataSls1SALES  &  "'contribDollars': " & Round(rs("TotSales"),0) & "}," 
			Else
				RemainderDollarDiff = RemainderDollarDiff +  + rs("TotSales")		
				
			End If
			
			ChartElementNumber = ChartElementNumber + 1
			
			rs.movenext
		Loop until rs.eof
		
		'am Charts
		amChartDataSls1SALES  = amChartDataSls1SALES  & "{'primary': 'Other',"
		amChartDataSls1SALES  = amChartDataSls1SALES  &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
		amChartDataSls1SALES  = amChartDataSls1SALES  &  "'contribDollars': " & Round((RemainderDollarDiff) ,0) & "}" 

	End If	
			
%>