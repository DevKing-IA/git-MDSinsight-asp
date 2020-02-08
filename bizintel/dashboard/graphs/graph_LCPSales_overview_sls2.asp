<%
amChartDataLCPSls2 = ""

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

%>
<div id="chartdivLCPSalesSls2" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<% 
 	'Get all Secondary Salesman


If Session("TimePeriod") = "LCP" Then
	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",SecondarySalesman "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY AR_Customer.SecondarySalesman "
	SQL = SQL & " ORDER BY SUM(TotalSales) DESC"
End If

If Session("TimePeriod") = "L6P" Then
	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",SecondarySalesman "
	SQL = SQL & " FROM CustCatPeriodSales "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber >= " & PeriodSeqBeingEvaluated - 6
	SQL = SQL & " GROUP BY AR_Customer.SecondarySalesman "
	SQL = SQL & " ORDER BY SUM(TotalSales) DESC"
End If

If Session("TimePeriod") = "FY" Then
	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",SecondarySalesman "
	SQL = SQL & " FROM CustCatPeriodSales "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber >= " & GetPeriodOneThisFiscalYearSeqNum()
	SQL = SQL & " GROUP BY AR_Customer.SecondarySalesman "
	SQL = SQL & " ORDER BY SUM(TotalSales) DESC"
End If

	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
	
		ChartElementNumber = 1
		ChartRemainder = 100
		amChartDataLCPSls2 = ""
		RemainderDollarDiff = 0
		
		Do
		
			DollarDiff = rs("TotSales")
			
			ContributionPercent = (DollarDiff / GrandTotSalesRef) * 100

			'Now handle the part for the chart (Hah! "The part for the chart")
			If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
				ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
				'am Charts
				If Instr(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman"))," ") <> 0 Then 
					amChartDataLCPSls2  = amChartDataLCPSls2  & "{'LCPsecondary': '" & Left(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman")),Instr(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman"))," ")+1)  & "',"
				Else
					amChartDataLCPSls2  = amChartDataLCPSls2  & "{'LCPsecondary': '" & GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman")) & "',"										
				End If
				amChartDataLCPSls2  = amChartDataLCPSls2  &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
				amChartDataLCPSls2  = amChartDataLCPSls2  &  "'contribDollars': " & Round(DollarDiff ,0) & "}," 
			Else
				RemainderDollarDiff = RemainderDollarDiff + DollarDiff	
			End If
			
			ChartElementNumber = ChartElementNumber + 1
			
			rs.movenext
		Loop until rs.eof
		
		'am Charts
		amChartDataLCPSls2  = amChartDataLCPSls2  & "{'LCPsecondary': 'Other',"
		amChartDataLCPSls2  = amChartDataLCPSls2  &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
		amChartDataLCPSls2  = amChartDataLCPSls2  &  "'contribDollars': " & Round((RemainderDollarDiff) ,0) & "}" 

	End If	
			
%>