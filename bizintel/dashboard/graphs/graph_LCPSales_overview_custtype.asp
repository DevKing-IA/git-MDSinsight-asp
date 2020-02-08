<%
amChartDataLCPCustType = ""

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

%>
<div id="chartdivLCPSalesCustType" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<% 
 	'Get all Codes

If Session("TimePeriod") = "LCP" Then
	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",AR_Customer.CustType  "
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY AR_Customer.CustType"
	SQL = SQL & " ORDER BY SUM(TotalSales) DESC"
End If

If Session("TimePeriod") = "L6P" Then
	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",AR_Customer.CustType  "
	SQL = SQL & " FROM CustCatPeriodSales "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber >= " & PeriodSeqBeingEvaluated - 6
	SQL = SQL & " GROUP BY AR_Customer.CustType"
	SQL = SQL & " ORDER BY SUM(TotalSales) DESC"
End If

If Session("TimePeriod") = "FY" Then
	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & ",AR_Customer.CustType  "
	SQL = SQL & " FROM CustCatPeriodSales "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales.CustNum"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber >= " & GetPeriodOneThisFiscalYearSeqNum() 
	SQL = SQL & " GROUP BY AR_Customer.CustType"
	SQL = SQL & " ORDER BY SUM(TotalSales) DESC"
End If


	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		
		ChartElementNumber = 1
		ChartRemainder = 100
		amChartDataLCPType = ""
		RemainderDollarDiff = 0
		
		Do
		
			DollarDiff = rs("TotSales")
			
			ContributionPercent = (DollarDiff / GrandTotSalesRef ) * 100

			'Now handle the part for the chart (Hah! "The part for the chart")
			If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
				ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
				'am Charts
				amChartDataLCPCustType  = amChartDataLCPCustType  & "{'LCPcusttype': '" & GetCustTypeByCode(rs("CustType")) & "',"
				amChartDataLCPCustType  = amChartDataLCPCustType  &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
				amChartDataLCPCustType  = amChartDataLCPCustType  &  "'contribDollars': " & Round(DollarDiff ,0) & "}," 
			Else
				RemainderDollarDiff = RemainderDollarDiff + DollarDiff	
			End If

			
			ChartElementNumber = ChartElementNumber + 1
			
			rs.movenext
		Loop until rs.eof
		
		'am Charts
		amChartDataLCPCustType  = amChartDataLCPCustType  & "{'LCPcusttype': 'Other',"
		amChartDataLCPCustType  = amChartDataLCPCustType  &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
		amChartDataLCPCustType  = amChartDataLCPCustType  &  "'contribDollars': " & Round((RemainderDollarDiff) ,0) & "}" 

	End If	
			
%>