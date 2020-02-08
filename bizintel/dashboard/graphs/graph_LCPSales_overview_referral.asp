<%
amChartDataLCPReferral = ""

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
%>

<div id="chartdivLCPSalesRef" style="width: 100%; height: 350px; margin: 0 auto"></div>	

<% 
 	'Get all Referral Codes

	GrandTotSalesRef = 0
'	Session("TimePeriod") = "LCP"
	Session("TimePeriod") = "FY"
'	Session("TimePeriod") = "L6P"	

If Session("TimePeriod") = "LCP" Then
	SQL = "SELECT SUM(TotalSales) AS GrandTotSales FROM CustCatPeriodSales_ReportData WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
End IF
	
If Session("TimePeriod") = "L6P" Then
	SQL = "SELECT SUM(TotalSales) AS GrandTotSales FROM CustCatPeriodSales WHERE ThisPeriodSequenceNumber >= " & PeriodSeqBeingEvaluated - 6
End If
	
If Session("TimePeriod") = "FY" Then
	SQL = "SELECT SUM(TotalSales) AS GrandTotSales FROM CustCatPeriodSales WHERE ThisPeriodSequenceNumber >= " & GetPeriodOneThisFiscalYearSeqNum()
End If	


	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then GrandTotSalesRef = rs("GrandTotSales")


If Session("TimePeriod") = "LCP" Then
	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & " ,Referal.Description2 As ReferralDesc2"
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " INNER JOIN Referal ON Referal.ReferalCode = AR_Customer.ReferalCode"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY Referal.Description2"
	SQL = SQL & " ORDER BY SUM(TotalSales) DESC"
End If

If Session("TimePeriod") = "L6P" Then
	SQL = "SELECT SUM(TotalSales) AS TotSales ,Referal.Description2 As ReferralDesc2 FROM CustCatPeriodSales INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales.CustNum INNER JOIN "
	SQL = SQL & " Referal ON Referal.ReferalCode = AR_Customer.ReferalCode WHERE ThisPeriodSequenceNumber >= " & PeriodSeqBeingEvaluated - 6 & " GROUP BY Referal.Description2 ORDER BY SUM(TotalSales) DESC"
End If

If Session("TimePeriod") = "FY" Then
	SQL = "SELECT SUM(TotalSales) AS TotSales ,Referal.Description2 As ReferralDesc2 FROM CustCatPeriodSales INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales.CustNum INNER JOIN "
	SQL = SQL & " Referal ON Referal.ReferalCode = AR_Customer.ReferalCode WHERE ThisPeriodSequenceNumber >= " & GetPeriodOneThisFiscalYearSeqNum() & " GROUP BY Referal.Description2 ORDER BY SUM(TotalSales) DESC"
End If




	Set rs = cnn8.Execute(SQL)
	
'Response.Write(GrandTotSalesRef )	
	
	If not rs.EOF Then
	
	
		ChartElementNumber = 1
		ChartRemainder = 100
		amChartDataLCPReferral = ""
		RemainderDollarDiff = 0
		
		Do
		
			DollarDiff = rs("TotSales") 

			ContributionPercent = (DollarDiff / GrandTotSalesRef ) * 100
			
			'Now handle the part for the chart (Hah! "The part for the chart")
			If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
				ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
				'am Charts
				amChartDataLCPReferral = amChartDataLCPReferral & "{'LCPreferral': '" & rs("ReferralDesc2") & "',"
				amChartDataLCPReferral = amChartDataLCPReferral &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
				amChartDataLCPReferral = amChartDataLCPReferral &  "'contribDollars': " & Round(DollarDiff ,0) & "}," 
				
			Else
				RemainderDollarDiff = RemainderDollarDiff + DollarDiff	
			End If
			
			ChartElementNumber = ChartElementNumber + 1
			rs.movenext
			
		Loop until rs.eof
		
		'am Charts
		amChartDataLCPReferral = amChartDataLCPReferral & "{'LCPreferral': 'Other',"
		amChartDataLCPReferral = amChartDataLCPReferral &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
		amChartDataLCPReferral = amChartDataLCPReferral &  "'contribDollars': " & Round((RemainderDollarDiff) ,0) & "}" 
		
	End If	

		
%>