<%
amChartDataReferralSALES = ""

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

%>
<div id="chartdivRefSALES" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<% 
 	'Get all Referral Codes

	GrandTotalAllSalesRef = 0

	SQL = "SELECT SUM(TotalSales) AS TotSales "
	SQL = SQL & " ,Referal.Description2 As ReferralDesc2"
	SQL = SQL & " FROM CustCatPeriodSales_ReportData "
	SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
	SQL = SQL & " INNER JOIN Referal ON Referal.ReferalCode = AR_Customer.ReferalCode"
	SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
	SQL = SQL & " GROUP BY Referal.Description2 ORDER BY TotSales DESC"

	Set rs = cnn8.Execute(SQL)
	
	If not rs.EOF Then
	
		'Need to get totals first
		Do
	
			GrandTotalAllSalesRef = GrandTotalAllSalesRef + rs("TotSales")

			rs.MoveNext
			
		Loop While Not rs.Eof

		rs.MoveFirst
		
		ChartElementNumber = 1
		ChartDataReferral = ""
		ChartRemainder = 100
		amChartDataReferralSALES = ""
		RemainderDollarDiff = 0
		
		
		Do
		
			ContributionPercent = (rs("TotSales") / GrandTotalAllSalesRef ) * 100
			
			'Now handle the part for the chart (Hah! "The part for the chart")
			If ChartElementNumber < 6  and Round(ContributionPercent) > 4.99 Then 
				ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
				'am Charts
				amChartDataReferralSALES = amChartDataReferralSALES & "{'referral': '" & rs("ReferralDesc2") & "',"
				amChartDataReferralSALES = amChartDataReferralSALES &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
				amChartDataReferralSALES = amChartDataReferralSALES &  "'contribDollars': " & Round(rs("TotSales"),0) & "}," 
				
			Else
				RemainderDollarDiff = RemainderDollarDiff + rs("TotSales")	
			End If
			
			ChartElementNumber = ChartElementNumber + 1
			rs.movenext
			
		Loop until rs.eof
		
		'am Charts
		amChartDataReferralSALES = amChartDataReferralSALES & "{'referral': 'Other',"
		amChartDataReferralSALES = amChartDataReferralSALES &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
		amChartDataReferralSALES = amChartDataReferralSALES &  "'contribDollars': " & Round((RemainderDollarDiff) ,0) & "}" 
		
	End If	
			
%>