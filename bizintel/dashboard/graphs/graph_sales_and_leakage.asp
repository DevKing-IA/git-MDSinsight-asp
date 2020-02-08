<%
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rsCompanyLeakage = Server.CreateObject("ADODB.Recordset")
rsCompanyLeakage.CursorLocation = 3
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3

%>
<div id="chartdivSAL" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<%
'Create a temporary table

On Error Resume Next

SQLForChart = "DROP TABLE zBizIntelChartTable1_" & Session("UserNo")
Set rsCompanyLeakage = cnn8.Execute(SQLForChart)


SQLForChart = "CREATE TABLE zBizIntelChartTable1_" & Session("UserNo") & " ( "
SQLForChart = SQLForChart & "[PeriodSequenceNumber] [int] NULL,"
SQLForChart = SQLForChart & "[ThisPeriodTotalSalesDollars] [money] NULL,"
SQLForChart = SQLForChart & "[PriorPeriod1TotalSalesDollars] [money] NULL,"
SQLForChart = SQLForChart & "[PriorPeriod2TotalSalesDollars] [money] NULL,"
SQLForChart = SQLForChart & "[PriorPeriod3TotalSalesDollars] [money] NULL,"
SQLForChart = SQLForChart & "[Period] [int] NULL,"
SQLForChart = SQLForChart & "[PeriodYear] [int] NULL"
SQLForChart = SQLForChart & ")"

Set rsCompanyLeakage = cnn8.Execute(SQLForChart)
On Error Goto 0

LastPeriodSalesHolder = 0

For i = GetLastClosedPeriodSeqNum() - 15 to GetLastClosedPeriodSeqNum() ' Important, needs to be 15

		PeriodSeqBeingEvaluated = i
		
		WorkDaysIn3PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -3), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1
		WorkDaysIn12PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -12), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1)) + 1 
		WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated )) 
		WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated +1)) + 1 
		WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1),Date()) + 1
		
		'******************************************************************************
		'This is the first pass which just puts in the total Sales $ for a given period
		'******************************************************************************
		SQLForChart = "SELECT SUM(TotalSales) AS LCPSales "
		SQLForChart = SQLForChart & ", SUM(PriorPeriod1Sales) AS TotPriorPeriod1Sales "
		SQLForChart = SQLForChart & ", SUM(PriorPeriod2Sales) AS TotPriorPeriod2Sales "
		SQLForChart = SQLForChart & ", SUM(PriorPeriod3Sales) AS TotPriorPeriod3Sales "
		SQLForChart = SQLForChart & " FROM CustCatPeriodSales "
		SQLForChart = SQLForChart & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
		If LimitSelection = 1 Then
			SQLForChart = SQLForChart & " AND TotalSales < [3PriorPeriodsAeverage] "
			SQLForChart = SQLForChart & " AND [3PriorPeriodsAeverage] - TotalSales > " & FilterSalesDollars 
			SQLForChart = SQLForChart & " AND (CASE WHEN [3PriorPeriodsAeverage] <> 0 THEN (((TotalSales - [3PriorPeriodsAeverage] ) / [3PriorPeriodsAeverage]) * 100) * -1 END) >= " & FilterPercentage 
		End If

		'Response.Write("<br>" & SQLForChart & "<br>")	
	
		Set rs = cnn8.Execute(SQLForChart)			

		'Insert the record here
		SQLForChart = "INSERT INTO zBizIntelChartTable1_" & Session("UserNo") & " (PeriodSequenceNumber,ThisPeriodTotalSalesDollars,PriorPeriod1TotalSalesDollars,PriorPeriod2TotalSalesDollars,PriorPeriod3TotalSalesDollars)"
		SQLForChart = SQLForChart & " VALUES ("
		SQLForChart = SQLForChart & PeriodSeqBeingEvaluated 
		SQLForChart = SQLForChart & "," & rs("LCPSales")
		SQLForChart = SQLForChart & "," & rs("TotPriorPeriod1Sales")
		SQLForChart = SQLForChart & "," & rs("TotPriorPeriod2Sales")
		SQLForChart = SQLForChart & "," & rs("TotPriorPeriod3Sales")
		SQLForChart = SQLForChart & ")" 
		
		'Response.Write("<br>" & SQLForChart & "<br>")			
		Set rsCompanyLeakage = cnn8.Execute(SQLForChart)					
		
Next


SQLForChart = "UPDATE zBizIntelChartTable1_" & Session("UserNo") & " SET Period = "
SQLForChart = SQLForChart & " (SELECT Period FROM BillingPeriodHistory"
SQLForChart = SQLForChart & " WHERE BillperSequence = zBizIntelChartTable1_" & Session("UserNo") & ".PeriodSequenceNumber)"
'Response.Write("<br>" & SQLForChart & "<br>")			
Set rsCompanyLeakage = cnn8.Execute(SQLForChart)	

SQLForChart = "UPDATE zBizIntelChartTable1_" & Session("UserNo") & " SET PeriodYear = "
SQLForChart = SQLForChart & " (SELECT Year FROM BillingPeriodHistory"
SQLForChart = SQLForChart & " WHERE BillperSequence = zBizIntelChartTable1_" & Session("UserNo") & ".PeriodSequenceNumber)"
'Response.Write("<br>" & SQLForChart & "<br>")			
Set rsCompanyLeakage = cnn8.Execute(SQLForChart)
	

'Now Build the Chart Data Variable

amChartDataSAL = ""

For i = GetLastClosedPeriodSeqNum() - 12 to GetLastClosedPeriodSeqNum() ' Important, needs to be 15

	SQLForChart = "SELECT * FROM zBizIntelChartTable1_" & Session("UserNo") & " WHERE PeriodSequenceNumber = " & i
	
	'Response.Write("<br>" & SQLForChart & "<br>")
	
	Set rs = cnn8.Execute(SQLForChart)	
	
	DaysInPeriod =  NumberOfWorkDays(GetPeriodBeginDateBySeq(rs("PeriodSequenceNumber")), GetPeriodEndDateBySeq(rs("PeriodSequenceNumber")))
	
	DaysInPeriodString = "(" & DaysInPeriod & " days)"
	
	amChartDataSAL = amChartDataSAL  & "{'year': 'P" & rs("Period") & "/" & rs("PeriodYear") & "<br>" & DaysInPeriodString & "',"
	
	amChartDataSAL = amChartDataSAL  & "'sales': " & Round(rs("ThisPeriodTotalSalesDollars"),0) & ","        

	amChartDataSAL = amChartDataSAL  & "'ppsales': " & Round(rs("PriorPeriod1TotalSalesDollars"),0) & "," 
	
	amChartDataSAL = amChartDataSAL  & "'p3pavgsales': " & Round((rs("PriorPeriod1TotalSalesDollars") + rs("PriorPeriod2TotalSalesDollars") + rs("PriorPeriod3TotalSalesDollars")) / 3,0) & "}," 
	        
Next

'Strip last comma
amChartDataSAL = Left(amChartDataSAL, len(amChartDataSAL)-1)


SQLForChart = "DROP TABLE zBizIntelChartTable1_" & Session("UserNo")
'Set rsCompanyLeakage = cnn8.Execute(SQLForChart)

cnn8.Close
Set cnn8 = Nothing
Set rsCompanyLeakage = Nothing
Set rs = Nothing
%>