<div id="chartdivThisPerProj" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<%

PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum() 

WorkDaysIn3PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -3), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1
WorkDaysIn12PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -12), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1)) + 1 
WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated )) 
WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated +1)) + 1 
WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1),Date()) + 1

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3

' Get Last Closed Period and Three Prior Period Average
SQLForChart = "SELECT SUM(TotalSales) AS LCPSales, SUM([3PriorPeriodsAeverage]) AS P3PSales FROM CustCatPeriodSales_ReportData WHERE ThisPeriodSequenceNumber = " & GetLastClosedPeriodSeqNum()


'Response.Write(SQLForChartCust)

Set rs = cnn8.Execute(SQLForChart)

If not rs.EOF Then
	GrandTotLCPSales = Round(rs("LCPSales"),0)
	GrandTotP3PAvgSales = Round(rs("P3PSales"),0)
End If


' Get projection
CurrentDollars = GetCurrent_PostedTotal + GetCurrent_UnPostedTotal()

CurrentADS = CurrentDollars / WorkDaysSoFar 
									
NextPeriodProj = Round(CurrentADS * WorkDaysInCurrentPeriod,0)


cnn8.Close
Set cnn8 = Nothing
Set rs = Nothing
	

'Now Build the Chart Data Variable

amChartDataThisPerProj = ""

amChartDataThisPerProj = amChartDataThisPerProj & "{"
amChartDataThisPerProj = amChartDataThisPerProj & "'period': '" & GetPeriodAndYearBySeq(PeriodSeqBeingEvaluated) & "',"
amChartDataThisPerProj = amChartDataThisPerProj & "'dollars': '" & GrandTotLCPSales & "',"
amChartDataThisPerProj = amChartDataThisPerProj & "'color': '#999999'"
amChartDataThisPerProj = amChartDataThisPerProj & "},{"
amChartDataThisPerProj = amChartDataThisPerProj & "'period': 'P3P avg',"
amChartDataThisPerProj = amChartDataThisPerProj & "'dollars': '" & GrandTotP3PAvgSales & "',"
amChartDataThisPerProj = amChartDataThisPerProj & "'color': '#999999'"
amChartDataThisPerProj = amChartDataThisPerProj & "},{"
amChartDataThisPerProj = amChartDataThisPerProj & "'period': 'This Period',"
amChartDataThisPerProj = amChartDataThisPerProj & "'dollars': '" & NextPeriodProj & "',"
amChartDataThisPerProj = amChartDataThisPerProj & "'color': '#8A0CCF'"
amChartDataThisPerProj = amChartDataThisPerProj & "}"


Function GetCurrent_PostedTotal()
	
	resultGetCurrent_PostedTotal = 0

	Set cnnGetCurrent_PostedTotal = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal.open Session("ClientCnnString")
		

	SQLGetCurrent_PostedTotal = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_PostedTotal = SQLGetCurrent_PostedTotal & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1

	Set rsGetCurrent_PostedTotal = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal = cnnGetCurrent_PostedTotal.Execute(SQLGetCurrent_PostedTotal)

	If not rsGetCurrent_PostedTotal.EOF Then resultGetCurrent_PostedTotal = rsGetCurrent_PostedTotal("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal) Then resultGetCurrent_PostedTotal = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal.Close
	set rsGetCurrent_PostedTotal= Nothing
	cnnGetCurrent_PostedTotal.Close	
	set cnnGetCurrent_PostedTotal= Nothing
	
	GetCurrent_PostedTotal = resultGetCurrent_PostedTotal

End Function


Function GetCurrent_UnPostedTotal()

	resultGetCurrent_UnPostedTotal = 0

	Set cnnGetCurrent_UnPostedTotal = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnPostedTotal.open Session("ClientCnnString")
		

	SQLGetCurrent_UnPostedTotal = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_UnPostedTotal = SQLGetCurrent_UnPostedTotal & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1

	Set rsGetCurrent_UnPostedTotal = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnPostedTotal.CursorLocation = 3 
	Set rsGetCurrent_UnPostedTotal = cnnGetCurrent_UnPostedTotal.Execute(SQLGetCurrent_UnPostedTotal)

	If not rsGetCurrent_UnPostedTotal.EOF Then resultGetCurrent_UnPostedTotal = rsGetCurrent_UnPostedTotal("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnPostedTotal) Then resultGetCurrent_UnPostedTotal = 0 ' In case there are no results
	
	rsGetCurrent_UnPostedTotal.Close
	set rsGetCurrent_UnPostedTotal= Nothing
	cnnGetCurrent_UnPostedTotal.Close	
	set cnnGetCurrent_UnPostedTotal= Nothing

	GetCurrent_UnPostedTotal = resultGetCurrent_UnPostedTotal

End Function

%>