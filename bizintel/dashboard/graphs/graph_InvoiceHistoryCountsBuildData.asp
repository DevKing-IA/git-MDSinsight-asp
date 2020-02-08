<%
		
	firstDayOfMonthDate = loopMonth & "/1/" & loopYear
	lastDayOfMonth = GetLastDayofMonth(firstDayOfMonthDate)
	lastDayOfMonthDate = loopMonth & "/" & lastDayOfMonth  & "/" & loopYear
	
	'*******************************************************************************************************
	'GET ALL ORDERS IN INVOICEHISTORY FOR CURRENT LOOP MONTH
	'*******************************************************************************************************

	SQLForTotalOrders = "SELECT COUNT(*) AS TOTALORDERS FROM InvoiceHistory WHERE DATEPART(month, IvsDate) = '" & loopMonth & "' AND "
	SQLForTotalOrders = SQLForTotalOrders & " DATEPART(year, IvsDate) = '" & loopYear & "' AND IvsTotalAmt > 0 AND (IvsType = 'T' OR IvsType = 'G') "
	
	Set rsInvHistoryCountGraph = cnnInvHistoryCountGraph.Execute(SQLForTotalOrders)
	
	If NOT rsInvHistoryCountGraph.EOF Then
		numTotalOrdersThisMonth = rsInvHistoryCountGraph("TOTALORDERS")
	Else
		numTotalOrdersThisMonth = 0	
	End If

	
	'*******************************************************************************************************
	'GET ALL WEBSITE ORDERS IN INVOICEHISTORY FOR CURRENT LOOP MONTH
	'*******************************************************************************************************

	SQLForTotalOrders = "SELECT COUNT(*) AS TOTALWEBORDERS FROM InvoiceHistory WHERE DATEPART(month, IvsDate) = '" & loopMonth & "' AND "
	SQLForTotalOrders = SQLForTotalOrders & " DATEPART(year, IvsDate) = '" & loopYear & "' AND LoginName = 'websel' AND IvsTotalAmt > 0 AND (IvsType = 'T' OR IvsType = 'G') "
	
	Set rsInvHistoryCountGraph = cnnInvHistoryCountGraph.Execute(SQLForTotalOrders)
	
	If NOT rsInvHistoryCountGraph.EOF Then
		numWebsiteOrdersThisMonth = rsInvHistoryCountGraph("TOTALWEBORDERS")
	Else
		numWebsiteOrdersThisMonth = 0	
	End If
	
	If numTotalOrdersThisMonth > 0 Then
		numWebsiteOrdersPercentOfTotalOrders = Round(((numWebsiteOrdersThisMonth/numTotalOrdersThisMonth) * 100),0)
	Else
		numWebsiteOrdersPercentOfTotalOrders = 0
	End If

	
	'*******************************************************************************************************
	'GET ALL TELSEL ORDERS IN INVOICEHISTORY FOR CURRENT LOOP MONTH
	'*******************************************************************************************************

	SQLForTotalOrders = "SELECT COUNT(*) AS TOTALTELSELORDERS FROM InvoiceHistory WHERE DATEPART(month, IvsDate) = '" & loopMonth & "' AND "
	SQLForTotalOrders = SQLForTotalOrders & " (DATEPART(month, IvsDate) = '" & loopMonth & "') AND (DATEPART(year, IvsDate) = '" & loopYear & "') "
	SQLForTotalOrders = SQLForTotalOrders & " AND (LoginName <> 'websel' AND LoginName <> 'api' OR LoginName IS NULL) AND IvsTotalAmt > 0 AND (IvsType = 'T' OR IvsType = 'G') "	
	
	Set rsInvHistoryCountGraph = cnnInvHistoryCountGraph.Execute(SQLForTotalOrders)
	
	If NOT rsInvHistoryCountGraph.EOF Then
		numTelselOrdersThisMonth = rsInvHistoryCountGraph("TOTALTELSELORDERS")
	Else
		numTelselOrdersThisMonth = 0	
	End If
	
	If numTotalOrdersThisMonth > 0 Then
		numTelselOrdersPercentOfTotalOrders = Round(((numTelselOrdersThisMonth/numTotalOrdersThisMonth) * 100),0)
	Else
		numTelselOrdersPercentOfTotalOrders = 0
	End If


	
	'*******************************************************************************************************
	'GET ALL API ORDERS IN INVOICEHISTORY FOR CURRENT LOOP MONTH
	'*******************************************************************************************************

	SQLForTotalOrders = "SELECT COUNT(*) AS TOTALAPIORDERS FROM InvoiceHistory WHERE DATEPART(month, IvsDate) = '" & loopMonth & "' AND "
	SQLForTotalOrders = SQLForTotalOrders & " DATEPART(year, IvsDate) = '" & loopYear & "' AND LoginName = 'api' "
	
	Set rsInvHistoryCountGraph = cnnInvHistoryCountGraph.Execute(SQLForTotalOrders)
	
	If NOT rsInvHistoryCountGraph.EOF Then
		numAPIOrdersThisMonth = rsInvHistoryCountGraph("TOTALAPIORDERS")
	Else
		numAPIOrdersThisMonth = 0	
	End If

	If numTotalOrdersThisMonth > 0 Then
		numAPIOrdersPercentOfTotalOrders = Round(((numAPIOrdersThisMonth/numTotalOrdersThisMonth) * 100),0)
	Else
		numAPIOrdersPercentOfTotalOrders = 0
	End If

	
	'*******************************************************************************************************
	'LASTLY, BUILD CHART DATA FOR CURRENT LOOP MONTH
	'*******************************************************************************************************

	
	amChartDataInvHistCounts = amChartDataInvHistCounts  & "{'month': '" & loopMonth & "/" & loopYear & "',"   
	
	amChartDataInvHistCounts = amChartDataInvHistCounts & "'monthsingle': " & loopMonth & ","
	
	amChartDataInvHistCounts = amChartDataInvHistCounts & "'year': " & loopYear & ","	

	amChartDataInvHistCounts = amChartDataInvHistCounts  & "'totalorders': " & numTotalOrdersThisMonth & "," 
	
	amChartDataInvHistCounts = amChartDataInvHistCounts  & "'websiteorders': " & numWebsiteOrdersThisMonth & ","
	
	amChartDataInvHistCounts = amChartDataInvHistCounts  & "'websiteorderspercent': " & numWebsiteOrdersPercentOfTotalOrders & ","
	
	amChartDataInvHistCounts = amChartDataInvHistCounts  & "'telselorders': " & numTelselOrdersThisMonth & ","
	
	amChartDataInvHistCounts = amChartDataInvHistCounts  & "'telselorderspercent': " & numTelselOrdersPercentOfTotalOrders & ","
	
	amChartDataInvHistCounts = amChartDataInvHistCounts  & "'apiorders': " & numAPIOrdersThisMonth & "," 
	
	amChartDataInvHistCounts = amChartDataInvHistCounts  & "'apiorderspercent': " & numAPIOrdersPercentOfTotalOrders & "},"
		
%>
