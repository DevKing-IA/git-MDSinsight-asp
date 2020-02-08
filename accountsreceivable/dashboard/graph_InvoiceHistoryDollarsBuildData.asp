<%
		
	firstDayOfMonthDate = loopMonth & "/1/" & loopYear
	lastDayOfMonth = GetLastDayofMonth(firstDayOfMonthDate)
	lastDayOfMonthDate = loopMonth & "/" & lastDayOfMonth  & "/" & loopYear
	
	'*******************************************************************************************************
	'GET ALL ORDERS IN INVOICEHISTORY FOR CURRENT LOOP MONTH
	'*******************************************************************************************************

	SQLForTotalOrders = "SELECT Sum(IvsTotalAmt-IvsSalesTax-IvsDepositChg) AS TotalSales FROM InvoiceHistory WHERE DATEPART(month, IvsDate) = '" & loopMonth & "' AND "
	SQLForTotalOrders = SQLForTotalOrders & " DATEPART(year, IvsDate) = '" & loopYear & "' AND IvsTotalAmt > 0 AND (IvsType = 'T' OR IvsType = 'G') "
	
	Set rsInvHistoryDollarsGraph = cnnInvHistoryDollarsGraph.Execute(SQLForTotalOrders)
	
	If NOT rsInvHistoryDollarsGraph.EOF Then
		numTotalOrderDollarsThisMonth = rsInvHistoryDollarsGraph("TotalSales")
	Else
		numTotalOrderDollarsThisMonth = 0	
	End If

	If IsNull(numTotalOrderDollarsThisMonth) OR IsEmpty(numTotalOrderDollarsThisMonth) OR numTotalOrderDollarsThisMonth = "" OR Len(numTotalOrderDollarsThisMonth) < 1 Then
		numTotalOrderDollarsThisMonth = 0
	End If
	
	'*******************************************************************************************************
	'GET ALL WEBSITE ORDERS IN INVOICEHISTORY FOR CURRENT LOOP MONTH
	'*******************************************************************************************************

	SQLForTotalOrders = "SELECT Sum(IvsTotalAmt-IvsSalesTax-IvsDepositChg) AS TotalSales FROM InvoiceHistory WHERE DATEPART(month, IvsDate) = '" & loopMonth & "' AND "
	SQLForTotalOrders = SQLForTotalOrders & " DATEPART(year, IvsDate) = '" & loopYear & "' AND LoginName = 'websel' AND IvsTotalAmt > 0 AND (IvsType = 'T' OR IvsType = 'G') "
	
	Set rsInvHistoryDollarsGraph = cnnInvHistoryDollarsGraph.Execute(SQLForTotalOrders)
	
	If NOT rsInvHistoryDollarsGraph.EOF Then
		numWebsiteOrderDollarsThisMonth = rsInvHistoryDollarsGraph("TotalSales")
	Else
		numWebsiteOrderDollarsThisMonth = 0	
	End If
	
	If IsNull(numWebsiteOrderDollarsThisMonth) OR IsEmpty(numWebsiteOrderDollarsThisMonth) OR numWebsiteOrderDollarsThisMonth = "" OR Len(numWebsiteOrderDollarsThisMonth) < 1 Then
		numWebsiteOrderDollarsThisMonth = 0
	End If
	

	
	'*******************************************************************************************************
	'GET ALL TELSEL ORDERS IN INVOICEHISTORY FOR CURRENT LOOP MONTH
	'*******************************************************************************************************

	SQLForTotalOrders = "SELECT Sum(IvsTotalAmt-IvsSalesTax-IvsDepositChg) AS TotalSales FROM InvoiceHistory WHERE DATEPART(month, IvsDate) = '" & loopMonth & "' AND "
	SQLForTotalOrders = SQLForTotalOrders & " (DATEPART(month, IvsDate) = '" & loopMonth & "') AND (DATEPART(year, IvsDate) = '" & loopYear & "') "
	SQLForTotalOrders = SQLForTotalOrders & " AND (LoginName <> 'websel' AND LoginName <> 'api' OR LoginName IS NULL) AND IvsTotalAmt > 0 AND (IvsType = 'T' OR IvsType = 'G') "	
	
	Set rsInvHistoryDollarsGraph = cnnInvHistoryDollarsGraph.Execute(SQLForTotalOrders)
	
	If NOT rsInvHistoryDollarsGraph.EOF Then
		numTelselOrderDollarsThisMonth = rsInvHistoryDollarsGraph("TotalSales")
	Else
		numTelselOrderDollarsThisMonth = 0	
	End If
	
	If IsNull(numTelselOrderDollarsThisMonth) OR IsEmpty(numTelselOrderDollarsThisMonth) OR numTelselOrderDollarsThisMonth = "" OR Len(numTelselOrderDollarsThisMonth) < 1 Then
		numTelselOrderDollarsThisMonth = 0
	End If

	
	'*******************************************************************************************************
	'GET ALL API ORDERS IN INVOICEHISTORY FOR CURRENT LOOP MONTH
	'*******************************************************************************************************

	SQLForTotalOrders = "SELECT Sum(IvsTotalAmt-IvsSalesTax-IvsDepositChg) AS TotalSales FROM InvoiceHistory WHERE DATEPART(month, IvsDate) = '" & loopMonth & "' AND "
	SQLForTotalOrders = SQLForTotalOrders & " DATEPART(year, IvsDate) = '" & loopYear & "' AND LoginName = 'api' "
	
	Set rsInvHistoryDollarsGraph = cnnInvHistoryDollarsGraph.Execute(SQLForTotalOrders)
	
	If NOT rsInvHistoryDollarsGraph.EOF Then
		numAPIOrderDollarsThisMonth = rsInvHistoryDollarsGraph("TotalSales")
	Else
		numAPIOrderDollarsThisMonth = 0	
	End If

	If IsNull(numAPIOrderDollarsThisMonth) OR IsEmpty(numAPIOrderDollarsThisMonth) OR numAPIOrderDollarsThisMonth = "" OR Len(numAPIOrderDollarsThisMonth) < 1 Then
		numAPIOrderDollarsThisMonth = 0
	End If

	
	'*******************************************************************************************************
	'LASTLY, BUILD CHART DATA FOR CURRENT LOOP MONTH
	'*******************************************************************************************************

	
	amChartDataInvHistDollars = amChartDataInvHistDollars & "{'month': '" & loopMonth & "/" & loopYear & "',"   
	
	amChartDataInvHistDollars = amChartDataInvHistDollars & "'monthsingle': " & loopMonth & ","
	
	amChartDataInvHistDollars = amChartDataInvHistDollars & "'year': " & loopYear & ","	

	amChartDataInvHistDollars = amChartDataInvHistDollars & "'totalordersdollars': " & numTotalOrderDollarsThisMonth & "," 
	
	amChartDataInvHistDollars = amChartDataInvHistDollars & "'websiteordersdollars': " & numWebsiteOrderDollarsThisMonth & ","
	
	amChartDataInvHistDollars = amChartDataInvHistDollars & "'telselordersdollars': " & numTelselOrderDollarsThisMonth & ","
	
	amChartDataInvHistDollars = amChartDataInvHistDollars & "'apiordersdollars': " & numAPIOrderDollarsThisMonth & "},"
		
%>
