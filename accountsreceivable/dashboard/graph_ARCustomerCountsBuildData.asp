<%
		
	firstDayOfMonthDate = loopMonth & "/1/" & loopYear
	lastDayOfMonth = GetLastDayofMonth(firstDayOfMonthDate)
	lastDayOfMonthDate = loopMonth & "/" & lastDayOfMonth  & "/" & loopYear
	
	'*******************************************************************************************************
	'FIRST GET THE TOTAL INACTIVE CUSTOMERS FOR CURRENT LOOP MONTH
	'*******************************************************************************************************
	
	SQLForInactiveCustomers = "SELECT COUNT(*) AS TOTALINACTIVE FROM AR_CustomerInactive WHERE DATEPART(month, RecordCreationDateTime) = '" & loopMonth & "' AND "
	SQLForInactiveCustomers = SQLForInactiveCustomers & " DATEPART(year, RecordCreationDateTime) = '" & loopYear & "' "
	
	Set rsARCustomerCountGraph = cnnARCustomerCountGraph.Execute(SQLForInactiveCustomers)
	
	If NOT rsARCustomerCountGraph.EOF Then
		numInactiveAccountsThisMonth = rsARCustomerCountGraph("TOTALINACTIVE")
	Else
		numInactiveAccountsThisMonth = 0	
	End If


	'*******************************************************************************************************
	'THEN GET THE TOTAL NEW ACTIVE CUSTOMERS FOR CURRENT LOOP MONTH
	'*******************************************************************************************************
	
	SQLActiveCustomers = "SELECT COUNT(*) AS NEWCUSTOMERS FROM AR_Customer WHERE DATEPART(month, InstallDate) = '" & loopMonth & "' AND "
	SQLActiveCustomers = SQLActiveCustomers & " DATEPART(year, InstallDate) = '" & loopYear & "' "

	Set rsARCustomerCountGraph = cnnARCustomerCountGraph.Execute(SQLActiveCustomers)
	
	If NOT rsARCustomerCountGraph.EOF Then
		numNewActiveAccountsThisMonth = rsARCustomerCountGraph("NEWCUSTOMERS")
	Else
		numNewActiveAccountsThisMonth = 0
	End If
	

	'*******************************************************************************************************
	'THEN GET THE TOTAL CUSTOMERS FOR CURRENT LOOP MONTH
	'*******************************************************************************************************
	
	SQLTotalCustomers = "SELECT * FROM AR_CustomerCounts WHERE RecordCreationDateTime = "
	SQLTotalCustomers = SQLTotalCustomers & " (SELECT MAX(RecordCreationDateTime) AS Expr1 FROM AR_CustomerCounts "
	SQLTotalCustomers = SQLTotalCustomers & " WHERE RecordCreationDateTime BETWEEN '" & firstDayOfMonthDate  & "' AND '" & lastDayOfMonthDate & "') "

	Set rsARCustomerCountGraph = cnnARCustomerCountGraph.Execute(SQLTotalCustomers)
	
	If NOT rsARCustomerCountGraph.EOF Then
		numTotalAccountsThisMonth = rsARCustomerCountGraph("numTotalAccounts")
	Else
		numTotalAccountsThisMonth = 0
	End If
		
	
	'*******************************************************************************************************
	'LASTLY, BUILD CHART DATA FOR CURRENT LOOP MONTH
	'*******************************************************************************************************

	amChartDataARCustCounts = amChartDataARCustCounts  & "{'month': '" & loopMonth & "/" & loopYear & "',"   
		
	amChartDataARCustCounts = amChartDataARCustCounts  & "'monthsingle': " & loopMonth & ","
	
	amChartDataARCustCounts = amChartDataARCustCounts  & "'year': " & loopYear & ","

	amChartDataARCustCounts = amChartDataARCustCounts  & "'totalaccounts': " & numTotalAccountsThisMonth & "," 
	
	amChartDataARCustCounts = amChartDataARCustCounts  & "'activeaccounts': " & numNewActiveAccountsThisMonth & ","
	
	amChartDataARCustCounts = amChartDataARCustCounts  & "'inactiveaccounts': " & numInactiveAccountsThisMonth & "}," 
		
%>
