<%



Function GetCurrent_PostedTotal_ByReferralDesc2(passedReferralDesc2)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	
	resultGetCurrent_PostedTotal_ByReferralDesc2 = 0

	Set cnnGetCurrent_PostedTotal_ByReferralDesc2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByReferralDesc2.open Session("ClientCnnString")
		

	SQLGetCurrent_PostedTotal_ByReferralDesc2 = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_PostedTotal_ByReferralDesc2 = SQLGetCurrent_PostedTotal_ByReferralDesc2 & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_PostedTotal_ByReferralDesc2 = SQLGetCurrent_PostedTotal_ByReferralDesc2 & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_PostedTotal_ByReferralDesc2 = SQLGetCurrent_PostedTotal_ByReferralDesc2 & " (SELECT CustNum FROM AR_Customer WHERE ReferalCode IN (SELECT ReferalCode FROM Referal WHERE Description2 = '" & passedReferralDesc2 & "'))"


	Set rsGetCurrent_PostedTotal_ByReferralDesc2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByReferralDesc2.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByReferralDesc2 = cnnGetCurrent_PostedTotal_ByReferralDesc2.Execute(SQLGetCurrent_PostedTotal_ByReferralDesc2)

	If not rsGetCurrent_PostedTotal_ByReferralDesc2.EOF Then resultGetCurrent_PostedTotal_ByReferralDesc2 = rsGetCurrent_PostedTotal_ByReferralDesc2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByReferralDesc2) Then resultGetCurrent_PostedTotal_ByReferralDesc2 = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByReferralDesc2.Close
	set rsGetCurrent_PostedTotal_ByReferralDesc2= Nothing
	cnnGetCurrent_PostedTotal_ByReferralDesc2.Close	
	set cnnGetCurrent_PostedTotal_ByReferralDesc2= Nothing

	
	GetCurrent_PostedTotal_ByReferralDesc2 = resultGetCurrent_PostedTotal_ByReferralDesc2

End Function

Function GetCurrent_UnPostedTotal_ByReferralDesc2(passedReferralDesc2)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	
	resultGetCurrent_UnPostedTotal_ByReferralDesc2 = 0

	Set cnnGetCurrent_UnPostedTotal_ByReferralDesc2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnPostedTotal_ByReferralDesc2.open Session("ClientCnnString")
		

	SQLGetCurrent_UnPostedTotal_ByReferralDesc2 = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_UnPostedTotal_ByReferralDesc2 = SQLGetCurrent_UnPostedTotal_ByReferralDesc2 & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_UnPostedTotal_ByReferralDesc2 = SQLGetCurrent_UnPostedTotal_ByReferralDesc2 & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_UnPostedTotal_ByReferralDesc2 = SQLGetCurrent_UnPostedTotal_ByReferralDesc2 & " (SELECT CustNum FROM AR_Customer WHERE ReferalCode IN (SELECT ReferalCode FROM Referal WHERE Description2 = '" & passedReferralDesc2 & "'))"

	Set rsGetCurrent_UnPostedTotal_ByReferralDesc2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnPostedTotal_ByReferralDesc2.CursorLocation = 3 
	Set rsGetCurrent_UnPostedTotal_ByReferralDesc2 = cnnGetCurrent_UnPostedTotal_ByReferralDesc2.Execute(SQLGetCurrent_UnPostedTotal_ByReferralDesc2)

	If not rsGetCurrent_UnPostedTotal_ByReferralDesc2.EOF Then resultGetCurrent_UnPostedTotal_ByReferralDesc2 = rsGetCurrent_UnPostedTotal_ByReferralDesc2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnPostedTotal_ByReferralDesc2) Then resultGetCurrent_UnPostedTotal_ByReferralDesc2 = 0 ' In case there are no results
	
	rsGetCurrent_UnPostedTotal_ByReferralDesc2.Close
	set rsGetCurrent_UnPostedTotal_ByReferralDesc2= Nothing
	cnnGetCurrent_UnPostedTotal_ByReferralDesc2.Close	
	set cnnGetCurrent_UnPostedTotal_ByReferralDesc2= Nothing

	
	GetCurrent_UnPostedTotal_ByReferralDesc2 = resultGetCurrent_UnPostedTotal_ByReferralDesc2

End Function


Function GetCurrent_PostedTotal_ByCustType(passedCustType)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	
	resultGetCurrent_PostedTotal_ByCustType = 0

	Set cnnGetCurrent_PostedTotal_ByCustType = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByCustType.open Session("ClientCnnString")
		

	SQLGetCurrent_PostedTotal_ByCustType = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_PostedTotal_ByCustType = SQLGetCurrent_PostedTotal_ByCustType & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_PostedTotal_ByCustType = SQLGetCurrent_PostedTotal_ByCustType & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_PostedTotal_ByCustType = SQLGetCurrent_PostedTotal_ByCustType & " (SELECT CustNum FROM AR_Customer WHERE CustType ='" & passedCustType & "')"


	Set rsGetCurrent_PostedTotal_ByCustType = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByCustType.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByCustType = cnnGetCurrent_PostedTotal_ByCustType.Execute(SQLGetCurrent_PostedTotal_ByCustType)

	If not rsGetCurrent_PostedTotal_ByCustType.EOF Then resultGetCurrent_PostedTotal_ByCustType = rsGetCurrent_PostedTotal_ByCustType("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByCustType) Then resultGetCurrent_PostedTotal_ByCustType = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByCustType.Close
	set rsGetCurrent_PostedTotal_ByCustType= Nothing
	cnnGetCurrent_PostedTotal_ByCustType.Close	
	set cnnGetCurrent_PostedTotal_ByCustType= Nothing

	
	GetCurrent_PostedTotal_ByCustType = resultGetCurrent_PostedTotal_ByCustType

End Function

Function GetCurrent_UnPostedTotal_ByCustType(passedCustType)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	resultGetCurrent_UnPostedTotal_ByCustType = 0

	Set cnnGetCurrent_UnPostedTotal_ByCustType = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnPostedTotal_ByCustType.open Session("ClientCnnString")
		

	SQLGetCurrent_UnPostedTotal_ByCustType = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_UnPostedTotal_ByCustType = SQLGetCurrent_UnPostedTotal_ByCustType & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_UnPostedTotal_ByCustType = SQLGetCurrent_UnPostedTotal_ByCustType & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_UnPostedTotal_ByCustType = SQLGetCurrent_UnPostedTotal_ByCustType & " (SELECT CustNum FROM AR_Customer WHERE CustType ='" & passedCustType & "')"

	Set rsGetCurrent_UnPostedTotal_ByCustType = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnPostedTotal_ByCustType.CursorLocation = 3 
	Set rsGetCurrent_UnPostedTotal_ByCustType = cnnGetCurrent_UnPostedTotal_ByCustType.Execute(SQLGetCurrent_UnPostedTotal_ByCustType)

	If not rsGetCurrent_UnPostedTotal_ByCustType.EOF Then resultGetCurrent_UnPostedTotal_ByCustType = rsGetCurrent_UnPostedTotal_ByCustType("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnPostedTotal_ByCustType) Then resultGetCurrent_UnPostedTotal_ByCustType = 0 ' In case there are no results
	
	rsGetCurrent_UnPostedTotal_ByCustType.Close
	set rsGetCurrent_UnPostedTotal_ByCustType= Nothing
	cnnGetCurrent_UnPostedTotal_ByCustType.Close	
	set cnnGetCurrent_UnPostedTotal_ByCustType= Nothing

	
	GetCurrent_UnPostedTotal_ByCustType = resultGetCurrent_UnPostedTotal_ByCustType

End Function


Function GetCurrent_PostedTotal_ByPrimary(passedPrimary)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	
	resultGetCurrent_PostedTotal_ByPrimary = 0

	Set cnnGetCurrent_PostedTotal_ByPrimary = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByPrimary.open Session("ClientCnnString")
		

	SQLGetCurrent_PostedTotal_ByPrimary = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_PostedTotal_ByPrimary = SQLGetCurrent_PostedTotal_ByPrimary & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_PostedTotal_ByPrimary = SQLGetCurrent_PostedTotal_ByPrimary & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_PostedTotal_ByPrimary = SQLGetCurrent_PostedTotal_ByPrimary & " (SELECT CustNum FROM AR_Customer WHERE Salesman  ='" & passedPrimary & "')"


	Set rsGetCurrent_PostedTotal_ByPrimary = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByPrimary.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByPrimary = cnnGetCurrent_PostedTotal_ByPrimary.Execute(SQLGetCurrent_PostedTotal_ByPrimary)

	If not rsGetCurrent_PostedTotal_ByPrimary.EOF Then resultGetCurrent_PostedTotal_ByPrimary = rsGetCurrent_PostedTotal_ByPrimary("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByPrimary) Then resultGetCurrent_PostedTotal_ByPrimary = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByPrimary.Close
	set rsGetCurrent_PostedTotal_ByPrimary= Nothing
	cnnGetCurrent_PostedTotal_ByPrimary.Close	
	set cnnGetCurrent_PostedTotal_ByPrimary= Nothing

	
	GetCurrent_PostedTotal_ByPrimary = resultGetCurrent_PostedTotal_ByPrimary

End Function

Function GetCurrent_UnPostedTotal_ByPrimary(passedPrimary)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	resultGetCurrent_UnPostedTotal_ByPrimary = 0

	Set cnnGetCurrent_UnPostedTotal_ByPrimary = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnPostedTotal_ByPrimary.open Session("ClientCnnString")
		

	SQLGetCurrent_UnPostedTotal_ByPrimary = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_UnPostedTotal_ByPrimary = SQLGetCurrent_UnPostedTotal_ByPrimary & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_UnPostedTotal_ByPrimary = SQLGetCurrent_UnPostedTotal_ByPrimary & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_UnPostedTotal_ByPrimary = SQLGetCurrent_UnPostedTotal_ByPrimary & " (SELECT CustNum FROM AR_Customer WHERE Salesman  ='" & passedPrimary & "')"

	Set rsGetCurrent_UnPostedTotal_ByPrimary = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnPostedTotal_ByPrimary.CursorLocation = 3 
	Set rsGetCurrent_UnPostedTotal_ByPrimary = cnnGetCurrent_UnPostedTotal_ByPrimary.Execute(SQLGetCurrent_UnPostedTotal_ByPrimary)

	If not rsGetCurrent_UnPostedTotal_ByPrimary.EOF Then resultGetCurrent_UnPostedTotal_ByPrimary = rsGetCurrent_UnPostedTotal_ByPrimary("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnPostedTotal_ByPrimary) Then resultGetCurrent_UnPostedTotal_ByPrimary = 0 ' In case there are no results
	
	rsGetCurrent_UnPostedTotal_ByPrimary.Close
	set rsGetCurrent_UnPostedTotal_ByPrimary= Nothing
	cnnGetCurrent_UnPostedTotal_ByPrimary.Close	
	set cnnGetCurrent_UnPostedTotal_ByPrimary= Nothing

	
	GetCurrent_UnPostedTotal_ByPrimary = resultGetCurrent_UnPostedTotal_ByPrimary

End Function



Function GetCurrent_PostedTotal_BySecondary(passedSecondary)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	
	resultGetCurrent_PostedTotal_BySecondary = 0

	Set cnnGetCurrent_PostedTotal_BySecondary = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_BySecondary.open Session("ClientCnnString")
		

	SQLGetCurrent_PostedTotal_BySecondary = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_PostedTotal_BySecondary = SQLGetCurrent_PostedTotal_BySecondary & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_PostedTotal_BySecondary = SQLGetCurrent_PostedTotal_BySecondary & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_PostedTotal_BySecondary = SQLGetCurrent_PostedTotal_BySecondary & " (SELECT CustNum FROM AR_Customer WHERE SecondarySalesman  ='" & passedSecondary & "')"


	Set rsGetCurrent_PostedTotal_BySecondary = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_BySecondary.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_BySecondary = cnnGetCurrent_PostedTotal_BySecondary.Execute(SQLGetCurrent_PostedTotal_BySecondary)

	If not rsGetCurrent_PostedTotal_BySecondary.EOF Then resultGetCurrent_PostedTotal_BySecondary = rsGetCurrent_PostedTotal_BySecondary("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_BySecondary) Then resultGetCurrent_PostedTotal_BySecondary = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_BySecondary.Close
	set rsGetCurrent_PostedTotal_BySecondary= Nothing
	cnnGetCurrent_PostedTotal_BySecondary.Close	
	set cnnGetCurrent_PostedTotal_BySecondary= Nothing

	
	GetCurrent_PostedTotal_BySecondary = resultGetCurrent_PostedTotal_BySecondary

End Function

Function GetCurrent_UnPostedTotal_BySecondary(passedSecondary)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	resultGetCurrent_UnPostedTotal_BySecondary = 0

	Set cnnGetCurrent_UnPostedTotal_BySecondary = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnPostedTotal_BySecondary.open Session("ClientCnnString")
		

	SQLGetCurrent_UnPostedTotal_BySecondary = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_UnPostedTotal_BySecondary = SQLGetCurrent_UnPostedTotal_BySecondary & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_UnPostedTotal_BySecondary = SQLGetCurrent_UnPostedTotal_BySecondary & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_UnPostedTotal_BySecondary = SQLGetCurrent_UnPostedTotal_BySecondary & " (SELECT CustNum FROM AR_Customer WHERE SecondarySalesman  ='" & passedSecondary & "')"

	Set rsGetCurrent_UnPostedTotal_BySecondary = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnPostedTotal_BySecondary.CursorLocation = 3 
	Set rsGetCurrent_UnPostedTotal_BySecondary = cnnGetCurrent_UnPostedTotal_BySecondary.Execute(SQLGetCurrent_UnPostedTotal_BySecondary)

	If not rsGetCurrent_UnPostedTotal_BySecondary.EOF Then resultGetCurrent_UnPostedTotal_BySecondary = rsGetCurrent_UnPostedTotal_BySecondary("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnPostedTotal_BySecondary) Then resultGetCurrent_UnPostedTotal_BySecondary = 0 ' In case there are no results
	
	rsGetCurrent_UnPostedTotal_BySecondary.Close
	set rsGetCurrent_UnPostedTotal_BySecondary= Nothing
	cnnGetCurrent_UnPostedTotal_BySecondary.Close	
	set cnnGetCurrent_UnPostedTotal_BySecondary= Nothing

	
	GetCurrent_UnPostedTotal_BySecondary = resultGetCurrent_UnPostedTotal_BySecondary

End Function

%>
