<%
'************************************************************
'List of all the functions & subs
'************************************************************
'Function GetPartnerNameByAPIKey(passedAPIKey)
'Function GetNumberOfOrdersByDate(passedDate)
'Function GetNumberOfInvoicesByDate(passedDate)
'Function GetNumberOfRAsByDate(passedDate)
'Function GetNumberOfCMsByDate(passedDate)
'Function GetNumberOfSummaryInvoicesByDate(passedDate)
'Func GetAPIRepostURL()
'Func GetAPIRepostInvoicesURL()
'Func GetRAIDByOrderID(passedOrderID)
'Func GetOrderDriverNotesByOrderID(passedOrderID)
'Func GetOrderWarehouseNotesByOrderID(passedOrderID)
'Func GetRADriverNotesByOrderID(passedOrderID)
'Func GetRAWarehouseNotesByOrderID(passedOrderID)
'Func GetRARANotesByOrderID(passedOrderID)
'Func GetAPIRepostSumInvURL()
'Func GetAPIRepostRAURL()
'************************************************************
'End List of all the functions & subs
'************************************************************

'*******************************************************************************************************************************
'*******************************************************************************************************************************
Function GetAPIRepostInvoicesURL()

	resultGetAPIRepostInvoicesURL = ""

	Set cnnGetAPIRepostInvoicesURL = Server.CreateObject("ADODB.Connection")
	cnnGetAPIRepostInvoicesURL.open Session("ClientCnnString")
		
	SQLGetAPIRepostInvoicesURL = "SELECT * FROM Settings_Global"
 
	Set rsGetAPIRepostInvoicesURL = Server.CreateObject("ADODB.Recordset")
	rsGetAPIRepostInvoicesURL.CursorLocation = 3 
	Set rsGetAPIRepostInvoicesURL = cnnGetAPIRepostInvoicesURL.Execute(SQLGetAPIRepostInvoicesURL)
			 
	If not rsGetAPIRepostInvoicesURL.EOF Then  
		If Not IsNull(rsGetAPIRepostInvoicesURL("InvoiceAPIRepostURL")) Then resultGetAPIRepostInvoicesURL = rsGetAPIRepostInvoicesURL("InvoiceAPIRepostURL")
	End If
	
	set rsGetAPIRepostInvoicesURL= Nothing
	cnnGetAPIRepostInvoicesURL.Close	
	set cnnGetAPIRepostInvoicesURL= Nothing
	
	GetAPIRepostInvoicesURL = resultGetAPIRepostInvoicesURL
	
End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************



'*******************************************************************************************************************************
'*******************************************************************************************************************************

Function GetAPIRepostURL()

	resultGetAPIRepostURL = ""

	Set cnnGetAPIRepostURL = Server.CreateObject("ADODB.Connection")
	cnnGetAPIRepostURL.open Session("ClientCnnString")
		
	SQLGetAPIRepostURL = "SELECT * FROM Settings_Global"
 
	Set rsGetAPIRepostURL = Server.CreateObject("ADODB.Recordset")
	rsGetAPIRepostURL.CursorLocation = 3 
	Set rsGetAPIRepostURL = cnnGetAPIRepostURL.Execute(SQLGetAPIRepostURL)
			 
	If not rsGetAPIRepostURL.EOF Then  
		If Not IsNull(rsGetAPIRepostURL("OrderAPIRepostURL")) Then resultGetAPIRepostURL = rsGetAPIRepostURL("OrderAPIRepostURL")
	End If
	
	set rsGetAPIRepostURL= Nothing
	cnnGetAPIRepostURL.Close	
	set cnnGetAPIRepostURL= Nothing
	
	GetAPIRepostURL = resultGetAPIRepostURL
	
End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************




'*******************************************************************************************************************************
'*******************************************************************************************************************************

Function GetPartnerNameByAPIKey(passedAPIKey)

	Set cnnGetPartnerNameByAPIKey  = Server.CreateObject("ADODB.Connection")
	cnnGetPartnerNameByAPIKey.open Session("ClientCnnString")

	resultGetPartnerNameByAPIKey = 0
		
	SQLGetPartnerNameByAPIKey  = "SELECT partnerCompanyName FROM IC_Partners WHERE partnerAPIKey = '" & passedAPIKey & "'"
	 
	Set rsGetPartnerNameByAPIKey  = Server.CreateObject("ADODB.Recordset")
	rsGetPartnerNameByAPIKey.CursorLocation = 3 
	
	rsGetPartnerNameByAPIKey.Open SQLGetPartnerNameByAPIKey,cnnGetPartnerNameByAPIKey 
			
	resultGetPartnerNameByAPIKey = rsGetPartnerNameByAPIKey("partnerCompanyName")
	
	rsGetPartnerNameByAPIKey.Close
	set rsGetPartnerNameByAPIKey = Nothing
	cnnGetPartnerNameByAPIKey.Close	
	set cnnGetPartnerNameByAPIKey = Nothing
	
	GetPartnerNameByAPIKey = resultGetPartnerNameByAPIKey
	
End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************




'*******************************************************************************************************************************
'*******************************************************************************************************************************

Function GetNumberOfOrdersByDate(passedDate)

	result = 0

	currentDay = day(passedDate)
	currentMonth = month(passedDate)
	currentYear = year(passedDate)
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQLDailyAPIOrders = "SELECT COUNT(*) AS NumOrders, SUM(OrderSubTotal) AS Subtotal, SUM(Tax) AS TaxTot, SUM(ShippingCharge) AS ShipTot, "
	SQLDailyAPIOrders = SQLDailyAPIOrders & " SUM(DepositCharge) AS DepositTot, SUM(FuelSurcharge) AS FuelTot, SUM(CouponCharge) AS CouponTot, SUM(GrandTotal) AS GranTot"
	SQLDailyAPIOrders = SQLDailyAPIOrders & " FROM            API_OR_OrderHeader"
	SQLDailyAPIOrders = SQLDailyAPIOrders & " WHERE        (InternalRecordIdentifier IN"
	SQLDailyAPIOrders = SQLDailyAPIOrders & " (SELECT        MAX(InternalRecordIdentifier) AS Expr1"
	SQLDailyAPIOrders = SQLDailyAPIOrders & " FROM            API_OR_OrderHeader AS API_OR_OrderHeader_1"
	SQLDailyAPIOrders = SQLDailyAPIOrders & " WHERE        (DAY(OrderDate) = " & CurrentDay & " AND MONTH(OrderDate) =  " & CurrentMonth & " AND YEAR(OrderDate) =  " & CurrentYear & " AND (Voided = 0)) "
	SQLDailyAPIOrders = SQLDailyAPIOrders & " GROUP BY OrderID)) "	

	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQLDailyAPIOrders)

	If not rsBoost1.eof then 
		result = rsBoost1("NumOrders")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetNumberOfOrdersByDate = result

End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************



'*******************************************************************************************************************************
'*******************************************************************************************************************************

Function GetNumberOfInvoicesByDate(passedDate)

	result = 0
	
	currentDay = day(passedDate)
	currentMonth = month(passedDate)
	currentYear = year(passedDate)	
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQLDailyAPIInvoices = "SELECT COUNT(*) AS NumInv, SUM(InvoiceSubTotal) AS Subtotal, SUM(Tax) AS TaxTot, SUM(ShippingCharge) AS ShipTot, "
	SQLDailyAPIInvoices = SQLDailyAPIInvoices & " SUM(DepositCharge) AS DepositTot, SUM(FuelSurcharge) AS FuelTot, SUM(CouponCharge) AS CouponTot, SUM(GrandTotal) AS GranTot "
	SQLDailyAPIInvoices = SQLDailyAPIInvoices & " FROM            API_IN_InvoiceHeader "
	SQLDailyAPIInvoices = SQLDailyAPIInvoices & " WHERE        (InternalRecordIdentifier IN "
	SQLDailyAPIInvoices = SQLDailyAPIInvoices & " (SELECT        MAX(InternalRecordIdentifier) AS Expr1 "
	SQLDailyAPIInvoices = SQLDailyAPIInvoices & " FROM            API_IN_InvoiceHeader AS API_IN_InvoiceHeader_1 "
	SQLDailyAPIInvoices = SQLDailyAPIInvoices & " WHERE        (DAY(RecordCreationDateTime) = " & CurrentDay & " AND MONTH(RecordCreationDateTime) =  " & CurrentMonth & " AND YEAR(RecordCreationDateTime) =  " & CurrentYear & " AND (Voided = 0)) "
	SQLDailyAPIInvoices = SQLDailyAPIInvoices & " GROUP BY InvoiceID))	"	

	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQLDailyAPIInvoices)

	If not rsBoost1.eof then 
		result = rsBoost1("NumInv")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetNumberOfInvoicesByDate = result

End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************



'*******************************************************************************************************************************
'*******************************************************************************************************************************

Function GetNumberOfRAsByDate(passedDate)

	result = 0
	
	currentDay = day(passedDate)
	currentMonth = month(passedDate)
	currentYear = year(passedDate)	
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQLDailyAPIReturnAuths = " SELECT  COUNT(*) AS NumRA, SUM(SubTotal) AS Subtotal, SUM(Tax) AS TaxTot, SUM(ShippingCharge) AS shiptot, "
	SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " SUM(DepositCharge) AS DepositTot, SUM(FuelSurcharge) AS FuelTot, SUM(CouponCharge) AS CouponTot, "
	SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " SUM(GrandTotal) AS GranTot "
	SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " FROM            API_OR_RAHeader "
	SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " WHERE        (InternalRecordIdentifier IN "
	SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " (SELECT        MAX(InternalRecordIdentifier) AS Expr1 "
	SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " FROM            API_OR_RAHeader AS API_OR_RAHeader_1 "
	SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " WHERE        (DAY(RecordCreationDateTime) = " & CurrentDay & " AND MONTH(RecordCreationDateTime) =  " & CurrentMonth & " AND YEAR(RecordCreationDateTime) =  " & CurrentYear & " AND (Voided = 0)) "
	SQLDailyAPIReturnAuths = SQLDailyAPIReturnAuths & " GROUP BY RAID))	"	
				
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQLDailyAPIReturnAuths)

	If not rsBoost1.eof then 
		result = rsBoost1("NumRA")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetNumberOfRAsByDate = result

End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************



'*******************************************************************************************************************************
'*******************************************************************************************************************************

Function GetNumberOfCMsByDate(passedDate)

	result = 0
	
	currentDay = day(passedDate)
	currentMonth = month(passedDate)
	currentYear = year(passedDate)	
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQLDailyAPICreditMemos = "SELECT COUNT(*) AS NumInv, SUM(CMSubTotal) AS Subtotal, SUM(Tax) AS TaxTot, SUM(ShippingCharge) AS ShipTot, "
	SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " SUM(DepositCharge) AS DepositTot, SUM(FuelSurcharge) AS FuelTot, SUM(CouponCharge) AS CouponTot, SUM(GrandTotal) AS GranTot "
	SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " FROM            API_IN_CMHeader "
	SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " WHERE        (InternalRecordIdentifier IN "
	SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " (SELECT        MAX(InternalRecordIdentifier) AS Expr1 "
	SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " FROM            API_IN_CMHeader AS API_IN_CMHeader_1 "
	SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " WHERE        (DAY(RecordCreationDateTime) = " & CurrentDay & " AND MONTH(RecordCreationDateTime) =  " & CurrentMonth & " AND YEAR(RecordCreationDateTime) =  " & CurrentYear & " AND (Voided = 0)) "
	SQLDailyAPICreditMemos = SQLDailyAPICreditMemos & " GROUP BY CMID))	"	

	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQLDailyAPICreditMemos)

	If not rsBoost1.eof then 
		result = rsBoost1("NumInv")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetNumberOfCMsByDate = result

End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************



'*******************************************************************************************************************************
'*******************************************************************************************************************************

Function GetNumberOfSummaryInvoicesByDate(passedDate)

	result = 0
	
	currentDay = day(passedDate)
	currentMonth = month(passedDate)
	currentYear = year(passedDate)	
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQLDailyAPISummInv = "SELECT COUNT(*) AS NumInv, SUM(Sub_Total) AS Subtotal, SUM(Total_Tax) AS TaxTot, SUM(Shipping_Charge) AS ShipTot, "
	SQLDailyAPISummInv = SQLDailyAPISummInv & " SUM(DepositCharge) AS DepositTot, SUM(FuelSurcharge) AS FuelTot, SUM(CouponCharge) AS CouponTot, SUM(Grand_Total) AS GranTot "
	SQLDailyAPISummInv = SQLDailyAPISummInv & " FROM            API_IN_SummaryInvoiceHeader "
	SQLDailyAPISummInv = SQLDailyAPISummInv & " WHERE        (InternalRecordIdentifier IN "
	SQLDailyAPISummInv = SQLDailyAPISummInv & " (SELECT        MAX(InternalRecordIdentifier) AS Expr1 "
	SQLDailyAPISummInv = SQLDailyAPISummInv & " FROM            API_IN_SummaryInvoiceHeader AS API_IN_SummaryInvoiceHeader_1 "
	SQLDailyAPISummInv = SQLDailyAPISummInv & " WHERE        (DAY(RecordCreationDateTime) = " & CurrentDay & " AND MONTH(RecordCreationDateTime) =  " & CurrentMonth & " AND YEAR(RecordCreationDateTime) =  " & CurrentYear & " AND (Voided = 0)) "
	SQLDailyAPISummInv = SQLDailyAPISummInv & " GROUP BY SumInvID))	"	

	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQLDailyAPISummInv)

	If not rsBoost1.eof then 
		result = rsBoost1("NumInv")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetNumberOfSummaryInvoicesByDate = result

End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************




'*******************************************************************************************************************************
'*******************************************************************************************************************************
Function GetRAIDByOrderID(passedOrderID)

	Set cnnGetRAIDByOrderID  = Server.CreateObject("ADODB.Connection")
	cnnGetRAIDByOrderID.open Session("ClientCnnString")

	resultGetRAIDByOrderID = ""
		
	SQLGetRAIDByOrderID  = "SELECT RAID FROM API_OR_RAHeader WHERE OrderID = '" & passedOrderID & "'"
	 
	Set rsGetRAIDByOrderID  = Server.CreateObject("ADODB.Recordset")
	rsGetRAIDByOrderID.CursorLocation = 3 
	
	rsGetRAIDByOrderID.Open SQLGetRAIDByOrderID,cnnGetRAIDByOrderID 
			
	resultGetRAIDByOrderID = rsGetRAIDByOrderID("RAID")
	
	rsGetRAIDByOrderID.Close
	set rsGetRAIDByOrderID = Nothing
	cnnGetRAIDByOrderID.Close	
	set cnnGetRAIDByOrderID = Nothing
	
	GetRAIDByOrderID = resultGetRAIDByOrderID


End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************



'*******************************************************************************************************************************
'*******************************************************************************************************************************
Function GetOrderDriverNotesByOrderID(passedOrderID)

	Set cnnGetOrderDriverNotesByOrderID  = Server.CreateObject("ADODB.Connection")
	cnnGetOrderDriverNotesByOrderID.open Session("ClientCnnString")

	resultGetOrderDriverNotesByOrderID = ""
		
	SQLGetOrderDriverNotesByOrderID  = "SELECT DriverNotes FROM API_OR_OrderHeader WHERE OrderID = '" & passedOrderID & "'"
	 
	Set rsGetOrderDriverNotesByOrderID  = Server.CreateObject("ADODB.Recordset")
	rsGetOrderDriverNotesByOrderID.CursorLocation = 3 
	
	rsGetOrderDriverNotesByOrderID.Open SQLGetOrderDriverNotesByOrderID,cnnGetOrderDriverNotesByOrderID 
			
	resultGetOrderDriverNotesByOrderID = rsGetOrderDriverNotesByOrderID("DriverNotes")
	
	rsGetOrderDriverNotesByOrderID.Close
	set rsGetOrderDriverNotesByOrderID = Nothing
	cnnGetOrderDriverNotesByOrderID.Close	
	set cnnGetOrderDriverNotesByOrderID = Nothing
	
	GetOrderDriverNotesByOrderID = resultGetOrderDriverNotesByOrderID

End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************


'*******************************************************************************************************************************
'*******************************************************************************************************************************
Function GetOrderWarehouseNotesByOrderID(passedOrderID)

	Set cnnGetOrderWarehouseNotesByOrderID  = Server.CreateObject("ADODB.Connection")
	cnnGetOrderWarehouseNotesByOrderID.open Session("ClientCnnString")

	resultGetOrderWarehouseNotesByOrderID = ""
		
	SQLGetOrderWarehouseNotesByOrderID  = "SELECT WarehouseNotes FROM API_OR_OrderHeader WHERE OrderID = '" & passedOrderID & "'"
	 
	Set rsGetOrderWarehouseNotesByOrderID  = Server.CreateObject("ADODB.Recordset")
	rsGetOrderWarehouseNotesByOrderID.CursorLocation = 3 
	
	rsGetOrderWarehouseNotesByOrderID.Open SQLGetOrderWarehouseNotesByOrderID,cnnGetOrderWarehouseNotesByOrderID 
			
	resultGetOrderWarehouseNotesByOrderID = rsGetOrderWarehouseNotesByOrderID("WarehouseNotes")
	
	rsGetOrderWarehouseNotesByOrderID.Close
	set rsGetOrderWarehouseNotesByOrderID = Nothing
	cnnGetOrderWarehouseNotesByOrderID.Close	
	set cnnGetOrderWarehouseNotesByOrderID = Nothing
	
	GetOrderWarehouseNotesByOrderID = resultGetOrderWarehouseNotesByOrderID

End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************


'*******************************************************************************************************************************
'*******************************************************************************************************************************
Function GetRADriverNotesByOrderID(passedOrderID)

	Set cnnGetRADriverNotesByOrderID  = Server.CreateObject("ADODB.Connection")
	cnnGetRADriverNotesByOrderID.open Session("ClientCnnString")

	resultGetRADriverNotesByOrderID = ""
		
	SQLGetRADriverNotesByOrderID  = "SELECT DriverNotes FROM API_OR_RAHeader WHERE OrderID = '" & passedOrderID & "'"
	 
	Set rsGetRADriverNotesByOrderID  = Server.CreateObject("ADODB.Recordset")
	rsGetRADriverNotesByOrderID.CursorLocation = 3 
	
	rsGetRADriverNotesByOrderID.Open SQLGetRADriverNotesByOrderID,cnnGetRADriverNotesByOrderID 
			
	resultGetRADriverNotesByOrderID = rsGetRADriverNotesByOrderID("DriverNotes")
	
	rsGetRADriverNotesByOrderID.Close
	set rsGetRADriverNotesByOrderID = Nothing
	cnnGetRADriverNotesByOrderID.Close	
	set cnnGetRADriverNotesByOrderID = Nothing
	
	GetRADriverNotesByOrderID = resultGetRADriverNotesByOrderID

End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************


'*******************************************************************************************************************************
'*******************************************************************************************************************************
Function GetRAWarehouseNotesByOrderID(passedOrderID)

	Set cnnGetRAWarehouseNotesByOrderID  = Server.CreateObject("ADODB.Connection")
	cnnGetRAWarehouseNotesByOrderID.open Session("ClientCnnString")

	resultGetRAWarehouseNotesByOrderID = ""
		
	SQLGetRAWarehouseNotesByOrderID  = "SELECT WarehouseNotes FROM API_OR_RAHeader WHERE OrderID = '" & passedOrderID & "'"
	 
	Set rsGetRAWarehouseNotesByOrderID  = Server.CreateObject("ADODB.Recordset")
	rsGetRAWarehouseNotesByOrderID.CursorLocation = 3 
	
	rsGetRAWarehouseNotesByOrderID.Open SQLGetRAWarehouseNotesByOrderID,cnnGetRAWarehouseNotesByOrderID 
			
	resultGetRAWarehouseNotesByOrderID = rsGetRAWarehouseNotesByOrderID("WarehouseNotes")
	
	rsGetRAWarehouseNotesByOrderID.Close
	set rsGetRAWarehouseNotesByOrderID = Nothing
	cnnGetRAWarehouseNotesByOrderID.Close	
	set cnnGetRAWarehouseNotesByOrderID = Nothing
	
	GetRAWarehouseNotesByOrderID = resultGetRAWarehouseNotesByOrderID

End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************




'*******************************************************************************************************************************
'*******************************************************************************************************************************
Function GetRARANotesByOrderID(passedOrderID)

	Set cnnGetRARANotesByOrderID  = Server.CreateObject("ADODB.Connection")
	cnnGetRARANotesByOrderID.open Session("ClientCnnString")

	resultGetRARANotesByOrderID = ""
		
	SQLGetRARANotesByOrderID  = "SELECT RA_Notes FROM API_OR_RAHeader WHERE OrderID = '" & passedOrderID & "'"
	 
	Set rsGetRARANotesByOrderID  = Server.CreateObject("ADODB.Recordset")
	rsGetRARANotesByOrderID.CursorLocation = 3 
	
	rsGetRARANotesByOrderID.Open SQLGetRARANotesByOrderID,cnnGetRARANotesByOrderID 
			
	resultGetRARANotesByOrderID = rsGetRARANotesByOrderID("RA_Notes")
	
	rsGetRARANotesByOrderID.Close
	set rsGetRARANotesByOrderID = Nothing
	cnnGetRARANotesByOrderID.Close	
	set cnnGetRARANotesByOrderID = Nothing
	
	GetRARANotesByOrderID = resultGetRARANotesByOrderID

End Function

'*******************************************************************************************************************************
'*******************************************************************************************************************************


'*******************************************************************************************************************************
'*******************************************************************************************************************************

Function GetAPIRepostSumInvURL()

	resultGetAPIRepostSumInvURL = ""

	Set cnnGetAPIRepostSumInvURL = Server.CreateObject("ADODB.Connection")
	cnnGetAPIRepostSumInvURL.open Session("ClientCnnString")
		
	SQLGetAPIRepostSumInvURL = "SELECT * FROM Settings_Global"
 
	Set rsGetAPIRepostSumInvURL = Server.CreateObject("ADODB.Recordset")
	rsGetAPIRepostSumInvURL.CursorLocation = 3 
	Set rsGetAPIRepostSumInvURL = cnnGetAPIRepostSumInvURL.Execute(SQLGetAPIRepostSumInvURL)
			 
	If not rsGetAPIRepostSumInvURL.EOF Then  
		If Not IsNull(rsGetAPIRepostSumInvURL("SumInvAPIRepostURL")) Then resultGetAPIRepostSumInvURL = rsGetAPIRepostSumInvURL("SumInvAPIRepostURL")
	End If
	
	set rsGetAPIRepostSumInvURL= Nothing
	cnnGetAPIRepostSumInvURL.Close	
	set cnnGetAPIRepostSumInvURL= Nothing
	
	GetAPIRepostSumInvURL = resultGetAPIRepostSumInvURL
	
End Function
'*******************************************************************************************************************************
'*******************************************************************************************************************************

Function GetAPIRepostRAURL()

	resultGetAPIRepostRAURL = ""

	Set cnnGetAPIRepostRAURL = Server.CreateObject("ADODB.Connection")
	cnnGetAPIRepostRAURL.open Session("ClientCnnString")
		
	SQLGetAPIRepostRAURL = "SELECT * FROM Settings_Global"
 
	Set rsGetAPIRepostRAURL = Server.CreateObject("ADODB.Recordset")
	rsGetAPIRepostRAURL.CursorLocation = 3 
	Set rsGetAPIRepostRAURL = cnnGetAPIRepostRAURL.Execute(SQLGetAPIRepostRAURL)
			 
	If not rsGetAPIRepostRAURL.EOF Then  
		If Not IsNull(rsGetAPIRepostRAURL("RAAPIRepostURL")) Then resultGetAPIRepostRAURL = rsGetAPIRepostRAURL("RAAPIRepostURL")
	End If
	
	set rsGetAPIRepostRAURL= Nothing
	cnnGetAPIRepostRAURL.Close	
	set cnnGetAPIRepostRAURL= Nothing
	
	GetAPIRepostRAURL = resultGetAPIRepostRAURL
	
End Function



%>