<%
'***************************************************
'List of all the functions & subs
'***************************************************
'Func GetLastPurchaseDateByCustByItem(passedCust,passedSKU)
'Func GetMTDPurchaseQtyByCustByItem(passedCust,passedSKU)
'Func GetYTDPurchaseQtyByCustByItem(passedCust,passedSKU)
'Func GetCustChainNum(passedCustID)
'Func VPCGetFirstRangeStart(passedReportNumber)
'Func VPCGetFirstRangeEnd(passedReportNumber)
'Func VPCGetSecondRangeStart(passedReportNumber)
'Func VPCGetSecondRangeEnd(passedReportNumber)
'Func VPCGetDatesOrPeriods(passedReportNumber)
'Func VPCGetSecondPeriodSeqNO(passedReportNumber)
'Func VPCGetFirstPeriodSeqNO(passedReportNumber)
'Func VPCReportTableSet(passedReportNumber)
'Func SalesByDayGetRangeStart(passedReportNumber)
'Func SalesByDayGetRangeEnd(passedReportNumber)
'Func SalesByDayGetDatesOrPeriods(passedReportNumber)
'Func SalesByDayGetPeriodSeqNO(passedReportNumber)
'Func SalesByDayReportTableSet(passedReportNumber)
'Func CustHasCategoryAnalNotes(passedCustID,PassedCategory)
'Func NoteNewCatAnalForUser(passedCustNum,passedEntryDateTime,passedCategory)
'Sub MARKNoteNewForUserCatAnal(passedCustNum,passedCategory)
'Func GetUnposedSalesByCustByProd(passedCustID,passedprodSKU)
'Func GetUnposedCasesByCustByProd(passedCustID,passedprodSKU)
'Func GetCurrent_UnpostedTotal_ByCustByCat(passedCustID,passedPeriodSeqBeingEvaluated,passedCategory)
'Func GetCurrent_PostedTotal_ByCustByCat(passedCustID,passedPeriodSeqBeingEvaluated,passedCategory)
'Func CustHasMCSNotes(passedCust)
'Func CustHasMESNotes(passedCust)
'Func NoteNewMCSForUser(passedCust)
'Sub MARKNoteNewForUserMCS(passedCustNum)
'Func GetMostRecentMCSNote(passedCust)
'Func GetMostRecentMESNote(passedCust)
'Func GetMostRecentMCSNoteUserNo(passedCust)
'Func GetMostRecentMESNoteUserNo(passedCust)
'Func NumberOfMCSActionsWithMCSReason(passedMCSReasonIntRecID)
'Func GetMCSReasonByReasonNum(passedMCSReasonIntRecID)
'Func TotalSalesByCustByMonthByYear_RentalsOnly(passedCustID,passedMonth,passedYear)
'Func TotalPostedLVFByCustByMonthByYear(passedCustID,passedMonth,passedYear)
'Func TotalCostByCustByMonthByYear_NoRent(passedCustID,passedMonth,passedYear)
'Func TotalSalesByCustByMonthByYear_NoRentals(passedCustID,passedMonth,passedYear)
'Func TotalSalesByCustByMonthByYear_RentalsOnly(passedCustID,passedMonth,passedYear)
'Func GetCurrent_UnpostedTotal_ByCust(passedCustID,passedPeriodBeingEvaluated)
'Func GetCurrent_PostedTotal_ByCust(passedCustID,passedPeriodBeingEvaluated)
'Func TotalCat21ByCustByMonthByYear(passedCustID,passedMonth,passedYear)
'Func GetMESNotesStatus(passedCustID, MESMonth)
'Func GetMESNotesNoActionStatus(passedCustID, MESMonth)
'Func GetMCSNotesStatus(passedCustID, MCSMonth)
'Func PendingLVFByCust(passedCustID)
'Func GetLastMCSActionNoteByMonthByYearByCust(passedCust, passedMonth, passedYear)
'Func GetLastMCSActionNoteReasonByMonthByYearByCust(passedCust, passedMonth, passedYear)
'Func GetLastMCSActionByMonthByYearByCust(passedCust, passedMonth, passedYear)
'Func GetLastMCSActionDateByMonthByYearByCust(passedCust, passedMonth, passedYear)
'Func TotalUNPostedLVFByCustByMonthByYear(passedCustID,passedMonth,passedYear)
'Func GetLastWebOrderDateFromOCSAccess(passedCustID)
'Func CustHasWebUserID(passedCustID)
'Func GetCurrentPeriod_PostedTotal()
'Func GetCurrentPeriod_UnPostedTotal()	
'Func GetCurrentPeriod_PostedTotalSls2()
'Func GetCurrentPeriod_UnPostedTotalSls2()	
'Func GetCurrentPeriod_PostedRentalsSls2(passedSecondarySalesman)
'Func GetCurrentPeriod_UnPostedRentalsSls2(passedSecondarySalesman)
'Func GetCurrentPeriod_PostedRentals()
'Func GetCurrentPeriod_UnPostedRentals()
'Func GetCurrentPeriod_PostedTotalSls()
'Func GetCurrentPeriod_UnPostedTotalSls(passedSalesman)
'Func TotalCostByPeriodSeqPrior12P(passedPeriodSeq,passedCustID)
'Func TotalCostByPeriodSeqPrior3P(passedPeriodSeq,passedCustID)
'Func TotalTPLYAllCats(passedPeriodSeq,passedCustID)
'Func TotalCostByPeriodSeq(passedPeriodSeq,passedCustID)
'Func GetCurrentPeriod_PostedProdSalesByCust(passedCustomerID)
'Func GetCurrentPeriod_PostedRentalsByCust(passedCustomerID)
'Func GetCurrentPeriod_UnPostedProdSalesByCust(passedCustomerID)
'Func GetCurrentPeriod_UnPostedRentalsByCust(passedCustomerID)
'Func GetCurrentPeriod_PostedTotalCustTyp(passedCustType)
'Func GetCurrentPeriod_UnPostedTotalCustType(passedCustType)
'Func GetCurrentPeriod_PostedRentalsCustType(passedCustType)
'Func GetCurrentPeriod_UnPostedRentalsCustType(passedCustType)
'Func GetCurrentPeriod_PostedRentalsCustType(passedCustType)
'Func GetCurrentPeriod_UnPostedRentalsCustType(passedCustType)
'Func GetCurrentPeriod_PostedTotalCustType(passedCustType)
'Func GetCurrentPeriod_UnPostedTotalCustType(passedCustType)
'Func GetCurrentPeriod_PostedTotalreferralDesc2(passedReferralDesc2)
'Func GetCurrentPeriod_UnPostedTotalReferralDesc2(passedReferralDesc2)
'Func GetCurrentPeriod_PostedRentalsReferralDesc2(passedReferralDesc2)
'Func GetCurrentPeriod_UnPostedRentalsReferralDesc2(passedReferralDesc2)
'***************************************************
'End List of all the functions & subs
'***************************************************

Function GetLastPurchaseDateByCustByItem(passedCust,passedSKU)

	Set cnnGetLastPurchaseDateByCustByItem = Server.CreateObject("ADODB.Connection")
	cnnGetLastPurchaseDateByCustByItem.open Session("ClientCnnString")

	resultGetLastPurchaseDateByCustByItem = ""
		
	SQLGetLastPurchaseDateByCustByItem = "SELECT TOP 1 * FROM InvoiceHistoryDetail "
	SQLGetLastPurchaseDateByCustByItem = SQLGetLastPurchaseDateByCustByItem & "WHERE CustNum = '" & passedCust & "' AND partnum = '" & passedSKU &  "' "
	SQLGetLastPurchaseDateByCustByItem = SQLGetLastPurchaseDateByCustByItem & "ORDER BY ivsDate DESC"
	 
	Set rsGetLastPurchaseDateByCustByItem = Server.CreateObject("ADODB.Recordset")
	rsGetLastPurchaseDateByCustByItem.CursorLocation = 3 
	Set rsGetLastPurchaseDateByCustByItem= cnnGetLastPurchaseDateByCustByItem.Execute(SQLGetLastPurchaseDateByCustByItem)
	
	If not rsGetLastPurchaseDateByCustByItem.eof then resultGetLastPurchaseDateByCustByItem =  rsGetLastPurchaseDateByCustByItem("IVSDATE")
		
	rsGetLastPurchaseDateByCustByItem.Close
	set rsGetLastPurchaseDateByCustByItem= Nothing
	cnnGetLastPurchaseDateByCustByItem.Close	
	set cnnGetLastPurchaseDateByCustByItem = Nothing
	
	GetLastPurchaseDateByCustByItem = resultGetLastPurchaseDateByCustByItem 
	
End Function

Function GetMTDPurchaseQtyByCustByItem(passedCust,passedSKU)

	Set cnnGetMTDPurchaseQtyByCustByItem = Server.CreateObject("ADODB.Connection")
	cnnGetMTDPurchaseQtyByCustByItem.open Session("ClientCnnString")

	resultGetMTDPurchaseQtyByCustByItem = 0
		
	SQLGetMTDPurchaseQtyByCustByItem = "SELECT SUM (itemQuantity) as Expr1 FROM InvoiceHistoryDetail "
	SQLGetMTDPurchaseQtyByCustByItem = SQLGetMTDPurchaseQtyByCustByItem & "WHERE CustNum = '" & passedCust & "' AND partnum = '" & passedSKU &  "' "
	SQLGetMTDPurchaseQtyByCustByItem = SQLGetMTDPurchaseQtyByCustByItem & "AND Year(ivsDate) = Year(getdate()) AND Month(ivsDate) = Month(getdate())"
	 
	Set rsGetMTDPurchaseQtyByCustByItem = Server.CreateObject("ADODB.Recordset")
	rsGetMTDPurchaseQtyByCustByItem.CursorLocation = 3 
	Set rsGetMTDPurchaseQtyByCustByItem= cnnGetMTDPurchaseQtyByCustByItem.Execute(SQLGetMTDPurchaseQtyByCustByItem)
	
	If not rsGetMTDPurchaseQtyByCustByItem.eof then 
		If Not IsNull(rsGetMTDPurchaseQtyByCustByItem("Expr1")) Then resultGetMTDPurchaseQtyByCustByItem =  rsGetMTDPurchaseQtyByCustByItem("Expr1")
	End If	
	rsGetMTDPurchaseQtyByCustByItem.Close
	set rsGetMTDPurchaseQtyByCustByItem= Nothing
	cnnGetMTDPurchaseQtyByCustByItem.Close	
	set cnnGetMTDPurchaseQtyByCustByItem = Nothing
	
	GetMTDPurchaseQtyByCustByItem = resultGetMTDPurchaseQtyByCustByItem 
	
End Function

Function GetYTDPurchaseQtyByCustByItem(passedCust,passedSKU)

	Set cnnGetYTDPurchaseQtyByCustByItem = Server.CreateObject("ADODB.Connection")
	cnnGetYTDPurchaseQtyByCustByItem.open Session("ClientCnnString")

	resultGetYTDPurchaseQtyByCustByItem = 0
		
	SQLGetYTDPurchaseQtyByCustByItem = "SELECT SUM (itemQuantity) as Expr1 FROM InvoiceHistoryDetail "
	SQLGetYTDPurchaseQtyByCustByItem = SQLGetYTDPurchaseQtyByCustByItem & "WHERE CustNum = '" & passedCust & "' AND partnum = '" & passedSKU &  "' "
	SQLGetYTDPurchaseQtyByCustByItem = SQLGetYTDPurchaseQtyByCustByItem & "AND Year(ivsDate) = Year(getdate())"
	 
	Set rsGetYTDPurchaseQtyByCustByItem = Server.CreateObject("ADODB.Recordset")
	rsGetYTDPurchaseQtyByCustByItem.CursorLocation = 3 
	Set rsGetYTDPurchaseQtyByCustByItem= cnnGetYTDPurchaseQtyByCustByItem.Execute(SQLGetYTDPurchaseQtyByCustByItem)
	
	If not rsGetYTDPurchaseQtyByCustByItem.eof then
		If Not IsNull(rsGetYTDPurchaseQtyByCustByItem("Expr1")) Then resultGetYTDPurchaseQtyByCustByItem =  rsGetYTDPurchaseQtyByCustByItem("Expr1")
	End IF
	
	rsGetYTDPurchaseQtyByCustByItem.Close
	set rsGetYTDPurchaseQtyByCustByItem= Nothing
	cnnGetYTDPurchaseQtyByCustByItem.Close	
	set cnnGetYTDPurchaseQtyByCustByItem = Nothing
	
	GetYTDPurchaseQtyByCustByItem = resultGetYTDPurchaseQtyByCustByItem 
	
End Function


Function GetCustChainNum(passedCustID)

    resultGetCustChainNum = 0

    Set cnnGetCustChainNum = Server.CreateObject("ADODB.Connection")
    cnnGetCustChainNum.open Session("ClientCnnString")
                                    
    SQLGetCustChainNum = "Select * from AR_Customer where CustNum = '" & passedCustID & "'"
    
    Set rsGetCustChainNum = Server.CreateObject("ADODB.Recordset")
    rsGetCustChainNum.CursorLocation = 3 
    Set rsGetCustChainNum= cnnGetCustChainNum.Execute(SQLGetCustChainNum)
                    
    If not rsGetCustChainNum.eof then resultGetCustChainNum = rsGetCustChainNum("ChainNum")
    
    set rsGetCustChainNum= Nothing
    cnnGetCustChainNum.Close      
    set cnnGetCustChainNum= Nothing
    
    GetCustChainNum = resultGetCustChainNum
                
End Function

Function VPCGetFirstRangeStart(passedReportNumber)

	resultVPCGetFirstRangeStart = ""
	
	SQLVPCGetFirstRangeStart = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")
	
	Set cnnVPCGetFirstRangeStart = Server.CreateObject("ADODB.Connection")
	cnnVPCGetFirstRangeStart.open (Session("ClientCnnString"))
	Set rsVPCGetFirstRangeStart = Server.CreateObject("ADODB.Recordset")
	Set rsVPCGetFirstRangeStart = cnnVPCGetFirstRangeStart.Execute(SQLVPCGetFirstRangeStart)

	If Not rsVPCGetFirstRangeStart.Eof Then 
	
		If Not IsNUll(rsVPCGetFirstRangeStart("ReportSpecificData9")) Then
			If rsVPCGetFirstRangeStart("ReportSpecificData9") <> "" Then resultVPCGetFirstRangeStart = rsVPCGetFirstRangeStart("ReportSpecificData9") 
		End If
	
	End IF
	
	Set rsVPCGetFirstRangeStart = Nothing
	cnnVPCGetFirstRangeStart.Close
	Set cnnVPCGetFirstRangeStart = Nothing

	VPCGetFirstRangeStart = resultVPCGetFirstRangeStart 
	
End Function

Function VPCGetFirstRangeEnd(passedReportNumber)

	resultVPCGetFirstRangeEnd = ""
	
	SQLVPCGetFirstRangeEnd = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")
	
	Set cnnVPCGetFirstRangeEnd = Server.CreateObject("ADODB.Connection")
	cnnVPCGetFirstRangeEnd.open (Session("ClientCnnString"))
	Set rsVPCGetFirstRangeEnd = Server.CreateObject("ADODB.Recordset")
	Set rsVPCGetFirstRangeEnd = cnnVPCGetFirstRangeEnd.Execute(SQLVPCGetFirstRangeEnd)

	If Not rsVPCGetFirstRangeEnd.Eof Then 
	
		If Not IsNUll(rsVPCGetFirstRangeEnd("ReportSpecificData10")) Then
			If rsVPCGetFirstRangeEnd("ReportSpecificData10") <> "" Then resultVPCGetFirstRangeEnd = rsVPCGetFirstRangeEnd("ReportSpecificData10") 
		End If
	
	End IF
	
	Set rsVPCGetFirstRangeEnd = Nothing
	cnnVPCGetFirstRangeEnd.Close
	Set cnnVPCGetFirstRangeEnd = Nothing

	VPCGetFirstRangeEnd = resultVPCGetFirstRangeEnd 
	
End Function

Function VPCGetSecondRangeStart(passedReportNumber)

	resultVPCGetSecondRangeStart = ""
	
	SQLVPCGetSecondRangeStart = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")
	
	Set cnnVPCGetSecondRangeStart = Server.CreateObject("ADODB.Connection")
	cnnVPCGetSecondRangeStart.open (Session("ClientCnnString"))
	Set rsVPCGetSecondRangeStart = Server.CreateObject("ADODB.Recordset")
	Set rsVPCGetSecondRangeStart = cnnVPCGetSecondRangeStart.Execute(SQLVPCGetSecondRangeStart)

	If Not rsVPCGetSecondRangeStart.Eof Then 
	
		If Not IsNUll(rsVPCGetSecondRangeStart("ReportSpecificData11")) Then
			If rsVPCGetSecondRangeStart("ReportSpecificData11") <> "" Then resultVPCGetSecondRangeStart = rsVPCGetSecondRangeStart("ReportSpecificData11") 
		End If
	
	End IF
	
	Set rsVPCGetSecondRangeStart = Nothing
	cnnVPCGetSecondRangeStart.Close
	Set cnnVPCGetSecondRangeStart = Nothing

	VPCGetSecondRangeStart = resultVPCGetSecondRangeStart 
	
End Function

Function VPCGetSecondRangeEnd(passedReportNumber)

	resultVPCGetSecondRangeEnd = ""
	
	SQLVPCGetSecondRangeEnd = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")
	
	Set cnnVPCGetSecondRangeEnd = Server.CreateObject("ADODB.Connection")
	cnnVPCGetSecondRangeEnd.open (Session("ClientCnnString"))
	Set rsVPCGetSecondRangeEnd = Server.CreateObject("ADODB.Recordset")
	Set rsVPCGetSecondRangeEnd = cnnVPCGetSecondRangeEnd.Execute(SQLVPCGetSecondRangeEnd)

	If Not rsVPCGetSecondRangeEnd.Eof Then 
	
		If Not IsNUll(rsVPCGetSecondRangeEnd("ReportSpecificData12")) Then
			If rsVPCGetSecondRangeEnd("ReportSpecificData12") <> "" Then resultVPCGetSecondRangeEnd = rsVPCGetSecondRangeEnd("ReportSpecificData12") 
		End If
	
	End IF
	
	Set rsVPCGetSecondRangeEnd = Nothing
	cnnVPCGetSecondRangeEnd.Close
	Set cnnVPCGetSecondRangeEnd = Nothing

	VPCGetSecondRangeEnd = resultVPCGetSecondRangeEnd 
	
End Function

Function VPCGetDatesOrPeriods(passedReportNumber)

	resultVPCGetDatesOrPeriods = ""
	
	SQLVPCGetDatesOrPeriods = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")
	
	Set cnnVPCGetDatesOrPeriods = Server.CreateObject("ADODB.Connection")
	cnnVPCGetDatesOrPeriods.open (Session("ClientCnnString"))
	Set rsVPCGetDatesOrPeriods = Server.CreateObject("ADODB.Recordset")
	Set rsVPCGetDatesOrPeriods = cnnVPCGetDatesOrPeriods.Execute(SQLVPCGetDatesOrPeriods)

	If Not rsVPCGetDatesOrPeriods.Eof Then 
	
		If Not IsNUll(rsVPCGetDatesOrPeriods("ReportSpecificData13")) Then
			If rsVPCGetDatesOrPeriods("ReportSpecificData13") <> "" Then resultVPCGetDatesOrPeriods = rsVPCGetDatesOrPeriods("ReportSpecificData13") 
		End If
	
	End IF
	
	Set rsVPCGetDatesOrPeriods = Nothing
	cnnVPCGetDatesOrPeriods.Close
	Set cnnVPCGetDatesOrPeriods = Nothing

	VPCGetDatesOrPeriods = resultVPCGetDatesOrPeriods 
	
End Function

Function VPCGetSecondPeriodSeqNO(passedReportNumber)

	resultVPCGetSecondPeriodSeqNO = ""
	
	SQLVPCGetSecondPeriodSeqNO = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")
	
	Set cnnVPCGetSecondPeriodSeqNO = Server.CreateObject("ADODB.Connection")
	cnnVPCGetSecondPeriodSeqNO.open (Session("ClientCnnString"))
	Set rsVPCGetSecondPeriodSeqNO = Server.CreateObject("ADODB.Recordset")
	Set rsVPCGetSecondPeriodSeqNO = cnnVPCGetSecondPeriodSeqNO.Execute(SQLVPCGetSecondPeriodSeqNO)

	If Not rsVPCGetSecondPeriodSeqNO.Eof Then 
	
		If Not IsNUll(rsVPCGetSecondPeriodSeqNO("ReportSpecificData8")) Then
			If rsVPCGetSecondPeriodSeqNO("ReportSpecificData8") <> "" Then resultVPCGetSecondPeriodSeqNO = rsVPCGetSecondPeriodSeqNO("ReportSpecificData8") 
		End If
	
	End IF
	
	Set rsVPCGetSecondPeriodSeqNO = Nothing
	cnnVPCGetSecondPeriodSeqNO.Close
	Set cnnVPCGetSecondPeriodSeqNO = Nothing

	VPCGetSecondPeriodSeqNO = resultVPCGetSecondPeriodSeqNO 
	
End Function

Function VPCGetFirstPeriodSeqNO(passedReportNumber)

	resultVPCGetFirstPeriodSeqNO = ""
	
	SQLVPCGetFirstPeriodSeqNO = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")
	
	Set cnnVPCGetFirstPeriodSeqNO = Server.CreateObject("ADODB.Connection")
	cnnVPCGetFirstPeriodSeqNO.open (Session("ClientCnnString"))
	Set rsVPCGetFirstPeriodSeqNO = Server.CreateObject("ADODB.Recordset")
	Set rsVPCGetFirstPeriodSeqNO = cnnVPCGetFirstPeriodSeqNO.Execute(SQLVPCGetFirstPeriodSeqNO)

	If Not rsVPCGetFirstPeriodSeqNO.Eof Then 
	
		If Not IsNUll(rsVPCGetFirstPeriodSeqNO("ReportSpecificData7")) Then
			If rsVPCGetFirstPeriodSeqNO("ReportSpecificData7") <> "" Then resultVPCGetFirstPeriodSeqNO = rsVPCGetFirstPeriodSeqNO("ReportSpecificData7") 
		End If
	
	End IF
	
	Set rsVPCGetFirstPeriodSeqNO = Nothing
	cnnVPCGetFirstPeriodSeqNO.Close
	Set cnnVPCGetFirstPeriodSeqNO = Nothing

	VPCGetFirstPeriodSeqNO = resultVPCGetFirstPeriodSeqNO 
	
End Function


Function VPCReportTableSet(passedReportNumber)

	resultVPCReportTableSet = False

	SQLVPCReportTableSet = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")

	Set cnnVPCReportTableSet = Server.CreateObject("ADODB.Connection")
	cnnVPCReportTableSet.open (Session("ClientCnnString"))
	Set rsVPCReportTableSet = Server.CreateObject("ADODB.Recordset")
	Set rsVPCReportTableSet = cnnVPCReportTableSet.Execute(SQLVPCReportTableSet)

	If NOT rsVPCReportTableSet.EOF Then resultVPCReportTableSet = True


	Set rsVPCReportTableSet = Nothing
	cnnVPCReportTableSet.Close
	Set cnnVPCReportTableSet = Nothing

	VPCReportTableSet = resultVPCReportTableSet 

End Function


Function SalesByDayGetRangeStart(passedReportNumber)

	resultSalesByDayGetRangeStart = ""
	
	SQLSalesByDayGetRangeStart = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")
	
	Set cnnSalesByDayGetRangeStart = Server.CreateObject("ADODB.Connection")
	cnnSalesByDayGetRangeStart.open (Session("ClientCnnString"))
	Set rsSalesByDayGetRangeStart = Server.CreateObject("ADODB.Recordset")
	Set rsSalesByDayGetRangeStart = cnnSalesByDayGetRangeStart.Execute(SQLSalesByDayGetRangeStart)

	If Not rsSalesByDayGetRangeStart.Eof Then 
	
		If Not IsNUll(rsSalesByDayGetRangeStart("ReportSpecificData3")) Then
			If rsSalesByDayGetRangeStart("ReportSpecificData3") <> "" Then resultSalesByDayGetRangeStart = rsSalesByDayGetRangeStart("ReportSpecificData3") 
		End If
	
	End IF
	
	Set rsSalesByDayGetRangeStart = Nothing
	cnnSalesByDayGetRangeStart.Close
	Set cnnSalesByDayGetRangeStart = Nothing

	SalesByDayGetRangeStart = resultSalesByDayGetRangeStart 
	
End Function



Function SalesByDayGetRangeEnd(passedReportNumber)

	resultSalesByDayGetRangeEnd = ""
	
	SQLSalesByDayGetRangeEnd = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")
	
	Set cnnSalesByDayGetRangeEnd = Server.CreateObject("ADODB.Connection")
	cnnSalesByDayGetRangeEnd.open (Session("ClientCnnString"))
	Set rsSalesByDayGetRangeEnd = Server.CreateObject("ADODB.Recordset")
	Set rsSalesByDayGetRangeEnd = cnnSalesByDayGetRangeEnd.Execute(SQLSalesByDayGetRangeEnd)

	If Not rsSalesByDayGetRangeEnd.Eof Then 
	
		If Not IsNUll(rsSalesByDayGetRangeEnd("ReportSpecificData4")) Then
			If rsSalesByDayGetRangeEnd("ReportSpecificData4") <> "" Then resultSalesByDayGetRangeEnd = rsSalesByDayGetRangeEnd("ReportSpecificData4") 
		End If
	
	End IF
	
	Set rsSalesByDayGetRangeEnd = Nothing
	cnnSalesByDayGetRangeEnd.Close
	Set cnnSalesByDayGetRangeEnd = Nothing

	SalesByDayGetRangeEnd = resultSalesByDayGetRangeEnd 
	
End Function


Function SalesByDayGetDatesOrPeriods(passedReportNumber)

	resultSalesByDayGetDatesOrPeriods = ""
	
	SQLSalesByDayGetDatesOrPeriods = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")
	
	Set cnnSalesByDayGetDatesOrPeriods = Server.CreateObject("ADODB.Connection")
	cnnSalesByDayGetDatesOrPeriods.open (Session("ClientCnnString"))
	Set rsSalesByDayGetDatesOrPeriods = Server.CreateObject("ADODB.Recordset")
	Set rsSalesByDayGetDatesOrPeriods = cnnSalesByDayGetDatesOrPeriods.Execute(SQLSalesByDayGetDatesOrPeriods)

	If Not rsSalesByDayGetDatesOrPeriods.Eof Then 
	
		If Not IsNUll(rsSalesByDayGetDatesOrPeriods("ReportSpecificData1")) Then
			If rsSalesByDayGetDatesOrPeriods("ReportSpecificData1") <> "" Then resultSalesByDayGetDatesOrPeriods = rsSalesByDayGetDatesOrPeriods("ReportSpecificData1") 
		End If
	
	End IF
	
	Set rsSalesByDayGetDatesOrPeriods = Nothing
	cnnSalesByDayGetDatesOrPeriods.Close
	Set cnnSalesByDayGetDatesOrPeriods = Nothing

	SalesByDayGetDatesOrPeriods = resultSalesByDayGetDatesOrPeriods 
	
End Function



Function SalesByDayGetPeriodSeqNO(passedReportNumber)

	resultSalesByDayGetPeriodSeqNO = ""
	
	SQLSalesByDayGetPeriodSeqNO = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")
	
	Set cnnSalesByDayGetPeriodSeqNO = Server.CreateObject("ADODB.Connection")
	cnnSalesByDayGetPeriodSeqNO.open (Session("ClientCnnString"))
	Set rsSalesByDayGetPeriodSeqNO = Server.CreateObject("ADODB.Recordset")
	Set rsSalesByDayGetPeriodSeqNO = cnnSalesByDayGetPeriodSeqNO.Execute(SQLSalesByDayGetPeriodSeqNO)

	If Not rsSalesByDayGetPeriodSeqNO.Eof Then 
	
		If Not IsNUll(rsSalesByDayGetPeriodSeqNO("ReportSpecificData2")) Then
			If rsSalesByDayGetPeriodSeqNO("ReportSpecificData2") <> "" Then resultSalesByDayGetPeriodSeqNO = rsSalesByDayGetPeriodSeqNO("ReportSpecificData2") 
		End If
	
	End IF
	
	Set rsSalesByDayGetPeriodSeqNO = Nothing
	cnnSalesByDayGetPeriodSeqNO.Close
	Set cnnSalesByDayGetPeriodSeqNO = Nothing

	SalesByDayGetPeriodSeqNO = resultSalesByDayGetPeriodSeqNO 
	
End Function



Function SalesByDayReportTableSet(passedReportNumber)

	resultSalesByDayReportTableSet = False

	SQLSalesByDayReportTableSet = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")

	Set cnnSalesByDayReportTableSet = Server.CreateObject("ADODB.Connection")
	cnnSalesByDayReportTableSet.open (Session("ClientCnnString"))
	Set rsSalesByDayReportTableSet = Server.CreateObject("ADODB.Recordset")
	Set rsSalesByDayReportTableSet = cnnSalesByDayReportTableSet.Execute(SQLSalesByDayReportTableSet)

	If NOT rsSalesByDayReportTableSet.EOF Then resultSalesByDayReportTableSet = True


	Set rsSalesByDayReportTableSet = Nothing
	cnnSalesByDayReportTableSet.Close
	Set cnnSalesByDayReportTableSet = Nothing

	SalesByDayReportTableSet = resultSalesByDayReportTableSet 

End Function

Function CustHasCategoryAnalNotes(passedCust,passedCategory)

	Set cnnCustHasCategoryAnalNotes = Server.CreateObject("ADODB.Connection")
	cnnCustHasCategoryAnalNotes.open Session("ClientCnnString")

	resultCustHasCategoryAnalNotes = False
		
	SQLCustHasCategoryAnalNotes = "SELECT TOP 1 * FROM AR_CustomerNotes "
	SQLCustHasCategoryAnalNotes = SQLCustHasCategoryAnalNotes & "WHERE CustID = '" & passedCust & "' AND Category = " & passedCategory
	 
	Set rsCustHasCategoryAnalNotes = Server.CreateObject("ADODB.Recordset")
	rsCustHasCategoryAnalNotes.CursorLocation = 3 
	Set rsCustHasCategoryAnalNotes= cnnCustHasCategoryAnalNotes.Execute(SQLCustHasCategoryAnalNotes)
	
	If not rsCustHasCategoryAnalNotes.eof then resultCustHasCategoryAnalNotes =  True
		
	rsCustHasCategoryAnalNotes.Close
	set rsCustHasCategoryAnalNotes= Nothing
	cnnCustHasCategoryAnalNotes.Close	
	set cnnCustHasCategoryAnalNotes = Nothing
	
	CustHasCategoryAnalNotes = resultCustHasCategoryAnalNotes 
	
End Function

Function NoteNewCatAnalForUser(passedCustNum,passedCategory)

	resultNoteNewCatAnalForUser = False
	
	SQLNoteNewCatAnalForUser = "SELECT * FROM AR_CustomerNotesUserViewed Where CustID ='" & passedCustNum & "' AND UserNo = " & Session("Userno") & " AND Category = " & passedCategory
	
	Set cnnNoteNewCatAnalForUser = Server.CreateObject("ADODB.Connection")
	cnnNoteNewCatAnalForUser.open (Session("ClientCnnString"))
	Set rsNoteNewCatAnalForUser = Server.CreateObject("ADODB.Recordset")
	rsNoteNewCatAnalForUser.CursorLocation = 3 
	Set rsNoteNewCatAnalForUser = cnnNoteNewCatAnalForUser.Execute(SQLNoteNewCatAnalForUser)

	Set rsNoteCatAnal = Server.CreateObject("ADODB.Recordset")
	rsNoteCatAnal.CursorLocation = 3 

	If not rsNoteNewCatAnalForUser.EOF Then
		'OK, so see when the last note was created, not by us
		SQLCustHasCategoryAnalNotes = "SELECT TOP 1 RecordCreationDateTime FROM AR_CustomerNotes "
		SQLCustHasCategoryAnalNotes = SQLCustHasCategoryAnalNotes & "WHERE CustID = '" & passedCust & "' AND Category = " & passedCategory
		SQLCustHasCategoryAnalNotes = SQLCustHasCategoryAnalNotes & " ORDER BY RecordCreationDateTime DESC"
		
		Set rsNoteCatAnal = cnnNoteNewCatAnalForUser.Execute(SQLCustHasCategoryAnalNotes)
		If Not rsNoteCatAnal.Eof Then
			If rsNoteNewCatAnalForUser("DateLastViewed") < rsNoteCatAnal("RecordCreationDateTime")  Then resultNoteNewCatAnalForUser = True
		End If
	Else
		resultNoteNewCatAnalForUser = True 'Also true if they have never seen any of them
	End If
	cnnNoteNewCatAnalForUser.close
	set rsNoteNewCatAnalForUser = nothing
	set rsNoteCatAnal = nothing
	set cnnNoteNewCatAnalForUser= nothing	

	NoteNewCatAnalForUser = resultNoteNewCatAnalForUser

End Function

Sub MARKNoteNewForUserCatAnal(passedCustNum,passedCategory)

	SQLMARKNoteNewForUserCatAnal = "SELECT * FROM AR_CustomerNotesUserViewed Where CustID ='" & passedCustNum & "' AND UserNo = " & Session("Userno") & " AND Category = " & passedCategory
	
	Set cnnMARKNoteNewForUserCatAnal = Server.CreateObject("ADODB.Connection")
	cnnMARKNoteNewForUserCatAnal.open (Session("ClientCnnString"))
	Set rMARKNoteNewForUserCatAnal = Server.CreateObject("ADODB.Recordset")
	rMARKNoteNewForUserCatAnal.CursorLocation = 3 
	Set rMARKNoteNewForUserCatAnal = cnnMARKNoteNewForUserCatAnal.Execute(SQLMARKNoteNewForUserCatAnal)

	If rMARKNoteNewForUserCatAnal.EOF Then ' Nothing there so we need to insert
		SQLMARKNoteNewForUserCatAnal = "INSERT INTO AR_CustomerNotesUserViewed (CustID ,UserNo, Category) VALUES ('" & passedCustNum & "',"  & Session("UserNo") & "," & passedCategory & ")"
	Else
		SQLMARKNoteNewForUserCatAnal = "UPDATE AR_CustomerNotesUserViewed Set DateLastViewed = getdate() Where CustID ='" & passedCustNum & "' AND UserNo = " & Session("Userno") & " AND Category = " & passedCategory
	End If
	
	Set rMARKNoteNewForUserCatAnal = cnnMARKNoteNewForUserCatAnal.Execute(SQLMARKNoteNewForUserCatAnal)
		
	cnnMARKNoteNewForUserCatAnal.close
	set rMARKNoteNewForUserCatAnal = nothing
	set cnnMARKNoteNewForUserCatAnal= nothing	

End Sub

Function GetUnposedSalesByCustByProd(passedCustNum,passedProdSKU,passedPeriodSeq)

	resultGetUnposedSalesByCustByProd = 0
	
	If Session("CalcTax") = True Then
		SQLGetUnposedSalesByCustByProd = "SELECT SUM(CASE WHEN TaxablePart = 'Y' THEN (Qty * UnitPrice) + ((Qty * UnitPrice) * (TaxPercent / 100)) WHEN TaxablePart <> 'Y' THEN (Qty * UnitPrice) END ) AS TotSales FROM TelSelParts "
	Else		
		SQLGetUnposedSalesByCustByProd = "SELECT SUM((Qty * UnitPrice)) AS TotSales FROM TelSelParts "	
	End If		
	SQLGetUnposedSalesByCustByProd = SQLGetUnposedSalesByCustByProd & " INNER JOIN TelSel ON TelSelParts.InvoiceNo = TelSel.InvoiceNo "
	SQLGetUnposedSalesByCustByProd = SQLGetUnposedSalesByCustByProd & " WHERE TelselParts.CustNum = " & passedCustNum & " AND "
	SQLGetUnposedSalesByCustByProd = SQLGetUnposedSalesByCustByProd & " PartNumber = '" & passedProdSKU & "' AND "
	SQLGetUnposedSalesByCustByProd = SQLGetUnposedSalesByCustByProd & "TelselParts.PeriodSeqNo = " & passedPeriodSeq & " AND "
	SQLGetUnposedSalesByCustByProd = SQLGetUnposedSalesByCustByProd & "(InvoiceTFlag = 'O' OR InvoiceTFlag = 'T')"
'Response.Write(SQLGetUnposedSalesByCustByProd & "<BR>")
	Set cnnGetUnposedSalesByCustByProd = Server.CreateObject("ADODB.Connection")
	cnnGetUnposedSalesByCustByProd.open (Session("ClientCnnString"))
	Set rsGetUnposedSalesByCustByProd = Server.CreateObject("ADODB.Recordset")
	rsGetUnposedSalesByCustByProd.CursorLocation = 3 
	Set rsGetUnposedSalesByCustByProd = cnnGetUnposedSalesByCustByProd.Execute(SQLGetUnposedSalesByCustByProd)

	If not rsGetUnposedSalesByCustByProd.EOF Then
		resultGetUnposedSalesByCustByProd = rsGetUnposedSalesByCustByProd("TotSales")
	End If
	
	'Account for nulls returned by query
	If NOT IsNumeric(resultGetUnposedSalesByCustByProd) Then resultGetUnposedSalesByCustByProd = 0
	
	cnnGetUnposedSalesByCustByProd.close
	set rsGetUnposedSalesByCustByProd = nothing
	set cnnGetUnposedSalesByCustByProd= nothing	

	GetUnposedSalesByCustByProd = resultGetUnposedSalesByCustByProd

End Function

Function GetUnposedCasesByCustByProd(passedCustNum,passedProdSKU,passedPeriodSeq)

	resultGetUnposedCasesByCustByProd = 0

	Set cnnGetUnposedCasesByCustByProd = Server.CreateObject("ADODB.Connection")
	cnnGetUnposedCasesByCustByProd.open (Session("ClientCnnString"))
	Set rsGetUnposedCasesByCustByProd = Server.CreateObject("ADODB.Recordset")
	rsGetUnposedCasesByCustByProd.CursorLocation = 3 

	SQLGetUnposedCasesByCustByProd = "SELECT prodCaseConversionFactor from IC_Product WHERE prodSKU = '" & passedProdSKU & "'"

	Set rsGetUnposedCasesByCustByProd = cnnGetUnposedCasesByCustByProd.Execute(SQLGetUnposedCasesByCustByProd)

	If not rsGetUnposedCasesByCustByProd.EOF Then
		CaseConversionFactor = rsGetUnposedCasesByCustByProd("prodCaseConversionFactor")
	Else
		CaseConversionFactor = 1
	End If


	SQLGetUnposedCasesByCustByProd  = "SELECT SUM(CASE WHEN SalesUnit <> 'U' THEN Qty WHEN SalesUnit = 'U' THEN "
	SQLGetUnposedCasesByCustByProd = SQLGetUnposedCasesByCustByProd & " Round(cast(Qty as float) / cast(" & CaseConversionFactor & " as float),2)END ) AS TotCases "
	SQLGetUnposedCasesByCustByProd = SQLGetUnposedCasesByCustByProd & " FROM TelSelParts "
	SQLGetUnposedCasesByCustByProd = SQLGetUnposedCasesByCustByProd & " INNER JOIN TelSel ON TelSelParts.InvoiceNo = TelSel.InvoiceNo "
	SQLGetUnposedCasesByCustByProd = SQLGetUnposedCasesByCustByProd & " WHERE TelselParts.CustNum = " & passedCustNum & " AND "
	SQLGetUnposedCasesByCustByProd = SQLGetUnposedCasesByCustByProd & " PartNumber = '" & passedProdSKU & "' AND "
	SQLGetUnposedCasesByCustByProd = SQLGetUnposedCasesByCustByProd & "TelselParts.PeriodSeqNo = " & passedPeriodSeq & " AND "
	SQLGetUnposedCasesByCustByProd = SQLGetUnposedCasesByCustByProd & "(InvoiceTFlag = 'O' OR InvoiceTFlag = 'T')"

	Set rsGetUnposedCasesByCustByProd = cnnGetUnposedCasesByCustByProd.Execute(SQLGetUnposedCasesByCustByProd)

	If not rsGetUnposedCasesByCustByProd.EOF Then
		resultGetUnposedCasesByCustByProd = rsGetUnposedCasesByCustByProd("TotCases")
	End If
	
	'Account for nulls returned by query
	If NOT IsNumeric(resultGetUnposedCasesByCustByProd) Then resultGetUnposedCasesByCustByProd = 0
	
	cnnGetUnposedCasesByCustByProd.close
	set rsGetUnposedCasesByCustByProd = nothing
	set cnnGetUnposedCasesByCustByProd= nothing	

	GetUnposedCasesByCustByProd = resultGetUnposedCasesByCustByProd

End Function

Function GetCurrent_UnpostedTotal_ByCustByCat(passedCustID,passedPeriodSeqBeingEvaluated,passedCategory)

	resultGetCurrent_UnpostedTotal_ByCustByCat = 0

	Set cnnGetCurrent_UnpostedTotal_ByCustByCat = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnpostedTotal_ByCustByCat.open Session("ClientCnnString")
		
	If Session("CalcTax") = True Then
		SQLGetCurrent_UnpostedTotal_ByCustByCat = "SELECT SUM(CASE WHEN taxablePart = 'Y' THEN ExtendedPrice + (ExtendedPrice * (TaxPercent / 100)) WHEN taxablepart <> 'Y' THEN ExtendedPrice END ) AS TotalForCurrent FROM TelselParts "
	Else
		SQLGetCurrent_UnpostedTotal_ByCustByCat = "SELECT Sum(ExtendedPrice) AS TotalForCurrent FROM TelselParts "	
	End If		
	SQLGetCurrent_UnpostedTotal_ByCustByCat  = SQLGetCurrent_UnpostedTotal_ByCustByCat & " INNER JOIN TelSel ON TelSelParts.InvoiceNo = TelSel.InvoiceNo "
	SQLGetCurrent_UnpostedTotal_ByCustByCat = SQLGetCurrent_UnpostedTotal_ByCustByCat & " WHERE Category = " & passedCategory & " AND TelSelParts.CustNum = " & passedCustID & " AND "
	SQLGetCurrent_UnpostedTotal_ByCustByCat = SQLGetCurrent_UnpostedTotal_ByCustByCat & "PeriodSeqNo = " & passedPeriodSeqBeingEvaluated + 1 & " AND "
	SQLGetCurrent_UnpostedTotal_ByCustByCat = SQLGetCurrent_UnpostedTotal_ByCustByCat & "(InvoiceTFlag = 'O' OR InvoiceTFlag = 'T')"

'Response.Write(SQLGetCurrent_UnpostedTotal_ByCustByCat & "<br>") 

	Set rsGetCurrent_UnpostedTotal_ByCustByCat = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnpostedTotal_ByCustByCat.CursorLocation = 3 
	Set rsGetCurrent_UnpostedTotal_ByCustByCat = cnnGetCurrent_UnpostedTotal_ByCustByCat.Execute(SQLGetCurrent_UnpostedTotal_ByCustByCat)

	If not rsGetCurrent_UnpostedTotal_ByCustByCat.EOF Then resultGetCurrent_UnpostedTotal_ByCustByCat = rsGetCurrent_UnpostedTotal_ByCustByCat("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnpostedTotal_ByCustByCat) Then resultGetCurrent_UnpostedTotal_ByCustByCat = 0 ' In case there are no results
	
	rsGetCurrent_UnpostedTotal_ByCustByCat.Close
	set rsGetCurrent_UnpostedTotal_ByCustByCat= Nothing
	cnnGetCurrent_UnpostedTotal_ByCustByCat.Close	
	set cnnGetCurrent_UnpostedTotal_ByCustByCat= Nothing
	
	GetCurrent_UnpostedTotal_ByCustByCat = resultGetCurrent_UnpostedTotal_ByCustByCat
'Response.Write(resultGetCurrent_UnpostedTotal_ByCustByCat & "<br>")
End Function

Function GetCurrent_PostedTotal_ByCustByCat(passedCustID,passedPeriodSeqBeingEvaluated,passedCategory)

	resultGetCurrent_PostedTotal_ByCustByCat = 0
	periodToFind = GetPeriodBySeq(passedPeriodSeqBeingEvaluated+1)
	periodYearToFind = GetPeriodYearBySeq(passedPeriodSeqBeingEvaluated+1)



	Set cnnGetCurrent_PostedTotal_ByCustByCat = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByCustByCat.open Session("ClientCnnString")

	If Session("CalcTax") = True Then		
		SQLGetCurrent_PostedTotal_ByCustByCat = "SELECT SUM(CASE WHEN prodTaxable = 'Y' THEN (itemPrice*itemQuantity) + ((itemPrice*itemQuantity) * (prodTaxPercent / 100)) WHEN prodTaxable <> 'Y' THEN (itemPrice*itemQuantity) END ) AS TotalForCurrent "
	Else		
		SQLGetCurrent_PostedTotal_ByCustByCat = "SELECT SUM(itemPrice*itemQuantity) AS TotalForCurrent "
	End If		
	SQLGetCurrent_PostedTotal_ByCustByCat = SQLGetCurrent_PostedTotal_ByCustByCat & " FROM InvoiceHistoryDetail WHERE CustNum = " & passedCustID & " AND prodCategory = '" & passedCategory & "' "
	SQLGetCurrent_PostedTotal_ByCustByCat = SQLGetCurrent_PostedTotal_ByCustByCat & " AND Period = " & periodToFind  & " "
	SQLGetCurrent_PostedTotal_ByCustByCat = SQLGetCurrent_PostedTotal_ByCustByCat & " AND PeriodYear = " & periodYearToFind 
	
	

'Response.Write(SQLGetCurrent_PostedTotal_ByCustByCat & "<br>") 

	Set rsGetCurrent_PostedTotal_ByCustByCat = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByCustByCat.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByCustByCat = cnnGetCurrent_PostedTotal_ByCustByCat.Execute(SQLGetCurrent_PostedTotal_ByCustByCat)

	If not rsGetCurrent_PostedTotal_ByCustByCat.EOF Then resultGetCurrent_PostedTotal_ByCustByCat = rsGetCurrent_PostedTotal_ByCustByCat("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByCustByCat) Then resultGetCurrent_PostedTotal_ByCustByCat = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByCustByCat.Close
	set rsGetCurrent_PostedTotal_ByCustByCat= Nothing
	cnnGetCurrent_PostedTotal_ByCustByCat.Close	
	set cnnGetCurrent_PostedTotal_ByCustByCat= Nothing
	
	GetCurrent_PostedTotal_ByCustByCat = resultGetCurrent_PostedTotal_ByCustByCat

End Function

Function CustHasMCSNotes(passedCust)

	Set cnnCustHasMCSNotes = Server.CreateObject("ADODB.Connection")
	cnnCustHasMCSNotes.open Session("ClientCnnString")

	resultCustHasMCSNotes = False
		
	SQLCustHasMCSNotes = "SELECT TOP 1 * FROM AR_CustomerNotes "
	SQLCustHasMCSNotes = SQLCustHasMCSNotes & "WHERE CustID = '" & passedCust & "' AND NoteType = 'MCS'"
	 
	Set rsCustHasMCSNotes = Server.CreateObject("ADODB.Recordset")
	rsCustHasMCSNotes.CursorLocation = 3 
	Set rsCustHasMCSNotes= cnnCustHasMCSNotes.Execute(SQLCustHasMCSNotes)
	
	If not rsCustHasMCSNotes.eof then resultCustHasMCSNotes =  True
		
	rsCustHasMCSNotes.Close
	set rsCustHasMCSNotes= Nothing
	cnnCustHasMCSNotes.Close	
	set cnnCustHasMCSNotes = Nothing
	
	CustHasMCSNotes = resultCustHasMCSNotes 
	
End Function

Function NoteNewMCSForUser(passedCustNum)

	resultNoteNewMCSForUser = False
	
	SQLNoteNewMCSForUser = "SELECT * FROM AR_CustomerNotesUserViewed Where CustID ='" & passedCustNum & "' AND UserNo = " & Session("Userno") & " AND NoteType = 'MCS'"
	
	Set cnnNoteNewMCSForUser = Server.CreateObject("ADODB.Connection")
	cnnNoteNewMCSForUser.open (Session("ClientCnnString"))
	Set rsNoteNewMCSForUser = Server.CreateObject("ADODB.Recordset")
	rsNoteNewMCSForUser.CursorLocation = 3 
	Set rsNoteNewMCSForUser = cnnNoteNewMCSForUser.Execute(SQLNoteNewMCSForUser)

	Set rsNoteCatAnal = Server.CreateObject("ADODB.Recordset")
	rsNoteCatAnal.CursorLocation = 3 

	If not rsNoteNewMCSForUser.EOF Then
		'OK, so see when the last note was created, not by us
		SQLCustHasCategoryAnalNotes = "SELECT TOP 1 RecordCreationDateTime FROM AR_CustomerNotes "
		SQLCustHasCategoryAnalNotes = SQLCustHasCategoryAnalNotes & "WHERE CustID = '" & passedCust & "' AND NoteType = 'MCS'"
		SQLCustHasCategoryAnalNotes = SQLCustHasCategoryAnalNotes & " ORDER BY RecordCreationDateTime DESC"
		
		Set rsNoteCatAnal = cnnNoteNewMCSForUser.Execute(SQLCustHasCategoryAnalNotes)
		If Not rsNoteCatAnal.Eof Then
			If rsNoteNewMCSForUser("DateLastViewed") < rsNoteCatAnal("RecordCreationDateTime")  Then resultNoteNewMCSForUser = True
		End If
	Else
		resultNoteNewMCSForUser = True 'Also true if they have never seen any of them
	End If
	cnnNoteNewMCSForUser.close
	set rsNoteNewMCSForUser = nothing
	set rsNoteCatAnal = nothing
	set cnnNoteNewMCSForUser= nothing	

	NoteNewMCSForUser = resultNoteNewMCSForUser

End Function

Sub MARKNoteNewForUserMCS(passedCustNum)

	SQLMARKNoteNewForUserMCS = "SELECT * FROM AR_CustomerNotesUserViewed Where CustID ='" & passedCustNum & "' AND UserNo = " & Session("Userno") & " AND NoteType = 'MCS'"
	
	Set cnnMARKNoteNewForUserMCS = Server.CreateObject("ADODB.Connection")
	cnnMARKNoteNewForUserMCS.open (Session("ClientCnnString"))
	Set rMARKNoteNewForUserMCS = Server.CreateObject("ADODB.Recordset")
	rMARKNoteNewForUserMCS.CursorLocation = 3 
	Set rMARKNoteNewForUserMCS = cnnMARKNoteNewForUserMCS.Execute(SQLMARKNoteNewForUserMCS)

	If rMARKNoteNewForUserMCS.EOF Then ' Nothing there so we need to insert
		SQLMARKNoteNewForUserMCS = "INSERT INTO AR_CustomerNotesUserViewed (CustID ,UserNo, Category) VALUES ('" & passedCustNum & "',"  & Session("UserNo") & "," & -2 & ")"
	Else
		SQLMARKNoteNewForUserMCS = "UPDATE AR_CustomerNotesUserViewed Set DateLastViewed = getdate() Where CustID ='" & passedCustNum & "' AND UserNo = " & Session("Userno") & " AND Category = -2"
	End If
	
	Set rMARKNoteNewForUserMCS = cnnMARKNoteNewForUserMCS.Execute(SQLMARKNoteNewForUserMCS)
		
	cnnMARKNoteNewForUserMCS.close
	set rMARKNoteNewForUserMCS = nothing
	set cnnMARKNoteNewForUserMCS= nothing	

End Sub

Function GetMostRecentMCSNote(passedCust)

	Set cnnGetMostRecentMCSNote = Server.CreateObject("ADODB.Connection")
	cnnGetMostRecentMCSNote.open Session("ClientCnnString")

	resultGetMostRecentMCSNote = ""
		
	SQLGetMostRecentMCSNote = "SELECT TOP 1 * FROM AR_CustomerNotes "
	SQLGetMostRecentMCSNote = SQLGetMostRecentMCSNote & "WHERE CustID = '" & passedCust & "' AND NoteType = 'MCS' ORDER BY RecordCreationDateTime DESC"
	 
	Set rsGetMostRecentMCSNote = Server.CreateObject("ADODB.Recordset")
	rsGetMostRecentMCSNote.CursorLocation = 3 
	Set rsGetMostRecentMCSNote= cnnGetMostRecentMCSNote.Execute(SQLGetMostRecentMCSNote)
	
	If not rsGetMostRecentMCSNote.eof then resultGetMostRecentMCSNote =  rsGetMostRecentMCSNote("Note")
		
	rsGetMostRecentMCSNote.Close
	set rsGetMostRecentMCSNote= Nothing
	cnnGetMostRecentMCSNote.Close	
	set cnnGetMostRecentMCSNote = Nothing
	
	GetMostRecentMCSNote = resultGetMostRecentMCSNote 
	
End Function

Function GetMostRecentMCSNoteUserNo(passedCust)

	Set cnnGetMostRecentMCSNoteUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetMostRecentMCSNoteUserNo.open Session("ClientCnnString")

	resultGetMostRecentMCSNoteUserNo = ""
		
	SQLGetMostRecentMCSNoteUserNo = "SELECT TOP 1 * FROM AR_CustomerNotes "
	SQLGetMostRecentMCSNoteUserNo = SQLGetMostRecentMCSNoteUserNo & "WHERE CustID = '" & passedCust & "' AND NoteType = 'MCS' ORDER BY RecordCreationDateTime DESC"
	 
	Set rsGetMostRecentMCSNoteUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetMostRecentMCSNoteUserNo.CursorLocation = 3 
	Set rsGetMostRecentMCSNoteUserNo= cnnGetMostRecentMCSNoteUserNo.Execute(SQLGetMostRecentMCSNoteUserNo)
	
	If not rsGetMostRecentMCSNoteUserNo.eof then resultGetMostRecentMCSNoteUserNo =  rsGetMostRecentMCSNoteUserNo("EnteredByUserNo")
		
	rsGetMostRecentMCSNoteUserNo.Close
	set rsGetMostRecentMCSNoteUserNo= Nothing
	cnnGetMostRecentMCSNoteUserNo.Close	
	set cnnGetMostRecentMCSNoteUserNo = Nothing
	
	GetMostRecentMCSNoteUserNo = resultGetMostRecentMCSNoteUserNo 
	
End Function


Function NumberOfMCSActionsWithMCSReason(passedMCSReasonIntRecID)

	Set cnnNumberOfMCSActionsByMCSReason = Server.CreateObject("ADODB.Connection")
	cnnNumberOfMCSActionsByMCSReason.open Session("ClientCnnString")

	resultNumberOfMCSActionsByMCSReason = 0
	
	SQLNumberOfMCSActionsByMCSReason = "SELECT * FROM BI_MCSActions WHERE MCSReasonIntRecID = " & passedMCSReasonIntRecID
	 
	Set rsNumberOfMCSActionsByMCSReason = Server.CreateObject("ADODB.Recordset")
	rsNumberOfMCSActionsByMCSReason.CursorLocation = 3 
	
	rsNumberOfMCSActionsByMCSReason.Open SQLNumberOfMCSActionsByMCSReason,cnnNumberOfMCSActionsByMCSReason 
			
	resultNumberOfMCSActionsByMCSReason = rsNumberOfMCSActionsByMCSReason.RecordCount
	
	rsNumberOfMCSActionsByMCSReason.Close
	set rsNumberOfMCSActionsByMCSReason = Nothing
	cnnNumberOfMCSActionsByMCSReason.Close	
	set cnnNumberOfMCSActionsByMCSReason = Nothing
	
	NumberOfMCSActionsByMCSReason = resultNumberOfMCSActionsByMCSReason 
	
End Function


Function GetMCSReasonByReasonNum(passedMCSReasonIntRecID)

	resultGetMCSReasonByReasonNum = 0

	Set cnnGetMCSReasonByReasonNum = Server.CreateObject("ADODB.Connection")
	cnnGetMCSReasonByReasonNum.open Session("ClientCnnString")
		
	SQLGetMCSReasonByReasonNum = "SELECT * FROM BI_MCSReasons WHERE InternalRecordIdentifier = " & passedMCSReasonIntRecID
	 
	Set rsGetMCSReasonByReasonNum = Server.CreateObject("ADODB.Recordset")
	rsGetMCSReasonByReasonNum.CursorLocation = 3 
	Set rsGetMCSReasonByReasonNum = cnnGetMCSReasonByReasonNum.Execute(SQLGetMCSReasonByReasonNum)
			
	If not rsGetMCSReasonByReasonNum.EOF Then resultGetMCSReasonByReasonNum = rsGetMCSReasonByReasonNum("Reason")
	
	rsGetMCSReasonByReasonNum.Close
	set rsGetMCSReasonByReasonNum= Nothing
	cnnGetMCSReasonByReasonNum.Close	
	set cnnGetMCSReasonByReasonNum= Nothing
	
	GetMCSReasonByReasonNum = resultGetMCSReasonByReasonNum
	
End Function


Function TotalSalesByCustByMonthByYear_RentalsOnly(passedCustID,passedMonth,passedYear)

	resultTotalSalesByCustByMonthByYear_RentalsOnly = ""

	Set cnnTotalSalesByCustByMonthByYear_RentalsOnly = Server.CreateObject("ADODB.Connection")
	cnnTotalSalesByCustByMonthByYear_RentalsOnly.open Session("ClientCnnString")
		
	SQLTotalSalesByCustByMonthByYear_RentalsOnly = "SELECT Sum(IvsTotalAmt-IvsSalesTax-IvsDepositChg) AS TotalSales FROM InvoiceHistory WHERE CustNum = '" & passedCustID & "' "
	SQLTotalSalesByCustByMonthByYear_RentalsOnly = SQLTotalSalesByCustByMonthByYear_RentalsOnly  & " AND Month(IvsDate) = " & passedMonth
	SQLTotalSalesByCustByMonthByYear_RentalsOnly = SQLTotalSalesByCustByMonthByYear_RentalsOnly  & " AND Year(IvsDate) = " & passedYear
	SQLTotalSalesByCustByMonthByYear_RentalsOnly = SQLTotalSalesByCustByMonthByYear_RentalsOnly  & " AND IvsType = 'G'"

	Set rsTotalSalesByCustByMonthByYear_RentalsOnly = Server.CreateObject("ADODB.Recordset")
	rsTotalSalesByCustByMonthByYear_RentalsOnly.CursorLocation = 3 
	Set rsTotalSalesByCustByMonthByYear_RentalsOnly = cnnTotalSalesByCustByMonthByYear_RentalsOnly.Execute(SQLTotalSalesByCustByMonthByYear_RentalsOnly)

	If not rsTotalSalesByCustByMonthByYear_RentalsOnly.EOF Then
		resultTotalSalesByCustByMonthByYear_RentalsOnly = rsTotalSalesByCustByMonthByYear_RentalsOnly("TotalSales")
	Else
		resultTotalSalesByCustByMonthByYear_RentalsOnly = 0 ' Because there were no sales
	End If
	If Not IsNumeric(resultTotalSalesByCustByMonthByYear_RentalsOnly) Then resultTotalSalesByCustByMonthByYear_RentalsOnly = 0 'To account for null result

	rsTotalSalesByCustByMonthByYear_RentalsOnly.Close
	set rsTotalSalesByCustByMonthByYear_RentalsOnly= Nothing
	cnnTotalSalesByCustByMonthByYear_RentalsOnly.Close	
	set cnnTotalSalesByCustByMonthByYear_RentalsOnly= Nothing
	
	TotalSalesByCustByMonthByYear_RentalsOnly = resultTotalSalesByCustByMonthByYear_RentalsOnly

End Function

Function TotalPostedLVFByCustByMonthByYear(passedCustID,passedMonth,passedYear)

	resultTotalPostedLVFByCustByMonthByYear = ""

	Set cnnTotalPostedLVFByCustByMonthByYear = Server.CreateObject("ADODB.Connection")
	cnnTotalPostedLVFByCustByMonthByYear.open Session("ClientCnnString")
		
	SQLTotalPostedLVFByCustByMonthByYear = "SELECT Sum(itemQuantity * itemPrice) AS TotalLVF FROM InvoiceHistoryDetail WHERE CustNum = '" & passedCustID & "' "
	SQLTotalPostedLVFByCustByMonthByYear = SQLTotalPostedLVFByCustByMonthByYear  & " AND Month(IvsDate) = " & passedMonth
	SQLTotalPostedLVFByCustByMonthByYear = SQLTotalPostedLVFByCustByMonthByYear  & " AND Year(IvsDate) = " & passedYear
	SQLTotalPostedLVFByCustByMonthByYear = SQLTotalPostedLVFByCustByMonthByYear  & " AND partnum = 'LVM'"

	Set rsTotalPostedLVFByCustByMonthByYear = Server.CreateObject("ADODB.Recordset")
	rsTotalPostedLVFByCustByMonthByYear.CursorLocation = 3 
	Set rsTotalPostedLVFByCustByMonthByYear = cnnTotalPostedLVFByCustByMonthByYear.Execute(SQLTotalPostedLVFByCustByMonthByYear)

	If not rsTotalPostedLVFByCustByMonthByYear.EOF Then
		resultTotalPostedLVFByCustByMonthByYear = rsTotalPostedLVFByCustByMonthByYear("TotalLVF")
	Else
		resultTotalPostedLVFByCustByMonthByYear = 0 ' Because there were no sales
	End If
	If Not IsNumeric(resultTotalPostedLVFByCustByMonthByYear) Then resultTotalPostedLVFByCustByMonthByYear = 0 'To account for null result

	rsTotalPostedLVFByCustByMonthByYear.Close
	set rsTotalPostedLVFByCustByMonthByYear= Nothing
	cnnTotalPostedLVFByCustByMonthByYear.Close	
	set cnnTotalPostedLVFByCustByMonthByYear= Nothing
	
	TotalPostedLVFByCustByMonthByYear = resultTotalPostedLVFByCustByMonthByYear

End Function

Function TotalXSFByCustByMonthByYear(passedCustID,passedMonth,passedYear)

	resultTotalXSFByCustByMonthByYear = 0

	Set cnnTotalXSFByCustByMonthByYear = Server.CreateObject("ADODB.Connection")
	cnnTotalXSFByCustByMonthByYear.open Session("ClientCnnString")
		
	SQLTotalXSFByCustByMonthByYear = "SELECT Sum(itemQuantity * itemPrice) AS TotalLVF FROM InvoiceHistoryDetail WHERE CustNum = '" & passedCustID & "' "
	SQLTotalXSFByCustByMonthByYear = SQLTotalXSFByCustByMonthByYear  & " AND Month(IvsDate) = " & passedMonth
	SQLTotalXSFByCustByMonthByYear = SQLTotalXSFByCustByMonthByYear  & " AND Year(IvsDate) = " & passedYear
	SQLTotalXSFByCustByMonthByYear = SQLTotalXSFByCustByMonthByYear  & " AND partnum LIKE 'XSF%'"

	Set rsTotalXSFByCustByMonthByYear = Server.CreateObject("ADODB.Recordset")
	rsTotalXSFByCustByMonthByYear.CursorLocation = 3 
	Set rsTotalXSFByCustByMonthByYear = cnnTotalXSFByCustByMonthByYear.Execute(SQLTotalXSFByCustByMonthByYear)

	If not rsTotalXSFByCustByMonthByYear.EOF Then
		resultTotalXSFByCustByMonthByYear = rsTotalXSFByCustByMonthByYear("TotalLVF")
	Else
		resultTotalXSFByCustByMonthByYear = 0 ' Because there were no sales
	End If
	If Not IsNumeric(resultTotalXSFByCustByMonthByYear) Then resultTotalXSFByCustByMonthByYear = 0 'To account for null result

	rsTotalXSFByCustByMonthByYear.Close
	set rsTotalXSFByCustByMonthByYear= Nothing
	cnnTotalXSFByCustByMonthByYear.Close	
	set cnnTotalXSFByCustByMonthByYear= Nothing
	
	TotalXSFByCustByMonthByYear = resultTotalXSFByCustByMonthByYear

End Function

Function TotalCostByCustByMonthByYear_NoRent(passedCustID,passedMonth,passedYear)

	resultTotalCostByCustByMonthByYear_NoRent = ""

	Set cnnTotalCostByCustByMonthByYear_NoRent = Server.CreateObject("ADODB.Connection")
	cnnTotalCostByCustByMonthByYear_NoRent.open Session("ClientCnnString")
		
	SQLTotalCostByCustByMonthByYear_NoRent = "SELECT Sum(itemQuantity*itemCost) AS PeriodTotCost FROM InvoiceHistoryDetail WHERE "
	SQLTotalCostByCustByMonthByYear_NoRent = SQLTotalCostByCustByMonthByYear_NoRent & "CustNum = '" & passedCustID & "'"
	SQLTotalCostByCustByMonthByYear_NoRent = SQLTotalCostByCustByMonthByYear_NoRent & " AND Month(IvsDate) = " & passedMonth
	SQLTotalCostByCustByMonthByYear_NoRent = SQLTotalCostByCustByMonthByYear_NoRent & " AND Year(IvsDate) = " & passedYear
	SQLTotalCostByCustByMonthByYear_NoRent = SQLTotalCostByCustByMonthByYear_NoRent & " AND prodCategory <> 21 " 
		 
	Set rsTotalCostByCustByMonthByYear_NoRent = Server.CreateObject("ADODB.Recordset")
	rsTotalCostByCustByMonthByYear_NoRent.CursorLocation = 3 
	Set rsTotalCostByCustByMonthByYear_NoRent = cnnTotalCostByCustByMonthByYear_NoRent.Execute(SQLTotalCostByCustByMonthByYear_NoRent)

	If not rsTotalCostByCustByMonthByYear_NoRent.EOF Then resultTotalCostByCustByMonthByYear_NoRent = rsTotalCostByCustByMonthByYear_NoRent("PeriodTotCost")

	rsTotalCostByCustByMonthByYear_NoRent.Close
	set rsTotalCostByCustByMonthByYear_NoRent= Nothing
	cnnTotalCostByCustByMonthByYear_NoRent.Close	
	set cnnTotalCostByCustByMonthByYear_NoRent= Nothing
	
	TotalCostByCustByMonthByYear_NoRent = resultTotalCostByCustByMonthByYear_NoRent

End Function

Function TotalSalesByCustByMonthByYear_NoRentals(passedCustID,passedMonth,passedYear)

	resultTotalSalesByCustByMonthByYear_NoRentals = ""

	Set cnnTotalSalesByCustByMonthByYear_NoRentals = Server.CreateObject("ADODB.Connection")
	cnnTotalSalesByCustByMonthByYear_NoRentals.open Session("ClientCnnString")
		
	SQLTotalSalesByCustByMonthByYear_NoRentals = "SELECT Sum(IvsTotalAmt-IvsSalesTax-IvsDepositChg) AS TotalSales FROM InvoiceHistory WHERE CustNum = '" & passedCustID & "' "
	SQLTotalSalesByCustByMonthByYear_NoRentals = SQLTotalSalesByCustByMonthByYear_NoRentals  & " AND Month(IvsDate) = " & passedMonth
	SQLTotalSalesByCustByMonthByYear_NoRentals = SQLTotalSalesByCustByMonthByYear_NoRentals  & " AND Year(IvsDate) = " & passedYear
	SQLTotalSalesByCustByMonthByYear_NoRentals = SQLTotalSalesByCustByMonthByYear_NoRentals  & " AND IvsType <> 'G' AND IvsType <> 'E'"

	Set rsTotalSalesByCustByMonthByYear_NoRentals = Server.CreateObject("ADODB.Recordset")
	rsTotalSalesByCustByMonthByYear_NoRentals.CursorLocation = 3 
	Set rsTotalSalesByCustByMonthByYear_NoRentals = cnnTotalSalesByCustByMonthByYear_NoRentals.Execute(SQLTotalSalesByCustByMonthByYear_NoRentals)

	If not rsTotalSalesByCustByMonthByYear_NoRentals.EOF Then
		resultTotalSalesByCustByMonthByYear_NoRentals = rsTotalSalesByCustByMonthByYear_NoRentals("TotalSales")
	Else
		resultTotalSalesByCustByMonthByYear_NoRentals = 0 ' Because there were no sales
	End If
	If Not IsNumeric(resultTotalSalesByCustByMonthByYear_NoRentals) Then resultTotalSalesByCustByMonthByYear_NoRentals = 0 'To account for null result

	rsTotalSalesByCustByMonthByYear_NoRentals.Close
	set rsTotalSalesByCustByMonthByYear_NoRentals= Nothing
	cnnTotalSalesByCustByMonthByYear_NoRentals.Close	
	set cnnTotalSalesByCustByMonthByYear_NoRentals= Nothing
	
	TotalSalesByCustByMonthByYear_NoRentals = resultTotalSalesByCustByMonthByYear_NoRentals

End Function

Function TotalSalesByCustByMonthByYear_RentalsOnly(passedCustID,passedMonth,passedYear)

	resultTotalSalesByCustByMonthByYear_RentalsOnly = ""

	Set cnnTotalSalesByCustByMonthByYear_RentalsOnly = Server.CreateObject("ADODB.Connection")
	cnnTotalSalesByCustByMonthByYear_RentalsOnly.open Session("ClientCnnString")
		
	SQLTotalSalesByCustByMonthByYear_RentalsOnly = "SELECT Sum(IvsTotalAmt-IvsSalesTax-IvsDepositChg) AS TotalSales FROM InvoiceHistory WHERE CustNum = '" & passedCustID & "' "
	SQLTotalSalesByCustByMonthByYear_RentalsOnly = SQLTotalSalesByCustByMonthByYear_RentalsOnly  & " AND Month(IvsDate) = " & passedMonth
	SQLTotalSalesByCustByMonthByYear_RentalsOnly = SQLTotalSalesByCustByMonthByYear_RentalsOnly  & " AND Year(IvsDate) = " & passedYear
	SQLTotalSalesByCustByMonthByYear_RentalsOnly = SQLTotalSalesByCustByMonthByYear_RentalsOnly  & " AND IvsType = 'G'"

	Set rsTotalSalesByCustByMonthByYear_RentalsOnly = Server.CreateObject("ADODB.Recordset")
	rsTotalSalesByCustByMonthByYear_RentalsOnly.CursorLocation = 3 
	Set rsTotalSalesByCustByMonthByYear_RentalsOnly = cnnTotalSalesByCustByMonthByYear_RentalsOnly.Execute(SQLTotalSalesByCustByMonthByYear_RentalsOnly)

	If not rsTotalSalesByCustByMonthByYear_RentalsOnly.EOF Then
		resultTotalSalesByCustByMonthByYear_RentalsOnly = rsTotalSalesByCustByMonthByYear_RentalsOnly("TotalSales")
	Else
		resultTotalSalesByCustByMonthByYear_RentalsOnly = 0 ' Because there were no sales
	End If
	If Not IsNumeric(resultTotalSalesByCustByMonthByYear_RentalsOnly) Then resultTotalSalesByCustByMonthByYear_RentalsOnly = 0 'To account for null result

	rsTotalSalesByCustByMonthByYear_RentalsOnly.Close
	set rsTotalSalesByCustByMonthByYear_RentalsOnly= Nothing
	cnnTotalSalesByCustByMonthByYear_RentalsOnly.Close	
	set cnnTotalSalesByCustByMonthByYear_RentalsOnly= Nothing
	
	TotalSalesByCustByMonthByYear_RentalsOnly = resultTotalSalesByCustByMonthByYear_RentalsOnly

End Function

Function GetCurrent_UnpostedTotal_ByCust(passedCustID,passedPeriodBeingEvaluated)

	StartDateToFind = GetPeriodBeginDateBySeq(passedPeriodBeingEvaluated+1)
	EndDateToFind = GetPeriodEndDateBySeq(passedPeriodBeingEvaluated+1)
	
	resultGetCurrent_UnpostedTotal_ByCust = 0

	Set cnnGetCurrent_UnpostedTotal_ByCust = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnpostedTotal_ByCust.open Session("ClientCnnString")
		
	SQLGetCurrent_UnpostedTotal_ByCust = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatMonth WHERE CustID='" & passedCustID & "' AND "
	SQLGetCurrent_UnpostedTotal_ByCust = SQLGetCurrent_UnpostedTotal_ByCust & "PostedOrUnposted = 'U' "
	SQLGetCurrent_UnpostedTotal_ByCust = SQLGetCurrent_UnpostedTotal_ByCust & " AND CategoryID <> 21"

	Set rsGetCurrent_UnpostedTotal_ByCust = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnpostedTotal_ByCust.CursorLocation = 3 
	Set rsGetCurrent_UnpostedTotal_ByCust = cnnGetCurrent_UnpostedTotal_ByCust.Execute(SQLGetCurrent_UnpostedTotal_ByCust)

	If not rsGetCurrent_UnpostedTotal_ByCust.EOF Then resultGetCurrent_UnpostedTotal_ByCust = rsGetCurrent_UnpostedTotal_ByCust("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnpostedTotal_ByCust) Then resultGetCurrent_UnpostedTotal_ByCust = 0 ' In case there are no results
	
	rsGetCurrent_UnpostedTotal_ByCust.Close
	set rsGetCurrent_UnpostedTotal_ByCust= Nothing
	cnnGetCurrent_UnpostedTotal_ByCust.Close	
	set cnnGetCurrent_UnpostedTotal_ByCust= Nothing
	
	GetCurrent_UnpostedTotal_ByCust = resultGetCurrent_UnpostedTotal_ByCust 

End Function

Function GetCurrent_PostedTotal_ByCust(passedCustID,passedPeriodBeingEvaluated)


	StartDateToFind = GetPeriodBeginDateBySeq(passedPeriodBeingEvaluated+1)
	EndDateToFind = GetPeriodEndDateBySeq(passedPeriodBeingEvaluated+1)
	
	resultGetCurrent_PostedTotal_ByCust = 0

	Set cnnGetCurrent_PostedTotal_ByCust = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByCust.open Session("ClientCnnString")
		
	SQLGetCurrent_PostedTotal_ByCust = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatMonth WHERE CustID='" & passedCustID & "' AND "
	SQLGetCurrent_PostedTotal_ByCust = SQLGetCurrent_PostedTotal_ByCust & "PostedOrUnposted = 'P' "
	SQLGetCurrent_PostedTotal_ByCust = SQLGetCurrent_PostedTotal_ByCust & " AND CategoryID <> 21"
	
	Set rsGetCurrent_PostedTotal_ByCust = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByCust.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByCust = cnnGetCurrent_PostedTotal_ByCust.Execute(SQLGetCurrent_PostedTotal_ByCust)

	If not rsGetCurrent_PostedTotal_ByCust.EOF Then resultGetCurrent_PostedTotal_ByCust = rsGetCurrent_PostedTotal_ByCust("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByCust) Then resultGetCurrent_PostedTotal_ByCust = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByCust.Close
	set rsGetCurrent_PostedTotal_ByCust= Nothing
	cnnGetCurrent_PostedTotal_ByCust.Close	
	set cnnGetCurrent_PostedTotal_ByCust= Nothing

	
	GetCurrent_PostedTotal_ByCust = resultGetCurrent_PostedTotal_ByCust

End Function

Function TotalCat21ByCustByMonthByYear(passedCustID,passedMonth,passedYear)

	resultTotalCat21ByCustByMonthByYear = ""

	Set cnnTotalCat21ByCustByMonthByYear = Server.CreateObject("ADODB.Connection")
	cnnTotalCat21ByCustByMonthByYear.open Session("ClientCnnString")
		
	SQLTotalCat21ByCustByMonthByYear = "SELECT Sum(itemQuantity * itemPrice) AS TotalLVF FROM InvoiceHistoryDetail WHERE CustNum = '" & passedCustID & "' "
	SQLTotalCat21ByCustByMonthByYear = SQLTotalCat21ByCustByMonthByYear  & " AND Month(IvsDate) = " & passedMonth
	SQLTotalCat21ByCustByMonthByYear = SQLTotalCat21ByCustByMonthByYear  & " AND Year(IvsDate) = " & passedYear
	SQLTotalCat21ByCustByMonthByYear = SQLTotalCat21ByCustByMonthByYear  & " AND prodCategory = 21"

	Set rsTotalCat21ByCustByMonthByYear = Server.CreateObject("ADODB.Recordset")
	rsTotalCat21ByCustByMonthByYear.CursorLocation = 3 
	Set rsTotalCat21ByCustByMonthByYear = cnnTotalCat21ByCustByMonthByYear.Execute(SQLTotalCat21ByCustByMonthByYear)

	If not rsTotalCat21ByCustByMonthByYear.EOF Then
		resultTotalCat21ByCustByMonthByYear = rsTotalCat21ByCustByMonthByYear("TotalLVF")
	Else
		resultTotalCat21ByCustByMonthByYear = 0 ' Because there were no sales
	End If
	If Not IsNumeric(resultTotalCat21ByCustByMonthByYear) Then resultTotalCat21ByCustByMonthByYear = 0 'To account for null result

	rsTotalCat21ByCustByMonthByYear.Close
	set rsTotalCat21ByCustByMonthByYear= Nothing
	cnnTotalCat21ByCustByMonthByYear.Close	
	set cnnTotalCat21ByCustByMonthByYear= Nothing
	
	TotalCat21ByCustByMonthByYear = resultTotalCat21ByCustByMonthByYear

End Function


Function GetMESNotesStatus(passedCustID, MESMonth)

	resultGetMESNotesStatus = 0

	Set cnnGetMESNotesStatus = Server.CreateObject("ADODB.Connection")
	cnnGetMESNotesStatus.open Session("ClientCnnString")
		
	SQLGetMESNotesStatus = "SELECT CustID, MESMonth, Action FROM BI_MESActions WHERE CustID = " & passedCustID & " AND MESMonth = '" & MESMonth & "'"

	Set rsGetMESNotesStatus = Server.CreateObject("ADODB.Recordset")
	rsGetMESNotesStatus.CursorLocation = 3 
	Set rsGetMESNotesStatus = cnnGetMESNotesStatus.Execute(SQLGetMESNotesStatus)
			 
	If not rsGetMESNotesStatus.EOF Then 
		if rsGetMESNotesStatus("Action") = "no_action_necessary" Then
			resultGetMESNotesStatus = 2
		Else
			resultGetMESNotesStatus = 1
		End If
	end if

	rsGetMESNotesStatus.Close
	
	set rsGetMESNotesStatus= Nothing
	cnnGetMESNotesStatus.Close	
	set cnnGetMESNotesStatus= Nothing
	
	GetMESNotesStatus = resultGetMESNotesStatus
	
End Function

Function GetMESNotesNoActionStatus(passedCustID, MESMonth)

	GetMESNotesNoActionStatus = 0

	Set cnnGetMESNotesStatus = Server.CreateObject("ADODB.Connection")
	cnnGetMESNotesStatus.open Session("ClientCnnString")
		
	SQLGetMESNotesStatus = "SELECT CustID, MESMonth, Action FROM BI_MESActions WHERE CustID = " & passedCustID & " AND MESMonth = '" & MESMonth & "' AND Action='no_action_necessary'"

	Set rsGetMESNotesStatus = Server.CreateObject("ADODB.Recordset")
	rsGetMESNotesStatus.CursorLocation = 3 
	Set rsGetMESNotesStatus = cnnGetMESNotesStatus.Execute(SQLGetMESNotesStatus)
			 
	If not rsGetMESNotesStatus.EOF Then 
		if rsGetMESNotesStatus("Action") = "no_action_necessary" Then
			GetMESNotesNoActionStatus = 2
		Else
			GetMESNotesNoActionStatus = 1
		End If
	end if

	rsGetMESNotesStatus.Close
	
	set rsGetMESNotesStatus= Nothing
	cnnGetMESNotesStatus.Close	
	set cnnGetMESNotesStatus= Nothing
	
	GetMESNotesNoActionStatus = GetMESNotesNoActionStatus
	
End Function

Function GetMCSNotesStatus(passedCustID, MCSMonth)

	resultGetMCSNotesStatus = 0

	Set cnnGetMCSNotesStatus = Server.CreateObject("ADODB.Connection")
	cnnGetMCSNotesStatus.open Session("ClientCnnString")
		
	SQLGetMCSNotesStatus = "SELECT CustID, MCSMonth, Action FROM BI_MCSActions WHERE CustID = " & passedCustID & " AND MCSMonth = '" & MCSMonth & "'"

	Set rsGetMCSNotesStatus = Server.CreateObject("ADODB.Recordset")
	rsGetMCSNotesStatus.CursorLocation = 3 
	Set rsGetMCSNotesStatus = cnnGetMCSNotesStatus.Execute(SQLGetMCSNotesStatus)
			 
	If not rsGetMCSNotesStatus.EOF Then 
		if rsGetMCSNotesStatus("Action") = "no_action_necessary" Then
			resultGetMCSNotesStatus = 2
		Else
			resultGetMCSNotesStatus = 1
		End If
	end if

	rsGetMCSNotesStatus.Close
	
	set rsGetMCSNotesStatus= Nothing
	cnnGetMCSNotesStatus.Close	
	set cnnGetMCSNotesStatus= Nothing
	
	GetMCSNotesStatus = resultGetMCSNotesStatus
	
End Function

Function GetMCSNotesNoActionStatus(passedCustID, MCSMonth)

	GetMCSNotesNoActionStatus = 0

	Set cnnGetMCSNotesStatus = Server.CreateObject("ADODB.Connection")
	cnnGetMCSNotesStatus.open Session("ClientCnnString")
		
	SQLGetMCSNotesStatus = "SELECT CustID, MCSMonth, Action FROM BI_MCSActions WHERE CustID = " & passedCustID & " AND MCSMonth = '" & MCSMonth & "' AND Action='no_action_necessary'"

	Set rsGetMCSNotesStatus = Server.CreateObject("ADODB.Recordset")
	rsGetMCSNotesStatus.CursorLocation = 3 
	Set rsGetMCSNotesStatus = cnnGetMCSNotesStatus.Execute(SQLGetMCSNotesStatus)
			 
	If not rsGetMCSNotesStatus.EOF Then 
		if rsGetMCSNotesStatus("Action") = "no_action_necessary" Then
			GetMCSNotesNoActionStatus = 2
		Else
			GetMCSNotesNoActionStatus = 1
		End If
	end if

	rsGetMCSNotesStatus.Close
	
	set rsGetMCSNotesStatus= Nothing
	cnnGetMCSNotesStatus.Close	
	set cnnGetMCSNotesStatus= Nothing
	
	GetMCSNotesNoActionStatus = GetMCSNotesNoActionStatus
	
End Function

Function CustHasMESNotes(passedCust)

	Set cnnCustHasMESNotes = Server.CreateObject("ADODB.Connection")
	cnnCustHasMESNotes.open Session("ClientCnnString")

	resultCustHasMESNotes = False
		
	SQLCustHasMESNotes = "SELECT TOP 1 * FROM AR_CustomerNotes "
	SQLCustHasMESNotes = SQLCustHasMESNotes & "WHERE CustID = '" & passedCust & "' AND NoteType = 'MES'"
	 
	Set rsCustHasMESNotes = Server.CreateObject("ADODB.Recordset")
	rsCustHasMESNotes.CursorLocation = 3 
	Set rsCustHasMESNotes= cnnCustHasMESNotes.Execute(SQLCustHasMESNotes)
	
	If not rsCustHasMESNotes.eof then resultCustHasMESNotes =  True
		
	rsCustHasMESNotes.Close
	set rsCustHasMESNotes= Nothing
	cnnCustHasMESNotes.Close	
	set cnnCustHasMESNotes = Nothing
	
	CustHasMESNotes = resultCustHasMESNotes 
	
End Function

Function GetMostRecentMESNoteUserNo(passedCust)

	Set cnnGetMostRecentMESNoteUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetMostRecentMESNoteUserNo.open Session("ClientCnnString")

	resultGetMostRecentMESNoteUserNo = ""
		
	SQLGetMostRecentMESNoteUserNo = "SELECT TOP 1 * FROM AR_CustomerNotes "
	SQLGetMostRecentMESNoteUserNo = SQLGetMostRecentMESNoteUserNo & "WHERE CustID = '" & passedCust & "' AND NoteType = 'MES' ORDER BY RecordCreationDateTime DESC"
	 
	Set rsGetMostRecentMESNoteUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetMostRecentMESNoteUserNo.CursorLocation = 3 
	Set rsGetMostRecentMESNoteUserNo= cnnGetMostRecentMESNoteUserNo.Execute(SQLGetMostRecentMESNoteUserNo)
	
	If not rsGetMostRecentMESNoteUserNo.eof then resultGetMostRecentMESNoteUserNo =  rsGetMostRecentMESNoteUserNo("EnteredByUserNo")
		
	rsGetMostRecentMESNoteUserNo.Close
	set rsGetMostRecentMESNoteUserNo= Nothing
	cnnGetMostRecentMESNoteUserNo.Close	
	set cnnGetMostRecentMESNoteUserNo = Nothing
	
	GetMostRecentMESNoteUserNo = resultGetMostRecentMESNoteUserNo 
	
End Function

Function GetMostRecentMESNote(passedCust)

	Set cnnGetMostRecentMESNote = Server.CreateObject("ADODB.Connection")
	cnnGetMostRecentMESNote.open Session("ClientCnnString")

	resultGetMostRecentMESNote = ""
		
	SQLGetMostRecentMESNote = "SELECT TOP 1 * FROM AR_CustomerNotes "
	SQLGetMostRecentMESNote = SQLGetMostRecentMESNote & "WHERE CustID = '" & passedCust & "' AND NoteType = 'MES' ORDER BY RecordCreationDateTime DESC"
	 
	Set rsGetMostRecentMESNote = Server.CreateObject("ADODB.Recordset")
	rsGetMostRecentMESNote.CursorLocation = 3 
	Set rsGetMostRecentMESNote= cnnGetMostRecentMESNote.Execute(SQLGetMostRecentMESNote)
	
	If not rsGetMostRecentMESNote.eof then resultGetMostRecentMESNote =  rsGetMostRecentMESNote("Note")
		
	rsGetMostRecentMESNote.Close
	set rsGetMostRecentMESNote= Nothing
	cnnGetMostRecentMESNote.Close	
	set cnnGetMostRecentMESNote = Nothing
	
	GetMostRecentMESNote = resultGetMostRecentMESNote 
	
End Function

Function PendingLVFByCust(passedCust)

	Set cnnPendingLVFByCust = Server.CreateObject("ADODB.Connection")
	cnnPendingLVFByCust.open Session("ClientCnnString")

	resultPendingLVFByCust = 0
		
	SQLPendingLVFByCust = "SELECT TOP 1 * FROM BI_MCSActions "
	SQLPendingLVFByCust = SQLPendingLVFByCust & "WHERE CustID = '" & passedCust & "' AND ActionNotes Like '%Send Invoice to client%'"
	SQLPendingLVFByCust = SQLPendingLVFByCust & " AND Month(RecordCreationDateTime) = Month(getdate()) AND Year(RecordCreationDateTime) = Year(getdate()) "
	 
	Set rsPendingLVFByCust = Server.CreateObject("ADODB.Recordset")
	rsPendingLVFByCust.CursorLocation = 3 
	Set rsPendingLVFByCust= cnnPendingLVFByCust.Execute(SQLPendingLVFByCust)
	
	If not rsPendingLVFByCust.eof then 
	
		resultPendingLVFByCust =  rsPendingLVFByCust("ActionNotes")
		positionOfDollarSign = InStr(resultPendingLVFByCust,"$")
		lengthOfActionNotes = Len(resultPendingLVFByCust)
		resultPendingLVFByCust = Right(resultPendingLVFByCust, lengthOfActionNotes - positionOfDollarSign)
		
		'..........Send invoice to client for the amount of $199.9
	End If
		
	rsPendingLVFByCust.Close
	set rsPendingLVFByCust= Nothing
	cnnPendingLVFByCust.Close	
	set cnnPendingLVFByCust = Nothing
	
	PendingLVFByCust = resultPendingLVFByCust 
	
End Function

Function GetLastMCSActionDateByMonthByYearByCust(passedCust, passedMonth, passedYear)

	Set cnnGetLastMCSActionDateByMonthByYearByCust = Server.CreateObject("ADODB.Connection")
	cnnGetLastMCSActionDateByMonthByYearByCust.open Session("ClientCnnString")

	resultGetLastMCSActionDateByMonthByYearByCust = ""
		
	SQLGetLastMCSActionDateByMonthByYearByCust = "SELECT TOP 1 * FROM BI_MCSActions "
	SQLGetLastMCSActionDateByMonthByYearByCust = SQLGetLastMCSActionDateByMonthByYearByCust & "WHERE CustID = '" & passedCust & "' AND "
	SQLGetLastMCSActionDateByMonthByYearByCust = SQLGetLastMCSActionDateByMonthByYearByCust & " (MONTH(RecordCreationDateTime)) = " & passedMonth & "  AND "
	SQLGetLastMCSActionDateByMonthByYearByCust = SQLGetLastMCSActionDateByMonthByYearByCust & " (YEAR(RecordCreationDateTime)) = " & passedYear 
	SQLGetLastMCSActionDateByMonthByYearByCust = SQLGetLastMCSActionDateByMonthByYearByCust & " ORDER BY RecordCreationDateTime DESC"
	 
	Set rsGetLastMCSActionDateByMonthByYearByCust = Server.CreateObject("ADODB.Recordset")
	rsGetLastMCSActionDateByMonthByYearByCust.CursorLocation = 3 
	Set rsGetLastMCSActionDateByMonthByYearByCust= cnnGetLastMCSActionDateByMonthByYearByCust.Execute(SQLGetLastMCSActionDateByMonthByYearByCust)
	
	If not rsGetLastMCSActionDateByMonthByYearByCust.eof then resultGetLastMCSActionDateByMonthByYearByCust =  rsGetLastMCSActionDateByMonthByYearByCust("RecordCreationDateTime")
		
	rsGetLastMCSActionDateByMonthByYearByCust.Close
	set rsGetLastMCSActionDateByMonthByYearByCust= Nothing
	cnnGetLastMCSActionDateByMonthByYearByCust.Close	
	set cnnGetLastMCSActionDateByMonthByYearByCust = Nothing
	
	GetLastMCSActionDateByMonthByYearByCust = resultGetLastMCSActionDateByMonthByYearByCust 
	
End Function

Function GetLastMCSActionByMonthByYearByCust(passedCust, passedMonth, passedYear)

	Set cnnGetLastMCSActionByMonthByYearByCust = Server.CreateObject("ADODB.Connection")
	cnnGetLastMCSActionByMonthByYearByCust.open Session("ClientCnnString")

	resultGetLastMCSActionByMonthByYearByCust = ""
		
	SQLGetLastMCSActionByMonthByYearByCust = "SELECT TOP 1 * FROM BI_MCSActions "
	SQLGetLastMCSActionByMonthByYearByCust = SQLGetLastMCSActionByMonthByYearByCust & "WHERE CustID = '" & passedCust & "' AND "
	SQLGetLastMCSActionByMonthByYearByCust = SQLGetLastMCSActionByMonthByYearByCust & " (MONTH(RecordCreationDateTime)) = " & passedMonth & "  AND "
	SQLGetLastMCSActionByMonthByYearByCust = SQLGetLastMCSActionByMonthByYearByCust & " (YEAR(RecordCreationDateTime)) = " & passedYear 
	SQLGetLastMCSActionByMonthByYearByCust = SQLGetLastMCSActionByMonthByYearByCust & " ORDER BY RecordCreationDateTime DESC"
	 
	Set rsGetLastMCSActionByMonthByYearByCust = Server.CreateObject("ADODB.Recordset")
	rsGetLastMCSActionByMonthByYearByCust.CursorLocation = 3 
	Set rsGetLastMCSActionByMonthByYearByCust= cnnGetLastMCSActionByMonthByYearByCust.Execute(SQLGetLastMCSActionByMonthByYearByCust)
	
	If not rsGetLastMCSActionByMonthByYearByCust.eof then resultGetLastMCSActionByMonthByYearByCust =  rsGetLastMCSActionByMonthByYearByCust("Action")
		
	rsGetLastMCSActionByMonthByYearByCust.Close
	set rsGetLastMCSActionByMonthByYearByCust= Nothing
	cnnGetLastMCSActionByMonthByYearByCust.Close	
	set cnnGetLastMCSActionByMonthByYearByCust = Nothing
	
	GetLastMCSActionByMonthByYearByCust = resultGetLastMCSActionByMonthByYearByCust 
	
End Function

Function GetLastMCSActionNoteByMonthByYearByCust(passedCust, passedMonth, passedYear)

	Set cnnGetLastMCSActionNoteByMonthByYearByCust = Server.CreateObject("ADODB.Connection")
	cnnGetLastMCSActionNoteByMonthByYearByCust.open Session("ClientCnnString")

	resultGetLastMCSActionNoteByMonthByYearByCust = ""
		
	SQLGetLastMCSActionNoteByMonthByYearByCust = "SELECT TOP 1 * FROM BI_MCSActions "
	SQLGetLastMCSActionNoteByMonthByYearByCust = SQLGetLastMCSActionNoteByMonthByYearByCust & "WHERE CustID = '" & passedCust & "' AND "
	SQLGetLastMCSActionNoteByMonthByYearByCust = SQLGetLastMCSActionNoteByMonthByYearByCust & " (MONTH(RecordCreationDateTime)) = " & passedMonth & "  AND "
	SQLGetLastMCSActionNoteByMonthByYearByCust = SQLGetLastMCSActionNoteByMonthByYearByCust & " (YEAR(RecordCreationDateTime)) = " & passedYear 
	
	'FixIt
	'This line can come out once the bug if found that inserts undefined into the actions table
	SQLGetLastMCSActionNoteByMonthByYearByCust = SQLGetLastMCSActionNoteByMonthByYearByCust & " AND Action <> 'undefined' "
	
	SQLGetLastMCSActionNoteByMonthByYearByCust = SQLGetLastMCSActionNoteByMonthByYearByCust & " ORDER BY RecordCreationDateTime DESC"
	 
	Set rsGetLastMCSActionNoteByMonthByYearByCust = Server.CreateObject("ADODB.Recordset")
	rsGetLastMCSActionNoteByMonthByYearByCust.CursorLocation = 3 
	Set rsGetLastMCSActionNoteByMonthByYearByCust= cnnGetLastMCSActionNoteByMonthByYearByCust.Execute(SQLGetLastMCSActionNoteByMonthByYearByCust)
	
	If not rsGetLastMCSActionNoteByMonthByYearByCust.eof then resultGetLastMCSActionNoteByMonthByYearByCust =  rsGetLastMCSActionNoteByMonthByYearByCust("ActionNotes")
		
	rsGetLastMCSActionNoteByMonthByYearByCust.Close
	set rsGetLastMCSActionNoteByMonthByYearByCust= Nothing
	cnnGetLastMCSActionNoteByMonthByYearByCust.Close	
	set cnnGetLastMCSActionNoteByMonthByYearByCust = Nothing
	
	GetLastMCSActionNoteByMonthByYearByCust = resultGetLastMCSActionNoteByMonthByYearByCust 
	
End Function


Function GetLastMCSActionNoteReasonByMonthByYearByCust(passedCust, passedMonth, passedYear)

	Set cnnGetLastMCSActionNoteReasonByMonthByYearByCust = Server.CreateObject("ADODB.Connection")
	cnnGetLastMCSActionNoteReasonByMonthByYearByCust.open Session("ClientCnnString")

	resultGetLastMCSActionNoteReasonByMonthByYearByCust = ""
		
	SQLGetLastMCSActionNoteReasonByMonthByYearByCust = "SELECT TOP 1 * FROM BI_MCSActions "
	SQLGetLastMCSActionNoteReasonByMonthByYearByCust = SQLGetLastMCSActionNoteReasonByMonthByYearByCust & "WHERE CustID = '" & passedCust & "' AND "
	SQLGetLastMCSActionNoteReasonByMonthByYearByCust = SQLGetLastMCSActionNoteReasonByMonthByYearByCust & " (MONTH(RecordCreationDateTime)) = " & passedMonth & "  AND "
	SQLGetLastMCSActionNoteReasonByMonthByYearByCust = SQLGetLastMCSActionNoteReasonByMonthByYearByCust & " (YEAR(RecordCreationDateTime)) = " & passedYear 
	
	'FixIt
	'This line can come out once the bug if found that inserts undefined into the actions table
	SQLGetLastMCSActionNoteReasonByMonthByYearByCust = SQLGetLastMCSActionNoteReasonByMonthByYearByCust & " AND Action <> 'undefined' "
	
	SQLGetLastMCSActionNoteReasonByMonthByYearByCust = SQLGetLastMCSActionNoteReasonByMonthByYearByCust & " ORDER BY RecordCreationDateTime DESC"
	 
	Set rsGetLastMCSActionNoteReasonByMonthByYearByCust = Server.CreateObject("ADODB.Recordset")
	rsGetLastMCSActionNoteReasonByMonthByYearByCust.CursorLocation = 3 
	Set rsGetLastMCSActionNoteReasonByMonthByYearByCust= cnnGetLastMCSActionNoteReasonByMonthByYearByCust.Execute(SQLGetLastMCSActionNoteReasonByMonthByYearByCust)
	
	If not rsGetLastMCSActionNoteReasonByMonthByYearByCust.eof then resultGetLastMCSActionNoteReasonByMonthByYearByCust =  rsGetLastMCSActionNoteReasonByMonthByYearByCust("MCSReasonIntRecID")
		
	rsGetLastMCSActionNoteReasonByMonthByYearByCust.Close
	set rsGetLastMCSActionNoteReasonByMonthByYearByCust= Nothing
	cnnGetLastMCSActionNoteReasonByMonthByYearByCust.Close	
	set cnnGetLastMCSActionNoteReasonByMonthByYearByCust = Nothing
	
	GetLastMCSActionNoteReasonByMonthByYearByCust = resultGetLastMCSActionNoteReasonByMonthByYearByCust 
	
End Function


Function TotalUNPostedLVFByCustByMonthByYear(passedCustID,passedMonth,passedYear)

	resultTotalUNPostedLVFByCustByMonthByYear = ""

	Set cnnTotalUNPostedLVFByCustByMonthByYear = Server.CreateObject("ADODB.Connection")
	cnnTotalUNPostedLVFByCustByMonthByYear.open Session("ClientCnnString")
	
	SQLTotalUNPostedLVFByCustByMonthByYear  = "SELECT Sum(ExtendedPrice) AS TotalUnPostedLVF FROM TelselParts "	
	SQLTotalUNPostedLVFByCustByMonthByYear = SQLTotalUNPostedLVFByCustByMonthByYear & " INNER JOIN TelSel ON TelSelParts.InvoiceNo = TelSel.InvoiceNo "
	SQLTotalUNPostedLVFByCustByMonthByYear = SQLTotalUNPostedLVFByCustByMonthByYear & " WHERE PartNumber = 'LVM' "
	SQLTotalUNPostedLVFByCustByMonthByYear = SQLTotalUNPostedLVFByCustByMonthByYear & " AND TelSelParts.CustNum = " & passedCustID 
	SQLTotalUNPostedLVFByCustByMonthByYear = SQLTotalUNPostedLVFByCustByMonthByYear & " AND Month(TelselParts.InvoiceDate) = " & passedMonth 
	SQLTotalUNPostedLVFByCustByMonthByYear = SQLTotalUNPostedLVFByCustByMonthByYear & " AND Year(TelselParts.InvoiceDate) = " & passedYear 
	SQLTotalUNPostedLVFByCustByMonthByYear = SQLTotalUNPostedLVFByCustByMonthByYear & " AND (TelSel.InvoiceTFlag = 'O' OR TelSel.InvoiceTFlag = 'T')"

'response.write(SQLTotalUNPostedLVFByCustByMonthByYear & "<br>")

	Set rsTotalUNPostedLVFByCustByMonthByYear = Server.CreateObject("ADODB.Recordset")
	rsTotalUNPostedLVFByCustByMonthByYear.CursorLocation = 3 
	Set rsTotalUNPostedLVFByCustByMonthByYear = cnnTotalUNPostedLVFByCustByMonthByYear.Execute(SQLTotalUNPostedLVFByCustByMonthByYear)

	If not rsTotalUNPostedLVFByCustByMonthByYear.EOF Then
		resultTotalUNPostedLVFByCustByMonthByYear = rsTotalUNPostedLVFByCustByMonthByYear("TotalUnPostedLVF")
	Else
		resultTotalUNPostedLVFByCustByMonthByYear = 0 ' Because there were no sales
	End If
	If Not IsNumeric(resultTotalUNPostedLVFByCustByMonthByYear) Then resultTotalUNPostedLVFByCustByMonthByYear = 0 'To account for null result

	rsTotalUNPostedLVFByCustByMonthByYear.Close
	set rsTotalUNPostedLVFByCustByMonthByYear= Nothing
	cnnTotalUNPostedLVFByCustByMonthByYear.Close	
	set cnnTotalUNPostedLVFByCustByMonthByYear= Nothing
	
	TotalUNPostedLVFByCustByMonthByYear = resultTotalUNPostedLVFByCustByMonthByYear

End Function

Function GetLastWebOrderDateFromOCSAccess(passedCustID)

	resultGetLastWebOrderDateFromOCSAccess = ""
			
	Set cnnInsight = Server.CreateObject("ADODB.Connection")
	cnnInsight.open (InsightCnnString)
	Set rsInsight = Server.CreateObject("ADODB.Recordset")
	rsInsight.CursorLocation = 3 
		
	SQLInsight = "SELECT * FROM tblServerInfo where clientKey='"& MUV_READ("CLIENTID") &"'"
	Set rsInsight = cnnInsight.Execute(SQLInsight)
	
	If NOT rsInsight.EOF Then
	
		OCSAccessCnnString = "Driver={SQL Server};Server=" & rsInsight.Fields("OCSAccess_dbServer")
		OCSAccessCnnString = OCSAccessCnnString & ";Database=" & rsInsight.Fields("OCSAccess_dbCatalog")
		OCSAccessCnnString = OCSAccessCnnString & ";Uid=" & rsInsight.Fields("OCSAccess_dbLogin")
		OCSAccessCnnString = OCSAccessCnnString & ";Pwd=" & rsInsight.Fields("OCSAccess_dbPassword") & ";"
		
		set rsInsight = Nothing
		cnnInsight.Close
	
		Set cnnOCSAccess = Server.CreateObject("ADODB.Connection")
		cnnOCSAccess.open (OCSAccessCnnString)
		Set rsOCSAccess = Server.CreateObject("ADODB.Recordset")
		rsOCSAccess.CursorLocation = 3 
		
		SQLOCSAccess = "SELECT MAX(OrderDate) As MxDate FROM tblOrders WHERE CustID ='"& passedCustID & "'"
		Set rsOCSAccess = cnnOCSAccess.Execute(SQLOCSAccess)
	
		If NOT rsOCSAccess.EOF Then
			resultGetLastWebOrderDateFromOCSAccess = rsOCSAccess("MxDate")
		End If

	End If
	
	GetLastWebOrderDateFromOCSAccess = resultGetLastWebOrderDateFromOCSAccess 

End Function

Function CustHasWebUserID(passedCustID)

	resultCustHasWebUserID = False
			
	Set cnnInsight = Server.CreateObject("ADODB.Connection")
	cnnInsight.open (InsightCnnString)
	Set rsInsight = Server.CreateObject("ADODB.Recordset")
	rsInsight.CursorLocation = 3 
		
	SQLInsight = "SELECT * FROM tblServerInfo where clientKey='"& MUV_READ("CLIENTID") &"'"
	Set rsInsight = cnnInsight.Execute(SQLInsight)
	
	If NOT rsInsight.EOF Then
	
		OCSAccessCnnString = "Driver={SQL Server};Server=" & rsInsight.Fields("OCSAccess_dbServer")
		OCSAccessCnnString = OCSAccessCnnString & ";Database=" & rsInsight.Fields("OCSAccess_dbCatalog")
		OCSAccessCnnString = OCSAccessCnnString & ";Uid=" & rsInsight.Fields("OCSAccess_dbLogin")
		OCSAccessCnnString = OCSAccessCnnString & ";Pwd=" & rsInsight.Fields("OCSAccess_dbPassword") & ";"
		
		set rsInsight = Nothing
		cnnInsight.Close
	
		Set cnnOCSAccess = Server.CreateObject("ADODB.Connection")
		cnnOCSAccess.open (OCSAccessCnnString)
		Set rsOCSAccess = Server.CreateObject("ADODB.Recordset")
		rsOCSAccess.CursorLocation = 3 
		
		SQLOCSAccess = "SELECT * FROM tblUser WHERE CustID ='"& passedCustID & "'"
		Set rsOCSAccess = cnnOCSAccess.Execute(SQLOCSAccess)
	
		If NOT rsOCSAccess.EOF Then
			resultCustHasWebUserID = True
		End If

	End If
	
	CustHasWebUserID = resultCustHasWebUserID 

End Function

Function GetCurrentPeriod_PostedTotal()

	resultGetCurrentPeriod_PostedTotal = 0

	Set cnnGetCurrentPeriod_PostedTotal = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedTotal.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedTotal = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedTotal = SQLGetCurrentPeriod_PostedTotal & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1

	Set rsGetCurrentPeriod_PostedTotal = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedTotal.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedTotal = cnnGetCurrentPeriod_PostedTotal.Execute(SQLGetCurrentPeriod_PostedTotal)

	If not rsGetCurrentPeriod_PostedTotal.EOF Then resultGetCurrentPeriod_PostedTotal = rsGetCurrentPeriod_PostedTotal("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedTotal) Then resultGetCurrentPeriod_PostedTotal = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedTotal.Close
	set rsGetCurrentPeriod_PostedTotal= Nothing
	cnnGetCurrentPeriod_PostedTotal.Close	
	set cnnGetCurrentPeriod_PostedTotal= Nothing
	
	GetCurrentPeriod_PostedTotal = resultGetCurrentPeriod_PostedTotal

End Function

Function GetCurrentPeriod_PostedTotalSls2(passedSecondarySalesman)

	resultGetCurrentPeriod_PostedTotalSls2 = 0

	Set cnnGetCurrentPeriod_PostedTotalSls2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedTotalSls2.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedTotalSls2 = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedTotalSls2 = SQLGetCurrentPeriod_PostedTotalSls2 & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedTotalSls2 = SQLGetCurrentPeriod_PostedTotalSls2 & " AND CustID IN "
	SQLGetCurrentPeriod_PostedTotalSls2 = SQLGetCurrentPeriod_PostedTotalSls2 & " (SELECT CustNum FROM AR_Customer WHERE SecondarySalesman = " & passedSecondarySalesman & ") "

	Set rsGetCurrentPeriod_PostedTotalSls2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedTotalSls2.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedTotalSls2 = cnnGetCurrentPeriod_PostedTotalSls2.Execute(SQLGetCurrentPeriod_PostedTotalSls2)

	If not rsGetCurrentPeriod_PostedTotalSls2.EOF Then resultGetCurrentPeriod_PostedTotalSls2 = rsGetCurrentPeriod_PostedTotalSls2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedTotalSls2) Then resultGetCurrentPeriod_PostedTotalSls2 = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedTotalSls2.Close
	set rsGetCurrentPeriod_PostedTotalSls2= Nothing
	cnnGetCurrentPeriod_PostedTotalSls2.Close	
	set cnnGetCurrentPeriod_PostedTotalSls2= Nothing
	
	GetCurrentPeriod_PostedTotalSls2 = resultGetCurrentPeriod_PostedTotalSls2

End Function

Function GetCurrentPeriod_PostedRentalsSls2(passedSecondarySalesman)

	resultGetCurrentPeriod_PostedRentalsSls2 = 0

	Set cnnGetCurrentPeriod_PostedRentalsSls2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedRentalsSls2.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedRentalsSls2 = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedRentalsSls2 = SQLGetCurrentPeriod_PostedRentalsSls2 & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedRentalsSls2 = SQLGetCurrentPeriod_PostedRentalsSls2 & " AND CustID IN "
	SQLGetCurrentPeriod_PostedRentalsSls2 = SQLGetCurrentPeriod_PostedRentalsSls2 & " (SELECT CustNum FROM AR_Customer WHERE SecondarySalesman = " & passedSecondarySalesman & ") "

	Set rsGetCurrentPeriod_PostedRentalsSls2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedRentalsSls2.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedRentalsSls2 = cnnGetCurrentPeriod_PostedRentalsSls2.Execute(SQLGetCurrentPeriod_PostedRentalsSls2)

	If not rsGetCurrentPeriod_PostedRentalsSls2.EOF Then resultGetCurrentPeriod_PostedRentalsSls2 = rsGetCurrentPeriod_PostedRentalsSls2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedRentalsSls2) Then resultGetCurrentPeriod_PostedRentalsSls2 = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedRentalsSls2.Close
	set rsGetCurrentPeriod_PostedRentalsSls2= Nothing
	cnnGetCurrentPeriod_PostedRentalsSls2.Close	
	set cnnGetCurrentPeriod_PostedRentalsSls2= Nothing
	
	GetCurrentPeriod_PostedRentalsSls2 = resultGetCurrentPeriod_PostedRentalsSls2

End Function


Function GetCurrentPeriod_UnPostedTotal()

	resultGetCurrentPeriod_UnPostedTotal = 0

	Set cnnGetCurrentPeriod_UnPostedTotal = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedTotal.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedTotal = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedTotal = SQLGetCurrentPeriod_UnPostedTotal & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1

	Set rsGetCurrentPeriod_UnPostedTotal = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedTotal.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedTotal = cnnGetCurrentPeriod_UnPostedTotal.Execute(SQLGetCurrentPeriod_UnPostedTotal)

	If not rsGetCurrentPeriod_UnPostedTotal.EOF Then resultGetCurrentPeriod_UnPostedTotal = rsGetCurrentPeriod_UnPostedTotal("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedTotal) Then resultGetCurrentPeriod_UnPostedTotal = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedTotal.Close
	set rsGetCurrentPeriod_UnPostedTotal= Nothing
	cnnGetCurrentPeriod_UnPostedTotal.Close	
	set cnnGetCurrentPeriod_UnPostedTotal= Nothing

	GetCurrentPeriod_UnPostedTotal = resultGetCurrentPeriod_UnPostedTotal

End Function

Function GetCurrentPeriod_UnPostedTotalSls2(passedSecondarySalesman)

	resultGetCurrentPeriod_UnPostedTotalSls2 = 0

	Set cnnGetCurrentPeriod_UnPostedTotalSls2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedTotalSls2.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedTotalSls2 = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedTotalSls2 = SQLGetCurrentPeriod_UnPostedTotalSls2 & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedTotalSls2 = SQLGetCurrentPeriod_UnPostedTotalSls2 & " AND CustID IN "
	SQLGetCurrentPeriod_UnPostedTotalSls2 = SQLGetCurrentPeriod_UnPostedTotalSls2 & " (SELECT CustNum FROM AR_Customer WHERE SecondarySalesman = " & passedSecondarySalesman & ") "
	

	Set rsGetCurrentPeriod_UnPostedTotalSls2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedTotalSls2.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedTotalSls2 = cnnGetCurrentPeriod_UnPostedTotalSls2.Execute(SQLGetCurrentPeriod_UnPostedTotalSls2)

	If not rsGetCurrentPeriod_UnPostedTotalSls2.EOF Then resultGetCurrentPeriod_UnPostedTotalSls2 = rsGetCurrentPeriod_UnPostedTotalSls2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedTotalSls2) Then resultGetCurrentPeriod_UnPostedTotalSls2 = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedTotalSls2.Close
	set rsGetCurrentPeriod_UnPostedTotalSls2= Nothing
	cnnGetCurrentPeriod_UnPostedTotalSls2.Close	
	set cnnGetCurrentPeriod_UnPostedTotalSls2= Nothing

	GetCurrentPeriod_UnPostedTotalSls2 = resultGetCurrentPeriod_UnPostedTotalSls2

End Function

Function GetCurrentPeriod_UnPostedRentalsSls2(passedSecondarySalesman)

	resultGetCurrentPeriod_UnPostedRentalsSls2 = 0

	Set cnnGetCurrentPeriod_UnPostedRentalsSls2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedRentalsSls2.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedRentalsSls2 = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedRentalsSls2 = SQLGetCurrentPeriod_UnPostedRentalsSls2 & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedRentalsSls2 = SQLGetCurrentPeriod_UnPostedRentalsSls2 & " AND CustID IN "
	SQLGetCurrentPeriod_UnPostedRentalsSls2 = SQLGetCurrentPeriod_UnPostedRentalsSls2 & " (SELECT CustNum FROM AR_Customer WHERE SecondarySalesman = " & passedSecondarySalesman & ") "
	

	Set rsGetCurrentPeriod_UnPostedRentalsSls2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedRentalsSls2.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedRentalsSls2 = cnnGetCurrentPeriod_UnPostedRentalsSls2.Execute(SQLGetCurrentPeriod_UnPostedRentalsSls2)

	If not rsGetCurrentPeriod_UnPostedRentalsSls2.EOF Then resultGetCurrentPeriod_UnPostedRentalsSls2 = rsGetCurrentPeriod_UnPostedRentalsSls2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedRentalsSls2) Then resultGetCurrentPeriod_UnPostedRentalsSls2 = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedRentalsSls2.Close
	set rsGetCurrentPeriod_UnPostedRentalsSls2= Nothing
	cnnGetCurrentPeriod_UnPostedRentalsSls2.Close	
	set cnnGetCurrentPeriod_UnPostedRentalsSls2= Nothing

	GetCurrentPeriod_UnPostedRentalsSls2 = resultGetCurrentPeriod_UnPostedRentalsSls2

End Function


Function GetCurrentPeriod_PostedRentals()

	resultGetCurrentPeriod_PostedRentals = 0

	Set cnnGetCurrentPeriod_PostedRentals = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedRentals.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedRentals = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedRentals = SQLGetCurrentPeriod_PostedRentals & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1

	Set rsGetCurrentPeriod_PostedRentals = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedRentals.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedRentals = cnnGetCurrentPeriod_PostedRentals.Execute(SQLGetCurrentPeriod_PostedRentals)

	If not rsGetCurrentPeriod_PostedRentals.EOF Then resultGetCurrentPeriod_PostedRentals = rsGetCurrentPeriod_PostedRentals("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedRentals) Then resultGetCurrentPeriod_PostedRentals = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedRentals.Close
	set rsGetCurrentPeriod_PostedRentals= Nothing
	cnnGetCurrentPeriod_PostedRentals.Close	
	set cnnGetCurrentPeriod_PostedRentals= Nothing
	
	GetCurrentPeriod_PostedRentals = resultGetCurrentPeriod_PostedRentals

End Function

Function GetCurrentPeriod_UnPostedRentals()

	resultGetCurrentPeriod_UnPostedRentals = 0

	Set cnnGetCurrentPeriod_UnPostedRentals = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedRentals.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedRentals = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedRentals = SQLGetCurrentPeriod_UnPostedRentals & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	

	Set rsGetCurrentPeriod_UnPostedRentals = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedRentals.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedRentals = cnnGetCurrentPeriod_UnPostedRentals.Execute(SQLGetCurrentPeriod_UnPostedRentals)

	If not rsGetCurrentPeriod_UnPostedRentals.EOF Then resultGetCurrentPeriod_UnPostedRentals = rsGetCurrentPeriod_UnPostedRentals("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedRentals) Then resultGetCurrentPeriod_UnPostedRentals = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedRentals.Close
	set rsGetCurrentPeriod_UnPostedRentals= Nothing
	cnnGetCurrentPeriod_UnPostedRentals.Close	
	set cnnGetCurrentPeriod_UnPostedRentals= Nothing

	GetCurrentPeriod_UnPostedRentals = resultGetCurrentPeriod_UnPostedRentals

End Function



Function GetCurrentPeriod_PostedTotalSls(passedSalesman)

	resultGetCurrentPeriod_PostedTotalSls = 0

	Set cnnGetCurrentPeriod_PostedTotalSls = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedTotalSls.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedTotalSls = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedTotalSls = SQLGetCurrentPeriod_PostedTotalSls & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedTotalSls = SQLGetCurrentPeriod_PostedTotalSls & " AND CustID IN "
	SQLGetCurrentPeriod_PostedTotalSls = SQLGetCurrentPeriod_PostedTotalSls & " (SELECT CustNum FROM AR_Customer WHERE Salesman = " & passedSalesman & ") "

	Set rsGetCurrentPeriod_PostedTotalSls = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedTotalSls.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedTotalSls = cnnGetCurrentPeriod_PostedTotalSls.Execute(SQLGetCurrentPeriod_PostedTotalSls)

	If not rsGetCurrentPeriod_PostedTotalSls.EOF Then resultGetCurrentPeriod_PostedTotalSls = rsGetCurrentPeriod_PostedTotalSls("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedTotalSls) Then resultGetCurrentPeriod_PostedTotalSls = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedTotalSls.Close
	set rsGetCurrentPeriod_PostedTotalSls= Nothing
	cnnGetCurrentPeriod_PostedTotalSls.Close	
	set cnnGetCurrentPeriod_PostedTotalSls= Nothing
	
	GetCurrentPeriod_PostedTotalSls = resultGetCurrentPeriod_PostedTotalSls

End Function


Function GetCurrentPeriod_UnPostedTotalSls(passedSalesman)

	resultGetCurrentPeriod_UnPostedTotalSls = 0

	Set cnnGetCurrentPeriod_UnPostedTotalSls = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedTotalSls.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedTotalSls = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedTotalSls = SQLGetCurrentPeriod_UnPostedTotalSls & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedTotalSls = SQLGetCurrentPeriod_UnPostedTotalSls & " AND CustID IN "
	SQLGetCurrentPeriod_UnPostedTotalSls = SQLGetCurrentPeriod_UnPostedTotalSls & " (SELECT CustNum FROM AR_Customer WHERE Salesman = " & passedSalesman & ") "
	

	Set rsGetCurrentPeriod_UnPostedTotalSls = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedTotalSls.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedTotalSls = cnnGetCurrentPeriod_UnPostedTotalSls.Execute(SQLGetCurrentPeriod_UnPostedTotalSls)

	If not rsGetCurrentPeriod_UnPostedTotalSls.EOF Then resultGetCurrentPeriod_UnPostedTotalSls = rsGetCurrentPeriod_UnPostedTotalSls("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedTotalSls) Then resultGetCurrentPeriod_UnPostedTotalSls = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedTotalSls.Close
	set rsGetCurrentPeriod_UnPostedTotalSls= Nothing
	cnnGetCurrentPeriod_UnPostedTotalSls.Close	
	set cnnGetCurrentPeriod_UnPostedTotalSls= Nothing

	GetCurrentPeriod_UnPostedTotalSls = resultGetCurrentPeriod_UnPostedTotalSls

End Function

Function GetCurrentPeriod_PostedRentalsSls(passedSalesman)

	resultGetCurrentPeriod_PostedRentalsSls = 0

	Set cnnGetCurrentPeriod_PostedRentalsSls = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedRentalsSls.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedRentalsSls = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedRentalsSls = SQLGetCurrentPeriod_PostedRentalsSls & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedRentalsSls = SQLGetCurrentPeriod_PostedRentalsSls & " AND CustID IN "
	SQLGetCurrentPeriod_PostedRentalsSls = SQLGetCurrentPeriod_PostedRentalsSls & " (SELECT CustNum FROM AR_Customer WHERE Salesman = " & passedSalesman & ") "

	Set rsGetCurrentPeriod_PostedRentalsSls = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedRentalsSls.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedRentalsSls = cnnGetCurrentPeriod_PostedRentalsSls.Execute(SQLGetCurrentPeriod_PostedRentalsSls)

	If not rsGetCurrentPeriod_PostedRentalsSls.EOF Then resultGetCurrentPeriod_PostedRentalsSls = rsGetCurrentPeriod_PostedRentalsSls("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedRentalsSls) Then resultGetCurrentPeriod_PostedRentalsSls = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedRentalsSls.Close
	set rsGetCurrentPeriod_PostedRentalsSls= Nothing
	cnnGetCurrentPeriod_PostedRentalsSls.Close	
	set cnnGetCurrentPeriod_PostedRentalsSls= Nothing
	
	GetCurrentPeriod_PostedRentalsSls = resultGetCurrentPeriod_PostedRentalsSls

End Function

Function GetCurrentPeriod_UnPostedRentalsSls(passedSalesman)

	resultGetCurrentPeriod_UnPostedRentalsSls = 0

	Set cnnGetCurrentPeriod_UnPostedRentalsSls = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedRentalsSls.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedRentalsSls = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedRentalsSls = SQLGetCurrentPeriod_UnPostedRentalsSls & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedRentalsSls = SQLGetCurrentPeriod_UnPostedRentalsSls & " AND CustID IN "
	SQLGetCurrentPeriod_UnPostedRentalsSls = SQLGetCurrentPeriod_UnPostedRentalsSls & " (SELECT CustNum FROM AR_Customer WHERE Salesman = " & passedSalesman & ") "
	

	Set rsGetCurrentPeriod_UnPostedRentalsSls = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedRentalsSls.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedRentalsSls = cnnGetCurrentPeriod_UnPostedRentalsSls.Execute(SQLGetCurrentPeriod_UnPostedRentalsSls)

	If not rsGetCurrentPeriod_UnPostedRentalsSls.EOF Then resultGetCurrentPeriod_UnPostedRentalsSls = rsGetCurrentPeriod_UnPostedRentalsSls("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedRentalsSls) Then resultGetCurrentPeriod_UnPostedRentalsSls = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedRentalsSls.Close
	set rsGetCurrentPeriod_UnPostedRentalsSls= Nothing
	cnnGetCurrentPeriod_UnPostedRentalsSls.Close	
	set cnnGetCurrentPeriod_UnPostedRentalsSls= Nothing

	GetCurrentPeriod_UnPostedRentalsSls = resultGetCurrentPeriod_UnPostedRentalsSls

End Function

Function TotalCostByPeriodSeqPrior12P(passedPeriodSeq,passedCustID)

	resultTotalCostByPeriodSeqPrior12P = ""

	Set cnnTotalCostByPeriodSeqPrior12P = Server.CreateObject("ADODB.Connection")
	cnnTotalCostByPeriodSeqPrior12P.open Session("ClientCnnString")
		
	SQLTotalCostByPeriodSeqPrior12P = "SELECT Sum(TotalCost) AS PeriodTotCost FROM CustCatPeriodSales WHERE CustNum = '" & passedCustID & "' "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "AND (ThisPeriodSequenceNumber = " & passedPeriodSeq -1 & " OR "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -2 & " OR "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -3 & " OR "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -4 & " OR "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -5 & " OR "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -6 & " OR "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -7 & " OR "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -8 & " OR "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -9 & " OR "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -10 & " OR "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -11 & " OR "
	SQLTotalCostByPeriodSeqPrior12P = SQLTotalCostByPeriodSeqPrior12P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -12 & ") "

	Set rsTotalCostByPeriodSeqPrior12P = Server.CreateObject("ADODB.Recordset")
	rsTotalCostByPeriodSeqPrior12P.CursorLocation = 3 
	Set rsTotalCostByPeriodSeqPrior12P = cnnTotalCostByPeriodSeqPrior12P.Execute(SQLTotalCostByPeriodSeqPrior12P)

	If not rsTotalCostByPeriodSeqPrior12P.EOF Then resultTotalCostByPeriodSeqPrior12P = rsTotalCostByPeriodSeqPrior12P("PeriodTotCost")

	rsTotalCostByPeriodSeqPrior12P.Close
	set rsTotalCostByPeriodSeqPrior12P= Nothing
	cnnTotalCostByPeriodSeqPrior12P.Close	
	set cnnTotalCostByPeriodSeqPrior12P= Nothing
	
	TotalCostByPeriodSeqPrior12P = resultTotalCostByPeriodSeqPrior12P

End Function

Function TotalCostByPeriodSeqPrior3P(passedPeriodSeq,passedCustID)

	resultTotalCostByPeriodSeqPrior3P = ""

	Set cnnTotalCostByPeriodSeqPrior3P = Server.CreateObject("ADODB.Connection")
	cnnTotalCostByPeriodSeqPrior3P.open Session("ClientCnnString")
		
	SQLTotalCostByPeriodSeqPrior3P = "SELECT Sum(TotalCost) AS PeriodTotCost FROM CustCatPeriodSales WHERE CustNum = '" & passedCustID & "' "
	SQLTotalCostByPeriodSeqPrior3P = SQLTotalCostByPeriodSeqPrior3P  & "AND (ThisPeriodSequenceNumber = " & passedPeriodSeq -1 & " OR "
	SQLTotalCostByPeriodSeqPrior3P = SQLTotalCostByPeriodSeqPrior3P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -2 & " OR "
	SQLTotalCostByPeriodSeqPrior3P = SQLTotalCostByPeriodSeqPrior3P  & "ThisPeriodSequenceNumber = " & passedPeriodSeq -3 & ")"

	Set rsTotalCostByPeriodSeqPrior3P = Server.CreateObject("ADODB.Recordset")
	rsTotalCostByPeriodSeqPrior3P.CursorLocation = 3 
	Set rsTotalCostByPeriodSeqPrior3P = cnnTotalCostByPeriodSeqPrior3P.Execute(SQLTotalCostByPeriodSeqPrior3P)

	If not rsTotalCostByPeriodSeqPrior3P.EOF Then resultTotalCostByPeriodSeqPrior3P = rsTotalCostByPeriodSeqPrior3P("PeriodTotCost")

	rsTotalCostByPeriodSeqPrior3P.Close
	set rsTotalCostByPeriodSeqPrior3P= Nothing
	cnnTotalCostByPeriodSeqPrior3P.Close	
	set cnnTotalCostByPeriodSeqPrior3P= Nothing
	
	TotalCostByPeriodSeqPrior3P = resultTotalCostByPeriodSeqPrior3P

End Function

Function TotalTPLYAllCats(passedPeriodSeq,passedCustID)

	resultTotalTPLYAllCats = ""

	Set cnnTotalTPLYAllCats = Server.CreateObject("ADODB.Connection")
	cnnTotalTPLYAllCats.open Session("ClientCnnString")
		
	'SQLTotalTPLYAllCats = "SELECT SUM(TotalSales) AS TPLY FROM CustCatPeriodSales WHERE CustNum = '" & passedCustID & "' AND ThisPeriodSequenceNumber = " & passedPeriodSeq - 12
	
	SQLTotalTPLYAllCats = "SELECT Sum(ThisPeriodLastYearSales) AS TPLY FROM CustCatPeriodSales_ReportData WHERE CustNum = '" & passedCustID & "' AND ThisPeriodSequenceNumber = " & passedPeriodSeq
	
	Set rsTotalTPLYAllCats = Server.CreateObject("ADODB.Recordset")
	rsTotalTPLYAllCats.CursorLocation = 3 
	Set rsTotalTPLYAllCats = cnnTotalTPLYAllCats.Execute(SQLTotalTPLYAllCats)

	If not rsTotalTPLYAllCats.EOF Then resultTotalTPLYAllCats = rsTotalTPLYAllCats("TPLY")

	rsTotalTPLYAllCats.Close
	set rsTotalTPLYAllCats= Nothing
	cnnTotalTPLYAllCats.Close	
	set cnnTotalTPLYAllCats= Nothing
	
	TotalTPLYAllCats = resultTotalTPLYAllCats

End Function

Function TotalCostByPeriodSeq(passedPeriodSeq,passedCustID)

	resultTotalCostByPeriodSeq = ""

	Set cnnTotalCostByPeriodSeq = Server.CreateObject("ADODB.Connection")
	cnnTotalCostByPeriodSeq.open Session("ClientCnnString")
		
	SQLTotalCostByPeriodSeq = "SELECT Sum(TotalCost) AS PeriodTotCost FROM CustCatPeriodSales_ReportData WHERE CustNum = '" & passedCustID & "' AND ThisPeriodSequenceNumber = " & passedPeriodSeq
 
	Set rsTotalCostByPeriodSeq = Server.CreateObject("ADODB.Recordset")
	rsTotalCostByPeriodSeq.CursorLocation = 3 
	Set rsTotalCostByPeriodSeq = cnnTotalCostByPeriodSeq.Execute(SQLTotalCostByPeriodSeq)

	If not rsTotalCostByPeriodSeq.EOF Then resultTotalCostByPeriodSeq = rsTotalCostByPeriodSeq("PeriodTotCost")

	rsTotalCostByPeriodSeq.Close
	set rsTotalCostByPeriodSeq= Nothing
	cnnTotalCostByPeriodSeq.Close	
	set cnnTotalCostByPeriodSeq= Nothing
	
	TotalCostByPeriodSeq = resultTotalCostByPeriodSeq

End Function


Function GetCurrentPeriod_PostedRentalsByCust(passedCustomerID)

	resultGetCurrentPeriod_PostedRentalsByCust = 0

	Set cnnGetCurrentPeriod_PostedRentalsByCust = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedRentalsByCust.open Session("ClientCnnString")
		
	SQLGetCurrentPeriod_PostedRentalsByCust = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedRentalsByCust = SQLGetCurrentPeriod_PostedRentalsByCust & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedRentalsByCust = SQLGetCurrentPeriod_PostedRentalsByCust & " AND CustID = '" & passedCustomerID & "'"

	Set rsGetCurrentPeriod_PostedRentalsByCust = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedRentalsByCust.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedRentalsByCust = cnnGetCurrentPeriod_PostedRentalsByCust.Execute(SQLGetCurrentPeriod_PostedRentalsByCust)

	If not rsGetCurrentPeriod_PostedRentalsByCust.EOF Then resultGetCurrentPeriod_PostedRentalsByCust = rsGetCurrentPeriod_PostedRentalsByCust("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedRentalsByCust) Then resultGetCurrentPeriod_PostedRentalsByCust = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedRentalsByCust.Close
	set rsGetCurrentPeriod_PostedRentalsByCust= Nothing
	cnnGetCurrentPeriod_PostedRentalsByCust.Close	
	set cnnGetCurrentPeriod_PostedRentalsByCust= Nothing
	
	GetCurrentPeriod_PostedRentalsByCust = resultGetCurrentPeriod_PostedRentalsByCust

End Function


Function GetCurrentPeriod_UnPostedRentalsByCust(passedCustomerID)

	resultGetCurrentPeriod_UnPostedRentalsByCust = 0

	Set cnnGetCurrentPeriod_UnPostedRentalsByCust = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedRentalsByCust.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedRentalsByCust = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedRentalsByCust = SQLGetCurrentPeriod_UnPostedRentalsByCust & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedRentalsByCust = SQLGetCurrentPeriod_UnPostedRentalsByCust & " AND CustID = '" & passedCustomerID & "'"
	

	Set rsGetCurrentPeriod_UnPostedRentalsByCust = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedRentalsByCust.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedRentalsByCust = cnnGetCurrentPeriod_UnPostedRentalsByCust.Execute(SQLGetCurrentPeriod_UnPostedRentalsByCust)

	If not rsGetCurrentPeriod_UnPostedRentalsByCust.EOF Then resultGetCurrentPeriod_UnPostedRentalsByCust = rsGetCurrentPeriod_UnPostedRentalsByCust("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedRentalsByCust) Then resultGetCurrentPeriod_UnPostedRentalsByCust = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedRentalsByCust.Close
	set rsGetCurrentPeriod_UnPostedRentalsByCust= Nothing
	cnnGetCurrentPeriod_UnPostedRentalsByCust.Close	
	set cnnGetCurrentPeriod_UnPostedRentalsByCust= Nothing

	GetCurrentPeriod_UnPostedRentalsByCust = resultGetCurrentPeriod_UnPostedRentalsByCust

End Function

Function GetCurrentPeriod_UnPostedProdSalesByCust(passedCustomerID)

	resultGetCurrentPeriod_UnPostedProdSalesByCust = 0

	Set cnnGetCurrentPeriod_UnPostedProdSalesByCust = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedProdSalesByCust.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedProdSalesByCust = "SELECT SUM(CASE WHEN CategoryID <> 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedProdSalesByCust = SQLGetCurrentPeriod_UnPostedProdSalesByCust & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedProdSalesByCust = SQLGetCurrentPeriod_UnPostedProdSalesByCust & " AND CustID = '" & passedCustomerID & "'"

	

	Set rsGetCurrentPeriod_UnPostedProdSalesByCust = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedProdSalesByCust.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedProdSalesByCust = cnnGetCurrentPeriod_UnPostedProdSalesByCust.Execute(SQLGetCurrentPeriod_UnPostedProdSalesByCust)

	If not rsGetCurrentPeriod_UnPostedProdSalesByCust.EOF Then resultGetCurrentPeriod_UnPostedProdSalesByCust = rsGetCurrentPeriod_UnPostedProdSalesByCust("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedProdSalesByCust) Then resultGetCurrentPeriod_UnPostedProdSalesByCust = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedProdSalesByCust.Close
	set rsGetCurrentPeriod_UnPostedProdSalesByCust= Nothing
	cnnGetCurrentPeriod_UnPostedProdSalesByCust.Close	
	set cnnGetCurrentPeriod_UnPostedProdSalesByCust= Nothing

	GetCurrentPeriod_UnPostedProdSalesByCust = resultGetCurrentPeriod_UnPostedProdSalesByCust

End Function

Function GetCurrentPeriod_PostedProdSalesByCust(passedCustomerID)


	resultGetCurrentPeriod_PostedProdSalesByCust = 0

	Set cnnGetCurrentPeriod_PostedProdSalesByCust = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedProdSalesByCust.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedProdSalesByCust = "SELECT SUM(CASE WHEN CategoryID <> 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedProdSalesByCust = SQLGetCurrentPeriod_PostedProdSalesByCust & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedProdSalesByCust = SQLGetCurrentPeriod_PostedProdSalesByCust & " AND CustID = '" & passedCustomerID & "'"

'Response.Write(SQLGetCurrentPeriod_PostedProdSalesByCust)

	Set rsGetCurrentPeriod_PostedProdSalesByCust = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedProdSalesByCust.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedProdSalesByCust = cnnGetCurrentPeriod_PostedProdSalesByCust.Execute(SQLGetCurrentPeriod_PostedProdSalesByCust)

	If not rsGetCurrentPeriod_PostedProdSalesByCust.EOF Then resultGetCurrentPeriod_PostedProdSalesByCust = rsGetCurrentPeriod_PostedProdSalesByCust("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedProdSalesByCust) Then resultGetCurrentPeriod_PostedProdSalesByCust = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedProdSalesByCust.Close
	set rsGetCurrentPeriod_PostedProdSalesByCust= Nothing
	cnnGetCurrentPeriod_PostedProdSalesByCust.Close	
	set cnnGetCurrentPeriod_PostedProdSalesByCust= Nothing
	
	GetCurrentPeriod_PostedProdSalesByCust = resultGetCurrentPeriod_PostedProdSalesByCust

End Function

Function GetCurrentPeriod_PostedTotalCustTyp(passedCustType)

	resultGetCurrentPeriod_PostedTotalCustTyp = 0

	Set cnnGetCurrentPeriod_PostedTotalCustTyp = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedTotalCustTyp.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedTotalCustTyp = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedTotalCustTyp = SQLGetCurrentPeriod_PostedTotalCustTyp & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedTotalCustTyp = SQLGetCurrentPeriod_PostedTotalCustTyp & " AND CustID IN "
	SQLGetCurrentPeriod_PostedTotalCustTyp = SQLGetCurrentPeriod_PostedTotalCustTyp & " (SELECT CustNum FROM AR_Customer WHERE CustType = " & passedCustType & ") "

	Set rsGetCurrentPeriod_PostedTotalCustTyp = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedTotalCustTyp.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedTotalCustTyp = cnnGetCurrentPeriod_PostedTotalCustTyp.Execute(SQLGetCurrentPeriod_PostedTotalCustTyp)

	If not rsGetCurrentPeriod_PostedTotalCustTyp.EOF Then resultGetCurrentPeriod_PostedTotalCustTyp = rsGetCurrentPeriod_PostedTotalCustTyp("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedTotalCustTyp) Then resultGetCurrentPeriod_PostedTotalCustTyp = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedTotalCustTyp.Close
	set rsGetCurrentPeriod_PostedTotalCustTyp= Nothing
	cnnGetCurrentPeriod_PostedTotalCustTyp.Close	
	set cnnGetCurrentPeriod_PostedTotalCustTyp= Nothing
	
	GetCurrentPeriod_PostedTotalCustTyp = resultGetCurrentPeriod_PostedTotalCustTyp

End Function


Function GetCurrentPeriod_UnPostedTotalCustType(passedCustType)

	resultGetCurrentPeriod_UnPostedTotalCustType = 0

	Set cnnGetCurrentPeriod_UnPostedTotalCustType = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedTotalCustType.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedTotalCustType = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedTotalCustType = SQLGetCurrentPeriod_UnPostedTotalCustType & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedTotalCustType = SQLGetCurrentPeriod_UnPostedTotalCustType & " AND CustID IN "
	SQLGetCurrentPeriod_UnPostedTotalCustType = SQLGetCurrentPeriod_UnPostedTotalCustType & " (SELECT CustNum FROM AR_Customer WHERE CustType = " & passedCustType & ") "
	

	Set rsGetCurrentPeriod_UnPostedTotalCustType = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedTotalCustType.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedTotalCustType = cnnGetCurrentPeriod_UnPostedTotalCustType.Execute(SQLGetCurrentPeriod_UnPostedTotalCustType)

	If not rsGetCurrentPeriod_UnPostedTotalCustType.EOF Then resultGetCurrentPeriod_UnPostedTotalCustType = rsGetCurrentPeriod_UnPostedTotalCustType("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedTotalCustType) Then resultGetCurrentPeriod_UnPostedTotalCustType = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedTotalCustType.Close
	set rsGetCurrentPeriod_UnPostedTotalCustType= Nothing
	cnnGetCurrentPeriod_UnPostedTotalCustType.Close	
	set cnnGetCurrentPeriod_UnPostedTotalCustType= Nothing

	GetCurrentPeriod_UnPostedTotalCustType = resultGetCurrentPeriod_UnPostedTotalCustType

End Function

Function GetCurrentPeriod_PostedRentalsCustType(passedCustType)

	resultGetCurrentPeriod_PostedRentalsCustType = 0

	Set cnnGetCurrentPeriod_PostedRentalsCustType = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedRentalsCustType.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedRentalsCustType = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedRentalsCustType = SQLGetCurrentPeriod_PostedRentalsCustType & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedRentalsCustType = SQLGetCurrentPeriod_PostedRentalsCustType & " AND CustID IN "
	SQLGetCurrentPeriod_PostedRentalsCustType = SQLGetCurrentPeriod_PostedRentalsCustType & " (SELECT CustNum FROM AR_Customer WHERE CustType = " & passedCustType & ") "

	Set rsGetCurrentPeriod_PostedRentalsCustType = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedRentalsCustType.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedRentalsCustType = cnnGetCurrentPeriod_PostedRentalsCustType.Execute(SQLGetCurrentPeriod_PostedRentalsCustType)

	If not rsGetCurrentPeriod_PostedRentalsCustType.EOF Then resultGetCurrentPeriod_PostedRentalsCustType = rsGetCurrentPeriod_PostedRentalsCustType("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedRentalsCustType) Then resultGetCurrentPeriod_PostedRentalsCustType = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedRentalsCustType.Close
	set rsGetCurrentPeriod_PostedRentalsCustType= Nothing
	cnnGetCurrentPeriod_PostedRentalsCustType.Close	
	set cnnGetCurrentPeriod_PostedRentalsCustType= Nothing
	
	GetCurrentPeriod_PostedRentalsCustType = resultGetCurrentPeriod_PostedRentalsCustType

End Function


Function GetCurrentPeriod_UnPostedRentalsCustType(passedCustType)

	resultGetCurrentPeriod_UnPostedRentalsCustType = 0

	Set cnnGetCurrentPeriod_UnPostedRentalsCustType = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedRentalsCustType.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedRentalsCustType = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedRentalsCustType = SQLGetCurrentPeriod_UnPostedRentalsCustType & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedRentalsCustType = SQLGetCurrentPeriod_UnPostedRentalsCustType & " AND CustID IN "
	SQLGetCurrentPeriod_UnPostedRentalsCustType = SQLGetCurrentPeriod_UnPostedRentalsCustType & " (SELECT CustNum FROM AR_Customer WHERE CustType = " & passedCustType & ") "
	

	Set rsGetCurrentPeriod_UnPostedRentalsCustType = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedRentalsCustType.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedRentalsCustType = cnnGetCurrentPeriod_UnPostedRentalsCustType.Execute(SQLGetCurrentPeriod_UnPostedRentalsCustType)

	If not rsGetCurrentPeriod_UnPostedRentalsCustType.EOF Then resultGetCurrentPeriod_UnPostedRentalsCustType = rsGetCurrentPeriod_UnPostedRentalsCustType("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedRentalsCustType) Then resultGetCurrentPeriod_UnPostedRentalsCustType = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedRentalsCustType.Close
	set rsGetCurrentPeriod_UnPostedRentalsCustType= Nothing
	cnnGetCurrentPeriod_UnPostedRentalsCustType.Close	
	set cnnGetCurrentPeriod_UnPostedRentalsCustType= Nothing

	GetCurrentPeriod_UnPostedRentalsCustType = resultGetCurrentPeriod_UnPostedRentalsCustType

End Function

Function GetCurrentPeriod_PostedRentalsCustType(passedCustType)

	resultGetCurrentPeriod_PostedRentalsCustType = 0

	Set cnnGetCurrentPeriod_PostedRentalsCustType = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedRentalsCustType.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedRentalsCustType = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedRentalsCustType = SQLGetCurrentPeriod_PostedRentalsCustType & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedRentalsCustType = SQLGetCurrentPeriod_PostedRentalsCustType & " AND CustID IN "
	SQLGetCurrentPeriod_PostedRentalsCustType = SQLGetCurrentPeriod_PostedRentalsCustType & " (SELECT CustNum FROM AR_Customer WHERE CustType = " & passedCustType & ") "

	Set rsGetCurrentPeriod_PostedRentalsCustType = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedRentalsCustType.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedRentalsCustType = cnnGetCurrentPeriod_PostedRentalsCustType.Execute(SQLGetCurrentPeriod_PostedRentalsCustType)

	If not rsGetCurrentPeriod_PostedRentalsCustType.EOF Then resultGetCurrentPeriod_PostedRentalsCustType = rsGetCurrentPeriod_PostedRentalsCustType("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedRentalsCustType) Then resultGetCurrentPeriod_PostedRentalsCustType = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedRentalsCustType.Close
	set rsGetCurrentPeriod_PostedRentalsCustType= Nothing
	cnnGetCurrentPeriod_PostedRentalsCustType.Close	
	set cnnGetCurrentPeriod_PostedRentalsCustType= Nothing
	
	GetCurrentPeriod_PostedRentalsCustType = resultGetCurrentPeriod_PostedRentalsCustType

End Function

Function GetCurrentPeriod_UnPostedRentalsCustType(passedCustType)

	resultGetCurrentPeriod_UnPostedRentalsCustType = 0

	Set cnnGetCurrentPeriod_UnPostedRentalsCustType = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedRentalsCustType.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedRentalsCustType = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedRentalsCustType = SQLGetCurrentPeriod_UnPostedRentalsCustType & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedRentalsCustType = SQLGetCurrentPeriod_UnPostedRentalsCustType & " AND CustID IN "
	SQLGetCurrentPeriod_UnPostedRentalsCustType = SQLGetCurrentPeriod_UnPostedRentalsCustType & " (SELECT CustNum FROM AR_Customer WHERE CustType = " & passedCustType & ") "
	

	Set rsGetCurrentPeriod_UnPostedRentalsCustType = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedRentalsCustType.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedRentalsCustType = cnnGetCurrentPeriod_UnPostedRentalsCustType.Execute(SQLGetCurrentPeriod_UnPostedRentalsCustType)

	If not rsGetCurrentPeriod_UnPostedRentalsCustType.EOF Then resultGetCurrentPeriod_UnPostedRentalsCustType = rsGetCurrentPeriod_UnPostedRentalsCustType("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedRentalsCustType) Then resultGetCurrentPeriod_UnPostedRentalsCustType = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedRentalsCustType.Close
	set rsGetCurrentPeriod_UnPostedRentalsCustType= Nothing
	cnnGetCurrentPeriod_UnPostedRentalsCustType.Close	
	set cnnGetCurrentPeriod_UnPostedRentalsCustType= Nothing

	GetCurrentPeriod_UnPostedRentalsCustType = resultGetCurrentPeriod_UnPostedRentalsCustType

End Function

Function GetCurrentPeriod_PostedTotalCustType(passedCustType)

	resultGetCurrentPeriod_PostedTotalCustType = 0

	Set cnnGetCurrentPeriod_PostedTotalCustType = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedTotalCustType.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedTotalCustType = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedTotalCustType = SQLGetCurrentPeriod_PostedTotalCustType & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedTotalCustType = SQLGetCurrentPeriod_PostedTotalCustType & " AND CustID IN "
	SQLGetCurrentPeriod_PostedTotalCustType = SQLGetCurrentPeriod_PostedTotalCustType & " (SELECT CustNum FROM AR_Customer WHERE CustType = " & passedCustType & ") "

	Set rsGetCurrentPeriod_PostedTotalCustType = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedTotalCustType.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedTotalCustType = cnnGetCurrentPeriod_PostedTotalCustType.Execute(SQLGetCurrentPeriod_PostedTotalCustType)

	If not rsGetCurrentPeriod_PostedTotalCustType.EOF Then resultGetCurrentPeriod_PostedTotalCustType = rsGetCurrentPeriod_PostedTotalCustType("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedTotalCustType) Then resultGetCurrentPeriod_PostedTotalCustType = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedTotalCustType.Close
	set rsGetCurrentPeriod_PostedTotalCustType= Nothing
	cnnGetCurrentPeriod_PostedTotalCustType.Close	
	set cnnGetCurrentPeriod_PostedTotalCustType= Nothing
	
	GetCurrentPeriod_PostedTotalCustType = resultGetCurrentPeriod_PostedTotalCustType

End Function


Function GetCurrentPeriod_UnPostedTotalCustType(passedCustType)

	resultGetCurrentPeriod_UnPostedTotalCustType = 0

	Set cnnGetCurrentPeriod_UnPostedTotalCustType = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedTotalCustType.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedTotalCustType = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedTotalCustType = SQLGetCurrentPeriod_UnPostedTotalCustType & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedTotalCustType = SQLGetCurrentPeriod_UnPostedTotalCustType & " AND CustID IN "
	SQLGetCurrentPeriod_UnPostedTotalCustType = SQLGetCurrentPeriod_UnPostedTotalCustType & " (SELECT CustNum FROM AR_Customer WHERE CustType = " & passedCustType & ") "
	

	Set rsGetCurrentPeriod_UnPostedTotalCustType = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedTotalCustType.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedTotalCustType = cnnGetCurrentPeriod_UnPostedTotalCustType.Execute(SQLGetCurrentPeriod_UnPostedTotalCustType)

	If not rsGetCurrentPeriod_UnPostedTotalCustType.EOF Then resultGetCurrentPeriod_UnPostedTotalCustType = rsGetCurrentPeriod_UnPostedTotalCustType("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedTotalCustType) Then resultGetCurrentPeriod_UnPostedTotalCustType = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedTotalCustType.Close
	set rsGetCurrentPeriod_UnPostedTotalCustType= Nothing
	cnnGetCurrentPeriod_UnPostedTotalCustType.Close	
	set cnnGetCurrentPeriod_UnPostedTotalCustType= Nothing

	GetCurrentPeriod_UnPostedTotalCustType = resultGetCurrentPeriod_UnPostedTotalCustType

End Function

Function GetCurrentPeriod_PostedTotalreferralDesc2(passedReferralDesc2)

	resultGetCurrentPeriod_PostedTotalreferralDesc2 = 0

	Set cnnGetCurrentPeriod_PostedTotalreferralDesc2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedTotalreferralDesc2.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedTotalreferralDesc2 = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedTotalreferralDesc2 = SQLGetCurrentPeriod_PostedTotalreferralDesc2 & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedTotalreferralDesc2 = SQLGetCurrentPeriod_PostedTotalreferralDesc2 & " AND ReferralDesc2 = '" & passedReferralDesc2 & "'"

	Set rsGetCurrentPeriod_PostedTotalreferralDesc2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedTotalreferralDesc2.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedTotalreferralDesc2 = cnnGetCurrentPeriod_PostedTotalreferralDesc2.Execute(SQLGetCurrentPeriod_PostedTotalreferralDesc2)

	If not rsGetCurrentPeriod_PostedTotalreferralDesc2.EOF Then resultGetCurrentPeriod_PostedTotalreferralDesc2 = rsGetCurrentPeriod_PostedTotalreferralDesc2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedTotalreferralDesc2) Then resultGetCurrentPeriod_PostedTotalreferralDesc2 = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedTotalreferralDesc2.Close
	set rsGetCurrentPeriod_PostedTotalreferralDesc2= Nothing
	cnnGetCurrentPeriod_PostedTotalreferralDesc2.Close	
	set cnnGetCurrentPeriod_PostedTotalreferralDesc2= Nothing
	
	GetCurrentPeriod_PostedTotalreferralDesc2 = resultGetCurrentPeriod_PostedTotalreferralDesc2

End Function


Function GetCurrentPeriod_UnPostedTotalReferralDesc2(passedReferralDesc2)

	resultGetCurrentPeriod_UnPostedTotalReferralDesc2 = 0

	Set cnnGetCurrentPeriod_UnPostedTotalReferralDesc2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedTotalReferralDesc2.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedTotalReferralDesc2 = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedTotalReferralDesc2 = SQLGetCurrentPeriod_UnPostedTotalReferralDesc2 & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedTotalReferralDesc2 = SQLGetCurrentPeriod_UnPostedTotalReferralDesc2 & " AND ReferralDesc2 = '" & passedReferralDesc2 & "'"	

	Set rsGetCurrentPeriod_UnPostedTotalReferralDesc2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedTotalReferralDesc2.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedTotalReferralDesc2 = cnnGetCurrentPeriod_UnPostedTotalReferralDesc2.Execute(SQLGetCurrentPeriod_UnPostedTotalReferralDesc2)

	If not rsGetCurrentPeriod_UnPostedTotalReferralDesc2.EOF Then resultGetCurrentPeriod_UnPostedTotalReferralDesc2 = rsGetCurrentPeriod_UnPostedTotalReferralDesc2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedTotalReferralDesc2) Then resultGetCurrentPeriod_UnPostedTotalReferralDesc2 = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedTotalReferralDesc2.Close
	set rsGetCurrentPeriod_UnPostedTotalReferralDesc2= Nothing
	cnnGetCurrentPeriod_UnPostedTotalReferralDesc2.Close	
	set cnnGetCurrentPeriod_UnPostedTotalReferralDesc2= Nothing

	GetCurrentPeriod_UnPostedTotalReferralDesc2 = resultGetCurrentPeriod_UnPostedTotalReferralDesc2

End Function

Function GetCurrentPeriod_PostedRentalsReferralDesc2(passedReferralDesc2)

	resultGetCurrentPeriod_PostedRentalsReferralDesc2 = 0

	Set cnnGetCurrentPeriod_PostedRentalsReferralDesc2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_PostedRentalsReferralDesc2.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_PostedRentalsReferralDesc2 = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_PostedRentalsReferralDesc2 = SQLGetCurrentPeriod_PostedRentalsReferralDesc2 & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_PostedRentalsReferralDesc2 = SQLGetCurrentPeriod_PostedRentalsReferralDesc2 & " AND ReferralDesc2 = '" & passedReferralDesc2 & "'"


	Set rsGetCurrentPeriod_PostedRentalsReferralDesc2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_PostedRentalsReferralDesc2.CursorLocation = 3 
	Set rsGetCurrentPeriod_PostedRentalsReferralDesc2 = cnnGetCurrentPeriod_PostedRentalsReferralDesc2.Execute(SQLGetCurrentPeriod_PostedRentalsReferralDesc2)

	If not rsGetCurrentPeriod_PostedRentalsReferralDesc2.EOF Then resultGetCurrentPeriod_PostedRentalsReferralDesc2 = rsGetCurrentPeriod_PostedRentalsReferralDesc2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_PostedRentalsReferralDesc2) Then resultGetCurrentPeriod_PostedRentalsReferralDesc2 = 0 ' In case there are no results
	
	rsGetCurrentPeriod_PostedRentalsReferralDesc2.Close
	set rsGetCurrentPeriod_PostedRentalsReferralDesc2= Nothing
	cnnGetCurrentPeriod_PostedRentalsReferralDesc2.Close	
	set cnnGetCurrentPeriod_PostedRentalsReferralDesc2= Nothing
	
	GetCurrentPeriod_PostedRentalsReferralDesc2 = resultGetCurrentPeriod_PostedRentalsReferralDesc2

End Function

Function GetCurrentPeriod_UnPostedRentalsReferralDesc2(passedReferralDesc2)

	resultGetCurrentPeriod_UnPostedRentalsReferralDesc2 = 0

	Set cnnGetCurrentPeriod_UnPostedRentalsReferralDesc2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrentPeriod_UnPostedRentalsReferralDesc2.open Session("ClientCnnString")
		

	SQLGetCurrentPeriod_UnPostedRentalsReferralDesc2 = "SELECT SUM(CASE WHEN CategoryID = 0 THEN TotalSales END) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrentPeriod_UnPostedRentalsReferralDesc2 = SQLGetCurrentPeriod_UnPostedRentalsReferralDesc2 & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & GetLastClosedPeriodSeqNum() + 1
	SQLGetCurrentPeriod_UnPostedRentalsReferralDesc2 = SQLGetCurrentPeriod_UnPostedRentalsReferralDesc2 & " AND ReferralDesc2 = '" & passedReferralDesc2 & "'"

	

	Set rsGetCurrentPeriod_UnPostedRentalsReferralDesc2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrentPeriod_UnPostedRentalsReferralDesc2.CursorLocation = 3 
	Set rsGetCurrentPeriod_UnPostedRentalsReferralDesc2 = cnnGetCurrentPeriod_UnPostedRentalsReferralDesc2.Execute(SQLGetCurrentPeriod_UnPostedRentalsReferralDesc2)

	If not rsGetCurrentPeriod_UnPostedRentalsReferralDesc2.EOF Then resultGetCurrentPeriod_UnPostedRentalsReferralDesc2 = rsGetCurrentPeriod_UnPostedRentalsReferralDesc2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrentPeriod_UnPostedRentalsReferralDesc2) Then resultGetCurrentPeriod_UnPostedRentalsReferralDesc2 = 0 ' In case there are no results
	
	rsGetCurrentPeriod_UnPostedRentalsReferralDesc2.Close
	set rsGetCurrentPeriod_UnPostedRentalsReferralDesc2= Nothing
	cnnGetCurrentPeriod_UnPostedRentalsReferralDesc2.Close	
	set cnnGetCurrentPeriod_UnPostedRentalsReferralDesc2= Nothing

	GetCurrentPeriod_UnPostedRentalsReferralDesc2 = resultGetCurrentPeriod_UnPostedRentalsReferralDesc2

End Function

%>