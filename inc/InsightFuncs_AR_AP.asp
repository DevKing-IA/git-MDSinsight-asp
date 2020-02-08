<%
'********************************
'List of all the functions & subs
'********************************
'Func NumberOfARCustAccounts()
'Func NumberOfActiveARCustAccounts()
'Func NumberOfInactiveARCustAccounts()
'Func NumberOfARCustAccountsNotX()
'Func NumberOfARCustAccountBillToLocationsByCustID(passedCustID)
'Func NumberOfARCustAccountShipToLocationsByCustID(passedCustID)
'Func NumberOfARCustAccountGeneralNotesByCustID(passedCustID)
'Func NumberOfARCustAccountAllNotesByCustID(passedCustID)
'Func NumberOfARCustContactsByCustID(passedCustID)
'Func NumberOfCustomersWithOrdersThisMonth()
'Func NumberOfCustomersWithOrdersPassedMonthYear(passedMonth, passedYear)
'Func NumberOfCustAccountsDefinedForPartner(passedPartnerIntRecID)
'Func NumberOfCustomersWithClassCode(passedCustClassCode)
'Func GetCustClassDescByIntRecID(passedIntRecID)
'Func GetClassCodeByIntRecID(passedIntRecID)
'Func GetCustClassByCustID(passedCustID)
'Func GetCustAbbrvNameByCustID(passedCustID)
'Func GetCustMESByCustID(passedCustID)
'Func GetCustMCSByCustID(passedCustID)
'Func GetNumberOfInHistDetailLinesByInvoiceNumber(passedInvoiceNumber)
'Func GetInvoiceExportedToSageLastDate(passedInvoiceID)
'Func GetCustRegionByCustID(passedCustID)
'Func GetCustRegionIntRecIDByCustID(passedCustID)
'Func GetRegionNameByRegionIntRecID(passedRegionIntRecID)
'Func GetQtyCustByRegion()
'Func GetCustIntRecIDByARCustContactIntRecID(passedCustID)
'Func GetCustIDByCustIntRecID(passedCustIntRecID)
'Func GetCustIntRecIDByCustID(passedCustID)
'Func GetPaymentMethodByIntRecID(passedIntRecID)
'Func GetCustNoteTypeByNoteIntRecID(passedNoteIntRecID)
'Func GetCustNoteCountByNoteType(passedNoteIntRecID, passedCustID, passedUserNo)
'Func GetCustNoteCountByNoteTypeAllUsers(passedNoteIntRecID, passedCustID)
'Func GetCustNoteCountByNoteTypeJustMe(passedNoteIntRecID, passedCustID)
'Func GetCustNoteTypeCanBeEdited(passedNoteIntRecID)
'Func HasNoteTypeBeenViewedByUser(passedCustNum,passedNoteTypeIntRecID)
'Func UserHasAnyUnviewedNotes(passedCustNum)
'Sub MarkNewNoteNoteTypeForUserAsRead(passedNoteTypeIntRecID, passedCustID)
'Func GetCustRefDesc2ByReferralCode(passedReferralCode)
'************************************
'End List of all the functions & subs
'************************************

Function NumberOfARCustAccounts()

	Set cnnNumberOfARCustAccounts  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfARCustAccounts.open Session("ClientCnnString")

	resultNumberOfARCustAccounts = 0
		
	SQLNumberOfARCustAccounts  = "SELECT COUNT(*) AS CUSTCOUNT FROM AR_Customer"
	 
	Set rsNumberOfARCustAccounts  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfARCustAccounts.CursorLocation = 3 
	
	rsNumberOfARCustAccounts.Open SQLNumberOfARCustAccounts,cnnNumberOfARCustAccounts 
			
	resultNumberOfARCustAccounts = rsNumberOfARCustAccounts("CUSTCOUNT")
	
	rsNumberOfARCustAccounts.Close
	set rsNumberOfARCustAccounts = Nothing
	cnnNumberOfARCustAccounts.Close	
	set cnnNumberOfARCustAccounts = Nothing
	
	NumberOfARCustAccounts = resultNumberOfARCustAccounts
	
End Function


Function NumberOfActiveARCustAccounts()

	Set cnnNumberOfActiveARCustAccounts  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfActiveARCustAccounts.open Session("ClientCnnString")

	resultNumberOfActiveARCustAccounts = 0
		
	SQLNumberOfActiveARCustAccounts  = "SELECT COUNT(*) AS CUSTCOUNT FROM AR_Customer WHERE AcctStatus = 'A'"
	 
	Set rsNumberOfActiveARCustAccounts  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfActiveARCustAccounts.CursorLocation = 3 
	
	rsNumberOfActiveARCustAccounts.Open SQLNumberOfActiveARCustAccounts,cnnNumberOfActiveARCustAccounts 
			
	resultNumberOfActiveARCustAccounts = rsNumberOfActiveARCustAccounts("CUSTCOUNT")
	
	rsNumberOfActiveARCustAccounts.Close
	set rsNumberOfActiveARCustAccounts = Nothing
	cnnNumberOfActiveARCustAccounts.Close	
	set cnnNumberOfActiveARCustAccounts = Nothing
	
	NumberOfActiveARCustAccounts = resultNumberOfActiveARCustAccounts
	
End Function


Function NumberOfInactiveARCustAccounts()

	Set cnnNumberOfInactiveARCustAccounts  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfInactiveARCustAccounts.open Session("ClientCnnString")

	resultNumberOfInactiveARCustAccounts = 0
		
	SQLNumberOfInactiveARCustAccounts  = "SELECT COUNT(*) AS CUSTCOUNT FROM AR_Customer WHERE AcctStatus = 'I'"
	 
	Set rsNumberOfInactiveARCustAccounts  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfInactiveARCustAccounts.CursorLocation = 3 
	
	rsNumberOfInactiveARCustAccounts.Open SQLNumberOfInactiveARCustAccounts,cnnNumberOfInactiveARCustAccounts 
			
	resultNumberOfInactiveARCustAccounts = rsNumberOfInactiveARCustAccounts("CUSTCOUNT")
	
	rsNumberOfInactiveARCustAccounts.Close
	set rsNumberOfInactiveARCustAccounts = Nothing
	cnnNumberOfInactiveARCustAccounts.Close	
	set cnnNumberOfInactiveARCustAccounts = Nothing
	
	NumberOfInactiveARCustAccounts = resultNumberOfInactiveARCustAccounts
	
End Function


Function NumberOfARCustAccountsNotX()

	Set cnnNumberOfARCustAccountsNotX  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfARCustAccountsNotX.open Session("ClientCnnString")

	resultNumberOfARCustAccountsNotX = 0
		
	SQLNumberOfARCustAccountsNotX  = "SELECT COUNT(*) AS CUSTCOUNT FROM AR_Customer WHERE AcctStatus <> 'X'"
	 
	Set rsNumberOfARCustAccountsNotX  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfARCustAccountsNotX.CursorLocation = 3 
	
	rsNumberOfARCustAccountsNotX.Open SQLNumberOfARCustAccountsNotX,cnnNumberOfARCustAccountsNotX 
			
	resultNumberOfARCustAccountsNotX = rsNumberOfARCustAccountsNotX("CUSTCOUNT")
	
	rsNumberOfARCustAccountsNotX.Close
	set rsNumberOfARCustAccountsNotX = Nothing
	cnnNumberOfARCustAccountsNotX.Close	
	set cnnNumberOfARCustAccountsNotX = Nothing
	
	NumberOfARCustAccountsNotX = resultNumberOfARCustAccountsNotX
	
End Function

Function NumberOfARCustAccountBillToLocationsByCustID(passedCustID)

	Set cnnNumberOfARCustAccountBillToLocationsByCustID  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfARCustAccountBillToLocationsByCustID.open Session("ClientCnnString")

	resultNumberOfARCustAccountBillToLocationsByCustID = 0
		
	SQLNumberOfARCustAccountBillToLocationsByCustID  = "SELECT COUNT(*) AS BILLTOCOUNT FROM AR_CustomerBillTo WHERE CustNum = '" & passedCustID & "'"
	 
	Set rsNumberOfARCustAccountBillToLocationsByCustID  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfARCustAccountBillToLocationsByCustID.CursorLocation = 3 
	
	rsNumberOfARCustAccountBillToLocationsByCustID.Open SQLNumberOfARCustAccountBillToLocationsByCustID,cnnNumberOfARCustAccountBillToLocationsByCustID 
			
	resultNumberOfARCustAccountBillToLocationsByCustID = rsNumberOfARCustAccountBillToLocationsByCustID("BILLTOCOUNT")
	
	rsNumberOfARCustAccountBillToLocationsByCustID.Close
	set rsNumberOfARCustAccountBillToLocationsByCustID = Nothing
	cnnNumberOfARCustAccountBillToLocationsByCustID.Close	
	set cnnNumberOfARCustAccountBillToLocationsByCustID = Nothing
	
	NumberOfARCustAccountBillToLocationsByCustID = resultNumberOfARCustAccountBillToLocationsByCustID
	
End Function

Function NumberOfARCustAccountShipToLocationsByCustID(passedCustID)

	Set cnnNumberOfARCustAccountShipToLocationsByCustID  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfARCustAccountShipToLocationsByCustID.open Session("ClientCnnString")

	resultNumberOfARCustAccountShipToLocationsByCustID = 0
		
	SQLNumberOfARCustAccountShipToLocationsByCustID  = "SELECT COUNT(*) AS ShipToCOUNT FROM AR_CustomerShipTo WHERE CustNum = '" & passedCustID & "'"
	 
	Set rsNumberOfARCustAccountShipToLocationsByCustID  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfARCustAccountShipToLocationsByCustID.CursorLocation = 3 
	
	rsNumberOfARCustAccountShipToLocationsByCustID.Open SQLNumberOfARCustAccountShipToLocationsByCustID,cnnNumberOfARCustAccountShipToLocationsByCustID 
			
	resultNumberOfARCustAccountShipToLocationsByCustID = rsNumberOfARCustAccountShipToLocationsByCustID("ShipToCOUNT")
	
	rsNumberOfARCustAccountShipToLocationsByCustID.Close
	set rsNumberOfARCustAccountShipToLocationsByCustID = Nothing
	cnnNumberOfARCustAccountShipToLocationsByCustID.Close	
	set cnnNumberOfARCustAccountShipToLocationsByCustID = Nothing
	
	NumberOfARCustAccountShipToLocationsByCustID = resultNumberOfARCustAccountShipToLocationsByCustID
	
End Function


Function NumberOfARCustAccountGeneralNotesByCustID(passedCustID)

	Set cnnNumberOfARCustAccountGeneralNotesByCustID  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfARCustAccountGeneralNotesByCustID.open Session("ClientCnnString")

	resultNumberOfARCustAccountGeneralNotesByCustID = 0
		
	SQLNumberOfARCustAccountGeneralNotesByCustID  = "SELECT COUNT(*) AS NoteCount FROM AR_CustomerNotes WHERE CustID = '" & passedCustID & "' AND NoteType = 'General'"
	 
	Set rsNumberOfARCustAccountGeneralNotesByCustID  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfARCustAccountGeneralNotesByCustID.CursorLocation = 3 
	
	rsNumberOfARCustAccountGeneralNotesByCustID.Open SQLNumberOfARCustAccountGeneralNotesByCustID,cnnNumberOfARCustAccountGeneralNotesByCustID 
			
	resultNumberOfARCustAccountGeneralNotesByCustID = rsNumberOfARCustAccountGeneralNotesByCustID("NoteCount")
	
	rsNumberOfARCustAccountGeneralNotesByCustID.Close
	set rsNumberOfARCustAccountGeneralNotesByCustID = Nothing
	cnnNumberOfARCustAccountGeneralNotesByCustID.Close	
	set cnnNumberOfARCustAccountGeneralNotesByCustID = Nothing
	
	NumberOfARCustAccountGeneralNotesByCustID = resultNumberOfARCustAccountGeneralNotesByCustID
	
End Function




Function NumberOfARCustAccountAllNotesByCustID(passedCustID)

	Set cnnNumberOfARCustAccountAllNotesByCustID  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfARCustAccountAllNotesByCustID.open Session("ClientCnnString")

	resultNumberOfARCustAccountAllNotesByCustID = 0
		
	SQLNumberOfARCustAccountAllNotesByCustID  = "SELECT COUNT(*) AS NoteCount FROM AR_CustomerNotes WHERE CustID = '" & passedCustID & "'"
	 
	Set rsNumberOfARCustAccountAllNotesByCustID  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfARCustAccountAllNotesByCustID.CursorLocation = 3 
	
	rsNumberOfARCustAccountAllNotesByCustID.Open SQLNumberOfARCustAccountAllNotesByCustID,cnnNumberOfARCustAccountAllNotesByCustID 
			
	resultNumberOfARCustAccountAllNotesByCustID = rsNumberOfARCustAccountAllNotesByCustID("NoteCount")
	
	rsNumberOfARCustAccountAllNotesByCustID.Close
	set rsNumberOfARCustAccountAllNotesByCustID = Nothing
	cnnNumberOfARCustAccountAllNotesByCustID.Close	
	set cnnNumberOfARCustAccountAllNotesByCustID = Nothing
	
	NumberOfARCustAccountAllNotesByCustID = resultNumberOfARCustAccountAllNotesByCustID
	
End Function


Function NumberOfARCustContactsByCustID(passedCustID)

	Set cnnNumberOfARCustContactsByCustID  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfARCustContactsByCustID.open Session("ClientCnnString")

	resultNumberOfARCustContactsByCustID = 0
	
	CustomerIntRecID = GetCustIntRecIDByARCustContactIntRecID(passedCustID)
		
	SQLNumberOfARCustContactsByCustID  = "SELECT COUNT(*) AS ContactCount FROM AR_CustomerContacts WHERE CustomerIntRecID = '" & CustomerIntRecID & "'"
	 
	Set rsNumberOfARCustContactsByCustID  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfARCustContactsByCustID.CursorLocation = 3 
	
	rsNumberOfARCustContactsByCustID.Open SQLNumberOfARCustContactsByCustID,cnnNumberOfARCustContactsByCustID 
			
	resultNumberOfARCustContactsByCustID = rsNumberOfARCustContactsByCustID("ContactCount")
	
	rsNumberOfARCustContactsByCustID.Close
	set rsNumberOfARCustContactsByCustID = Nothing
	cnnNumberOfARCustContactsByCustID.Close	
	set cnnNumberOfARCustContactsByCustID = Nothing
	
	NumberOfARCustContactsByCustID = resultNumberOfARCustContactsByCustID
	
End Function



Function NumberOfCustomersWithOrdersThisMonth()

	currentMonth = Day(Now())
	currentYear = Year(Now())

	Set cnnNumberOfCustomersWithOrdersThisMonth  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfCustomersWithOrdersThisMonth.open Session("ClientCnnString")

	resultNumberOfCustomersWithOrdersThisMonth = 0
		
	SQLNumberOfCustomersWithOrdersThisMonth  = "SELECT COUNT(*) AS CUSTCOUNT FROM InvoiceHistory WHERE MONTH(IvsDate) = " & currentMonth & " AND YEAR(IvsDate) = " & currentYear
	 
	Set rsNumberOfCustomersWithOrdersThisMonth  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfCustomersWithOrdersThisMonth.CursorLocation = 3 
	
	rsNumberOfCustomersWithOrdersThisMonth.Open SQLNumberOfCustomersWithOrdersThisMonth,cnnNumberOfCustomersWithOrdersThisMonth 
			
	resultNumberOfCustomersWithOrdersThisMonth = rsNumberOfCustomersWithOrdersThisMonth("CUSTCOUNT")
	
	rsNumberOfCustomersWithOrdersThisMonth.Close
	set rsNumberOfCustomersWithOrdersThisMonth = Nothing
	cnnNumberOfCustomersWithOrdersThisMonth.Close	
	set cnnNumberOfCustomersWithOrdersThisMonth = Nothing
	
	NumberOfCustomersWithOrdersThisMonth = resultNumberOfCustomersWithOrdersThisMonth
	
End Function



Function NumberOfCustomersWithOrdersPassedMonthYear(passedMonth, passedYear)

	Set cnnNumberOfCustomersWithOrdersPassedMonthYear  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfCustomersWithOrdersPassedMonthYear.open Session("ClientCnnString")

	resultNumberOfCustomersWithOrdersPassedMonthYear = 0
		
	SQLNumberOfCustomersWithOrdersPassedMonthYear  = "SELECT COUNT(*) AS ORDERCOUNT FROM InvoiceHistory WHERE MONTH(IvsDate) = " & passedMonth & " AND YEAR(IvsDate) = " & passedYear
	 
	Set rsNumberOfCustomersWithOrdersPassedMonthYear  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfCustomersWithOrdersPassedMonthYear.CursorLocation = 3 
	
	rsNumberOfCustomersWithOrdersPassedMonthYear.Open SQLNumberOfCustomersWithOrdersPassedMonthYear,cnnNumberOfCustomersWithOrdersPassedMonthYear 
			
	resultNumberOfCustomersWithOrdersPassedMonthYear = rsNumberOfCustomersWithOrdersPassedMonthYear("ORDERCOUNT")
	
	rsNumberOfCustomersWithOrdersPassedMonthYear.Close
	set rsNumberOfCustomersWithOrdersPassedMonthYear = Nothing
	cnnNumberOfCustomersWithOrdersPassedMonthYear.Close	
	set cnnNumberOfCustomersWithOrdersPassedMonthYear = Nothing
	
	NumberOfCustomersWithOrdersPassedMonthYear = resultNumberOfCustomersWithOrdersPassedMonthYear
	
End Function



Function NumberOfCustAccountsDefinedForPartner(passedPartnerIntRecID)

	Set cnnNumberOfCustAccountsDefinedForPartnerNum  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfCustAccountsDefinedForPartnerNum.open Session("ClientCnnString")

	resultNumberOfCustAccountsDefinedForPartnerNum = 0
		
	SQLNumberOfCustAccountsDefinedForPartnerNum  = "SELECT COUNT(*) AS CUSTCOUNT FROM AR_CustomerMapping WHERE partnerRecID = " & passedPartnerIntRecID
	 
	Set rsNumberOfCustAccountsDefinedForPartnerNum  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfCustAccountsDefinedForPartnerNum.CursorLocation = 3 
	
	rsNumberOfCustAccountsDefinedForPartnerNum.Open SQLNumberOfCustAccountsDefinedForPartnerNum,cnnNumberOfCustAccountsDefinedForPartnerNum 
			
	resultNumberOfCustAccountsDefinedForPartnerNum = rsNumberOfCustAccountsDefinedForPartnerNum("CUSTCOUNT")
	
	rsNumberOfCustAccountsDefinedForPartnerNum.Close
	set rsNumberOfCustAccountsDefinedForPartnerNum = Nothing
	cnnNumberOfCustAccountsDefinedForPartnerNum.Close	
	set cnnNumberOfCustAccountsDefinedForPartnerNum = Nothing
	
	NumberOfCustAccountsDefinedForPartner = resultNumberOfCustAccountsDefinedForPartnerNum
	
End Function

Function NumberOfCustomersWithClassCode(passedCustClassCode)

	Set cnnNumberOfCustomersWithClassCode  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfCustomersWithClassCode.open Session("ClientCnnString")

	resultNumberOfCustomersWithClassCode = 0
		
	SQLNumberOfCustomersWithClassCode  = "SELECT COUNT(*) AS CUSTCOUNT FROM AR_CUSTOMER WHERE ClassCode = '" & passedCustClassCode& "'"
	 
	Set rsNumberOfCustomersWithClassCode  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfCustomersWithClassCode.CursorLocation = 3 
	
	rsNumberOfCustomersWithClassCode.Open SQLNumberOfCustomersWithClassCode,cnnNumberOfCustomersWithClassCode 
			
	resultNumberOfCustomersWithClassCode = rsNumberOfCustomersWithClassCode("CUSTCOUNT")
	
	rsNumberOfCustomersWithClassCode.Close
	set rsNumberOfCustomersWithClassCode = Nothing
	cnnNumberOfCustomersWithClassCode.Close	
	set cnnNumberOfCustomersWithClassCode = Nothing
	
	NumberOfCustomersWithClassCode = resultNumberOfCustomersWithClassCode
	
End Function


Function GetCustClassDescByIntRecID(passedIntRecID)

	Set cnnGetCustClassDescByClassCode  = Server.CreateObject("ADODB.Connection")
	cnnGetCustClassDescByClassCode.open Session("ClientCnnString")

	resultGetCustClassDescByClassCode = ""
		
	SQLGetCustClassDescByClassCode  = "SELECT * FROM AR_CustomerClass WHERE InternalRecordIdentifier = " & passedIntRecID
	 
	Set rsGetCustClassDescByClassCode  = Server.CreateObject("ADODB.Recordset")
	rsGetCustClassDescByClassCode.CursorLocation = 3 
	
	rsGetCustClassDescByClassCode.Open SQLGetCustClassDescByClassCode,cnnGetCustClassDescByClassCode 
			
	resultGetCustClassDescByClassCode = rsGetCustClassDescByClassCode("ClassDescription")
	
	rsGetCustClassDescByClassCode.Close
	set rsGetCustClassDescByClassCode = Nothing
	cnnGetCustClassDescByClassCode.Close	
	set cnnGetCustClassDescByClassCode = Nothing
	
	GetCustClassDescByClassCode  = resultGetCustClassDescByClassCode 
	
End Function


Function GetClassCodeByIntRecID(passedIntRecID)

	Set cnnGetClassCodeByIntRecID  = Server.CreateObject("ADODB.Connection")
	cnnGetClassCodeByIntRecID.open Session("ClientCnnString")

	resultGetClassCodeByIntRecID = ""
		
	SQLGetClassCodeByIntRecID  = "SELECT * FROM AR_CustomerClass WHERE InternalRecordIdentifier = " & passedIntRecID
	 
	Set rsGetClassCodeByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetClassCodeByIntRecID.CursorLocation = 3 
	
	rsGetClassCodeByIntRecID.Open SQLGetClassCodeByIntRecID,cnnGetClassCodeByIntRecID 
			
	resultGetClassCodeByIntRecID = rsGetClassCodeByIntRecID("ClassCode")
	
	rsGetClassCodeByIntRecID.Close
	set rsGetClassCodeByIntRecID = Nothing
	cnnGetClassCodeByIntRecID.Close	
	set cnnGetClassCodeByIntRecID = Nothing
	
	GetClassCodeByIntRecID  = resultGetClassCodeByIntRecID 
	
End Function


Function GetCustClassByCustID(passedCustID)

	Set cnnGetCustClassByCustID  = Server.CreateObject("ADODB.Connection")
	cnnGetCustClassByCustID.open Session("ClientCnnString")

	resultGetCustClassByCustID = ""
		
	SQLGetCustClassByCustID  = "SELECT ClassCode FROM AR_Customer WHERE CustNum = '" & passedCustID & "'"
	 
	Set rsGetCustClassByCustID  = Server.CreateObject("ADODB.Recordset")
	rsGetCustClassByCustID.CursorLocation = 3 
	
	rsGetCustClassByCustID.Open SQLGetCustClassByCustID,cnnGetCustClassByCustID 
			
	If Not rsGetCustClassByCustID.EOF Then resultGetCustClassByCustID = rsGetCustClassByCustID("ClassCode")
	
	rsGetCustClassByCustID.Close
	set rsGetCustClassByCustID = Nothing
	cnnGetCustClassByCustID.Close	
	set cnnGetCustClassByCustID = Nothing
	
	GetCustClassByCustID  = resultGetCustClassByCustID 
	
End Function

Function GetCustAbbrvNameByCustID(passedCustID)

	resultGetCustAbbrvNameByCustID=""

	Set cnnGetCustAbbrvNameByCustID = Server.CreateObject("ADODB.Connection")
	cnnGetCustAbbrvNameByCustID.open Session("ClientCnnString")
	Set rsGetCustAbbrvNameByCustID = Server.CreateObject("ADODB.Recordset")
	rsGetCustAbbrvNameByCustID.CursorLocation = 3 
	

	SQLGetCustAbbrvNameByCustID = "SELECT AbbreviatedName FROM " & Session("SQL_Owner") & ".AR_CustomerExt WHERE CustID= '" & passedCustID & "'"
	 

	Set rsGetCustAbbrvNameByCustID= cnnGetCustAbbrvNameByCustID.Execute(SQLGetCustAbbrvNameByCustID)
	
	
	If not rsGetCustAbbrvNameByCustID.eof then resultGetCustAbbrvNameByCustID = rsGetCustAbbrvNameByCustID("AbbreviatedName")
	
	Set rsGetCustAbbrvNameByCustID= Nothing
	cnnGetCustAbbrvNameByCustID.Close
	Set cnnGetCustAbbrvNameByCustID= Nothing
	
	GetCustAbbrvNameByCustID = resultGetCustAbbrvNameByCustID
	
End Function

Function GetCustMESByCustID(passedCustID)

	resultGetCustMESByCustID=""

	Set cnnGetCustMESByCustID = Server.CreateObject("ADODB.Connection")
	cnnGetCustMESByCustID.open Session("ClientCnnString")
	Set rsGetCustMESByCustID = Server.CreateObject("ADODB.Recordset")
	rsGetCustMESByCustID.CursorLocation = 3 
	

	SQLGetCustMESByCustID = "SELECT MonthlyExpectedSalesDollars FROM AR_Customer WHERE CustNum= '" & passedCustID & "'"
	 

	Set rsGetCustMESByCustID= cnnGetCustMESByCustID.Execute(SQLGetCustMESByCustID)
	
	
	If not rsGetCustMESByCustID.eof then resultGetCustMESByCustID = rsGetCustMESByCustID("MonthlyExpectedSalesDollars")
	
	If Not IsNumeric(resultGetCustMESByCustID) Then resultGetCustMESByCustID = 0
	
	Set rsGetCustMESByCustID= Nothing
	cnnGetCustMESByCustID.Close
	Set cnnGetCustMESByCustID= Nothing
	
	GetCustMESByCustID = resultGetCustMESByCustID
	
End Function


Function GetCustMCSByCustID(passedCustID)

	resultGetCustMCSByCustID=""

	Set cnnGetCustMCSByCustID = Server.CreateObject("ADODB.Connection")
	cnnGetCustMCSByCustID.open Session("ClientCnnString")
	Set rsGetCustMCSByCustID = Server.CreateObject("ADODB.Recordset")
	rsGetCustMCSByCustID.CursorLocation = 3 
	

	SQLGetCustMCSByCustID = "SELECT MonthlyContractedSalesDollars FROM AR_Customer WHERE CustNum= '" & passedCustID & "'"
	 

	Set rsGetCustMCSByCustID= cnnGetCustMCSByCustID.Execute(SQLGetCustMCSByCustID)
	
	
	If not rsGetCustMCSByCustID.eof then resultGetCustMCSByCustID = rsGetCustMCSByCustID("MonthlyContractedSalesDollars")
	
	If Not IsNumeric(resultGetCustMCSByCustID) Then resultGetCustMCSByCustID = 0
	
	Set rsGetCustMCSByCustID= Nothing
	cnnGetCustMCSByCustID.Close
	Set cnnGetCustMCSByCustID= Nothing
	
	GetCustMCSByCustID = resultGetCustMCSByCustID
	
End Function

Function GetNumberOfInHistDetailLinesByInvoiceNumber(passedInvoiceNumber)
	
	resultGetNumberOfInHistDetailLinesByInvoiceNumber = ""
	
	Set cnnGetNumberOfInHistDetailLinesByInvoiceNumber = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfInHistDetailLinesByInvoiceNumber.open Session("ClientCnnString")
		
	SQLGetNumberOfInHistDetailLinesByInvoiceNumber = "SELECT COUNT(*) as Expr1 FROM In_InvoiceHistDetail WHERE InvoiceID = '" & passedInvoiceNumber & "'"
	 
	Set rsGetNumberOfInHistDetailLinesByInvoiceNumber = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfInHistDetailLinesByInvoiceNumber.CursorLocation = 3 
	Set rsGetNumberOfInHistDetailLinesByInvoiceNumber= cnnGetNumberOfInHistDetailLinesByInvoiceNumber.Execute(SQLGetNumberOfInHistDetailLinesByInvoiceNumber)
		
			
	If not rsGetNumberOfInHistDetailLinesByInvoiceNumber.eof then resultGetNumberOfInHistDetailLinesByInvoiceNumber = rsGetNumberOfInHistDetailLinesByInvoiceNumber("Expr1")
	
	set rsGetNumberOfInHistDetailLinesByInvoiceNumber= Nothing
	set cnnGetNumberOfInHistDetailLinesByInvoiceNumber= Nothing
	
	GetNumberOfInHistDetailLinesByInvoiceNumber = resultGetNumberOfInHistDetailLinesByInvoiceNumber 
	
End Function

Function NumberOfCustomersWithType(passedCustTypeCode)


	Set cnnNumberOfCustomersWithType  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfCustomersWithType.open Session("ClientCnnString")

	resultNumberOfCustomersWithType = 0
		
	SQLNumberOfCustomersWithType  = "SELECT COUNT(*) AS CUSTCOUNT FROM AR_Customer WHERE CustType = '" & passedCustTypeCode& "'"
	 
	Set rsNumberOfCustomersWithType  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfCustomersWithType.CursorLocation = 3 
	
	rsNumberOfCustomersWithType.Open SQLNumberOfCustomersWithType,cnnNumberOfCustomersWithType 
			
	resultNumberOfCustomersWithType = rsNumberOfCustomersWithType("CUSTCOUNT")
	
	rsNumberOfCustomersWithType.Close
	set rsNumberOfCustomersWithType = Nothing
	cnnNumberOfCustomersWithType.Close	
	set cnnNumberOfCustomersWithType = Nothing
	
	NumberOfCustomersWithType = resultNumberOfCustomersWithType
	
End Function


Function GetCustTypeDescByIntRecID(passedIntRecID)

	Set cnnGetCustClassDescByCustType  = Server.CreateObject("ADODB.Connection")
	cnnGetCustClassDescByCustType.open Session("ClientCnnString")

	resultGetCustClassDescByCustType = ""
		
	SQLGetCustClassDescByCustType  = "SELECT * FROM AR_CustomerType WHERE InternalRecordIdentifier = " & passedIntRecID
	 
	Set rsGetCustClassDescByCustType  = Server.CreateObject("ADODB.Recordset")
	rsGetCustClassDescByCustType.CursorLocation = 3 
	
	rsGetCustClassDescByCustType.Open SQLGetCustClassDescByCustType,cnnGetCustClassDescByCustType 
			
	resultGetCustClassDescByCustType = rsGetCustClassDescByCustType("TypeDescription")
	
	rsGetCustClassDescByCustType.Close
	set rsGetCustClassDescByCustType = Nothing
	cnnGetCustClassDescByCustType.Close	
	set cnnGetCustClassDescByCustType = Nothing
	
	GetCustTypeDescByIntRecID  = resultGetCustClassDescByCustType 
	
End Function


Function NumberOfCustomersWithType2(passedCustTypeCode)

	Set cnnNumberOfCustomersWithType  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfCustomersWithType.open Session("ClientCnnString")

	resultNumberOfCustomersWithType = 0
		
	SQLNumberOfCustomersWithType  = "SELECT COUNT(*) AS CUSTCOUNT FROM AR_Customer WHERE ReferalCode = '" & passedCustTypeCode& "'"
	 
	Set rsNumberOfCustomersWithType  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfCustomersWithType.CursorLocation = 3 
	
	rsNumberOfCustomersWithType.Open SQLNumberOfCustomersWithType,cnnNumberOfCustomersWithType 
			
	resultNumberOfCustomersWithType = rsNumberOfCustomersWithType("CUSTCOUNT")
	
	rsNumberOfCustomersWithType.Close
	set rsNumberOfCustomersWithType = Nothing
	cnnNumberOfCustomersWithType.Close	
	set cnnNumberOfCustomersWithType = Nothing
	
	NumberOfCustomersWithType2 = resultNumberOfCustomersWithType
	
End Function

Function NumberOfCustomersWithReferral(passedCustReferalCode)

	Set cnnNumberOfCustomersWithRef  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfCustomersWithRef.open Session("ClientCnnString")

	resultNumberOfCustomersWithRef = 0
		
	SQLNumberOfCustomersWithRef  = "SELECT COUNT(*) AS CUSTCOUNT FROM AR_Customer WHERE ReferalCode = '" & passedCustReferalCode& "'"
	 
	Set rsNumberOfCustomersWithRef  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfCustomersWithRef.CursorLocation = 3 
	
	rsNumberOfCustomersWithRef.Open SQLNumberOfCustomersWithRef,cnnNumberOfCustomersWithRef 
			
	resultNumberOfCustomersWithRef = rsNumberOfCustomersWithRef("CUSTCOUNT")
	
	rsNumberOfCustomersWithRef.Close
	set rsNumberOfCustomersWithRef = Nothing
	cnnNumberOfCustomersWithRef.Close	
	set cnnNumberOfCustomersWithRef = Nothing
	
	NumberOfCustomersWithReferral = resultNumberOfCustomersWithRef
	
End Function

Function GetCustRefDescByIntRecID(passedIntRecID)

	Set cnnGetCustClassDescByCustRef  = Server.CreateObject("ADODB.Connection")
	cnnGetCustClassDescByCustRef.open Session("ClientCnnString")

	resultGetCustClassDescByCustRef = ""
		
	SQLGetCustClassDescByCustType  = "SELECT * FROM AR_CustomerReferral WHERE InternalRecordIdentifier = " & passedIntRecID
	 
	Set rsGetCustClassDescByCustRef  = Server.CreateObject("ADODB.Recordset")
	rsGetCustClassDescByCustRef.CursorLocation = 3 
	
	rsGetCustClassDescByCustRef.Open SQLGetCustClassDescByCustType,cnnGetCustClassDescByCustRef 
			
	resultGetCustClassDescByCustRef = rsGetCustClassDescByCustRef("ReferralName")
	
	rsGetCustClassDescByCustRef.Close
	set rsGetCustClassDescByCustRef = Nothing
	cnnGetCustClassDescByCustRef.Close	
	set cnnGetCustClassDescByCustRef = Nothing
	
	GetCustRefDescByIntRecID  = resultGetCustClassDescByCustRef 
	
End Function

Function NumberOfCustomersWithChain(passedCustReferalCode)


	Set cnnNumberOfCustomersWithRef  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfCustomersWithRef.open Session("ClientCnnString")

	resultNumberOfCustomersWithRef = 0
		
	SQLNumberOfCustomersWithRef  = "SELECT COUNT(*) AS CUSTCOUNT FROM AR_Customer WHERE ChainNum = '" & passedCustReferalCode& "'"
	 
	Set rsNumberOfCustomersWithRef  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfCustomersWithRef.CursorLocation = 3 
	
	rsNumberOfCustomersWithRef.Open SQLNumberOfCustomersWithRef,cnnNumberOfCustomersWithRef 
			
	resultNumberOfCustomersWithRef = rsNumberOfCustomersWithRef("CUSTCOUNT")
	
	rsNumberOfCustomersWithRef.Close
	set rsNumberOfCustomersWithRef = Nothing
	cnnNumberOfCustomersWithRef.Close	
	set cnnNumberOfCustomersWithRef = Nothing
	
	NumberOfCustomersWithChain = resultNumberOfCustomersWithRef
	
End Function

Function GetCustChainByIntRecID(passedIntRecID)

	Set cnnGetCustClassDescByCustRef  = Server.CreateObject("ADODB.Connection")
	cnnGetCustClassDescByCustRef.open Session("ClientCnnString")

	resultGetCustClassDescByCustRef = ""
		
	SQLGetCustClassDescByCustType  = "SELECT * FROM AR_Chain WHERE InternalRecordIdentifier = " & passedIntRecID
	 
	Set rsGetCustClassDescByCustRef  = Server.CreateObject("ADODB.Recordset")
	rsGetCustClassDescByCustRef.CursorLocation = 3 
	
	rsGetCustClassDescByCustRef.Open SQLGetCustClassDescByCustType,cnnGetCustClassDescByCustRef 
			
	resultGetCustClassDescByCustRef = rsGetCustClassDescByCustRef("Description")
	
	rsGetCustClassDescByCustRef.Close
	set rsGetCustClassDescByCustRef = Nothing
	cnnGetCustClassDescByCustRef.Close	
	set cnnGetCustClassDescByCustRef = Nothing
	
	GetCustChainByIntRecID  = resultGetCustClassDescByCustRef 
	
End Function

Function NumberOfCustomersWithTerm(passedCustTermCode)

	Set cnnNumberOfCustomersWithTerm  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfCustomersWithTerm.open Session("ClientCnnString")

	resultNumberOfCustomersWithTerm = 0
		
	SQLNumberOfCustomersWithTerm  = "SELECT COUNT(*) AS CUSTCOUNT FROM AR_Customer WHERE TermsIntRecID = '" & passedCustTermCode& "'"
	 
	Set rsNumberOfCustomersWithTerm  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfCustomersWithTerm.CursorLocation = 3 
	
	rsNumberOfCustomersWithTerm.Open SQLNumberOfCustomersWithTerm,cnnNumberOfCustomersWithTerm 
			
	resultNumberOfCustomersWithTerm = rsNumberOfCustomersWithTerm("CUSTCOUNT")
	
	rsNumberOfCustomersWithTerm.Close
	set rsNumberOfCustomersWithTerm = Nothing
	cnnNumberOfCustomersWithTerm.Close	
	set cnnNumberOfCustomersWithTerm = Nothing
	
	NumberOfCustomersWithTerm = resultNumberOfCustomersWithTerm
	
End Function
Function GetCustTermDescByIntRecID(passedIntRecID)

	Set cnnGetCustDescByCustTerm  = Server.CreateObject("ADODB.Connection")
	cnnGetCustDescByCustTerm.open Session("ClientCnnString")

	resultGetCustDescByCustTerm = ""
		
	SQLGetCustDescByCustTerm  = "SELECT * FROM AR_Terms WHERE InternalRecordIdentifier = " & passedIntRecID
	 
	Set rsGetCustDescByCustTerm  = Server.CreateObject("ADODB.Recordset")
	rsGetCustDescByCustTerm.CursorLocation = 3 
	
	rsGetCustDescByCustTerm.Open SQLGetCustDescByCustTerm,cnnGetCustDescByCustTerm 
			
	resultGetCustDescByCustTerm = rsGetCustDescByCustTerm("Description")
	
	rsGetCustDescByCustTerm.Close
	set rsGetCustDescByCustTerm = Nothing
	cnnGetCustDescByCustTerm.Close	
	set cnnGetCustDescByCustTerm = Nothing
	
	GetCustTermDescByIntRecID  = resultGetCustDescByCustTerm 
	
End Function


Function GetInvoiceExportedToSageLastDate(passedInvoiceID)

	Set cnnGetInvoiceExportedToSageLastDate  = Server.CreateObject("ADODB.Connection")
	cnnGetInvoiceExportedToSageLastDate.open Session("ClientCnnString")

	resultGetInvoiceExportedToSageLastDate = ""
		
	SQLGetInvoiceExportedToSageLastDate  = "SELECT TOP 1 RecordCreationDateTime FROM IN_InvoicesExportedSage WHERE InvoiceID = '" & passedInvoiceID & "' ORDER BY RecordCreationDateTime DESC"
	 
	Set rsGetInvoiceExportedToSageLastDate  = Server.CreateObject("ADODB.Recordset")
	
	rsGetInvoiceExportedToSageLastDate.Open SQLGetInvoiceExportedToSageLastDate,cnnGetInvoiceExportedToSageLastDate 
			
	If NOT rsGetInvoiceExportedToSageLastDate.EOF Then resultGetInvoiceExportedToSageLastDate = rsGetInvoiceExportedToSageLastDate("RecordCreationDateTime")
	
	rsGetInvoiceExportedToSageLastDate.Close
	set rsGetInvoiceExportedToSageLastDate = Nothing
	cnnGetInvoiceExportedToSageLastDate.Close	
	set cnnGetInvoiceExportedToSageLastDate = Nothing
	
	GetInvoiceExportedToSageLastDate  = resultGetInvoiceExportedToSageLastDate 
	
End Function


Function GetCustRegionByCustID(passedCustID)

	Set cnnGetCustRegionByCustID  = Server.CreateObject("ADODB.Connection")
	cnnGetCustRegionByCustID.open Session("ClientCnnString")
	Set rsGetCustRegionByCustID  = Server.CreateObject("ADODB.Recordset")

	resultGetCustRegionByCustID = ""
		
	SQLGetCustRegionByCustID  = "SELECT City,[State],Zip FROM AR_Customer WHERE CustNum = '" & passedCustID & "'"
	 
	Set rsGetCustRegionByCustID = cnnGetCustRegionByCustID.Execute(SQLGetCustRegionByCustID)
	
	If NOT rsGetCustRegionByCustID.EOF Then
	
		'Get what we need: City, StateForCities, State, Zip Code
		Zip = rsGetCustRegionByCustID("Zip")
		City = rsGetCustRegionByCustID("City")
		If LEN(City) > 0 Then City = Replace(City,"'","''")
		State = rsGetCustRegionByCustID("State")
		If LEN(State) > 0 Then State = Replace(State,"'","''")
		
		'Zip First
		If LEN(Zip) > 0 <> "" Then
			SQLGetCustRegionByCustID  = "SELECT Region FROM AR_Regions WHERE "
			SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(ZipOrPostalCodes1, ' ,', ','), ', ', ','), ',') }, "
			SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "REPLACE(REPLACE(ZipOrPostalCodes2, ' ,', ','), ', ', ',')) } LIKE '%" & Zip & "%')"

			'Response.Write("<br>" & SQLGetCustRegionByCustID)
			
			Set rsGetCustRegionByCustID = cnnGetCustRegionByCustID.Execute(SQLGetCustRegionByCustID)

			If Not rsGetCustRegionByCustID.EOF Then resultGetCustRegionByCustID = rsGetCustRegionByCustID("Region")
		End If
		
		If resultGetCustRegionByCustID = "" Then ' No results yet
			If LEN(City) > 0 Then
				'If it is not set, now look for the city and state
				SQLGetCustRegionByCustID  = "SELECT Region FROM AR_Regions WHERE "
				SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "({ fn CONCAT({ fn CONCAT({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(Cities1, ' ,', ','), ', ', ','), ',') }, "
				SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "REPLACE(REPLACE(Cities2, ' ,', ','), ', ', ',')) }, ',') }, REPLACE(REPLACE(Cities3, ' ,', ','), ', ', ',')) "
				SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  &	"} LIKE '%," & City & "%' OR "
				SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "{ fn CONCAT({ fn CONCAT({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(Cities1, ' ,', ','), ', ', ','), ',') }, "
				SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "REPLACE(REPLACE(Cities2, ' ,', ','), ', ', ',')) }, ',') }, REPLACE(REPLACE(Cities3, ' ,', ','), ', ', ',')) "
				SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "} LIKE '%" & City & ",%') "
				SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "AND (StateForCities = '" & State & "')"
		
				''Response.Write("<br>" & SQLGetCustRegionByCustID)
				
				Set rsGetCustRegionByCustID = cnnGetCustRegionByCustID.Execute(SQLGetCustRegionByCustID)
		
				If Not rsGetCustRegionByCustID.EOF Then resultGetCustRegionByCustID = rsGetCustRegionByCustID("Region")
			End If
		End If
		
		If resultGetCustRegionByCustID = "" Then ' No results yet
			If LEN(State) > 0 Then
				'If it is not set, now look for the state
				SQLGetCustRegionByCustID  = "SELECT Region FROM AR_Regions WHERE "
				SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "REPLACE(REPLACE(StatesOrProvinces , ' ,', ','), ', ', ',') "
				SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  &	" LIKE '%" & State& "%'"

				'Response.Write("<br>" & SQLGetCustRegionByCustID)		
				Set rsGetCustRegionByCustID = cnnGetCustRegionByCustID.Execute(SQLGetCustRegionByCustID)
		
				If Not rsGetCustRegionByCustID.EOF Then resultGetCustRegionByCustID = rsGetCustRegionByCustID("Region")
			End If
	 	End If
	 	
		'Now we have to see if there is a catchall region
		If resultGetCustRegionByCustID = "" Then ' No results yet
			SQLGetCustRegionByCustID  = "SELECT Region FROM AR_Regions WHERE [CatchAllRegionIntRecIDs] IS NOT NULL and [CatchAllRegionIntRecIDs] <> """
			Set rsGetCustRegionByCustID = cnnGetCustRegionByCustID.Execute(SQLGetCustRegionByCustID)
		
			If Not rsGetCustRegionByCustID.EOF Then resultGetCustRegionByCustID = rsGetCustRegionByCustID("Region")

		End If

	End IF

	set rsGetCustRegionByCustID = Nothing
	cnnGetCustRegionByCustID.Close	
	set cnnGetCustRegionByCustID = Nothing
	

	GetCustRegionByCustID  = resultGetCustRegionByCustID 
	
End Function

FUNCTION GetRegionName(Zip, City,State)
	GetRegionName=""
	Set cnnGetCustRegionByCustID  = Server.CreateObject("ADODB.Connection")
	cnnGetCustRegionByCustID.open Session("ClientCnnString")
	Set rsGetCustRegionByCustID  = Server.CreateObject("ADODB.Recordset")
	
	
	If LEN(Zip) > 0 <> "" Then
			SQLGetCustRegionByCustID  = "SELECT Region FROM AR_Regions WHERE "
			SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(ZipOrPostalCodes1, ' ,', ','), ', ', ','), ',') }, "
			SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "REPLACE(REPLACE(ZipOrPostalCodes2, ' ,', ','), ', ', ',')) } = '" & Zip & "')"

			''Response.Write("<br>" & SQLGetCustRegionByCustID)
			
			Set rsGetCustRegionByCustID = cnnGetCustRegionByCustID.Execute(SQLGetCustRegionByCustID)

			If Not rsGetCustRegionByCustID.EOF Then GetRegionName = rsGetCustRegionByCustD("Region")
			GetRegionName = SQLGetCustRegionByCustID
		End If
		
		If LEN(City) > 0 Then
			'If it is not set, now look for the city and state
			SQLGetCustRegionByCustID  = "SELECT Region FROM AR_Regions WHERE "
			SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "({ fn CONCAT({ fn CONCAT({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(Cities1, ' ,', ','), ', ', ','), ',') }, "
			SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "REPLACE(REPLACE(Cities2, ' ,', ','), ', ', ',')) }, ',') }, REPLACE(REPLACE(Cities3, ' ,', ','), ', ', ',')) "
			SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  &	"} LIKE '%," & Replace(City,"'","''") & "%' OR "
			SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "{ fn CONCAT({ fn CONCAT({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(Cities1, ' ,', ','), ', ', ','), ',') }, "
			SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "REPLACE(REPLACE(Cities2, ' ,', ','), ', ', ',')) }, ',') }, REPLACE(REPLACE(Cities3, ' ,', ','), ', ', ',')) "
			SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "} LIKE '%" & Replace(City,"'","''") & ",%') "
			SQLGetCustRegionByCustID  = SQLGetCustRegionByCustID  & "AND (StateForCities = '" & Replace(State,"'","''") & "')"
	
			'Set rsGetCustRegionByCustID = cnnGetCustRegionByCustID.Execute(SQLGetCustRegionByCustID)
	
			If Not rsGetCustRegionByCustID.EOF Then GetRegionName = SQLGetCustRegionByCustID
			GetRegionName = SQLGetCustRegionByCustID
		End If
		
		If LEN(State) > 0 Then
			'If it is not set, now look for the state
			SQLGetCustRegionByCustID  = "SELECT Region FROM AR_Regions WHERE '" & Replace(State,"'","''") & "' IN (SELECT StatesOrProvinces FROM AR_Regions)"
			''Response.Write("<br>" & SQLGetCustRegionByCustID   )		
			Set rsGetCustRegionByCustID = cnnGetCustRegionByCustID.Execute(SQLGetCustRegionByCustID)
	
			If Not rsGetCustRegionByCustID.EOF Then GetRegionName = rsGetCustRegionByCustID("Region")
			GetRegionName = SQLGetCustRegionByCustID
		End If
	set rsGetCustRegionByCustID = Nothing
	cnnGetCustRegionByCustID.Close	
	set cnnGetCustRegionByCustID = Nothing
	
END FUNCTION

FUNCTION GetQtyCustByRegion()
	GetQtyCustByRegion=""
	
	
	Set cnnGetCustRegionByCustID  = Server.CreateObject("ADODB.Connection")
	cnnGetCustRegionByCustID.open Session("ClientCnnString")
	Set rsGetCustRegionByCustID  = Server.CreateObject("ADODB.Recordset")

	resultGetCustRegionByCustID = ""
		
		
	SQLGetCustQtyByRegion  = "SELECT COUNT(AR_Customer.CustNum) AS RegionQty, AR_Regions.Region As Region" 
	SQLGetCustQtyByRegion  =SQLGetCustQtyByRegion  & " FROM AR_Regions, AR_Customer,FS_CustomerFilters,IC_Filters"
	SQLGetCustQtyByRegion  =SQLGetCustQtyByRegion  & " WHERE ((AR_Customer.[State] Is NOT NULL AND AR_Customer.[State] IN (SELECT StatesOrProvinces FROM AR_Regions)) OR"
	SQLGetCustQtyByRegion  =SQLGetCustQtyByRegion  &  " (AR_Customer.Zip Is NOT NULL AND ({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(AR_Regions.ZipOrPostalCodes1, ' ,', ','), ', ', ','), ',') },REPLACE(REPLACE(AR_Regions.ZipOrPostalCodes2, ' ,', ','), ', ', ',')) } =  AR_Customer.Zip)) OR"
	SQLGetCustQtyByRegion  =SQLGetCustQtyByRegion  &  " (AR_Customer.City Is NOT NULL AND AR_Customer.[State] Is NOT NULL AND AR_Regions.StateForCities=AR_Customer.[State]"
	SQLGetCustQtyByRegion  =SQLGetCustQtyByRegion  &  " AND "
	SQLGetCustQtyByRegion  =SQLGetCustQtyByRegion  &  " (CHARINDEX(AR_Customer.City, { fn CONCAT({ fn CONCAT({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(AR_Regions.Cities1, ' ,', ','), ', ', ','), ',') },REPLACE(REPLACE(AR_Regions.Cities2, ' ,', ','), ', ', ',')) }, ',') }, REPLACE(REPLACE(AR_Regions.Cities3, ' ,', ','), ', ', ','))})>0"
	SQLGetCustQtyByRegion  =SQLGetCustQtyByRegion  &  " OR "
	SQLGetCustQtyByRegion  =SQLGetCustQtyByRegion  &  "CHARINDEX(AR_Customer.City, ','+{ fn CONCAT({ fn CONCAT({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(AR_Regions.Cities1, ' ,', ','), ', ', ','), ',') },REPLACE(REPLACE(AR_Regions.Cities2, ' ,', ','), ', ', ',')) }, ',') }, REPLACE(REPLACE(AR_Regions.Cities3, ' ,', ','), ', ', ','))})>0)))"
	SQLGetCustQtyByRegion  =SQLGetCustQtyByRegion  &  " AND IC_Filters.InternalRecordIdentifier=FS_CustomerFilters.FilterIntRecID AND AR_Customer.CustNum=FS_CustomerFilters.CustID"

	SQLGetCustQtyByRegion  =SQLGetCustQtyByRegion  &  " GROUP BY AR_Regions.Region"
	SQLGetCustQtyByRegion  =SQLGetCustQtyByRegion  &  " ORDER BY AR_Regions.Region" 
	 
	 
	Set rsGetCustRegionByCustID = cnnGetCustRegionByCustID.Execute(SQLGetCustQtyByRegion)
	DO WHILE NOT rsGetCustRegionByCustID.EOF
				
		IF LEN(GetQtyCustByRegion)>0 Then
			GetQtyCustByRegion=GetQtyCustByRegion & ","
		END IF
		GetQtyCustByRegion=GetQtyCustByRegion & "{""region"":""" & rsGetCustRegionByCustID("Region") & """,""qty"":""" & rsGetCustRegionByCustID("RegionQty") & """}"
		rsGetCustRegionByCustID.MoveNext
	LOOP
	set rsGetCustRegionByCustID = Nothing
	cnnGetCustRegionByCustID.Close	
	set cnnGetCustRegionByCustID = Nothing
	
	GetQtyCustByRegion="["+GetQtyCustByRegion+"]"
	
END FUNCTION


Function GetCustIntRecIDByARCustContactIntRecID(passedCustID)

	resultGetCustIntRecIDByARCustContactIntRecID=""

	Set cnnGetCustIntRecIDByARCustContactIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetCustIntRecIDByARCustContactIntRecID.open Session("ClientCnnString")
	Set rsGetCustIntRecIDByARCustContactIntRecID = Server.CreateObject("ADODB.Recordset")
	rsGetCustIntRecIDByARCustContactIntRecID.CursorLocation = 3 
	

	SQLGetCustIntRecIDByARCustContactIntRecID = "SELECT * FROM AR_Customer WHERE CustNum = '" & passedCustID & "'"
	 

	Set rsGetCustIntRecIDByARCustContactIntRecID= cnnGetCustIntRecIDByARCustContactIntRecID.Execute(SQLGetCustIntRecIDByARCustContactIntRecID)
	
	
	If not rsGetCustIntRecIDByARCustContactIntRecID.eof then resultGetCustIntRecIDByARCustContactIntRecID = rsGetCustIntRecIDByARCustContactIntRecID("InternalRecordIdentifier")
	
	Set rsGetCustIntRecIDByARCustContactIntRecID= Nothing
	cnnGetCustIntRecIDByARCustContactIntRecID.Close
	Set cnnGetCustIntRecIDByARCustContactIntRecID= Nothing
	
	GetCustIntRecIDByARCustContactIntRecID = resultGetCustIntRecIDByARCustContactIntRecID
	
End Function


Function GetCustIDByCustIntRecID(passedCustIntRecID)

	resultGetCustIDByCustIntRecID=""

	Set cnnGetCustIDByCustIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetCustIDByCustIntRecID.open Session("ClientCnnString")
	Set rsGetCustIDByCustIntRecID = Server.CreateObject("ADODB.Recordset")
	rsGetCustIDByCustIntRecID.CursorLocation = 3 
	

	SQLGetCustIDByCustIntRecID = "SELECT * FROM AR_Customer WHERE InternalRecordIdentifier = " & passedCustIntRecID
	 

	Set rsGetCustIDByCustIntRecID= cnnGetCustIDByCustIntRecID.Execute(SQLGetCustIDByCustIntRecID)
	
	
	If not rsGetCustIDByCustIntRecID.eof then resultGetCustIDByCustIntRecID = rsGetCustIDByCustIntRecID("CustNum")
	
	Set rsGetCustIDByCustIntRecID= Nothing
	cnnGetCustIDByCustIntRecID.Close
	Set cnnGetCustIDByCustIntRecID= Nothing
	
	GetCustIDByCustIntRecID = resultGetCustIDByCustIntRecID
	
End Function


Function GetCustIntRecIDByCustID(passedCustID)

	resultGetCustIntRecIDByCustID=""

	Set cnnGetCustIntRecIDByCustID = Server.CreateObject("ADODB.Connection")
	cnnGetCustIntRecIDByCustID.open Session("ClientCnnString")
	Set rsGetCustIntRecIDByCustID = Server.CreateObject("ADODB.Recordset")
	rsGetCustIntRecIDByCustID.CursorLocation = 3 
	

	SQLGetCustIntRecIDByCustID = "SELECT * FROM AR_Customer WHERE CustNum = '" & passedCustID & "'"
	 

	Set rsGetCustIntRecIDByCustID= cnnGetCustIntRecIDByCustID.Execute(SQLGetCustIntRecIDByCustID)
	
	
	If not rsGetCustIntRecIDByCustID.eof then resultGetCustIntRecIDByCustID = rsGetCustIntRecIDByCustID("InternalRecordIdentifier")
	
	Set rsGetCustIntRecIDByCustID= Nothing
	cnnGetCustIntRecIDByCustID.Close
	Set cnnGetCustIntRecIDByCustID= Nothing
	
	GetCustIntRecIDByCustID = resultGetCustIntRecIDByCustID
	
End Function

Function GetPaymentMethodByIntRecID(passedIntRecID)

	resultGetPaymentMethodByIntRecID=""

	Set cnnGetPaymentMethodByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetPaymentMethodByIntRecID.open Session("ClientCnnString")
	Set rsGetPaymentMethodByIntRecID = Server.CreateObject("ADODB.Recordset")
	rsGetPaymentMethodByIntRecID.CursorLocation = 3 
	

	SQLGetPaymentMethodByIntRecID = "SELECT * FROM AR_PaymentMethods WHERE InternalRecordIdentifier = " & passedIntRecID
	 

	Set rsGetPaymentMethodByIntRecID= cnnGetPaymentMethodByIntRecID.Execute(SQLGetPaymentMethodByIntRecID)
	
	
	If not rsGetPaymentMethodByIntRecID.eof then resultGetPaymentMethodByIntRecID = rsGetPaymentMethodByIntRecID("PayMethDescription")
	
	Set rsGetPaymentMethodByIntRecID= Nothing
	cnnGetPaymentMethodByIntRecID.Close
	Set cnnGetPaymentMethodByIntRecID= Nothing
	
	GetPaymentMethodByIntRecID = resultGetPaymentMethodByIntRecID
	
End Function



Function GetCustNoteTypeByNoteIntRecID(passedNoteIntRecID)

	resultGetCustNoteTypeByNoteIntRecID = ""

	Set cnnGetCustNoteTypeByNoteIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetCustNoteTypeByNoteIntRecID.open Session("ClientCnnString")
	Set rsGetCustNoteTypeByNoteIntRecID = Server.CreateObject("ADODB.Recordset")
	rsGetCustNoteTypeByNoteIntRecID.CursorLocation = 3 
	
	SQLGetCustNoteTypeByNoteIntRecID = "SELECT * FROM SC_NoteType WHERE InternalRecordIdentifier = " & passedNoteIntRecID

	Set rsGetCustNoteTypeByNoteIntRecID= cnnGetCustNoteTypeByNoteIntRecID.Execute(SQLGetCustNoteTypeByNoteIntRecID)
	
	If not rsGetCustNoteTypeByNoteIntRecID.eof then resultGetCustNoteTypeByNoteIntRecID = rsGetCustNoteTypeByNoteIntRecID("NoteType")
	
	Set rsGetCustNoteTypeByNoteIntRecID= Nothing
	cnnGetCustNoteTypeByNoteIntRecID.Close
	Set cnnGetCustNoteTypeByNoteIntRecID= Nothing
	
	GetCustNoteTypeByNoteIntRecID = resultGetCustNoteTypeByNoteIntRecID
	
End Function


Function GetCustNoteCountByNoteType(passedNoteIntRecID, passedCustID, passedUserNo)

	resultGetCustNoteCountByNoteType = 0

	Set cnnGetCustNoteCountByNoteType = Server.CreateObject("ADODB.Connection")
	cnnGetCustNoteCountByNoteType.open Session("ClientCnnString")
	Set rsGetCustNoteCountByNoteType = Server.CreateObject("ADODB.Recordset")
	rsGetCustNoteCountByNoteType.CursorLocation = 3 
	
	If passedNoteIntRecID = 0 Then
		If passedUserNo = 0 Then
			SQLGetCustNoteCountByNoteType = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID <> '' AND CustID = '" & passedCustID & "'"
		Else
			SQLGetCustNoteCountByNoteType = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID <> '' AND CustID = '" & passedCustID & "' AND EnteredByUserNo = " & passedUserNo
		End If
	Else
		If passedUserNo = 0 Then
			SQLGetCustNoteCountByNoteType = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID = " & passedNoteIntRecID & " AND CustID = '" & passedCustID & "'"
		Else
			SQLGetCustNoteCountByNoteType = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID = " & passedNoteIntRecID & " AND CustID = '" & passedCustID & "' AND EnteredByUserNo = " & passedUserNo
		End If
	End If

	Set rsGetCustNoteCountByNoteType= cnnGetCustNoteCountByNoteType.Execute(SQLGetCustNoteCountByNoteType)
	
	If not rsGetCustNoteCountByNoteType.eof then resultGetCustNoteCountByNoteType = rsGetCustNoteCountByNoteType("NoteTypeCount")
	
	Set rsGetCustNoteCountByNoteType= Nothing
	cnnGetCustNoteCountByNoteType.Close
	Set cnnGetCustNoteCountByNoteType= Nothing
	
	GetCustNoteCountByNoteType = resultGetCustNoteCountByNoteType
	
End Function



Function GetCustNoteCountByNoteTypeAllUsers(passedNoteIntRecID, passedCustID)

	resultGetCustNoteCountByNoteTypeAllUsers = 0

	Set cnnGetCustNoteCountByNoteTypeAllUsers = Server.CreateObject("ADODB.Connection")
	cnnGetCustNoteCountByNoteTypeAllUsers.open Session("ClientCnnString")
	Set rsGetCustNoteCountByNoteTypeAllUsers = Server.CreateObject("ADODB.Recordset")
	rsGetCustNoteCountByNoteTypeAllUsers.CursorLocation = 3 
	
	If passedNoteIntRecID = 0 Then
		SQLGetCustNoteCountByNoteTypeAllUsers = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID <> '' AND CustID = '" & passedCustID & "'"
	Else
		SQLGetCustNoteCountByNoteTypeAllUsers = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID = " & passedNoteIntRecID & " AND CustID = '" & passedCustID & "'"
	End If

	Set rsGetCustNoteCountByNoteTypeAllUsers= cnnGetCustNoteCountByNoteTypeAllUsers.Execute(SQLGetCustNoteCountByNoteTypeAllUsers)
	
	If not rsGetCustNoteCountByNoteTypeAllUsers.eof then resultGetCustNoteCountByNoteTypeAllUsers = rsGetCustNoteCountByNoteTypeAllUsers("NoteTypeCount")
	
	Set rsGetCustNoteCountByNoteTypeAllUsers= Nothing
	cnnGetCustNoteCountByNoteTypeAllUsers.Close
	Set cnnGetCustNoteCountByNoteTypeAllUsers= Nothing
	
	GetCustNoteCountByNoteTypeAllUsers = resultGetCustNoteCountByNoteTypeAllUsers
	
End Function



Function GetCustNoteCountByNoteTypeJustMe(passedNoteIntRecID, passedCustID)

	resultGetCustNoteCountByNoteTypeJustMe = 0

	Set cnnGetCustNoteCountByNoteTypeJustMe = Server.CreateObject("ADODB.Connection")
	cnnGetCustNoteCountByNoteTypeJustMe.open Session("ClientCnnString")
	Set rsGetCustNoteCountByNoteTypeJustMe = Server.CreateObject("ADODB.Recordset")
	rsGetCustNoteCountByNoteTypeJustMe.CursorLocation = 3 
	
	If passedNoteIntRecID = 0 Then
		SQLGetCustNoteCountByNoteTypeJustMe = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID <> '' AND CustID = '" & passedCustID & "' AND EnteredByUserNo = " & Session("UserNo")
	Else
		SQLGetCustNoteCountByNoteTypeJustMe = "SELECT COUNT(InternalRecordIdentifier) AS NoteTypeCount FROM AR_CustomerNotes WHERE NoteTypeIntRecID = " & passedNoteIntRecID & " AND CustID = '" & passedCustID & "' AND EnteredByUserNo = " & Session("UserNo")	End If

	Set rsGetCustNoteCountByNoteTypeJustMe= cnnGetCustNoteCountByNoteTypeJustMe.Execute(SQLGetCustNoteCountByNoteTypeJustMe)
	
	If not rsGetCustNoteCountByNoteTypeJustMe.eof then resultGetCustNoteCountByNoteTypeJustMe = rsGetCustNoteCountByNoteTypeJustMe("NoteTypeCount")
	
	Set rsGetCustNoteCountByNoteTypeJustMe= Nothing
	cnnGetCustNoteCountByNoteTypeJustMe.Close
	Set cnnGetCustNoteCountByNoteTypeJustMe= Nothing
	
	GetCustNoteCountByNoteTypeJustMe = resultGetCustNoteCountByNoteTypeJustMe
	
End Function


Function GetCustNoteTypeCanBeEdited(passedNoteIntRecID)

	resultGetCustNoteTypeCanBeEdited = 0

	Set cnnGetCustNoteTypeCanBeEdited = Server.CreateObject("ADODB.Connection")
	cnnGetCustNoteTypeCanBeEdited.open Session("ClientCnnString")
	Set rsGetCustNoteTypeCanBeEdited = Server.CreateObject("ADODB.Recordset")
	rsGetCustNoteTypeCanBeEdited.CursorLocation = 3 
	
	SQLGetCustNoteTypeCanBeEdited = "SELECT NoteTypeCanBeCreatedByUser FROM SC_NoteType WHERE InternalRecordIdentifier = " & passedNoteIntRecID

	Set rsGetCustNoteTypeCanBeEdited= cnnGetCustNoteTypeCanBeEdited.Execute(SQLGetCustNoteTypeCanBeEdited)
	
	If not rsGetCustNoteTypeCanBeEdited.eof then resultGetCustNoteTypeCanBeEdited = rsGetCustNoteTypeCanBeEdited("NoteTypeCanBeCreatedByUser")
	
	Set rsGetCustNoteTypeCanBeEdited= Nothing
	cnnGetCustNoteTypeCanBeEdited.Close
	Set cnnGetCustNoteTypeCanBeEdited= Nothing
	
	GetCustNoteTypeCanBeEdited = resultGetCustNoteTypeCanBeEdited
	
End Function





'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Function HasNoteTypeBeenViewedByUser(passedCustNum,passedNoteTypeIntRecID,passedShowOnlyMyNotes)

	resultHasNoteTypeBeenViewedByUser = "False"
	
	'ALl NOTES TABS SPECIAL CODE
	If passedNoteTypeIntRecID = 0 Then
	
		If UserHasAnyUnviewedNotes(passedCustNum) = "True" Then
			resultHasNoteTypeBeenViewedByUser = "False"
		Else
			resultHasNoteTypeBeenViewedByUser = "True"
		End If
		
	'ALL OTHER TABS
	Else
	
		If passedShowOnlyMyNotes = 0 Then
			TotalNotes = GetCustNoteCountByNoteTypeAllUsers(passedNoteTypeIntRecID, passedCustNum)
		Else
			TotalNotes = GetCustNoteCountByNoteTypeJustMe(passedNoteTypeIntRecID, passedCustNum)
		End If
			
		If TotalNotes > 0 Then
		
			SQLHasNoteTypeBeenViewedByUser = "SELECT * FROM AR_CustomerNotesUserViewed Where CustID ='" & passedCustNum & "' AND UserNo = " & Session("Userno") & " AND NoteTypeIntRecID = " & passedNoteTypeIntRecID
			
			Set cnnHasNoteTypeBeenViewedByUser = Server.CreateObject("ADODB.Connection")
			cnnHasNoteTypeBeenViewedByUser.open (Session("ClientCnnString"))
			Set rsHasNoteTypeBeenViewedByUser = Server.CreateObject("ADODB.Recordset")
			rsHasNoteTypeBeenViewedByUser.CursorLocation = 3 
			Set rsHasNoteTypeBeenViewedByUser = cnnHasNoteTypeBeenViewedByUser.Execute(SQLHasNoteTypeBeenViewedByUser)
		
			Set rsNote = Server.CreateObject("ADODB.Recordset")
			rsNote.CursorLocation = 3 
		
			If not rsHasNoteTypeBeenViewedByUser.EOF Then
				'OK, so see when the last note was created, not by us
				SQLCustHasNotes = "SELECT TOP 1 RecordCreationDateTime FROM AR_CustomerNotes "
				SQLCustHasNotes = SQLCustHasNotes & " WHERE CustID = '" & passedCustNum & "' AND NoteTypeIntRecID = " & passedNoteTypeIntRecID
				SQLCustHasNotes = SQLCustHasNotes & " ORDER BY RecordCreationDateTime DESC"
				
				Set rsNote = cnnHasNoteTypeBeenViewedByUser.Execute(SQLCustHasNotes)
				If Not rsNote.Eof Then
					If rsHasNoteTypeBeenViewedByUser("DateLastViewed") < rsNote("RecordCreationDateTime")  Then 
						resultHasNoteTypeBeenViewedByUser = "False"
					Else
						resultHasNoteTypeBeenViewedByUser = "True"
					End If
				End If
			Else
				resultHasNoteTypeBeenViewedByUser = "False" 'Also true if they have never seen any of them
			End If
			
			cnnHasNoteTypeBeenViewedByUser.close
			set rsHasNoteTypeBeenViewedByUser = nothing
			set rsNote = nothing
			set cnnHasNoteTypeBeenViewedByUser= nothing	
		
		Else
			resultHasNoteTypeBeenViewedByUser = "True"
		End If
		
	End If

	HasNoteTypeBeenViewedByUser = resultHasNoteTypeBeenViewedByUser

End Function

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Function UserHasAnyUnviewedNotes(passedCustNum)

	resultUserHasAnyUnviewedNotes = "False"
	
	SQLUserHasAnyUnviewedNotes = "SELECT MAX(DateLastViewed) AS DLV, CASE WHEN NoteTypeIntRecID IS NULL THEN 0 ELSE NoteTypeIntRecID END AS NoteTypeIntRecID  FROM AR_CustomerNotesUserViewed WHERE CustID ='" & passedCustNum & "' AND UserNo = " & Session("Userno") & " GROUP BY NoteTypeIntRecID  "
	Set cnnUserHasAnyUnviewedNotes = Server.CreateObject("ADODB.Connection")
	cnnUserHasAnyUnviewedNotes.open (Session("ClientCnnString"))
	Set rsUserHasAnyUnviewedNotes = Server.CreateObject("ADODB.Recordset")
	rsUserHasAnyUnviewedNotes.CursorLocation = 3 
	Set rsUserHasAnyUnviewedNotes = cnnUserHasAnyUnviewedNotes.Execute(SQLUserHasAnyUnviewedNotes)
	Set rsNote = Server.CreateObject("ADODB.Recordset")
	rsNote.CursorLocation = 3 
'response.write(SQLUserHasAnyUnviewedNotes)
	If not rsUserHasAnyUnviewedNotes.EOF Then
	
		Do While Not rsUserHasAnyUnviewedNotes.EOF 
		
			'OK, so see when the last note was created, not by us
			SQLCustHasNotes = "SELECT TOP 1 RecordCreationDateTime FROM AR_CustomerNotes "
			SQLCustHasNotes = SQLCustHasNotes & " WHERE CustID = '" & passedCustNum & "' "
			SQLCustHasNotes = SQLCustHasNotes & " AND NoteTypeIntRecID = " & rsUserHasAnyUnviewedNotes("NoteTypeIntRecID")
			SQLCustHasNotes = SQLCustHasNotes & " ORDER BY RecordCreationDateTime DESC"

			Set rsNote = cnnUserHasAnyUnviewedNotes.Execute(SQLCustHasNotes)
			
			If Not rsNote.Eof Then
				If  rsNote("RecordCreationDateTime") > rsUserHasAnyUnviewedNotes("DLV") Then
					resultUserHasAnyUnviewedNotes = "True"
					Exit Do
				End If
			End If
			
			rsUserHasAnyUnviewedNotes.MoveNext
		Loop

	Else
		resultUserHasAnyUnviewedNotes = "True"
	End If		

	' Now see if there are notes type that have NEVER been read
	'OK, so see when the last note was created, not by us
	SQLCustHasNotes = "SELECT CustID, NoteTypeIntRecID FROM AR_CustomerNotes WHERE CustID = '" & passedCustNum  & "' "
	SQLCustHasNotes = SQLCustHasNotes & " EXCEPT  "
	SQLCustHasNotes = SQLCustHasNotes & " SELECT CustID, NoteTypeIntRecID FROM AR_CustomerNotesUserViewed WHERE CustID = '" & passedCustNum & "' "

	Set rsNote = cnnUserHasAnyUnviewedNotes.Execute(SQLCustHasNotes)
	
	If Not rsNote.EOf Then
		resultUserHasAnyUnviewedNotes = "True"
	End If



	
	cnnUserHasAnyUnviewedNotes.close
	set rsUserHasAnyUnviewedNotes = nothing
	set rsNote = nothing
	set cnnUserHasAnyUnviewedNotes= nothing	

	UserHasAnyUnviewedNotes = resultUserHasAnyUnviewedNotes

End Function
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub MarkNewNoteNoteTypeForUserAsRead(passedNoteTypeIntRecID, passedCustID)

	SQLMarkNewNoteNoteTypeForUserAsRead = "SELECT * FROM AR_CustomerNotesUserViewed Where CustID ='" & passedCustID & "' AND UserNo = " & Session("Userno") & " AND NoteTypeIntRecID = " & passedNoteTypeIntRecID
	
	Set cnnMarkNewNoteNoteTypeForUserAsRead = Server.CreateObject("ADODB.Connection")
	cnnMarkNewNoteNoteTypeForUserAsRead.open (Session("ClientCnnString"))
	Set rMarkNewNoteNoteTypeForUserAsRead = Server.CreateObject("ADODB.Recordset")
	rMarkNewNoteNoteTypeForUserAsRead.CursorLocation = 3 
	Set rMarkNewNoteNoteTypeForUserAsRead = cnnMarkNewNoteNoteTypeForUserAsRead.Execute(SQLMarkNewNoteNoteTypeForUserAsRead)

	If rMarkNewNoteNoteTypeForUserAsRead.EOF Then ' Nothing there so we need to insert
		SQLMarkNewNoteNoteTypeForUserAsRead = "INSERT INTO AR_CustomerNotesUserViewed (CustID, UserNo, Category, NoteTypeIntRecID) VALUES ('" & passedCustID & "'," & Session("UserNo") & ",-2," & passedNoteTypeIntRecID & ")"
	Else
		SQLMarkNewNoteNoteTypeForUserAsRead = "UPDATE AR_CustomerNotesUserViewed Set DateLastViewed = getdate() Where CustID ='" & passedCustID & "' AND UserNo = " & Session("UserNo") & " AND NoteTypeIntRecID = " & passedNoteTypeIntRecID
	End If
	
	Set rMarkNewNoteNoteTypeForUserAsRead = cnnMarkNewNoteNoteTypeForUserAsRead.Execute(SQLMarkNewNoteNoteTypeForUserAsRead)
		
	cnnMarkNewNoteNoteTypeForUserAsRead.close
	set rMarkNewNoteNoteTypeForUserAsRead = nothing
	set cnnMarkNewNoteNoteTypeForUserAsRead= nothing	

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Function GetCustRegionIntRecIDByCustID(passedCustID)

	Set cnnGetCustRegionIntRecIDByCustID  = Server.CreateObject("ADODB.Connection")
	cnnGetCustRegionIntRecIDByCustID.open Session("ClientCnnString")
	Set rsGetCustRegionIntRecIDByCustID  = Server.CreateObject("ADODB.Recordset")

	resultGetCustRegionIntRecIDByCustID = ""
		
	SQLGetCustRegionIntRecIDByCustID  = "SELECT City,[State],Zip FROM AR_Customer WHERE CustNum = '" & passedCustID & "'"
	 
	Set rsGetCustRegionIntRecIDByCustID = cnnGetCustRegionIntRecIDByCustID.Execute(SQLGetCustRegionIntRecIDByCustID)
	
	If NOT rsGetCustRegionIntRecIDByCustID.EOF Then
	
		'Get what we need: City, StateForCities, State, Zip Code
		Zip = rsGetCustRegionIntRecIDByCustID("Zip")
		City = rsGetCustRegionIntRecIDByCustID("City")
		If LEN(City) > 0 Then City = Replace(City,"'","''")
		State = rsGetCustRegionIntRecIDByCustID("State")
		If LEN(State) > 0 Then State = Replace(State,"'","''")
		
		'Zip First
		If LEN(Zip) > 0 <> "" Then
			SQLGetCustRegionIntRecIDByCustID  = "SELECT InternalRecordIdentifier FROM AR_Regions WHERE "
			SQLGetCustRegionIntRecIDByCustID  = SQLGetCustRegionIntRecIDByCustID  & "({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(ZipOrPostalCodes1, ' ,', ','), ', ', ','), ',') }, "
			SQLGetCustRegionIntRecIDByCustID  = SQLGetCustRegionIntRecIDByCustID  & "REPLACE(REPLACE(ZipOrPostalCodes2, ' ,', ','), ', ', ',')) } LIKE '%" & Zip & "%')"

			'Response.Write("<br>" & SQLGetCustRegionIntRecIDByCustID)
			
			Set rsGetCustRegionIntRecIDByCustID = cnnGetCustRegionIntRecIDByCustID.Execute(SQLGetCustRegionIntRecIDByCustID)

			If Not rsGetCustRegionIntRecIDByCustID.EOF Then resultGetCustRegionIntRecIDByCustID = rsGetCustRegionIntRecIDByCustID("InternalRecordIdentifier")
		End If
		
		If resultGetCustRegionIntRecIDByCustID = "" Then ' No results yet
			If LEN(City) > 0 Then
				'If it is not set, now look for the city and state
				SQLGetCustRegionIntRecIDByCustID  = "SELECT InternalRecordIdentifier FROM AR_Regions WHERE "
				SQLGetCustRegionIntRecIDByCustID  = SQLGetCustRegionIntRecIDByCustID  & "({ fn CONCAT({ fn CONCAT({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(Cities1, ' ,', ','), ', ', ','), ',') }, "
				SQLGetCustRegionIntRecIDByCustID  = SQLGetCustRegionIntRecIDByCustID  & "REPLACE(REPLACE(Cities2, ' ,', ','), ', ', ',')) }, ',') }, REPLACE(REPLACE(Cities3, ' ,', ','), ', ', ',')) "
				SQLGetCustRegionIntRecIDByCustID  = SQLGetCustRegionIntRecIDByCustID  &	"} LIKE '%," & City & "%' OR "
				SQLGetCustRegionIntRecIDByCustID  = SQLGetCustRegionIntRecIDByCustID  & "{ fn CONCAT({ fn CONCAT({ fn CONCAT({ fn CONCAT(REPLACE(REPLACE(Cities1, ' ,', ','), ', ', ','), ',') }, "
				SQLGetCustRegionIntRecIDByCustID  = SQLGetCustRegionIntRecIDByCustID  & "REPLACE(REPLACE(Cities2, ' ,', ','), ', ', ',')) }, ',') }, REPLACE(REPLACE(Cities3, ' ,', ','), ', ', ',')) "
				SQLGetCustRegionIntRecIDByCustID  = SQLGetCustRegionIntRecIDByCustID  & "} LIKE '%" & City & ",%') "
				SQLGetCustRegionIntRecIDByCustID  = SQLGetCustRegionIntRecIDByCustID  & "AND (StateForCities = '" & State & "')"
		
				''Response.Write("<br>" & SQLGetCustRegionIntRecIDByCustID)
				
				Set rsGetCustRegionIntRecIDByCustID = cnnGetCustRegionIntRecIDByCustID.Execute(SQLGetCustRegionIntRecIDByCustID)
		
				If Not rsGetCustRegionIntRecIDByCustID.EOF Then resultGetCustRegionIntRecIDByCustID = rsGetCustRegionIntRecIDByCustID("InternalRecordIdentifier")
			End If
		End If
		
		If resultGetCustRegionIntRecIDByCustID = "" Then ' No results yet
			If LEN(State) > 0 Then
				'If it is not set, now look for the state
				SQLGetCustRegionIntRecIDByCustID  = "SELECT InternalRecordIdentifier FROM AR_Regions WHERE "
				SQLGetCustRegionIntRecIDByCustID  = SQLGetCustRegionIntRecIDByCustID  & "REPLACE(REPLACE(StatesOrProvinces , ' ,', ','), ', ', ',') "
				SQLGetCustRegionIntRecIDByCustID  = SQLGetCustRegionIntRecIDByCustID  &	" LIKE '%" & State& "%'"

				'Response.Write("<br>" & SQLGetCustRegionIntRecIDByCustID)		
				Set rsGetCustRegionIntRecIDByCustID = cnnGetCustRegionIntRecIDByCustID.Execute(SQLGetCustRegionIntRecIDByCustID)
		
				If Not rsGetCustRegionIntRecIDByCustID.EOF Then resultGetCustRegionIntRecIDByCustID = rsGetCustRegionIntRecIDByCustID("InternalRecordIdentifier")
			End If
	 	End If
	 	
	End IF
	
	set rsGetCustRegionIntRecIDByCustID = Nothing
	cnnGetCustRegionIntRecIDByCustID.Close	
	set cnnGetCustRegionIntRecIDByCustID = Nothing
	
	If resultGetCustRegionIntRecIDByCustID = "" THen resultGetCustRegionIntRecIDByCustID = 0

	GetCustRegionIntRecIDByCustID  = resultGetCustRegionIntRecIDByCustID 
	
End Function

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Function GetRegionNameByRegionIntRecID(passedRegionIntRecID)

	Set cnnGetRegionNameByRegionIntRecID  = Server.CreateObject("ADODB.Connection")
	cnnGetRegionNameByRegionIntRecID.open Session("ClientCnnString")
	Set rsGetRegionNameByRegionIntRecID  = Server.CreateObject("ADODB.Recordset")

	resultGetRegionNameByRegionIntRecID = ""
		
	SQLGetRegionNameByRegionIntRecID  = "SELECT Region FROM AR_Regions WHERE InternalRecordIdentifier = " & passedRegionIntRecID

	'Response.Write("<br>" & SQLGetRegionNameByRegionIntRecID)
	
	Set rsGetRegionNameByRegionIntRecID = cnnGetRegionNameByRegionIntRecID.Execute(SQLGetRegionNameByRegionIntRecID)

	If Not rsGetRegionNameByRegionIntRecID.EOF Then resultGetRegionNameByRegionIntRecID = rsGetRegionNameByRegionIntRecID("Region")
	
	set rsGetRegionNameByRegionIntRecID = Nothing
	cnnGetRegionNameByRegionIntRecID.Close	
	set cnnGetRegionNameByRegionIntRecID = Nothing
	
	If resultGetRegionNameByRegionIntRecID = "" THen resultGetRegionNameByRegionIntRecID = 0

	GetRegionNameByRegionIntRecID  = resultGetRegionNameByRegionIntRecID 
	
End Function

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Function GetCustRefDesc2ByReferralCode(passedReferralCode)

	Set cnnGetCustRefDesc2ByReferralCode  = Server.CreateObject("ADODB.Connection")
	cnnGetCustRefDesc2ByReferralCode.open Session("ClientCnnString")

	resultGetCustRefDesc2ByReferralCode = ""
		
	SQLGetCustRefDesc2ByReferralCode  = "SELECT * FROM Referal WHERE ReferalCode = " &  passedReferralCode
	 
	Set rsGetCustRefDesc2ByReferralCode  = Server.CreateObject("ADODB.Recordset")
	rsGetCustRefDesc2ByReferralCode.CursorLocation = 3 
	
	rsGetCustRefDesc2ByReferralCode.Open SQLGetCustRefDesc2ByReferralCode,cnnGetCustRefDesc2ByReferralCode 
			
	If Not rsGetCustRefDesc2ByReferralCode.Eof Then			
		resultGetCustRefDesc2ByReferralCode = rsGetCustRefDesc2ByReferralCode("Description2")
	End If
	
	rsGetCustRefDesc2ByReferralCode.Close
	set rsGetCustRefDesc2ByReferralCode = Nothing
	cnnGetCustRefDesc2ByReferralCode.Close	
	set cnnGetCustRefDesc2ByReferralCode = Nothing
	
	GetCustRefDesc2ByReferralCode2 = resultGetCustRefDesc2ByReferralCode
	
End Function


%>