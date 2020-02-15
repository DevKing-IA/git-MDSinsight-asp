<%
'Func MUV_Init ()
'Func MUV_Write (passedTag,passedValue)
'Func MUV_ReadAll ()
'Func MUV_Inspect (passedTAG)
'Func MUV_Remove (passedTag)
'Func MUV_Read (passedTag)
'Func MUV_ReadAndRemove (passedTag)
'Func DateOfLastPurchase(passedCust,passedSKU)
'Func HasCustomLoginPage(passedCustID)
'Func GetCustNameByCustNum(passedCustID)
'Func GetCustNumByCustIntRecID(passedCustInRecID)
'Func GetCustNameByCustIntRecID(passedCustIntRecID)
'Func GetSalesmanNameBySlsmnSequence(passedSlsmnSeq)
'Func GetSalesmanNameAndEmailBySlsmnSequence(passedSlsmnSeq)
'Func GetReferralNameByCode(passedReferralCode)
'Func GetCustTypeByCode(passedType)
'Func GetCategoryByID(passedCatID)
'Func GetCategoryIDByProdSKU(passedSKU)
'Func GetCustomerClassNameByID(passedCustClassCode)
'Func GetCustomerClassCodeByCustID(passedCustID)
'Func NumberOfLeakingCategories(passedCustID)
'Func GetCurrentPeriodAndYear()
'Func GetLastClosedPeriodAndYear()
'Func GetPriorClosedPeriodAndYear()
'Func GetLastClosedPeriod()
'Func GetLastClosedPeriodSeqNum()
'Func GetLastClosedReportPeriodIntRecID()
'Func GetCurrentReportPeriodIntRecID()
'Func GetFirstReportPeriodYear()
'Func GetFirstReportPeriodThisYearIntRecID()
'Func GetPriorClosedPeriodSeqNum()
'Func DiscrepancyDollars_vs_P3P(passedCustID)
'Func GetPeriodAndYearBySeq(passedBillPerSeq)
'Func GetPeriodYearBySeq(passedBillPerSeq)
'Func GetPeriodBySeq(passedBillPerSeq)
'Func GetPeriodBeginDateBySeq(passedBillPerSeq)
'Func GetPeriodEndDateBySeq(passedBillPerSeq)
'Func GetPeriodAndYearByIntRecID(passedReportPeriodIntRecID)
'Func GetPeriodYearByIntRecID(passedReportPeriodIntRecID)
'Func GetPeriodByIntRecID(passedReportPeriodIntRecID)
'Func GetPeriodBeginDateByIntRecID(passedReportPeriodIntRecID)
'Func GetPeriodEndDateByIntRecID(passedReportPeriodIntRecID)
'Func GetProdDescriptionFromInvDetsByPartnum(passedPartnum)
'Func GetPOSTParams(passedElementName)
'Func GetFromAddressByMsgID(passedMsgID)
'Func GetFromAddressByRecID(passedRecID)
'Func GetServiceTicketOpenDateTime(passedTicketNumber)
'Func GetHoldServiceTicketSubmittedDateTime(passedTicketNumber)
'Func GetServiceTicketCloseDateTime(passedTicketNumber)
'Func GetServiceTicketStatus(passedTicketNumber)
'Func GetNumberOfTicketsByDate(passedDate,passedMemoType)
'Func GetNumberOfNONDispatchedTicketsByDate(passedDate,passedMemoType)...
'Func FormatAsSortableDateTime(passedDateTime)
'Func ElapsedTimeCalcMethod()
'Func ServiceCallElapsedMinutes(passedMemoNumber)
'Func NumberofWorkMinutes(passedDate,passedTime,passedNormalBizDayStartTime,passedNormalBizDayEndTime)
'Func NumberofWorkMinutes_DateOpened(passedDate,passedTime,passedNorma...
'Func NumberofWorkMinutes_DateClosed(passedDate,passedTime,passedNorma...
'Func NumberofWorkMinutes_FullDay(passedDate,passedNormalBizDayStartTi...
'Func CustAROver(passedCustID,Age30_60_90)
'Func CustHasSalesInLast90Days(passedCustID)
'Func FormattedCustInfoByCustNum(passedCustID)
'Func GetCustTypeCodeByCustID(passedCustID)
'Func NumberOfOpenServiceCalls()
'Func NumberOfHoldServiceCalls()
'Func NumberOfServiceCallsNotDispatched()
'Func MarkAlertEmailSent(passedTicketNumber)
'Func MarkHoldAlertEmailSent(passedTicketNumber)
'Func MarkEscalationAlertEmailSent(passedTicketNumber)
'Func MarkHoldEscalationAlertEmailSent(passedTicketNumber)
'Func AlertEmailSent(passedTicketNumber)
'Func HoldAlertEmailSent(passedTicketNumber)
'Func EscalationAlertEmailSent(passedTicketNumber)
'Func HoldEscalationAlertEmailSent(passedTicketNumber)
'Func HoldEscalationAlertEmailSent(passedTicketNumber)
'Func BusinessDayStartorEndTime(passedSorE)
'Func NumOpenCallsByAcct(passedCustNum)
'Func GetTerm(passedGenericTerm)
'Func isMixedCase( str )
'Func TicketWasOnHold(passedTicketNumber)
'Func NumberOfArchivedNotes(passedCustNum)
'Func NumberOfCurrentNotes(passedCustNum)
'Func NumberOfAttachmentsNotes(passedCustNum)
'Func NumberOfServiceTicketsEver(passedCustNum)
'Func NumberOfServiceTicketsOpenForCust(passedCustNum)
'Func GetServiceTicketCurrentStage(passedTicketNumber)
'Func GetServiceTicketDispatchedTech(passedTicketNumber)
'Func GetServiceTicketDispatchedDateTime(passedTicketNumber)
'Func NumberOfServiceTicketsHOLDForCust(passedCustNum)
'Func GetServiceTicketSTAGEDateTime(passedTicketNumber,passedStage)
'Func GetServiceTicketSTAGEUser(passedTicketNumber,passedStage)
'Func TicketIsUrgent(passedTicketNumber)
'Func TicketOriginalDispatchDateTime(passedTicketNumber)
'Func NumberOfServiceTicketsDispatchedToTech(passedServiceTechNum)
'Func NumberOfServiceTicketsAcknowledgedByTech(passedServiceTechNum)
'Func NumberOfServiceTicketsAwaitingACKFromTech(passedServiceTechNum)
'Func NumberOfServiceTicketsClosedOrRedoByTech(passedServiceTechNum)
'Func advancedDispatchIsOn()
'Func GetServiceTicketCust(passedTicketNumber)
'Func TicketIsFilterChange(passedTicketNumber)
'Func Redispatch(passedTicketNumber)
'Func AwaitingRedispatch(passedTicketNumber)
'Func RemoveFromRedispatch(passedTicketNumber)
'Func GetChainDescByChainNum(passedChainNum)
'Func FilterChangeSubmitted(passedassetNumber,passedfilterchangedate)
'Func LastTechUserNo(passedTicketNumber)
'Func filterChangeModuleOn()
'Func GetCustRouteNum(passedCustID)
'Func GetMyFilterRoutes(passedUserNo)
'Func RemoveFilterChangeSubmitted(passedassetNumber,passedfilterchange...
'Func GetCustNumberByInvoiceNum(passedInvoiceNumber)
'Func prospectingModuleIsOn()
'Func CustPendingFilterChangeInfo(passedCustid)
'Func GetPONumberByInvoiceNum(passedInvoiceNumber)
'Func GetRouteNumByInvoiceNum(passedInvoiceNumber)
'Func GetRouteNameByRouteNum(passedRouteNumber)
'Func GetTermsNumByInvoiceNum(passedInvoiceNumber)
'Func GetTermsNameByTermsNum(passedTermsNumber)
'Func GetPrimarySalesmanByInvoiceNum(passedInvoiceNumber)
'Func GetSalesmanNameBySalesmanNum(passedSalesmanNumber)
'Func GetInvoiceDateByInvoiceNum(passedInvoiceNumber)
'Func GetSpecialCommentByCustNum(passedCustomerNumber)
'Func GetInvoiceSubTotsByIvsNum(passedInvoiceNumber,passedSubtotToGet)...
'Func GetTaxableFlagByIvsHistDetSequence(passedIvsHistDetSequence)
'Func GetInvoiceNumberByIvsSeq(passedIvsSeq)
'Func GetCustNumberByInvSeq(passedIvsSeq)
'Func GetPONumberByInvSeq(passedIvsSeq)
'Func GetRouteNumByInvSeq(passedIvsSeq)
'Func GetTermsNumByInvSeq(passedIvsSeq)
'Func GetPrimarySalesmanByInvSeq(passedIvsSeq)
'Func GetInvoiceDateByInvSeq(passedIvsSeq)
'Func GetInvoiceSubTotsByInvSeq(passedIvsSeq,passedSubtotToGet)
'Func GetAddressElementByCustNum(passedCustNum,passedElement)
'Func InvoiceProfitDollars(passedInvoiceNumber)
'Func GetNumberOfLinesByInvoiceNumber(passedInvoiceNumber)
'Func GetAlertType (passedAlertNumber)
'Func TZNow()
'Func BusinessDayEnd()
'Func GetServiceTicketCompanyName(passedTicketNumber)
'Func GetExtension(FileName)
'Func GetLastInvoiceFromWebDate(passedCustID)
'Func NumberOfWorkDays(passedStartDate, passedEdnDate)
'Func GetServiceTicketLastEntryDateTime(passedTicketNumber)
'Func EZTexting_Filter1(strInput)
'Func GetUserNoBySalesPersonNo(passedSlsmnNo)
'Func GetSalesPersonNoByUserNo(passedUserNo)
'Func NAGMasterON()
'Func fmt_mmddyy(passedinput)
'Func padDate(n, totalDigits) 
'Sub Write_API_AuditLog_Entry(passedIdentity,passedLogEntry,passedMode,passedModule)
'Func GetCompanyCountry()
'Func GetCustTypeByCustTypeNum(passedCustTypeNum)

Function EZTexting_Filter1(strInput)

	PassedstrInput = strInput
	PermittedCharacters = ".,:;!?()~=+-_/@$#&%'abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & chr(34)
	newChars = ""
	
	For x = 1 to Len(PassedstrInput)
		If Instr(PermittedCharacters,Mid(PassedstrInput,x,1)) <> 0 Then
			newChars = newChars & Mid(PassedstrInput,x,1)
		Else
			newChars = newChars & " " 
		End If
	Next 
	
	newChars = Replace(newChars ,"LOAN","LN")
	newChars = Replace(newChars ,"STOP","STP")

	EZTexting_Filter1 = newChars 
	
End Function

'***********************************
'***********************************
' Start SessionMultiUseVar Functions
'***********************************
'***********************************

Function MUV_Init ()
	Session("MultiUseVar1") = ""
    MUV_Init= 0 'Just a dummy value
End Function

Function MUV_Write (passedTag,passedValue)

	'Writes a tag & value to the multi use var
	'Will see if the tag exists first & if so, removes it before the write
	tmppassedTag = Ucase(passedTag)
	
	If MUV_Inspect(tmppassedTag) = False Then
		Session("MultiUseVar1") = Session("MultiUseVar1") & "{" & tmppassedTag& "}"
		Session("MultiUseVar1") = Session("MultiUseVar1") & passedValue
    	Session("MultiUseVar1") = Session("MultiUseVar1") & "{/" & tmppassedTag& "}"
    Else
    	dummy = MUV_Remove(tmppassedTag)
    	Session("MultiUseVar1") = Session("MultiUseVar1") & "{" & tmppassedTag& "}"
		Session("MultiUseVar1") = Session("MultiUseVar1") & passedValue
    	Session("MultiUseVar1") = Session("MultiUseVar1") & "{/" & tmppassedTag& "}"
    End If
    
    MUV_Write = 0 'Just a dummy value
    
End Function

Function MUV_ReadAll ()

    MUV_ReadAll = Session("MultiUseVar1") 
    
End Function

Function MUV_Inspect (passedTAG)

	' Check the multi use var to see if the passedTag exists
	' Simply returns True of False
	tmppassedTag = Ucase(passedTAG)
		
	If InStr(Session("MultiUseVar1"),"{"& tmppassedTag &"}") <> 0 Then result = True Else result = False
	
	MUV_Inspect = result
    
End Function

Function MUV_Remove (passedTag)

	'Removes the passedTag and value if present
	tmppassedTag = Ucase(passedTAG)
	''Response.Write("To remove:" & tmppassedTag & "<br>")
	''Response.Write("B remove:" & Session("MultiUseVar1") & "<br>")
	
	StartTagPOS = InStr(Session("MultiUseVar1"),"{"& tmppassedTag &"}")
	EndTagPOS = InStr(Session("MultiUseVar1"),"{/"& tmppassedTag &"}")
	LastTagPOS = InStr(Session("MultiUseVar1"),"{/"& tmppassedTag &"}") + Len("{/"& tmppassedTag &"}") - 1
	
	If StartTagPOS <> 0 Then
		TEMPVAR = Left(Session("MultiUseVar1"),StartTagPOS-1)
		TEMPVAR = TEMPVAR & Right(Session("MultiUseVar1"),Len(Session("MultiUseVar1")) - LastTagPOS)
		Session("MultiUseVar1") = TEMPVAR
	End IF	

	''Response.Write("A remove:" & Session("MultiUseVar1") & "<br>")

	MUV_Remove = 0 ' Just a dummy value

End Function

Function MUV_Read (passedTag)

	'Retruns the value of the passedtag if present
	tmppassedTag = Ucase(passedTAG)
	
	If MUV_Inspect(tmppassedTag) <> False Then
	
		StartTagPOS = InStr(Session("MultiUseVar1"),"{"& tmppassedTag &"}")
		EndTagPOS = InStr(Session("MultiUseVar1"),"{/"& tmppassedTag &"}")
		LastTagPOS = InStr(Session("MultiUseVar1"),"{/"& tmppassedTag &"}") + Len("{/"& tmppassedTag &"}") - 1
		
		result = Mid(Session("MultiUseVar1"),StartTagPOS+Len("{"& tmppassedTag &"}"),(EndTagPOS - (StartTagPOS + Len(tmppassedTag)))-2)
	
	Else
		result = ""
	End If
	
	MUV_Read = result

End Function

Function MUV_ReadAndRemove (passedTag)

	'Retruns the value of the passedtag if present
	tmppassedTag = Ucase(passedTAG)
	
	If MUV_Inspect(tmppassedTag) <> False Then
	
		StartTagPOS = InStr(Session("MultiUseVar1"),"{"& tmppassedTag &"}")
		EndTagPOS = InStr(Session("MultiUseVar1"),"{/"& tmppassedTag &"}")
		LastTagPOS = InStr(Session("MultiUseVar1"),"{/"& tmppassedTag &"}") + Len("{/"& tmppassedTag &"}") - 1
		
		result = Mid(Session("MultiUseVar1"),StartTagPOS+Len("{"& tmppassedTag &"}"),(EndTagPOS - (StartTagPOS + Len(tmppassedTag)))-2)
	
	Else
		result = ""
	End If
	
	MUV_ReadAndRemove = result

	dummy = MUV_Remove(passedTag)
	
End Function

'***********************************
'***********************************
'  End SessionMultiUseVar Functions
'***********************************
'***********************************



Sub CreateAuditLogEntry(passedElementOrEventName,passedElementOrEventNav,passedMajorMinor,passedSettingChange,passedDescription) 

	'Creates an entry in SC_AuditLog
	
	passedDescription= replace(passedDescription,"'","")
	
	Dim UserIPAddress
	
	UserIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If UserIPAddress = "" Then
		UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
	End If

	NameToLog = ""
	NameToLog = MUV_Read("DisplayName")
	If passedElementOrEventName = "Service Realtime Alert Sent" or passedElementOrEventName = "Service Realtime Escalation Alert Sent" Then NameToLog = "" ' So it will use SYSTEM
	If NameToLog = "" Then NameToLog = "System"
	
	SQL = "INSERT INTO SC_AuditLog (AuditElementOrEventNav,AuditElementOrEventName,AuditUserEmail, "
	SQL = SQL & "AuditDescription,AuditSettingChange,AuditIPAddress,AuditUserDisplayName,AuditMajorMinor)"
	SQL = SQL &  " VALUES ('" & passedElementOrEventNav & "'"
	SQL = SQL & ",'"  & passedElementOrEventName & "'"
	SQL = SQL & ",'"  & Session("userEmail") & "'"
	SQL = SQL & ",'"  & passedDescription & "'"		
	SQL = SQL & ","  & passedSettingChange
	SQL = SQL & ",'"  & UserIPAddress & "'"
	SQL = SQL & ",'"  & NameToLog & "'"
	SQL = SQL & ",'"  & passedMajorMinor & "')"
	
	'response.write(SQL)
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))

	Set rs8 = Server.CreateObject("ADODB.Recordset")
	rs8.CursorLocation = 3 
	Set rs8 = cnn8.Execute(SQL)
	set rs8 = Nothing
	
End Sub 

Sub CreateINSIGHTAuditLogEntry(passedIdentity,passedLogEntry,passedMode) 

	'Before we do anything, get the next entry THREAD
	Set cnnCreateINSIGHTAuditLogEntry = Server.CreateObject("ADODB.Connection")
	cnnCreateINSIGHTAuditLogEntry.open ("Driver={SQL Server};Server=66.201.99.15;Database=_BIInsight;Uid=biinsight;Pwd=Z32#kje4217;")
	Set rsCreateINSIGHTAuditLogEntry = Server.CreateObject("ADODB.Recordset")
	rsCreateINSIGHTAuditLogEntry.CursorLocation = 3 
		
	Set rsCreateINSIGHTAuditLogEntry = cnnCreateINSIGHTAuditLogEntry.Execute("SELECT MAX(EntryThread) As ThreadCount FROM API_AuditLog")

	If Not rsCreateINSIGHTAuditLogEntry.EOF Then
		ThreadVar =rsCreateINSIGHTAuditLogEntry("ThreadCount") + 1
	Else
		'Very unlikely
		ThreadVar = 1
	End If

	Set rsCreateINSIGHTAuditLogEntry = Nothing
	cnnCreateINSIGHTAuditLogEntry.Close
	Set cnnCreateINSIGHTAuditLogEntry = Nothing

	LogEntry2 = ""
	LogEntry3 = ""
	
	on error resume next
	'Creates an entry in API_AuditLog
	
	passedLogEntry= replace(passedLogEntry,"'","")
	
	If Len(passedLogEntry) > 16000 Then 
		LogEntry3 = Mid(passedLogEntry,16001,len(passedLogEntry)-16000)
		If Len(LogEntry3) > 8000 Then LogEntry3=Left(LogEntry3,8000) 
		LogEntry2 = Mid(passedLogEntry,8001,8000)
		passedLogEntry = Left(passedLogEntry,8000)
	ElseIf Len(passedLogEntry) > 8000 Then 
		LogEntry2 = Mid(passedLogEntry,8001,len(passedLogEntry)-8000) 
		If Len(LogEntry2) > 8000 Then LogEntry2 = Left(LogEntry2,8000)
		passedLogEntry = Left(passedLogEntry,8000)
	End If
	
	
	If LogEntry3 <> "" Then
		SQL = "INSERT INTO API_AuditLog([Identity],LogEntry,LogEntryPart2,LogEntryPart3,Mode,Serno,Thread)"
		SQL = SQL &  " VALUES ('" & passedIdentity & "'"
		SQL = SQL & ",'"  & passedLogEntry & "'"
		SQL = SQL & ",'"  & LogEntry2 & "'"
		SQL = SQL & ",'"  & LogEntry3 & "'"
		SQL = SQL & ",'"  & passedMode & "'"
		SQL = SQL & ",'"  & MUV_READ("SERNO") & "',ThreadVar)"

	ElseIf LogEntry2 = "" Then
		SQL = "INSERT INTO API_AuditLog([Identity],LogEntry,Mode,Serno,Thread)"
		SQL = SQL &  " VALUES ('" & passedIdentity & "'"
		SQL = SQL & ",'"  & passedLogEntry & "'"
		SQL = SQL & ",'"  & passedMode & "'"
		SQL = SQL & ",'"  & MUV_READ("SERNO") & "',ThreadVar)"

	Else
		SQL = "INSERT INTO API_AuditLog([Identity],LogEntry,LogEntryPart2,Mode,Serno,Thread)"
		SQL = SQL &  " VALUES ('" & passedIdentity & "'"
		SQL = SQL & ",'"  & passedLogEntry & "'"
		SQL = SQL & ",'"  & LogEntry2 & "'"
		SQL = SQL & ",'"  & passedMode & "'"
		SQL = SQL & ",'"  & MUV_READ("SERNO") & "',ThreadVar)"
	End If
	'response.write(SQL)
	
	Set cnnCreateINSIGHTAuditLogEntry = Server.CreateObject("ADODB.Connection")
	cnnCreateINSIGHTAuditLogEntry.open ("Driver={SQL Server};Server=66.201.99.15;Database=_BIInsight;Uid=biinsight;Pwd=Z32#kje4217;")
	Set rsCreateINSIGHTAuditLogEntry = Server.CreateObject("ADODB.Recordset")
	rsCreateINSIGHTAuditLogEntry.CursorLocation = 3 
		
	Set rsCreateINSIGHTAuditLogEntry = cnnCreateINSIGHTAuditLogEntry.Execute(SQL)

	Set rsCreateINSIGHTAuditLogEntry = Nothing
	cnnCreateINSIGHTAuditLogEntry.Close
	Set cnnCreateINSIGHTAuditLogEntry = Nothing

	On error goto 0
	
End Sub 






'*****************************************************************************************************************
'*****************************************************************************************************************


Function DateOfLastPurchase(passedCust,passedSKU)

	'Takes the passed Customer ID & SKU & reutrns the date of last purchase
	
	resultBoost=0
	LastPurchaseDate = ""

	
	Set rsBoost7 = Server.CreateObject("ADODB.Recordset")
	rsBoost7.CursorLocation = 3 

	SQL = "SELECT * FROM InvoiceHistoryDetail "
	SQL = SQL & "WHERE CustNum = '" & passedCust & "' AND partnum = '" & passedSKU &  "' "
	SQL = SQL & " AND ivsDate > '" & DateAdd("d", -365, Now()) & "' order by ivsDate desc"
	
		
	rsBoost7.Open SQL, Session("ClientCnnString")
	
	
	If Not rsBoost7.EOF Then LastPurchaseDate = rsBoost7("ivsDate")
		
			
	rsBoost7.Close
			
		
	resultBoost = LastPurchaseDate 
	
	DateOfLastPurchase= resultBoost
	
End Function



Function HasCustomLoginPage(passedCustID)

	resultLogin = 0
		
    '**************************************************************************
    'Get Temporary Connection String To Look Up User Settings
    '**************************************************************************
    
	SQL = "SELECT * FROM tblServerInfo where clientKey='"& passedCustID &"'"
	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open "Driver={SQL Server};Server=66.201.99.15;Database=_BIInsight;Uid=biinsight;Pwd=Z32#kje4217;"

	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3

	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection
	'and go back to login with QueryString
	If Recordset.recordcount <= 0 then
		Recordset.close
		Connection.close
		set Recordset=nothing
		set Connection=nothing
		Response.Redirect(baseURL)
	Else
		tmpCnnString = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
		tmpCnnString = tmpCnnString  & ";Database=" & Recordset.Fields("dbCatalog")
		tmpCnnString  = tmpCnnString  & ";Uid=" & Recordset.Fields("dbLogin")
		tmpCnnString  = tmpCnnString  & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
		tmpSQL_Owner = Recordset.Fields("dbLogin")
		resultLogin = Recordset.Fields("customLoginPage")
		Recordset.close
		Connection.close	
	End If	
	
	Session("ClientCnnString") = tmpCnnString 
		
    
    '**************************************************************************
    'Check To See If This Client ID has a custom login page
    '**************************************************************************
	
	HasCustomLoginPage = resultLogin 
	
End Function


Function GetCustNameByCustNum(passedCustID)

	resultGetCustNameByCustNum=""

	Set cnnGetCustNameByCustNum = Server.CreateObject("ADODB.Connection")
	cnnGetCustNameByCustNum.open Session("ClientCnnString")
	Set rsGetCustNameByCustNum = Server.CreateObject("ADODB.Recordset")
	rsGetCustNameByCustNum.CursorLocation = 3 
	

	SQLGetCustNameByCustNum = "Select * from " & Session("SQL_Owner") & ".AR_Customer where CustNum= '" & passedCustID & "'"
	 

	Set rsGetCustNameByCustNum= cnnGetCustNameByCustNum.Execute(SQLGetCustNameByCustNum)
	
	
	If not rsGetCustNameByCustNum.eof then resultGetCustNameByCustNum = rsGetCustNameByCustNum("Name")
	
	Set rsGetCustNameByCustNum= Nothing
	cnnGetCustNameByCustNum.Close
	Set cnnGetCustNameByCustNum= Nothing
	
	GetCustNameByCustNum = resultGetCustNameByCustNum
	
End Function


Function GetCustNumByCustIntRecID(passedCustInRecID)

	resultGetCustNumByCustIntRecID=""

	Set cnnGetCustNumByCustIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetCustNumByCustIntRecID.open Session("ClientCnnString")
	Set rsGetCustNumByCustIntRecID = Server.CreateObject("ADODB.Recordset")
	rsGetCustNumByCustIntRecID.CursorLocation = 3 
	

	SQLGetCustNumByCustIntRecID = "SELECT * FROM " & Session("SQL_Owner") & ".AR_Customer WHERE InternalRecordIdentifier = " & passedCustInRecID
	 

	Set rsGetCustNumByCustIntRecID= cnnGetCustNumByCustIntRecID.Execute(SQLGetCustNumByCustIntRecID)
	
	
	If not rsGetCustNumByCustIntRecID.eof then resultGetCustNumByCustIntRecID = rsGetCustNumByCustIntRecID("CustNum")
	
	Set rsGetCustNumByCustIntRecID= Nothing
	cnnGetCustNumByCustIntRecID.Close
	Set cnnGetCustNumByCustIntRecID= Nothing
	
	GetCustNumByCustIntRecID = resultGetCustNumByCustIntRecID
	
End Function




Function GetSalesmanNameBySlsmnSequence(passedSlsmnSeq)

	resultGetSalesmanNameBySlsmnSequence=""
		
	Set cnnGetSalesmanNameBySlsmnSequence = Server.CreateObject("ADODB.Connection")
	cnnGetSalesmanNameBySlsmnSequence.open Session("ClientcnnString")
	Set rsGetSalesmanNameBySlsmnSequence = Server.CreateObject("ADODB.Recordset")
	rsGetSalesmanNameBySlsmnSequence.CursorLocation = 3 
	
		
	If IsNull(passedSlsmnSeq) or passedSlsmnSeq = "" then
	
		resultGetSalesmanNameBySlsmnSequence=""
		
	Else
	
		SQLGetSalesmanNameBySlsmnSequence = "Select * from Salesman where SalesmanSequence = " & passedSlsmnSeq

		Set rsGetSalesmanNameBySlsmnSequence = cnnGetSalesmanNameBySlsmnSequence.Execute(SQLGetSalesmanNameBySlsmnSequence)
		
		If not rsGetSalesmanNameBySlsmnSequence.eof then
			resultGetSalesmanNameBySlsmnSequence = rsGetSalesmanNameBySlsmnSequence("Name")
		End If
		
		cnnGetSalesmanNameBySlsmnSequence.Close
		set rsGetSalesmanNameBySlsmnSequence = Nothing
		set cnnGetSalesmanNameBySlsmnSequence = Nothing
		
	End If
	
	GetSalesmanNameBySlsmnSequence= resultGetSalesmanNameBySlsmnSequence
	
End Function

Function GetSalesmanNameAndEmailBySlsmnSequence(passedSlsmnSeq)

	If IsNull(passedSlsmnSeq) or passedSlsmnSeq = "" Then
		resultGetSalesmanNameAndEmailBySlsmnSequence=""
	Else

		Set cnnGetSalesmanNameAndEmailBySlsmnSequence = Server.CreateObject("ADODB.Connection")
		cnnGetSalesmanNameAndEmailBySlsmnSequence.open Session("ClientCnnString")
	
		SQLGetSalesmanNameAndEmailBySlsmnSequence = "Select Name, emailAddress from Salesman where SalesmanSequence= " & passedSlsmnSeq
		 
		Set rsGetSalesmanNameAndEmailBySlsmnSequence = Server.CreateObject("ADODB.Recordset")
		rsGetSalesmanNameAndEmailBySlsmnSequence.CursorLocation = 3 
		Set rsGetSalesmanNameAndEmailBySlsmnSequence= cnnGetSalesmanNameAndEmailBySlsmnSequence.Execute(SQLGetSalesmanNameAndEmailBySlsmnSequence)
		
		
		If not rsGetSalesmanNameAndEmailBySlsmnSequence.eof then
			Nam = rsGetSalesmanNameAndEmailBySlsmnSequence("Name") 
			Eml = "Not Found"
			If Len(rsGetSalesmanNameAndEmailBySlsmnSequence("emailAddress")) > 1 Then Eml = rsGetSalesmanNameAndEmailBySlsmnSequence("emailAddress")
			resultGetSalesmanNameAndEmailBySlsmnSequence = Nam & "~" & Eml
		Else
			resultGetSalesmanNameAndEmailBySlsmnSequence="*Not Found*"
		End If
		
		set rsGetSalesmanNameAndEmailBySlsmnSequence= Nothing
		set cnnGetSalesmanNameAndEmailBySlsmnSequence= Nothing
		
	End If
	
	GetSalesmanNameAndEmailBySlsmnSequence = resultGetSalesmanNameAndEmailBySlsmnSequence
	
End Function


Function GetReferralNameByCode(passedReferralCode)

	If IsNull(passedReferralCode) Then passedReferralCode = 0
	
	If passedReferralCode = 0 Then
		resultGetReferralNameByCode=""
	Else
		Set cnnGetReferralNameByCode = Server.CreateObject("ADODB.Connection")
		cnnGetReferralNameByCode.open Session("ClientCnnString")
	
		SQLGetReferralNameByCode = "Select Name from Referal where ReferalCode = " & passedReferralCode
		 
		Set rsGetReferralNameByCode = Server.CreateObject("ADODB.Recordset")
		rsGetReferralNameByCode.CursorLocation = 3 
		Set rsGetReferralNameByCode= cnnGetReferralNameByCode.Execute(SQLGetReferralNameByCode)
			
		If not rsGetReferralNameByCode.eof then
			resultGetReferralNameByCode = rsGetReferralNameByCode("Name")
		Else
			resultGetReferralNameByCode = "Referral code " & passedReferralCode & " not found"
		End If
		
		set rsGetReferralNameByCode= Nothing
		set cnnGetReferralNameByCode= Nothing
	End If
		
	GetReferralNameByCode= resultGetReferralNameByCode
	
End Function

Function GetCustTypeByCode(passedType)

	If NOT IsNull(passedType) Then	
	
		Set cnn = Server.CreateObject("ADODB.Connection")
		cnn.open Session("ClientCnnString")
		
		resultBoost=""
			
		SQL = "Select Description from CustomerType where CustTypeSequence = " & passedType
		 
		Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
		rsBoost1.CursorLocation = 3 
		Set rsBoost1= cnn.Execute(SQL)
		
		
		If not rsBoost1.eof then
			resultBoost = rsBoost1("Description")
		End If
		
		set rsBoost1= Nothing
		set cnn= Nothing
	End If
	
	GetCustTypeByCode = resultBoost
	
End Function

Function GetCategoryByID(passedCatID)

	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	
	resultBoost="Cat " & passedCatID & " Not Found"
		
	SQL = "Select * from tblCategories where CategoryID = " & passedCatID
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
'	response.write(SQL)
	Set rsBoost1= cnn.Execute(SQL)
	
	
	If not rsBoost1.eof then resultBoost = rsBoost1("CategoryName")
	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetCategoryByID= resultBoost
	
End Function


Function GetCategoryIDByProdSKU(passedSKU)

	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	resultBoost =""
		
	SQL = "SELECT prodCategory FROM IC_Product where prodSKU = '" & (passedSKU) & "'"
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	
	If not rsBoost1.eof then resultBoost = rsBoost1("prodCategory")
	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetCategoryIDByProdSKU= resultBoost
	
End Function



Function GetCustomerClassNameByID(passedCustClassCode)

	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	resultBoost="Class Code " & passedCustClassCode & " Not Found"
		
	SQL = "SELECT * FROM AR_CustomerClass WHERE ClassCode = '" & passedCustClassCode & "'"
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	
	If not rsBoost1.eof then resultBoost = rsBoost1("ClassDescription")
	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetCustomerClassNameByID = resultBoost
	
End Function


Function GetCustomerClassCodeByCustID(passedCustID)

	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	resultBoost="Class Code " & passedCustClassCode & " Not Found"
		
	SQL = "SELECT * FROM IN_WebFulfillment WHERE CustID = '" & passedCustID & "'"
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	
	If not rsBoost1.eof then resultBoost = rsBoost1("CustClassCode")
	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetCustomerClassCodeByCustID = resultBoost
	
End Function



Function NumberOfLeakingCategories(passedCustID)

		'Returns the number of leaking categories for a given customer
		sCustID = passedCustID
		
		'Count The Number of Leaking Categories
		NumCatsLeaking = 0
		SQL4 = "SELECT Count (*) As Expr1 "
		SQL4 = SQL4 & "FROM " & Session("SQL_Owner") & ".CustCatPeriodSales_ReportData "
		SQL4 = SQL4 & "WHERE CustNum = '" & sCustID & "' AND"
		SQL4 = SQL4 & "((DiifThisPeriodVSLastYearDollars < " & ThresholdDollarsTPLY  & " AND DiifThisPeriodVSLastYearPercent < " & ThresholdPercentTPLY & ") OR "
		SQL4 = SQL4 & "(DiifThisPeriodVSLast3Dollars < " & ThresholdDollarsL3P & " AND DiifThisPeriodVSLast3Percent < " & ThresholdPercentL3P & ") OR "
		SQL4 = SQL4 & "(DiifThisPeriodVSLast12Dollars < " & ThresholdDollarsL12P & " AND DiifThisPeriodVSLast12Percent < " & ThresholdPercentL12P & "))"			

		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then NumCatsLeaking =  rs4("Expr1") 
		rs4.close
		
		NumberOfLeakingCategories = NumCatsLeaking 

End Function

Function GetCurrentPeriodAndYear()

		result = 0

		'Returns the current period
		SQL4 = "Select BillPerSequence from " & Session("SQL_Owner") & ".BillingPeriodHistory where CurrentIndicator = '*'"
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then BPS =  rs4("BillPerSequence") 
		rs4.close
		
		SQL4 = "Select * from " & Session("SQL_Owner") & ".BillingPeriodHistory where BillPerSequence = " & BPS
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =  rs4("Period")  & " - " & rs4("Year")
		rs4.close
		
		GetCurrentPeriodAndYear = result

End Function


Function GetLastClosedPeriodAndYear()

		result = 0

		'Returns the last closed period
		SQL4 = "Select BillPerSequence from " & Session("SQL_Owner") & ".BillingPeriodHistory where CurrentIndicator = '*'"
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then BPS =  rs4("BillPerSequence") 
		rs4.close
		
		SQL4 = "Select * from " & Session("SQL_Owner") & ".BillingPeriodHistory where BillPerSequence = " & BPS - 1
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =  rs4("Period")  & " - " & rs4("Year")
		rs4.close
		
		GetLastClosedPeriodAndYear = result

End Function

Function GetPriorClosedPeriodAndYear()

		result = 0

		'Returns the last closed period
		SQL4 = "Select BillPerSequence from " & Session("SQL_Owner") & ".BillingPeriodHistory where CurrentIndicator = '*'"
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then BPS =  rs4("BillPerSequence") 
		rs4.close
		
		SQL4 = "Select * from " & Session("SQL_Owner") & ".BillingPeriodHistory where BillPerSequence = " & BPS - 2
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =  rs4("Period")  & " - " & rs4("Year")
		rs4.close
		
		GetPriorClosedPeriodAndYear = result

End Function


Function GetLastClosedPeriod()

		result = 0

		'Returns the last closed period
		SQL4 = "Select BillPerSequence from " & Session("SQL_Owner") & ".BillingPeriodHistory where CurrentIndicator = '*'"
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then BPS =  rs4("BillPerSequence") 
		rs4.close
		
		SQL4 = "Select * from " & Session("SQL_Owner") & ".BillingPeriodHistory where BillPerSequence = " & BPS - 1
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =  rs4("Period")
		rs4.close
		
		GetLastClosedPeriod = result

End Function

Function GetLastClosedPeriodSeqNum()

		result = 0

		'Returns the last closed period
		SQL4 = "Select BillPerSequence from " & Session("SQL_Owner") & ".BillingPeriodHistory where CurrentIndicator = '*'"
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =  rs4("BillPerSequence") - 1
		rs4.close
	
		GetLastClosedPeriodSeqNum = result

End Function



Function GetLastClosedReportPeriodIntRecID()

		result = 0

		'Returns the last closed period
		SQL4 = "SELECT InternalRecordIdentifier FROM " & Session("SQL_Owner") & ".Settings_CompanyPeriods WHERE GETDATE() > BeginDate AND GETDATE() < EndDate"
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =  rs4("InternalRecordIdentifier") - 1
		rs4.close
	
		GetLastClosedReportPeriodIntRecID = result

End Function



Function GetCurrentReportPeriodIntRecID()

		result = 0

		'Returns the last closed period
		SQL4 = "SELECT InternalRecordIdentifier FROM " & Session("SQL_Owner") & ".Settings_CompanyPeriods WHERE GETDATE() > BeginDate AND GETDATE() < EndDate"
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =  rs4("InternalRecordIdentifier")
		rs4.close
	
		GetCurrentReportPeriodIntRecID = result

End Function


Function GetFirstReportPeriodThisYearIntRecID()

		result = 0
		currentYear = Year(Date())
		firstOfYear = cDate("1/1/" & currentYear)

		SQL4 = "SELECT InternalRecordIdentifier FROM " & Session("SQL_Owner") & ".Settings_CompanyPeriods WHERE '" & firstOfYear & "' >= BeginDate AND '" & firstOfYear & "' <= EndDate "
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =  rs4("InternalRecordIdentifier")
		rs4.close
	
		GetFirstReportPeriodThisYearIntRecID = result

End Function


Function GetFirstReportPeriodYear()

		result = ""

		SQL4 = "SELECT MIN(Year) as minYear FROM " & Session("SQL_Owner") & ".Settings_CompanyPeriods"
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =  rs4("minYear")
		rs4.close
	
		GetFirstReportPeriodYear = result

End Function

Function GetPriorClosedPeriodSeqNum()

		result = 0

		'Returns the last closed period
		SQL4 = "Select BillPerSequence from " & Session("SQL_Owner") & ".BillingPeriodHistory where CurrentIndicator = '*'"
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then BPS =  rs4("BillPerSequence") 
		rs4.close
		
		SQL4 = "Select * from " & Session("SQL_Owner") & ".BillingPeriodHistory where BillPerSequence = " & BPS - 2
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =  rs4("BillPerSequence") 
		rs4.close
		
		GetPriorClosedPeriodSeqNum= result

End Function

Function DiscrepancyDollars_vs_P3P(passedCustID)

		result = 0
		
		'Returns the total dollar diff between the last closed period and the P3P
		sCustID = passedCustID
		
		
		SQL4 = "SELECT Sum(DiifThisPeriodVSLast3Dollars) As SumDiff "
		SQL4 = SQL4 & "FROM " & Session("SQL_Owner") & ".CustCatPeriodSales_ReportData "
		SQL4 = SQL4 & "WHERE CustNum = '" & sCustID & "' AND"
		SQL4 = SQL4 & "(DiifThisPeriodVSLast3Dollars < " & ThresholdDollarsL3P & " AND DiifThisPeriodVSLast3Percent < " & ThresholdPercentL3P & ") "

		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then 
				result = rs4("SumDiff") 			
		Else
			result = 0
		End IF
		rs4.close
		
		If IsNull(result) Then result = 0
		
		result = FormatCurrency(result,2,,0)
		
		DiscrepancyDollars_vs_P3P = result 

End Function


Function GetPeriodAndYearBySeq(passedBillPerSeq)

		result = 0

		SQL4 = "Select * from " & Session("SQL_Owner") & ".BillingPeriodHistory where BillPerSequence = " & passedBillPerSeq
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =   rs4("Period")  & " - " & rs4("Year")
		rs4.close
		
		GetPeriodAndYearBySeq = result

End Function

Function GetPeriodYearBySeq(passedBillPerSeq)

		result = 0

		SQL4 = "Select * from " & Session("SQL_Owner") & ".BillingPeriodHistory where BillPerSequence = " & passedBillPerSeq
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result = rs4("Year")
		rs4.close
		
		GetPeriodYearBySeq = result

End Function

Function GetPeriodBySeq(passedBillPerSeq)

		result = 0

		SQL4 = "Select * from " & Session("SQL_Owner") & ".BillingPeriodHistory where BillPerSequence = " & passedBillPerSeq
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result = rs4("Period")
		rs4.close
		
		GetPeriodBySeq = result

End Function


Function GetPeriodBeginDateBySeq(passedBillPerSeq)

		result = 0

		SQL4 = "Select * from " & Session("SQL_Owner") & ".BillingPeriodHistory where BillPerSequence = " & passedBillPerSeq
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =   rs4("BeginDate")
		rs4.close
		
		GetPeriodBeginDateBySeq = result

End Function

Function GetPeriodEndDateBySeq(passedBillPerSeq)

		result = 0

		SQL4 = "Select * from " & Session("SQL_Owner") & ".BillingPeriodHistory where BillPerSequence = " & passedBillPerSeq
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =   rs4("EndDate")
		rs4.close
		
		GetPeriodEndDateBySeq= result

End Function





Function GetPeriodAndYearByIntRecID(passedReportPeriodIntRecID)

		result = 0

		SQL4 = "Select * from " & Session("SQL_Owner") & ".Settings_CompanyPeriods where InternalRecordIdentifier = " & passedReportPeriodIntRecID
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =   rs4("Period")  & " - " & rs4("Year")
		rs4.close
		
		GetPeriodAndYearByIntRecID = result

End Function

Function GetPeriodYearByIntRecID(passedReportPeriodIntRecID)

		result = 0

		SQL4 = "Select * from " & Session("SQL_Owner") & ".Settings_CompanyPeriods where InternalRecordIdentifier = " & passedReportPeriodIntRecID
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result = rs4("Year")
		rs4.close
		
		GetPeriodYearByIntRecID = result

End Function

Function GetPeriodByIntRecID(passedReportPeriodIntRecID)

		result = 0

		SQL4 = "Select * from " & Session("SQL_Owner") & ".Settings_CompanyPeriods where InternalRecordIdentifier = " & passedReportPeriodIntRecID
		'response.write("<br><br><br>" & SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result = rs4("Period")
		rs4.close
		
		GetPeriodByIntRecID = result

End Function


Function GetPeriodBeginDateByIntRecID(passedReportPeriodIntRecID)

		result = 0

		SQL4 = "Select * from " & Session("SQL_Owner") & ".Settings_CompanyPeriods where InternalRecordIdentifier = " & passedReportPeriodIntRecID
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =   rs4("BeginDate")
		rs4.close
		
		GetPeriodBeginDateByIntRecID = result

End Function

Function GetPeriodEndDateByIntRecID(passedReportPeriodIntRecID)

		result = 0

		SQL4 = "Select * from " & Session("SQL_Owner") & ".Settings_CompanyPeriods where InternalRecordIdentifier = " & passedReportPeriodIntRecID
		'response.write(SQL4)
		Set rs4 = Server.CreateObject("ADODB.Recordset")
		rs4.CursorLocation = 3
		rs4.Open SQL4 , Session("ClientCnnString")
		If Not rs4.Eof Then result =   rs4("EndDate")
		rs4.close
		
		GetPeriodEndDateByIntRecID= result

End Function



Function GetProdDescriptionFromInvDetsByPartnum(passedPartnum)

	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	
	resultBoost="Item# " & passedPartnum& " Not Found"
		
	SQL = "Select * from " & Session("SQL_Owner") & ".InvoiceHistoryDetail where partNum = '" & passedPartnum & "'"
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	
	If not rsBoost1.eof then resultBoost = rsBoost1("prodDescription")
	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetProdDescriptionFromInvDetsByPartnum = resultBoost
	
End Function

Function GetPOSTParams(passedElementName)

	result = ""

	SQL = "SELECT * FROM Settings_AR"
		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		Select Case(ucase(trim(passedElementName)))
			Case "SERNO"
				result = rs("POST_Serno")
			Case "MODE"
				result = rs("POST_Mode")
			Case "CUSTOMERURL1"
				result = rs("POST_CustomerURL1")
			Case "CUSTOMERURL2"
				result = rs("POST_CustomerURL2")
			Case "CUSTOMERURL1ONOFF"
				result = rs("POST_CustomerURL1ONOFF")
			Case "CUSTOMERURL2ONOFF"
				result = rs("POST_CustomerURL2ONOFF")
			Case "EMAILFORNON200RESPONSES"
				result = rs("EmailForNon200Responses") 			 
			Case Else
				result = "Error Bad Parameter Name"
		End Select
	End If
	

	SQL = "SELECT * FROM Settings_Global"
		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		Select Case(ucase(trim(passedElementName)))
			Case "SERNO"
				result = rs("POST_Serno")
			Case "MODE"
				result = rs("POST_Mode")
			Case "SERVICEMEMOURL1"
				result = rs("POST_ServiceMemoURL1")
			Case "ASSETLOCATIONURL1"
				result = rs("POST_AssetLocationURL1")
			Case "SERVICEMEMOURL2"
				result = rs("POST_ServiceMemoURL2")
			Case "ASSETLOCATIONURL2"
				result = rs("POST_AssetLocationURL2")
			Case "NEVERPUTONHOLD"
				result = rs("NeverPutOnHold")
			Case "SERVICEMEMOURL1ONOFF"
				result = rs("POST_ServiceMemoURL1ONOFF")
			Case "SERVICEMEMOURL2ONOFF"
				result = rs("POST_ServiceMemoURL2ONOFF")
			Case "ASSETLOCATIONURL1ONOFF"
				result = rs("POST_AssetLocationURL1ONOFF")
			Case "ASSETLOCATIONURL2ONOFF"
				result = rs("POST_AssetLocationURL2ONOFF")
			Case "SERVICEMEMOURL1MPLEXFORMAT"
				result = rs("POST_ServiceMemoURL1_MplexFormat")
			Case "EMAILFORNON200RESPONSES"
				result = rs("EmailForNon200Responses")
			Case "EWSPOSTURL"
				result = rs("EWSPostURL")
			Case "EWSDEFAULTAPPTDURATION"
				result = rs("EWSDefaultApptDuration")
			Case "EWSDEFAULTMEETINGDURATION"
				 result = rs("EWSDefaultMeetingDuration")
 			Case "REPOSTORDERMODE"
				 result = rs("OrderAPIRepostMode")
 			Case "REPOSTINVOICEMODE"
				 result = rs("InvoiceAPIRepostMode")
 			Case "REPOSTRAMODE"
				 result = rs("RAAPIRepostMode")
 			Case "REPOSTCMMODE"
				 result = rs("CMAPIRepostMode")
 			Case "REPOSTSUMINVMODE"
				 result = rs("SumInvAPIRepostMode")
 			Case "INVENTORYAPIREPOSTONHANDMODE"
				 result = rs("InventoryAPIRepostOnHandMode")
 			Case "INVENTORYAPIPOSTNHANDURL"
				 result = rs("InventoryAPIRepostOnHandURL")
 			Case "INVENTORYWEBAPPPOSTONHANDMODE"
				 result = rs("InventoryWebAppPostOnHandMode")
 			Case "INVENTORYWEBAPPPOSTONHANDURL"
				 result = rs("InventoryWebAppPostOnHandURL")
 			Case "BACKENDINVENTORYPOSTSMODE"
				 result = rs("BackendInventoryPostsMode")	
 			Case "BACKENDINVENTORYPOSTSURL"
				 result = rs("BackendInventoryPostsURL")				 			 
			Case Else
				result = "Error Bad Parameter Name"
		End Select
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

	GetPOSTParams = result

End Function


Function GetFromAddressByMsgID(passedMsgID)

	result = ""
	
	Set cnnGetFromAddressByMsgID = Server.CreateObject("ADODB.Connection")
	cnnGetFromAddressByMsgID.open Session("ClientCnnString")

	SQLGetFromAddressByMsgID = "Select * from PR_ProspectEmailLog where Msg_ID= '" & passedMsgID & "'"

	Set rsGetFromAddressByMsgID = Server.CreateObject("ADODB.Recordset")
	rsGetFromAddressByMsgID.CursorLocation = 3 
	Set rsGetFromAddressByMsgID= cnnGetFromAddressByMsgID.Execute(SQLGetFromAddressByMsgID)
	
	If not rsGetFromAddressByMsgID.eof then 
		result = rsGetFromAddressByMsgID("from_addr")
	End IF	
	set rsGetFromAddressByMsgID= Nothing
	set cnnGetFromAddressByMsgID= Nothing
	
	GetFromAddressByMsgID = result

End Function

Function GetFromAddressByRecID(passedRecID)

	result = ""
	
	Set cnnGetFromAddressByRecID = Server.CreateObject("ADODB.Connection")
	cnnGetFromAddressByRecID.open Session("ClientCnnString")

	SQLGetFromAddressByRecID = "Select * from PR_ProspectEmailLog where InternalRecordIdentifier = " & passedRecID

	Set rsGetFromAddressByRecID = Server.CreateObject("ADODB.Recordset")
	rsGetFromAddressByRecID.CursorLocation = 3 
	Set rsGetFromAddressByRecID= cnnGetFromAddressByRecID.Execute(SQLGetFromAddressByRecID)
	
	If not rsGetFromAddressByRecID.eof then 
		result = rsGetFromAddressByRecID("from_addr")
	End IF	
	set rsGetFromAddressByRecID= Nothing
	set cnnGetFromAddressByRecID= Nothing
	
	GetFromAddressByRecID = result

End Function



Function GetServiceTicketOpenDateTime(passedTicketNumber)

	result = "01/01/1980"
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "' AND RecordSubType='OPEN'"

	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		result = rsBoost1("SubmissionDateTime")
	Else
		SQL = "Select * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "' AND RecordSubType='HOLD'"
		Set rsBoost1= cnn.Execute(SQL)
		If not rsBoost1.eof then result = rsBoost1("SubmissionDateTime")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetServiceTicketOpenDateTime = result

End Function

Function GetHoldServiceTicketSubmittedDateTime(passedTicketNumber)

	result = "01/01/1980"
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "' AND RecordSubType='HOLD'"

	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		result = rsBoost1("SubmissionDateTime")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetHoldServiceTicketSubmittedDateTime = result

End Function



Function GetServiceTicketCloseDateTime(passedTicketNumber)

	result = "01/01/1980"
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "' AND (RecordSubType='CLOSE' or RecordSubType='CANCEL')"

	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		result = rsBoost1("SubmissionDateTime")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetServiceTicketCloseDateTime= result

End Function

Function GetServiceTicketStatus(passedTicketNumber)

	result = ""
	
	Set cnnStatus = Server.CreateObject("ADODB.Connection")
	cnnStatus.open Session("ClientCnnString")

	SQLstatus = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & passedTicketNumber & "' order by submissionDateTime desc"

	Set rsStatus = Server.CreateObject("ADODB.Recordset")
	rsStatus.CursorLocation = 3 
	Set rsStatus = cnnStatus.Execute(SQLstatus )
	
	If not rsStatus.eof then result = rsStatus("CurrentStatus")

	set rsStatus = Nothing
	set cnnStatus= Nothing
	
	GetServiceTicketStatus = result

End Function


Function GetNumberOfTicketsByDate(passedDate,passedMemoType)

	result = 0
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select Count(*) as Expr1 from FS_ServiceMemos where RecordSubType = '" & passedMemoType & "' AND Cast(SubmissionDateTime as Date) = '" & FormatDateTime(passedDate,2) & "'"

	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	'response.write(SQL)
	If not rsBoost1.eof then 
		result = rsBoost1("Expr1")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetNumberOfTicketsByDate = result

End Function

Function GetNumberOfNONDispatchedTicketsByDate(passedDate,passedMemoType)

	result = 0
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select Count(*) as Expr1 from FS_ServiceMemos where RecordSubType = '" & passedMemoType & "' AND Cast(SubmissionDateTime as Date) = '" & FormatDateTime(passedDate,2) & "' AND Dispatched <> 1"

	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		result = rsBoost1("Expr1")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetNumberOfNONDispatchedTicketsByDate= result

End Function

Function FormatAsSortableDateTime(passedDateTime)

	result = "19800101000000"
	
	'Create Custom Sort Key
	CustomSortValue = Year(passedDateTime)
	If Month(passedDateTime) < 10 Then 
		CustomSortValue = CustomSortValue & "0" &  Month(passedDateTime) 
	Else
		CustomSortValue = CustomSortValue & Month(passedDateTime) 
	End If		
	If Day(passedDateTime) < 10 Then 
		CustomSortValue = CustomSortValue  & "0" &  Day(passedDateTime) 
	Else
		CustomSortValue = CustomSortValue  &  Day(passedDateTime) 
	End If	
	CustomSortValue2 = FormatDateTime(passedDateTime,vbShortTime)
	CustomSortValue2 = Replace(CustomSortValue2,":","")
	
	'Now we have to manually so the seconds
	CustomSortValue3 = FormatDateTime(passedDateTime,vbLongTime)
	CustomSortValue3 = Replace(CustomSortValue3 ,"AM","")
	CustomSortValue3 = Replace(CustomSortValue3 ,"PM","")
	CustomSortValue3 = Trim(CustomSortValue3)
	CustomSortValue3 = Right(CustomSortValue3,2)
	
	CustomSortValue = CustomSortValue & CustomSortValue2 & CustomSortValue3 
	CustomSortValue = Replace(CustomSortValue," ","")
	
	result = CustomSortValue 
	
	FormatAsSortableDateTime = result

End Function

Function ElapsedTimeCalcMethod()

	result="Actual"
	
	SQL = "SELECT * FROM Settings_FieldService"
		
	Set cnn9 = Server.CreateObject("ADODB.Connection")
	cnn9.open (Session("ClientCnnString"))
	Set rs9 = Server.CreateObject("ADODB.Recordset")
	rs9.CursorLocation = 3 
	Set rs9 = cnn9.Execute(SQL)
		
	If not rs9.EOF Then result = rs9("ServiceDayElapsedTimeCalculationMethod")
	
	set rs9 = Nothing
	cnn9.close
	set cnn9 = Nothing

	ElapsedTimeCalcMethod = result

End Function


Function ServiceCallElapsedMinutes(passedMemoNumber)


	totalElapsedMinutes = 0
	
	'Get the normal business day start & end because it is used in a lot of places
	Set cnn10 = Server.CreateObject("ADODB.Connection")
	cnn10.open (Session("ClientCnnString"))
	Set rs10 = Server.CreateObject("ADODB.Recordset")
	rs10.CursorLocation = 3 
	
	SQL10 = "SELECT * FROM Settings_CompanyID"
	Set rs10 = cnn10.Execute(SQL10 )
	
	NormalBizDayStartTime = rs10.fields("BusinessDayStart")
	NormalBizDayEndTime = rs10.fields("BusinessDayEnd")
	MinutesInFullDay = dateDiff("n",NormalBizDayStartTime ,NormalBizDayEndTime)
	
	Set rs10 = Nothing
	cnn10.close	
	Set cnn10=Nothing


	OpenedDate = FormatDateTime(GetServiceTicketOpenDateTime(passedMemoNumber),2)
	OpenedTime = FormatDateTime(GetServiceTicketOpenDateTime(passedMemoNumber),4)
	OpenedDateTime = GetServiceTicketOpenDateTime(passedMemoNumber)
	
	ServiceTicketStatus = GetServiceTicketStatus(passedMemoNumber)
	

	If ServiceTicketStatus <> "OPEN" Then 
		ClosedDate =  FormatDateTime(GetServiceTicketCloseDateTime(passedMemoNumber),2)
		ClosedTime = FormatDateTime(GetServiceTicketCloseDateTime(passedMemoNumber),4)
		ClosedDateTime = GetServiceTicketCloseDateTime(passedMemoNumber)
	Else
		ClosedDate =  FormatDateTime(Now(),2)
		ClosedTime = FormatDateTime(Now(),4)
		ClosedDateTime = Now()
	End If
	
	'Response.Write("passedMemoNumber : " & passedMemoNumber & "<br>")
	'Response.Write("OpenedDate : " & OpenedDate & "<br>")
	'Response.Write("OpenedTime : " & OpenedTime & "<br>")
	'Response.Write("OpenedDateTime : " & OpenedDateTime & "<br>")
	'Response.Write("ClosedDate : " & ClosedDate & "<br>")
	'Response.Write("ClosedTime : " & ClosedTime & "<br>")
	'Response.Write("ClosedDateTime : " & ClosedDateTime & "<br>")
	
	'Response.Write("ServiceTicketStatus : " & ServiceTicketStatus & "<br>")
	
	
	'**************************************************************************************************************************
	'IN THIS CONDITION, THE SERVICE TICKET WAS OPENED AND CLOSED ON THE SAME DAY
	'IF SO THE DIFFERENCE IS JUST A TIME DIFFERENCE IN MINUTES ONLY
	'**************************************************************************************************************************
	If datediff("d",OpenedDateTime,ClosedDateTime) < 1 Then
	
		totalElapsedMinutes = datediff("n",OpenedDateTime,ClosedDateTime)
		
	End If
	
	'**************************************************************************************************************************
	'IN THIS CONDITION, THE SERVICE TICKET WAS OPENED AND CLOSED ON DIFFERENT DAYS - SPANNING MORE THAN ONE DAY
	'NOW WE NEED TO CALCULATE
	'**************************************************************************************************************************	
	If datediff("d",OpenedDateTime,ClosedDateTime) >= 1 Then
	
		NumberOfElements = datediff("d",OpenedDateTime,ClosedDateTime) 
		ReDim DaysArray(NumberOfElements)
		
		x = 0
		DateForEval = OpenedDateTime
		DateForEval = cdate(DateForEval)
		
		Do
			DaysArray(x) = DateAdd("d",1,DateForEval) 
			DateForEval = DateAdd("d",1,DateForEval) 
			
			x = x + 1
		
		Loop While x <= NumberOfElemenets
		
	End If

	ElapsedMinutes = 0
	
	For x = 0 To NumberOfElements - 1
	
		'*******************************************************************************************************
		'FOR EACH DATE, LOOKUP DATE IN COMPANY CALENDER TO SEE IF COMPANY WAS OPEN, CLOSED OR CLOSED EARLY
		'AND CALCULATE THE TOTAL MINUTES OPENED ON THAT DAY
		'*******************************************************************************************************
		'*******************************************************************************************************
		'If the date is the day the ticket is opened, calculate the time since the ticket was opened that day
		'If the date is the day the ticket is closed, calculate the time until the ticket was closed that day
		'Otherwise, the ticket was opened all day, calculate all the minutes the business was opened that day
		'*******************************************************************************************************
		
		'Response.Write("DaysArray(x): " & DaysArray(x) & "<br>")
		
		If datediff("d",DaysArray(x),OpenedDate) = 0 Then
			ElapsedMinutes = ElapsedMinutes + NumberofWorkMinutes(DaysArray(x),OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime)
		End If
		
		If datediff("d",DaysArray(x),ClosedDate) = 0 Then
			ElapsedMinutes = ElapsedMinutes + NumberofWorkMinutes(DaysArray(x),ClosedTime,NormalBizDayStartTime,NormalBizDayEndTime)
		End If
		
		If datediff("d",DaysArray(x),OpenedDate) <> 0 AND datediff("d",DaysArray(x),ClosedDate) <> 0 Then
			ElapsedMinutes = ElapsedMinutes + NumberofWorkMinutes(DaysArray(x),"04:00",NormalBizDayStartTime,NormalBizDayEndTime)
		End If



	Next


	totalElapsedMinutes = ElapsedMinutes
	
	ServiceCallElapsedMinutes = totalElapsedMinutes
		
	Exit Function
	
	
	
	
	If GetServiceTicketStatus(passedMemoNumber) <> "OPEN" Then ' See if it it closed or cancelled
	
			If GetServiceTicketStatus(passedMemoNumber) <> "HOLD" Then ' Not on hold
			
				'Get the number of minutes in a full day so we have it
			
				If datediff("d",GetServiceTicketOpenDateTime(passedMemoNumber),GetServiceTicketCloseDateTime(passedMemoNumber)) < 1 Then
						'It all happenned on the same day so just calc the minutes
						'Still have to use this function to see if is a da they were open
						OpenedDate = FormatDateTime(GetServiceTicketOpenDateTime(passedMemoNumber),2)
						OpenedTime = FormatDateTime(GetServiceTicketOpenDateTime(passedMemoNumber),4)
						If NumberofWorkMinutes_DateOpened(OpenedDate,OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime) <> 0 Then 'They were not closed
							ClosedDate =  FormatDateTime(GetServiceTicketCloseDateTime(passedMemoNumber),2)
							ClosedTimed = FormatDateTime(GetServiceTicketCloseDateTime(passedMemoNumber),4)
							If ClosedTimed > NormalBizDayEndTime Then ' It was closed after hours so just give them the minutes since opening that day
								totalElapsedMinutes =  NumberofWorkMinutes_DateOpened(OpenedDate,OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime)
							Else
								'Otherwise, it is just the actual minutes
								totalElapsedMinutes =  datediff("n",GetServiceTicketOpenDateTime(passedMemoNumber),GetServiceTicketCloseDateTime(passedMemoNumber))
							End If
						Else
							'They were closed
							totalElapsedMinutes = 0
						End If
						If debugmsg=1 then Response.Write(passedMemoNumber & "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz:" & totalElapsedMinutes  & "<br>")
				Else
						'Get the number of minutes remaining on the day it was opened
						OpenedDate = FormatDateTime(GetServiceTicketOpenDateTime(passedMemoNumber),2)
						OpenedTime = FormatDateTime(GetServiceTicketOpenDateTime(passedMemoNumber),4)
						
						If debugmsg=1 then Response.Write(passedMemoNumber & "XXXXXXXXXXXXXX<br>")
						If debugmsg=1 then Response.Write(passedMemoNumber & ":OpenedDate :" & OpenedDate & "<br>")
						If debugmsg=1 then Response.Write(passedMemoNumber & ":OpenedTime :" & OpenedTime & "<br>")
						totalElapsedMinutes =  NumberofWorkMinutes_DateOpened(OpenedDate,OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime)
						If debugmsg=1 then Response.Write( passedMemoNumber & ":Elapsed Min For Day Opened :" & totalElapsedMinutes & "<br>")
			
						'Now see if there are more days to add on
						If GetServiceTicketStatus(passedMemoNumber) = "CLOSE" or GetServiceTicketStatus(passedMemoNumber) = "CANCEL" Then'
							EndDate = GetServiceTicketCloseDateTime(passedMemoNumber)
						Else
							EndDate = Now()
						End If
						
						WorkingDate=DateAdd("d",1,GetServiceTicketOpenDateTime(passedMemoNumber))
						WorkingDate = FormatDateTime(WorkingDate,2)
						EndDate = FormatDateTime(EndDate,2)
						
						WorkingDate = cdate(WorkingDate)
						EndDate = cdate(EndDate)
						If debugmsg=1 then Response.Write( passedMemoNumber & "XXXXXXXXXXXXXXXXXXXXXXXXXbr>")
						If debugmsg=1 then Response.Write( passedMemoNumber & ":WorkingDate :" & WorkingDate & "<br>")
						If debugmsg=1 then Response.Write( passedMemoNumber & ":EndDate:" & EndDate& "<br>")
			
					
						Do While WorkingDate < EndDate
							totalElapsedMinutes  = totalElapsedMinutes + NumberofWorkMinutes_FullDay(WorkingDate,NormalBizDayStartTime,NormalBizDayEndTime)
							If debugmsg=1 then Response.Write( WorkingDate  & ":..loop..." & EndDate & ":..loop..." &  passedMemoNumber & ":..NumberofWorkMinutes_FullDay:" & NumberofWorkMinutes_FullDay(WorkingDate,NormalBizDayStartTime,NormalBizDayEndTime) & ":..loop...<br>")
							WorkingDate=DateAdd("d",1,WorkingDate)
						Loop
					
						' Now add in the minutes for the day that it was closed
						LastDayDate = FormatDateTime(GetServiceTicketCloseDateTime(passedMemoNumber),2)
						LastDayTime = FormatDateTime(GetServiceTicketCloseDateTime(passedMemoNumber),4)
						If debugmsg=1 then Response.Write( passedMemoNumber & ":LastDayDate :" & LastDayDate & "<br>")
						If debugmsg=1 then Response.Write( passedMemoNumber & ":LastDayTime :" & LastDayTime & "<br>")
						If debugmsg=1 then Response.Write( passedMemoNumber & ": NumberofWorkMinutes_DateClosed:" & NumberofWorkMinutes_DateClosed(LastDayDate,LastDayTime,NormalBizDayStartTime,NormalBizDayEndTime )& "<br>")
						totalElapsedMinutes = totalElapsedMinutes  + NumberofWorkMinutes_DateClosed(LastDayDate,LastDayTime,NormalBizDayStartTime,NormalBizDayEndTime ) 
						
				End If
			
			Else 'If we are here, the service ticket is on hold
			
				If debugmsg=1 then response.write("HOLD Ticket<br>")
				If debugmsg=1 then response.write(GetHoldServiceTicketSubmittedDateTime(passedMemoNumber) & "<br>")
				If debugmsg=1 then response.write(Now() & "<br>")
				If debugmsg=1 then response.write(datediff("d",GetHoldServiceTicketSubmittedDateTime(passedMemoNumber),Now()) & "<br>")
		
							
					If datediff("d",GetHoldServiceTicketSubmittedDateTime(passedMemoNumber),Now()) < 1 Then
							'It was only opened today so get the number of minutes from the time
							'it was open to Now() or until th end of the day if Now() is after hours
							If debugmsg=1 then Response.Write(passedMemoNumber & "11111111111111111111111111111111111111:" & totalElapsedMinutes  & "<br>")
							HoldDate = FormatDateTime(GetHoldServiceTicketSubmittedDateTime(passedMemoNumber),2)
							HoldTime = FormatDateTime(GetHoldServiceTicketSubmittedDateTime(passedMemoNumber),4)
							If NumberofWorkMinutes_DateOpened(HoldDate,HoldTime,NormalBizDayStartTime,NormalBizDayEndTime) <> 0 Then 'They were not closed
								If FormatDateTime(Now(),4) > NormalBizDayEndTime Then ' It was closed after hours so just give them the minutes since opening that day
									totalElapsedMinutes =  NumberofWorkMinutes_DateOpened(HoldDate,HoldTime,NormalBizDayStartTime,NormalBizDayEndTime)
								Else
									'Otherwise, it is just the actual minutes
									totalElapsedMinutes =  datediff("n",GetHoldServiceTicketSubmittedDateTime(passedMemoNumber),Now())
								End If
							Else
								'They were closed
								totalElapsedMinutes = 0
							End If
							If debugmsg=1 then Response.Write(passedMemoNumber & "tttttttttttttttttttttttttttzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz:" & totalElapsedMinutes  & "<br>")
					Else
							'Get the number of minutes remaining on the day it was opened
							HoldDate = FormatDateTime(GetHoldServiceTicketSubmittedDateTime(passedMemoNumber),2)
							HoldTime = FormatDateTime(GetHoldServiceTicketSubmittedDateTime(passedMemoNumber),4)
							
							If debugmsg=1 then Response.Write(passedMemoNumber & "XXXXXXXXXXXXXX<br>")
							If debugmsg=1 then Response.Write(passedMemoNumber & ":HoldDate :" & HolddDate & "<br>")
							If debugmsg=1 then Response.Write(passedMemoNumber & ":HoldTime :" & HoldTime & "<br>")
							totalElapsedMinutes =  NumberofWorkMinutes_DateOpened(HoldDate,HoldTime,NormalBizDayStartTime,NormalBizDayEndTime)
							If debugmsg=1 then Response.Write( passedMemoNumber & ":Elapsed Min For Day On Hold :" & totalElapsedMinutes & "<br>")
				
							'Now see if there are more days to add on
							EndDate = FormatDateTime(Now(),2)
							
							WorkingDate=DateAdd("d",1,GetHoldServiceTicketSubmittedDateTime(passedMemoNumber))
							WorkingDate = FormatDateTime(WorkingDate,2)
							EndDate = FormatDateTime(EndDate,2)
							
							WorkingDate = cdate(WorkingDate)
							EndDate = cdate(EndDate)
							If debugmsg=1 then Response.Write( passedMemoNumber & "XXXXXXXXXXXXXXXXXXXXXXXXX" & "<br>")
							If debugmsg=1 then Response.Write( passedMemoNumber & ":WorkingDate :" & WorkingDate & "<br>")
							If debugmsg=1 then Response.Write( passedMemoNumber & ":EndDate:" & EndDate& "<br>")
				
						
							Do While WorkingDate < EndDate
								totalElapsedMinutes  = totalElapsedMinutes + NumberofWorkMinutes_FullDay(WorkingDate,NormalBizDayStartTime,NormalBizDayEndTime)
								If debugmsg=1 then Response.Write( passedMemoNumber & ":" & WorkingDate  & ":..loop..." & EndDate & ":..loop..." &  passedMemoNumber & ":..NumberofWorkMinutes_FullDay:" & NumberofWorkMinutes_FullDay(WorkingDate,NormalBizDayStartTime,NormalBizDayEndTime) & ":..loop...<br>")
								WorkingDate=DateAdd("d",1,WorkingDate)
							Loop
						
							' Now add in the minutes for the day that it was closed
							If debugmsg=1 then Response.Write( passedMemoNumber & ":SoFar-before-LastDayDate :" & totalElapsedMinutes & "<br>")
		
							LastDayDate = FormatDateTime(Now(),2)
							LastDayTime = FormatDateTime(Now(),4)
							If debugmsg=1 then Response.Write( passedMemoNumber & ":LastDayDate :" & LastDayDate & "<br>")
							If debugmsg=1 then Response.Write( passedMemoNumber & ":LastDayTime :" & LastDayTime & "<br>")
							If debugmsg=1 then Response.Write( passedMemoNumber & ": NumberofWorkMinutes_DateClosed:" & NumberofWorkMinutes_DateClosed(LastDayDate,LastDayTime,NormalBizDayStartTime,NormalBizDayEndTime )& "<br>")
							totalElapsedMinutes = totalElapsedMinutes  + NumberofWorkMinutes_DateClosed(LastDayDate,LastDayTime,NormalBizDayStartTime,NormalBizDayEndTime ) 
							
					End If			
			
			End If
			
	Else ' Ticket is still open
	
		If debugmsg=1 then response.write("Open Ticket<br>")
		If debugmsg=1 then response.write(GetServiceTicketOpenDateTime(passedMemoNumber) & "<br>")
		If debugmsg=1 then response.write(Now() & "<br>")
		If debugmsg=1 then response.write(datediff("d",GetServiceTicketOpenDateTime(passedMemoNumber),Now()) & "<br>")

					
			If datediff("d",GetServiceTicketOpenDateTime(passedMemoNumber),Now()) < 1 Then
					'It was only opened today so get the number of minutes from the time
					'it was open to Now() or until th end of the day if Now() is after hours
					If debugmsg=1 then Response.Write(passedMemoNumber & "11111111111111111111111111111111111111:" & totalElapsedMinutes  & "<br>")
					OpenedDate = FormatDateTime(GetServiceTicketOpenDateTime(passedMemoNumber),2)
					OpenedTime = FormatDateTime(GetServiceTicketOpenDateTime(passedMemoNumber),4)
					If NumberofWorkMinutes_DateOpened(OpenedDate,OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime) <> 0 Then 'They were not closed
						If FormatDateTime(Now(),4) > NormalBizDayEndTime Then ' It was closed after hours so just give them the minutes since opening that day
							totalElapsedMinutes =  NumberofWorkMinutes_DateOpened(OpenedDate,OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime)
						Else
							'Otherwise, it is just the actual minutes
							totalElapsedMinutes =  datediff("n",GetServiceTicketOpenDateTime(passedMemoNumber),Now())
						End If
					Else
						'They were closed
						totalElapsedMinutes = 0
					End If
					If debugmsg=1 then Response.Write(passedMemoNumber & "sssssssssssssssssssssssssssssssssszzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz:" & totalElapsedMinutes  & "<br>")
			Else
					'Get the number of minutes remaining on the day it was opened
					OpenedDate = FormatDateTime(GetServiceTicketOpenDateTime(passedMemoNumber),2)
					OpenedTime = FormatDateTime(GetServiceTicketOpenDateTime(passedMemoNumber),4)
					
					If debugmsg=1 then Response.Write(passedMemoNumber & "XXXXXXXXXXXXXX<br>")
					If debugmsg=1 then Response.Write(passedMemoNumber & ":OpenedDate :" & OpenedDate & "<br>")
					If debugmsg=1 then Response.Write(passedMemoNumber & ":OpenedTime :" & OpenedTime & "<br>")
					totalElapsedMinutes =  NumberofWorkMinutes_DateOpened(OpenedDate,OpenedTime,NormalBizDayStartTime,NormalBizDayEndTime)
					If debugmsg=1 then Response.Write( passedMemoNumber & ":Elapsed Min For Day Opened :" & totalElapsedMinutes & "<br>")
		
					'Now see if there are more days to add on
					EndDate = FormatDateTime(Now(),2)
					
					WorkingDate=DateAdd("d",1,GetServiceTicketOpenDateTime(passedMemoNumber))
					WorkingDate = FormatDateTime(WorkingDate,2)
					EndDate = FormatDateTime(EndDate,2)
					
					WorkingDate = cdate(WorkingDate)
					EndDate = cdate(EndDate)
					If debugmsg=1 then Response.Write( passedMemoNumber & "XXXXXXXXXXXXXXXXXXXXXXXXX" & "<br>")
					If debugmsg=1 then Response.Write( passedMemoNumber & ":WorkingDate :" & WorkingDate & "<br>")
					If debugmsg=1 then Response.Write( passedMemoNumber & ":EndDate:" & EndDate& "<br>")
		
				
					Do While WorkingDate < EndDate
						totalElapsedMinutes  = totalElapsedMinutes + NumberofWorkMinutes_FullDay(WorkingDate,NormalBizDayStartTime,NormalBizDayEndTime)
						If debugmsg=1 then Response.Write( passedMemoNumber & ":" & WorkingDate  & ":..loop..." & EndDate & ":..loop..." &  passedMemoNumber & ":..NumberofWorkMinutes_FullDay:" & NumberofWorkMinutes_FullDay(WorkingDate,NormalBizDayStartTime,NormalBizDayEndTime) & ":..loop...<br>")
						WorkingDate=DateAdd("d",1,WorkingDate)
					Loop
				
					' Now add in the minutes for the day that it was closed
					If debugmsg=1 then Response.Write( passedMemoNumber & ":SoFar-before-LastDayDate :" & totalElapsedMinutes & "<br>")

					LastDayDate = FormatDateTime(Now(),2)
					LastDayTime = FormatDateTime(Now(),4)
					If debugmsg=1 then Response.Write( passedMemoNumber & ":LastDayDate :" & LastDayDate & "<br>")
					If debugmsg=1 then Response.Write( passedMemoNumber & ":LastDayTime :" & LastDayTime & "<br>")
					
					If debugmsg=1 then Response.Write("Parameters which will be passed below******<br>")
					If debugmsg=1 then Response.Write( passedMemoNumber & ":NormalBizDayStartTime:" & NormalBizDayStartTime & "<br>")
					If debugmsg=1 then Response.Write( passedMemoNumber & ":NormalBizDayEndTime :" & NormalBizDayEndTime & "<br>")
					If debugmsg=1 then Response.Write( passedMemoNumber & ": NumberofWorkMinutes_DateClosed:" & NumberofWorkMinutes_DateClosed(LastDayDate,LastDayTime,NormalBizDayStartTime,NormalBizDayEndTime )& "<br>")
					totalElapsedMinutes = totalElapsedMinutes  + NumberofWorkMinutes_DateClosed(LastDayDate,LastDayTime,NormalBizDayStartTime,NormalBizDayEndTime ) 
					
			End If
			

	
	
	End If

	ServiceCallElapsedMinutes = totalElapsedMinutes
	
End Function





Function NumberofWorkMinutes(passedDate,passedTime,passedNormalBizDayStartTime,passedNormalBizDayEndTime)


	NumberofWorkMinutesDate_result = 0

	SQL = "SELECT * FROM Settings_CompanyCalendar WHERE MonthNum='" & Month(passedDate) & "' AND DayNum ='" & Day(passedDate) & "' AND YearNum='" & Year(passedDate) & "'"

	Set cnn9 = Server.CreateObject("ADODB.Connection")
	cnn9.open (Session("ClientCnnString"))
	Set rs9 = Server.CreateObject("ADODB.Recordset")
	rs9.CursorLocation = 3 
	Set rs9 = cnn9.Execute(SQL)
		
	If not rs9.EOF Then
	
		Select Case rs9("OpenClosedCloseEarly")
		
			Case "Closed"
				'************************************************************
				'Do nothing, they are closed
				'************************************************************
				
			Case "Close Early"
			
				ClosingTime = cdate(rs9("ClosingTime"))
				
				If passedTime = "" Then
				
					NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,ClosingTime)
				Else
				
					'************************************************************
					'IF THE PASSED DATE IS TODAY AND WE ARE CLOSING EARLY TODAY
					'************************************************************
					If DateDiff("d",passedDate,Now()) = 0 Then
							
						'************************************************************		
						'IS THE PASSED TIME PRIOR TO THE CLOSE EARLY TIME
						'************************************************************
						If DateDiff("n",passedTime,ClosingTime) > 0 Then
						
							'************************************************************		
							'WE GOT IN HERE, SO PASSED TIME IS PRIOR TO CLOSE EARLY TIME
							'************************************************************
							
							'************************************************************		
							'IS THE CURRENT TIME AFTER THE CLOSE EARLY TIME
							'************************************************************
							If DateDiff("n",Time(),ClosingTime) < 0 Then
							
								'************************************************************		
								'IF IT IS, CALCULATE TIME UP TO CLOSE EARLY TIME
								'************************************************************
								NumberofWorkMinutesDate_result = DateDiff("n",passedTime,ClosingTime)
							Else
								'************************************************************		
								'OTHERWISE CALCULATE ELAPSED TIME SO FAR TODAY
								'************************************************************
								NumberofWorkMinutesDate_result = DateDiff("n",passedTime,Time())			
							End If
							
						Else
						
							'************************************************************		
							'THE PASSED TIME IS NOT AFTER THE CLOSE EARLY TIME
							'************************************************************
							NumberofWorkMinutesDate_result = DateDiff("n",passedTime,Time())			
						End If
						
					Else
					
						'************************************************************		
						'THE PASSED DATE IS NOT TODAY
						'************************************************************
						NumberofWorkMinutesDate_result = DateDiff("n",passedTime,ClosingTime)
						
					End If
					'************************************************************************************************
	
				End If
				
			End Select
	Else ' If not found, they were open
		If passedTime = "" Then
			NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,passedNormalBizDayEndTime)
		Else
			NumberofWorkMinutesDate_result = DateDiff("n",passedTime,passedNormalBizDayEndTime)
			If debugmsg=1 then Response.Write( passedMemoNumber & ":passedTime:" & passedTime & "<br>")
			If debugmsg=1 then Response.Write( passedMemoNumber & ":EndTime:" & passedNormalBizDayEndTime & "<br>")
		End IF	
	End If

	set rs9 = Nothing
	cnn9.close
	set cnn9 = Nothing
			

	If Weekday(passedDate,vbMonday) <= 5 Then
		NumberofWorkMinutes = NumberofWorkMinutesDate_result
	Else
		NumberofWorkMinutes = 0
	End If
	
End Function




Function NumberofWorkMinutes_DateOpened(passedDate,passedTime,passedNormalBizDayStartTime,passedNormalBizDayEndTime)

debugmsg=0

	If passedTime = "" OR passedTime <= passedNormalBizDayEndTime Then ' Otherwise it is after or before hours so return 0
		If debugmsg=1 then response.write("CCCCCCCCCCCCCCCCCCCCCCCCCCpassedTime" & passedTime & "<br>")
		If debugmsg=1 then response.write("CCCCCCCCCCCCCCCCCCCCCCCCCCEndTime:" & passedNormalBizDayStartTime&"<br>")

			SQL = "SELECT * FROM Settings_CompanyCalendar where Monthnum='" & Month(passedDate) & "' AND DayNum ='" & Day(passedDate) & "' AND YearNum='" & Year(passedDate) & "'"
		
			Set cnn9 = Server.CreateObject("ADODB.Connection")
			cnn9.open (Session("ClientCnnString"))
			Set rs9 = Server.CreateObject("ADODB.Recordset")
			rs9.CursorLocation = 3 
			Set rs9 = cnn9.Execute(SQL)
				
			If not rs9.EOF Then
				Select Case rs9("OpenClosedCloseEarly")
					Case "Closed"
						'Do nothing, they are closed
					Case "Close Early"
						ClosingTime = cdate(rs9("ClosingTime"))
						If passedTime = "" Then
							NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,ClosingTime)
						Else
							NumberofWorkMinutesDate_result = DateDiff("n",passedTime,ClosingTime)
						End If
					End Select
			Else ' If not found, they were open
				If passedTime = "" Then
					NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,passedNormalBizDayEndTime)
				Else
					NumberofWorkMinutesDate_result = DateDiff("n",passedTime,passedNormalBizDayEndTime)
					If debugmsg=1 then Response.Write( passedMemoNumber & ":passedTime:" & passedTime& "<br>")
					If debugmsg=1 then Response.Write( passedMemoNumber & ":EndTime:" & passedNormalBizDayEndTime& "<br>")
				End IF	
			End If
		
			set rs9 = Nothing
			cnn9.close
			set cnn9 = Nothing
			
	Else ' After hours
			If debugmsg=1 then response.write("CCCCCCCCCCCCCCCCCCCCCCCCCC<br>")
			NumberofWorkMinutesDate_result = 0
	End If
		
	NumberofWorkMinutes_DateOpened= NumberofWorkMinutesDate_result
	
End Function

Function NumberofWorkMinutes_DateClosed(passedDate,passedTime,passedNormalBizDayStartTime,passedNormalBizDayEndTime)

debugmsg=0
If debugmsg=1 then Response.Write( passedMemoNumber & "QQ:passedDate:" & passedDate& "<br>")
If debugmsg=1 then Response.Write( passedMemoNumber & "QQ:passedTime:" & passedTime& "<br>")
If debugmsg=1 then Response.Write( passedMemoNumber & "QQ:passedNormalBizDayStartTime:" & passedNormalBizDayStartTime& "<br>")
If debugmsg=1 then Response.Write( passedMemoNumber & "QQ:passedNormalBizDayEndTime:" & passedNormalBizDayEndTime& "<br>")
	
	SQL = "SELECT * FROM Settings_CompanyCalendar where Monthnum='" & Month(passedDate) & "' AND DayNum ='" & Day(passedDate) & "' AND YearNum='" & Year(passedDate) & "'"

If debugmsg=1 then Response.Write( SQL & "<BR>")

	Set cnn9 = Server.CreateObject("ADODB.Connection")
	cnn9.open (Session("ClientCnnString"))
	Set rs9 = Server.CreateObject("ADODB.Recordset")
	rs9.CursorLocation = 3 
	Set rs9 = cnn9.Execute(SQL)
		
	If not rs9.EOF Then
		Select Case rs9("OpenClosedCloseEarly")
			Case "Closed"
				'Do nothing, they are closed
			Case "Close Early"
				ClosingTime = cdate(rs9("ClosingTime"))
				If passedTime = "" Then
					NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,ClosingTime)
				Else
					NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,passedTime)
				End IF
			End Select
	Else ' AN EOF HERE MEANS THEY ARE OPEN
		If passedTime = "" Then
			NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,passedNormalBizDayEndTime)
		Else
			NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,passedTime)
			If debugmsg=1 then Response.Write( passedMemoNumber & "QQ:passedTime:" & passedTime& "<br>")
			If debugmsg=1 then Response.Write( passedMemoNumber & "QQ:NormalBizDayStartTime:" & passedNormalBizDayStartTime& "<br>")
		End IF	

 	End If

	set rs9 = Nothing
	cnn9.close
	set cnn9 = Nothing
		
	NumberofWorkMinutes_DateClosed= NumberofWorkMinutesDate_result
	
End Function

Function NumberofWorkMinutes_FullDay(passedDate,passedNormalBizDayStartTime,passedNormalBizDayEndTime)

debugmsg=0

'Response.Write( passedMemoNumber & ": GOT TO NumberofWorkMinutes_FullDay:"& "<br>")
'Response.Write( passedMemoNumber & ":NormalBizDayStartTime:" & passedNormalBizDayStartTime& "<br>")
'Response.Write( passedMemoNumber & ":NormalBizDayEndTime:" & passedNormalBizDayEndTime& "<br>")

	SQL = "SELECT * FROM Settings_CompanyCalendar where Monthnum='" & Month(passedDate) & "' AND DayNum ='" & Day(passedDate) & "' AND YearNum='" & Year(passedDate) & "'"

	Set cnn9 = Server.CreateObject("ADODB.Connection")
	cnn9.open (Session("ClientCnnString"))
	Set rs9 = Server.CreateObject("ADODB.Recordset")
	rs9.CursorLocation = 3 
	Set rs9 = cnn9.Execute(SQL)
		
	If not rs9.EOF Then
		Select Case rs9("OpenClosedCloseEarly")
			Case "Closed"
				'Do nothing, they are closed
			Case "Close Early"
				ClosingTime = cdate(rs9("ClosingTime"))
				NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,ClosingTime)
		End Select
	Else ' AN EOF HERE MEANS THEY ARE OPEN
		NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,passedNormalBizDayEndTime)
 	End If

	set rs9 = Nothing
	cnn9.close
	set cnn9 = Nothing
			
		
	NumberofWorkMinutes_FullDay= NumberofWorkMinutesDate_result
	
End Function

Function CustAROver(passedCustID,Age30_60_90)

	agedResult = 0
	MasterAccountNumber = 0
	
	Set cnnAged = Server.CreateObject("ADODB.Connection")
	cnnAged.open Session("ClientCnnString")
	
	SQLaged = "Select * from " & Session("SQL_Owner") & ".AR_Customer where CustNum= " & passedCustID 
	 
	Set rsAged = Server.CreateObject("ADODB.Recordset")
	rsAged.CursorLocation = 3 
	Set rsAged = cnnAged.Execute(SQLaged)
	
	If not rsAged.eof then 
		If rsAged.fields("MasterCustNum") <> "0" Then MasterAccountNumber = rsAged.fields("MasterCustNum")
		Select Case Age30_60_90
			Case 30
				agedResult = rsAged.fields("AcctOver30Ar") + rsAged.fields("AcctOver60Ar") + rsAged.fields("AcctOver90Ar") 	
			Case 60
				agedResult = rsAged.fields("AcctOver60Ar") + rsAged.fields("AcctOver90Ar") 
			Case 90
				agedResult = rsAged.fields("AcctOver90Ar") 			
		End Select
	End If
	
	
	If MasterAccountNumber <> 0 Then ' This is a mastered account, do it again but lookup the master
		SQLaged = "Select * from " & Session("SQL_Owner") & ".AR_Customer where CustNum= " & MasterAccountNumber 
		Set rsAged = cnnAged.Execute(SQLaged)
		If not rsAged.eof then 
			Select Case Age30_60_90
				Case 30
					agedResult = rsAged.fields("AcctOver30Ar") + rsAged.fields("AcctOver60Ar") + rsAged.fields("AcctOver90Ar") 	
				Case 60
					agedResult = rsAged.fields("AcctOver60Ar") + rsAged.fields("AcctOver90Ar") 
				Case 90
					agedResult = rsAged.fields("AcctOver90Ar") 			
			End Select
		End If
	End If

	set rsAged= Nothing
	cnnAged.Close
	set cnnAged = Nothing

	CustAROver = agedResult
	
End Function

Function CustHasSalesInLast90Days(passedCustID)

	Result90Days = False
	
	Set cnn90day = Server.CreateObject("ADODB.Connection")
	cnn90day.open Session("ClientCnnString")

	SQL90day = "Select Top 1 * from " & Session("SQL_Owner") & ".InvoiceHistory where CustNum= " & passedCustID &" Order by IvsDate Desc"
	 
	Set rs90day = Server.CreateObject("ADODB.Recordset")
	rs90day.CursorLocation = 3 
	Set rs90day = cnn90day.Execute(SQL90day)
	
	If not rs90day.eof then 
			If DateDiff("d",rs90day.fields("IvsDate"),Now()) > 89 Then Result90Days = True
	End If
	
	set rsAged= Nothing
	set cnnAged = Nothing
	
	CustHasSalesInLast90Days = Result90Days
	
End Function

Function FormattedCustInfoByCustNum(passedCustID)

	Set cnnFormattedCustInfoByCustNum = Server.CreateObject("ADODB.Connection")
	cnnFormattedCustInfoByCustNum.open Session("ClientCnnString")

	
	result = "*Not Found*"
		
	SQL = "Select * from " & Session("SQL_Owner") & ".AR_Customer where CustNum= '" & passedCustID & "'"
	 
	Set rsFormattedCustInfoByCustNum = Server.CreateObject("ADODB.Recordset")
	rsFormattedCustInfoByCustNum.CursorLocation = 3 
	Set rsFormattedCustInfoByCustNum = cnnFormattedCustInfoByCustNum.Execute(SQL)
	
	If not rsFormattedCustInfoByCustNum.eof then
		result = rsFormattedCustInfoByCustNum("Name") & "<br>"
		result = result  & rsFormattedCustInfoByCustNum("Addr1") & "<br>"
		If rsFormattedCustInfoByCustNum("Addr2") <> "" Then result = result  & rsFormattedCustInfoByCustNum("Addr2") & "<br>"
		result = result  & rsFormattedCustInfoByCustNum("CityStateZip")
		If rsFormattedCustInfoByCustNum("Phone") <> "" Then result = result  & "<br>" &  rsFormattedCustInfoByCustNum("Phone")
	End If
	
	set rsFormattedCustInfoByCustNum = Nothing
	cnnFormattedCustInfoByCustNum.close
	set cnnFormattedCustInfoByCustNum= Nothing
	
	FormattedCustInfoByCustNum = result 
	
End Function

Function GetCustTypeCodeByCustID(passedCustID)

	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	
	resultBoost = 0
		
	SQL = "Select * from AR_Customer where CustNum = '" & passedCustID & "'"
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	
	If not rsBoost1.eof then resultBoost = rsBoost1("CustType")
	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	GetCustTypeCodeByCustID = resultBoost
	
End Function


Function NumberOfOpenServiceCalls()

	Set cnnNumOp = Server.CreateObject("ADODB.Connection")
	cnnNumOp.open Session("ClientCnnString")

	
	resultNumOp = 0
		
	SQLNumOp = "Select Distinct MemoNumber from FS_ServiceMemos where CurrentStatus = 'OPEN'"
	 
	Set rsNumOp = Server.CreateObject("ADODB.Recordset")
	rsNumOp.CursorLocation = 3 
	
	rsNumOp.Open SQLNumOp , cnnNumOp
			
	resultNumOp = rsNumOp.RecordCount
	
	rsNumOp.Close
	set rsNumOp= Nothing
	cnnNumOp.Close	
	set cnnNumOp= Nothing
	
	NumberOfOpenServiceCalls = resultNumOp
	
End Function

Function NumberOfHoldServiceCalls()

	Set cnnNumOp = Server.CreateObject("ADODB.Connection")
	cnnNumOp.open Session("ClientCnnString")

	
	resultNumOp = 0
		
	SQLNumOp = "Select * from FS_ServiceMemos where CurrentStatus = 'HOLD'"
	 
	Set rsNumOp = Server.CreateObject("ADODB.Recordset")
	rsNumOp.CursorLocation = 3 
	
	rsNumOp.Open SQLNumOp , cnnNumOp
			
	resultNumOp = rsNumOp.RecordCount
	
	rsNumOp.Close
	set rsNumOp= Nothing
	cnnNumOp.Close	
	set cnnNumOp= Nothing
	
	NumberOfHoldServiceCalls = resultNumOp
	
End Function



Function NumberOfServiceCallsNotDispatched()

	Set cnnNumOp = Server.CreateObject("ADODB.Connection")
	cnnNumOp.open Session("ClientCnnString")

	resultNumOp = 0

	SQLNumOp = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND "
	SQLNumOp = SQLNumOp & "MemoNumber NOT IN "
	SQLNumOp = SQLNumOp & "(SELECT MemoNumber FROM FS_ServiceMemosDetail WHERE MemoStage = 'Dispatched')"
 
	'Response.Write(SQLNumOp )
	 
	Set rsNumOp = Server.CreateObject("ADODB.Recordset")
	rsNumOp.CursorLocation = 3 
	
	rsNumOp.Open SQLNumOp , cnnNumOp
	
	resultNumOp = rsNumOp.RecordCount
	
	rsNumOp.Close
	set rsNumOp= Nothing
	cnnNumOp.Close
	set cnnNumOp= Nothing
	
	NumberOfServiceCallsNotDispatched = resultNumOp
	
End Function

Function MarkAlertEmailSent(passedTicketNumber)

	result = 0
	
	Set cnnMarkAlertEmailSent = Server.CreateObject("ADODB.Connection")
	cnnMarkAlertEmailSent.open Session("ClientCnnString")

	SQLMarkAlertEmailSent = "Update FS_ServiceMemos Set AlertEmailSent = getdate() where MemoNumber = '" & passedTicketNumber & "'"

	Set rsMarkAlertEmailSent = Server.CreateObject("ADODB.Recordset")
	rsMarkAlertEmailSent.CursorLocation = 3 
	Set rsMarkAlertEmailSent= cnnMarkAlertEmailSent.Execute(SQLMarkAlertEmailSent)
	
	set rsMarkAlertEmailSent= Nothing
	set cnnMarkAlertEmailSent= Nothing
	
	MarkAlertEmailSent = result

End Function

Function MarkHoldAlertEmailSent(passedTicketNumber)

	result = 0
	
	Set cnnMarkAlertEmailSent = Server.CreateObject("ADODB.Connection")
	cnnMarkAlertEmailSent.open Session("ClientCnnString")

	SQLMarkAlertEmailSent = "Update FS_ServiceMemos Set HoldAlertEmailSent = getdate() where MemoNumber = '" & passedTicketNumber & "'"

	Set rsMarkAlertEmailSent = Server.CreateObject("ADODB.Recordset")
	rsMarkAlertEmailSent.CursorLocation = 3 
	Set rsMarkAlertEmailSent= cnnMarkAlertEmailSent.Execute(SQLMarkAlertEmailSent)
	
	set rsMarkAlertEmailSent= Nothing
	set cnnMarkAlertEmailSent= Nothing
	
	MarkHoldAlertEmailSent= result

End Function


Function MarkEscalationAlertEmailSent(passedTicketNumber)

	result = 0
	
	Set cnnMarkAlertEmailSent = Server.CreateObject("ADODB.Connection")
	cnnMarkAlertEmailSent.open Session("ClientCnnString")

	SQLMarkAlertEmailSent = "Update FS_ServiceMemos Set EscalationAlertEmailSent = getdate() where MemoNumber = '" & passedTicketNumber & "'"

	Set rsMarkAlertEmailSent = Server.CreateObject("ADODB.Recordset")
	rsMarkAlertEmailSent.CursorLocation = 3 
	Set rsMarkAlertEmailSent= cnnMarkAlertEmailSent.Execute(SQLMarkAlertEmailSent)
	
	set rsMarkAlertEmailSent= Nothing
	set cnnMarkAlertEmailSent= Nothing
	
	MarkEscalationAlertEmailSent= result

End Function

Function MarkHoldEscalationAlertEmailSent(passedTicketNumber)

	result = 0
	
	Set cnnMarkAlertEmailSent = Server.CreateObject("ADODB.Connection")
	cnnMarkAlertEmailSent.open Session("ClientCnnString")

	SQLMarkAlertEmailSent = "Update FS_ServiceMemos Set HoldEscalationAlertEmailSent = getdate() where MemoNumber = '" & passedTicketNumber & "'"

	Set rsMarkAlertEmailSent = Server.CreateObject("ADODB.Recordset")
	rsMarkAlertEmailSent.CursorLocation = 3 
	Set rsMarkAlertEmailSent= cnnMarkAlertEmailSent.Execute(SQLMarkAlertEmailSent)
	
	set rsMarkAlertEmailSent= Nothing
	set cnnMarkAlertEmailSent= Nothing
	
	MarkHoldEscalationAlertEmailSent= result

End Function

Function AlertEmailSent(passedTicketNumber)

	result = False
	
	Set cnnMarkAlertEmailSent = Server.CreateObject("ADODB.Connection")
	cnnMarkAlertEmailSent.open Session("ClientCnnString")

	SQLMarkAlertEmailSent = "Select * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "'"

	Set rsMarkAlertEmailSent = Server.CreateObject("ADODB.Recordset")
	rsMarkAlertEmailSent.CursorLocation = 3 
	Set rsMarkAlertEmailSent = cnnMarkAlertEmailSent.Execute(SQLMarkAlertEmailSent)
	
	If rsMarkAlertEmailSent("AlertEmailSent") <> "" Then result = True
	
	set rsMarkAlertEmailSent= Nothing
	set cnnMarkAlertEmailSent= Nothing
	
	AlertEmailSent = result

End Function

Function HoldAlertEmailSent(passedTicketNumber)

	result = False
	
	Set cnnMarkAlertEmailSent = Server.CreateObject("ADODB.Connection")
	cnnMarkAlertEmailSent.open Session("ClientCnnString")

	SQLMarkAlertEmailSent = "Select * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "'"

	Set rsMarkAlertEmailSent = Server.CreateObject("ADODB.Recordset")
	rsMarkAlertEmailSent.CursorLocation = 3 
	Set rsMarkAlertEmailSent = cnnMarkAlertEmailSent.Execute(SQLMarkAlertEmailSent)
	
	If rsMarkAlertEmailSent("HoldAlertEmailSent") <> "" Then result = True
	
	set rsMarkAlertEmailSent= Nothing
	set cnnMarkAlertEmailSent= Nothing
	
	HoldAlertEmailSent = result

End Function


Function EscalationAlertEmailSent(passedTicketNumber)

	result = 0
	
	Set cnnMarkAlertEmailSent = Server.CreateObject("ADODB.Connection")
	cnnMarkAlertEmailSent.open Session("ClientCnnString")

	SQLMarkAlertEmailSent = "Select * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "'"

	Set rsMarkAlertEmailSent = Server.CreateObject("ADODB.Recordset")
	rsMarkAlertEmailSent.CursorLocation = 3 
	Set rsMarkAlertEmailSent= cnnMarkAlertEmailSent.Execute(SQLMarkAlertEmailSent)

	If rsMarkAlertEmailSent("EscalationAlertEmailSent") <> "" Then result = True
		
	set rsMarkAlertEmailSent= Nothing
	set cnnMarkAlertEmailSent= Nothing
	
	EscalationAlertEmailSent = result

End Function

Function HoldEscalationAlertEmailSent(passedTicketNumber)

	result = 0
	
	Set cnnMarkAlertEmailSent = Server.CreateObject("ADODB.Connection")
	cnnMarkAlertEmailSent.open Session("ClientCnnString")

	SQLMarkAlertEmailSent = "Select * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "'"

	Set rsMarkAlertEmailSent = Server.CreateObject("ADODB.Recordset")
	rsMarkAlertEmailSent.CursorLocation = 3 
	Set rsMarkAlertEmailSent= cnnMarkAlertEmailSent.Execute(SQLMarkAlertEmailSent)

	If rsMarkAlertEmailSent("HoldEscalationAlertEmailSent") <> "" Then result = True
		
	set rsMarkAlertEmailSent= Nothing
	set cnnMarkAlertEmailSent= Nothing
	
	HoldEscalationAlertEmailSent = result

End Function

Function HoldEscalationAlertEmailSent(passedTicketNumber)

	result = 0
	
	Set cnnMarkAlertEmailSent = Server.CreateObject("ADODB.Connection")
	cnnMarkAlertEmailSent.open Session("ClientCnnString")

	SQLMarkAlertEmailSent = "Select * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "'"

	Set rsMarkAlertEmailSent = Server.CreateObject("ADODB.Recordset")
	rsMarkAlertEmailSent.CursorLocation = 3 
	Set rsMarkAlertEmailSent= cnnMarkAlertEmailSent.Execute(SQLMarkAlertEmailSent)

	If rsMarkAlertEmailSent("HoldEscalationAlertEmailSent") <> "" Then result = True
		
	set rsMarkAlertEmailSent= Nothing
	set cnnMarkAlertEmailSent= Nothing
	
	HoldEscalationAlertEmailSent = result

End Function

Function BusinessDayStartorEndTime(passedSorE)

	resultBusinessDayStartorEndTime = 0
	
	Set cnnBusinessDayStartorEndTime = Server.CreateObject("ADODB.Connection")
	cnnBusinessDayStartorEndTime.open (Session("ClientCnnString"))
	Set rsBusinessDayStartorEndTime = Server.CreateObject("ADODB.Recordset")
	rsBusinessDayStartorEndTime.CursorLocation = 3 
	SQLBusinessDayStartorEndTime = "SELECT * FROM Settings_CompanyID"
	
	Set rsBusinessDayStartorEndTime = cnnBusinessDayStartorEndTime.Execute(SQLBusinessDayStartorEndTime)


	If Ucase(passedSorE) = "S" Then resultBusinessDayStartorEndTime = rsBusinessDayStartorEndTime("BusinessDayStart")
	If Ucase(passedSorE) = "E" Then resultBusinessDayStartorEndTime = rsBusinessDayStartorEndTime("BusinessDayEnd")

	Set rsBusinessDayStartorEndTime = Nothing
	cnnBusinessDayStartorEndTime.close	
	Set cnnBusinessDayStartorEndTime=Nothing
	
	BusinessDayStartorEndTime = resultBusinessDayStartorEndTime 
	
End Function

Function NumOpenCallsByAcct(passedCustNum)

	resultNumOpenCallsByAcct = 0
	
	Set cnnNumOpenCallsByAcct = Server.CreateObject("ADODB.Connection")
	cnnNumOpenCallsByAcct.open Session("ClientCnnString")

	SQLNumOpenCallsByAcct = "Select Count(AccountNumber) AS Expr1 from FS_ServiceMemos where AccountNumber = '" & passedCustNum & "' AND CurrentStatus ='OPEN' AND RecordSubType <>'HOLD'"

	Set rsNumOpenCallsByAcct = Server.CreateObject("ADODB.Recordset")
	rsNumOpenCallsByAcct.CursorLocation = 3 
	Set rsNumOpenCallsByAcct= cnnNumOpenCallsByAcct.Execute(SQLNumOpenCallsByAcct)
	
	resultNumOpenCallsByAcct = rsNumOpenCallsByAcct("Expr1")
	
	set rsNumOpenCallsByAcct= Nothing
	cnnNumOpenCallsByAcct.close
	set cnnNumOpenCallsByAcct= Nothing
	
	NumOpenCallsByAcct = resultNumOpenCallsByAcct 

End Function

Function GetTerm(passedGenericTerm)

	resultGetTerm = passedGenericTerm ' Default back to the term passed if not found
		
	Set cnnGetTerm = Server.CreateObject("ADODB.Connection")
	cnnGetTerm.open Session("ClientCnnString")

	SQLGetTerm = "Select * from SC_Terminology where GenericTerm = '" & passedGenericTerm & "'"

	Set rsGetTerm = Server.CreateObject("ADODB.Recordset")
	rsGetTerm.CursorLocation = 3 
	Set rsGetTerm= cnnGetTerm.Execute(SQLGetTerm)
	
	If not rsGetTerm.Eof Then resultGetTerm = rsGetTerm("CustomTerm")
	
	set rsGetTerm= Nothing
	cnnGetTerm.close
	set cnnGetTerm= Nothing
	
	If isMixedCase(passedGenericTerm) = False Then ' it is all upper or all lower	
		Dim re
		Set re = New RegExp
		re.Pattern = "^[A-Z]$"
		If re.Test(Left(passedGenericTerm,1)) Then
		   resultGetTerm = Ucase(resultGetTerm)
		Else
		   resultGetTerm = Lcase(resultGetTerm)
		End If
	End If
	
	GetTerm = resultGetTerm 

End Function

Function isMixedCase( str )
		'detects mixed case strings using the english alphabet
		'ignores spaces and other non-alphabetic characters
		str = trim( str )
		isMixed = false
		lastCase = ""
		for i=1 to len( str )
			currentChar = mid( str, i, 1 )
			if asc( currentChar ) >= 65 and asc( currentChar ) <= 90 then
				if lastCase <> "" and lastCase <> "upper" then
					isMixed = true
					exit for
				else
					lastCase = "upper"
				end if
			else
				if asc( currentChar ) >= 97 and asc( currentChar ) <= 122 then
					'lower
					if lastCase <> "" and lastCase <> "lower" then
						isMixed = true
						exit for
					else
						lastCase = "lower"
					end if
				else
					lastCase = ""
				end if
			end if
		next
		if isMixed then
			isMixedCase = true
		else
			isMixedCase = false
		end if
End Function

Function TicketWasOnHold(passedTicketNumber)

	resultTicketWasOnHold = False
	
	Set cnnTicketWasOnHold = Server.CreateObject("ADODB.Connection")
	cnnTicketWasOnHold.open Session("ClientCnnString")

	SQLTicketWasOnHold = "Select * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "' AND RecordSubType='HOLD'"

	Set rsTicketWasOnHold = Server.CreateObject("ADODB.Recordset")
	rsTicketWasOnHold.CursorLocation = 3 
	Set rsTicketWasOnHold = cnnTicketWasOnHold.Execute(SQLTicketWasOnHold)
	
	If Not rsTicketWasOnHold.Eof Then resultTicketWasOnHold = True
	
	set rsTicketWasOnHold = Nothing
	set cnnTicketWasOnHold = Nothing
	
	TicketWasOnHold = resultTicketWasOnHold

End Function

Function NumberOfArchivedNotes(passedCustNum)

	resultNumberOfArchivedNotes = 0 

	SQLNumberOfArchivedNotes = "SELECT count(*) as Expr1 FROM tblCustomerNotes Where CustNum ='" & passedCustNum & "' AND Archived = 1"
	
	Set cnnNumberOfArchivedNotes = Server.CreateObject("ADODB.Connection")
	cnnNumberOfArchivedNotes.open (Session("ClientCnnString"))
	Set rNumberOfArchivedNotes = Server.CreateObject("ADODB.Recordset")
	rNumberOfArchivedNotes.CursorLocation = 3 
	Set rNumberOfArchivedNotes = cnnNumberOfArchivedNotes.Execute(SQLNumberOfArchivedNotes)

	If not rNumberOfArchivedNotes.EOF Then resultNumberOfArchivedNotes = rNumberOfArchivedNotes ("Expr1")
	
	cnnNumberOfArchivedNotes.close
	set rNumberOfArchivedNotes = nothing
	set cnnNumberOfArchivedNotes= nothing	
	
	NumberOfArchivedNotes = resultNumberOfArchivedNotes 
	
End Function

Function NumberOfCurrentNotes(passedCustNum)

	resultNumberOfCurrentNotes = 0 

	SQLNumberOfCurrentNotes = "SELECT count(*) as Expr1 FROM tblCustomerNotes Where CustNum ='" & passedCustNum & "' AND Archived <> 1"
	
	Set cnnNumberOfCurrentNotes = Server.CreateObject("ADODB.Connection")
	cnnNumberOfCurrentNotes.open (Session("ClientCnnString"))
	Set rNumberOfCurrentNotes = Server.CreateObject("ADODB.Recordset")
	rNumberOfCurrentNotes.CursorLocation = 3 
	Set rNumberOfCurrentNotes = cnnNumberOfCurrentNotes.Execute(SQLNumberOfCurrentNotes)

	If not rNumberOfCurrentNotes.EOF Then resultNumberOfCurrentNotes = rNumberOfCurrentNotes ("Expr1")
	
	cnnNumberOfCurrentNotes.close
	set rNumberOfCurrentNotes = nothing
	set cnnNumberOfCurrentNotes= nothing	
	
	NumberOfCurrentNotes = resultNumberOfCurrentNotes 
	
End Function

Function NumberOfAttachmentsNotes(passedCustNum)

	resultNumberOfAttachmentsNotes = 0 

	SQLNumberOfAttachmentsNotes = "SELECT count(*) as Expr1 FROM tblCustomerNotesAttachments Where CustNum ='" & passedCustNum & "'"
	
	Set cnnNumberOfAttachmentsNotes = Server.CreateObject("ADODB.Connection")
	cnnNumberOfAttachmentsNotes.open (Session("ClientCnnString"))
	Set rNumberOfAttachmentsNotes = Server.CreateObject("ADODB.Recordset")
	rNumberOfAttachmentsNotes.CursorLocation = 3 
	Set rNumberOfAttachmentsNotes = cnnNumberOfAttachmentsNotes.Execute(SQLNumberOfAttachmentsNotes)

	If not rNumberOfAttachmentsNotes.EOF Then resultNumberOfAttachmentsNotes = rNumberOfAttachmentsNotes ("Expr1")
	
	cnnNumberOfAttachmentsNotes.close
	set rNumberOfAttachmentsNotes = nothing
	set cnnNumberOfAttachmentsNotes= nothing	
	
	NumberOfAttachmentsNotes = resultNumberOfAttachmentsNotes 
	
End Function

Function NumberOfServiceTicketsEver(passedCustNum)

	resultNumberOfServiceTicketsEver = 0 

	SQLNumberOfServiceTicketsEver = "SELECT count(Distinct MemoNumber) as Expr1 FROM FS_ServiceMemos Where AccountNumber ='" & passedCustNum & "'"
	
	Set cnnNumberOfServiceTicketsEver = Server.CreateObject("ADODB.Connection")
	cnnNumberOfServiceTicketsEver.open (Session("ClientCnnString"))
	Set rNumberOfServiceTicketsEver = Server.CreateObject("ADODB.Recordset")
	rNumberOfServiceTicketsEver.CursorLocation = 3 
	Set rNumberOfServiceTicketsEver = cnnNumberOfServiceTicketsEver.Execute(SQLNumberOfServiceTicketsEver)

	If not rNumberOfServiceTicketsEver.EOF Then resultNumberOfServiceTicketsEver = rNumberOfServiceTicketsEver ("Expr1")
	
	cnnNumberOfServiceTicketsEver.close
	set rNumberOfServiceTicketsEver = nothing
	set cnnNumberOfServiceTicketsEver= nothing	
	
	NumberOfServiceTicketsEver = resultNumberOfServiceTicketsEver 
	
End Function

Function NumberOfServiceTicketsOpenForCust(passedCustNum)

	resultNumberOfServiceTicketsOpenForCust = 0 

	SQLNumberOfServiceTicketsOpenForCust = "SELECT count(Distinct Memonumber) as Expr1 From FS_ServiceMemos Where AccountNumber ='" & passedCustNum & "' AND CurrentStatus='OPEN'"
	
	Set cnnNumberOfServiceTicketsOpenForCust = Server.CreateObject("ADODB.Connection")
	cnnNumberOfServiceTicketsOpenForCust.open (Session("ClientCnnString"))
	Set rNumberOfServiceTicketsOpenForCust = Server.CreateObject("ADODB.Recordset")
	rNumberOfServiceTicketsOpenForCust.CursorLocation = 3 
	Set rNumberOfServiceTicketsOpenForCust = cnnNumberOfServiceTicketsOpenForCust.Execute(SQLNumberOfServiceTicketsOpenForCust)

	If not rNumberOfServiceTicketsOpenForCust.EOF Then resultNumberOfServiceTicketsOpenForCust = rNumberOfServiceTicketsOpenForCust ("Expr1")
	
	cnnNumberOfServiceTicketsOpenForCust.close
	set rNumberOfServiceTicketsOpenForCust = nothing
	set cnnNumberOfServiceTicketsOpenForCust= nothing	
	
	NumberOfServiceTicketsOpenForCust = resultNumberOfServiceTicketsOpenForCust 
	
End Function


Function GetServiceTicketCurrentStage(passedTicketNumber)

	'Use only when advanced dispatch module is on

	resultGetServiceTicketCurrentStatus = "Received"
	
	Set cnnGetServiceTicketCurrentStatus = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketCurrentStatus.open Session("ClientCnnString")

	SQLGetServiceTicketCurrentStatus = "Select * from FS_ServiceMemosDetail where MemoNumber = '" & passedTicketNumber & "' Order By SubmissionDateTime Desc"

	Set rsGetServiceTicketCurrentStatus = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketCurrentStatus.CursorLocation = 3 
	Set rsGetServiceTicketCurrentStatus = cnnGetServiceTicketCurrentStatus.Execute(SQLGetServiceTicketCurrentStatus)
	
	If not rsGetServiceTicketCurrentStatus.eof then 
		resultGetServiceTicketCurrentStatus = rsGetServiceTicketCurrentStatus("MemoStage")
	End IF	
	
	'Special code for Under Review
	'Bacause under review happens automatically, sometimes users cross each other
	'and when a record is marked released, it get marked under review because another
	'user was already sitting on the screen. If Under Review, do another lookup to make sure
	'it really hasn't been released
	If not rsGetServiceTicketCurrentStatus.eof then 
		If resultGetServiceTicketCurrentStatus = "Under Review" Then
			rsGetServiceTicketCurrentStatus.MoveFirst
				Do
					If rsGetServiceTicketCurrentStatus("MemoStage") = "Released" Then resultGetServiceTicketCurrentStatus = "Released"
					
					rsGetServiceTicketCurrentStatus.Movenext
				Loop Until rsGetServiceTicketCurrentStatus.Eof
		End If
	End If

	set rsGetServiceTicketCurrentStatus = Nothing
	cnnGetServiceTicketCurrentStatus.Close
	set cnnGetServiceTicketCurrentStatus = Nothing
	
	
	
	GetServiceTicketCurrentStage = resultGetServiceTicketCurrentStatus
	

End Function

Function GetServiceTicketDispatchedTech(passedTicketNumber)

	'Use only when advanced dispatch module is on

	resultGetServiceTicketDispatchedTech = ""
	
	Set cnnGetServiceTicketDispatchedTech = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketDispatchedTech.open Session("ClientCnnString")

	SQLGetServiceTicketDispatchedTech = "Select Top 1 * from FS_ServiceMemosDetail where MemoNumber = '" & passedTicketNumber & "' AND MemoStage = 'Dispatched' order by SubmissionDateTime desc"

	Set rsGetServiceTicketDispatchedTech = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketDispatchedTech.CursorLocation = 3 
	Set rsGetServiceTicketDispatchedTech = cnnGetServiceTicketDispatchedTech.Execute(SQLGetServiceTicketDispatchedTech)
	
	If not rsGetServiceTicketDispatchedTech.eof then 
		resultGetServiceTicketDispatchedTech = rsGetServiceTicketDispatchedTech("UserNoOfServiceTech")
	End IF	
	
	set rsGetServiceTicketDispatchedTech = Nothing
	cnnGetServiceTicketDispatchedTech.Close
	set cnnGetServiceTicketDispatchedTech = Nothing
	
	GetServiceTicketDispatchedTech= resultGetServiceTicketDispatchedTech

End Function



Function GetServiceTicketDispatchedDateTime(passedTicketNumber)

	'Use only when advanced dispatch module is on

	resultGetServiceTicketDispatchedDateTime = ""
	
	Set cnnGetServiceTicketDispatchedDateTime = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketDispatchedDateTime.open Session("ClientCnnString")

	SQLGetServiceTicketDispatchedDateTime = "Select Top 1 * from FS_ServiceMemosDetail where MemoNumber = '" & passedTicketNumber & "' AND MemoStage = 'Dispatched' Order By SubmissionDateTime Desc"

	Set rsGetServiceTicketDispatchedDateTime = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketDispatchedDateTime.CursorLocation = 3 
	Set rsGetServiceTicketDispatchedDateTime = cnnGetServiceTicketDispatchedDateTime.Execute(SQLGetServiceTicketDispatchedDateTime)
	
	If not rsGetServiceTicketDispatchedDateTime.eof then 
		resultGetServiceTicketDispatchedDateTime = rsGetServiceTicketDispatchedDateTime("SubmissionDateTime")
	End IF	
	
	set rsGetServiceTicketDispatchedDateTime = Nothing
	cnnGetServiceTicketDispatchedDateTime.Close
	set cnnGetServiceTicketDispatchedDateTime = Nothing
	
	GetServiceTicketDispatchedDateTime = resultGetServiceTicketDispatchedDateTime

End Function

Function NumberOfServiceTicketsHOLDForCust(passedCustNum)

	resultNumberOfServiceTicketsHOLDForCust = 0 

	SQLNumberOfServiceTicketsHOLDForCust = "SELECT count(Distinct Memonumber) as Expr1 From FS_ServiceMemos Where AccountNumber ='" & passedCustNum & "' AND CurrentStatus='HOLD'"
	
	Set cnnNumberOfServiceTicketsHOLDForCust = Server.CreateObject("ADODB.Connection")
	cnnNumberOfServiceTicketsHOLDForCust.open (Session("ClientCnnString"))
	Set rNumberOfServiceTicketsHOLDForCust = Server.CreateObject("ADODB.Recordset")
	rNumberOfServiceTicketsHOLDForCust.CursorLocation = 3 
	Set rNumberOfServiceTicketsHOLDForCust = cnnNumberOfServiceTicketsHOLDForCust.Execute(SQLNumberOfServiceTicketsHOLDForCust)

	If not rNumberOfServiceTicketsHOLDForCust.EOF Then resultNumberOfServiceTicketsHOLDForCust = rNumberOfServiceTicketsHOLDForCust ("Expr1")
	
	cnnNumberOfServiceTicketsHOLDForCust.close
	set rNumberOfServiceTicketsHOLDForCust = nothing
	set cnnNumberOfServiceTicketsHOLDForCust= nothing	
	
	NumberOfServiceTicketsHOLDForCust = resultNumberOfServiceTicketsHOLDForCust 
	
End Function

Function GetServiceTicketSTAGEDateTime(passedTicketNumber,passedStage)

	'Use only when advanced dispatch module is on

	resultGetServiceTicketSTAGEDateTime = ""
	
	Set cnnGetServiceTicketSTAGEDateTime = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketSTAGEDateTime.open Session("ClientCnnString")

	SQLGetServiceTicketSTAGEDateTime = "Select Top 1 * from FS_ServiceMemosDetail where MemoNumber = '" & passedTicketNumber & "' AND MemoStage = '" & passedStage & "' Order By SubmissionDateTime DEsc"

	Set rsGetServiceTicketSTAGEDateTime = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketSTAGEDateTime.CursorLocation = 3 
	Set rsGetServiceTicketSTAGEDateTime = cnnGetServiceTicketSTAGEDateTime.Execute(SQLGetServiceTicketSTAGEDateTime)
	
	If not rsGetServiceTicketSTAGEDateTime.eof then 
		resultGetServiceTicketSTAGEDateTime = rsGetServiceTicketSTAGEDateTime("SubmissionDateTime")
	Else
		'If not in there, then is received or released, get the OPEN record
		SQLGetServiceTicketSTAGEDateTime = "Select * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "' AND RecordSubType = 'OPEN'"
		Set rsGetServiceTicketSTAGEDateTime = cnnGetServiceTicketSTAGEDateTime.Execute(SQLGetServiceTicketSTAGEDateTime)
		If Not rsGetServiceTicketSTAGEDateTime.EOF Then resultGetServiceTicketSTAGEDateTime = rsGetServiceTicketSTAGEDateTime("RecordCreatedateTime")
	End IF	
	
	set rsGetServiceTicketSTAGEDateTime = Nothing
	cnnGetServiceTicketSTAGEDateTime.Close
	set cnnGetServiceTicketSTAGEDateTime = Nothing
	
	GetServiceTicketSTAGEDateTime = resultGetServiceTicketSTAGEDateTime

End Function

Function GetServiceTicketSTAGEUser(passedTicketNumber,passedStage)

	'Use only when advanced dispatch module is on

	resultGetServiceTicketSTAGEUser = ""
	
	Set cnnGetServiceTicketSTAGEUser = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketSTAGEUser.open Session("ClientCnnString")

	SQLGetServiceTicketSTAGEUser = "Select * from FS_ServiceMemosDetail where MemoNumber = '" & passedTicketNumber & "' AND MemoStage = '" & passedStage & "' Order By RecordCreatedDateTime Desc"

	Set rsGetServiceTicketSTAGEUser = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketSTAGEUser.CursorLocation = 3 
	Set rsGetServiceTicketSTAGEUser = cnnGetServiceTicketSTAGEUser.Execute(SQLGetServiceTicketSTAGEUser)
	
	If not rsGetServiceTicketSTAGEUser.eof then 
		resultGetServiceTicketSTAGEUser = GetUserDisplayNameByUserNo(rsGetServiceTicketSTAGEUser("UserNoSubmittingRecord"))
	End IF	
	
	set rsGetServiceTicketSTAGEUser = Nothing
	cnnGetServiceTicketSTAGEUser.Close
	set cnnGetServiceTicketSTAGEUser = Nothing
	
	GetServiceTicketSTAGEUser = resultGetServiceTicketSTAGEUser

End Function



Function TicketIsUrgent(passedTicketNumber)

	resultTicketIsUrgent = False
	
	Set cnnTicketIsUrgent = Server.CreateObject("ADODB.Connection")
	cnnTicketIsUrgent.open Session("ClientCnnString")

	'All recs should be the same so just get 1
	SQLTicketIsUrgent = "Select TOP 1 Urgent FROM FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "'"

	Set rsTicketIsUrgent = Server.CreateObject("ADODB.Recordset")
	rsTicketIsUrgent.CursorLocation = 3 
	Set rsTicketIsUrgent = cnnTicketIsUrgent.Execute(SQLTicketIsUrgent)
	
	If not rsTicketIsUrgent.eof then 
		If rsTicketIsUrgent("Urgent") = 1 Then resultTicketIsUrgent = True
	End IF	
	
	set rsTicketIsUrgent = Nothing
	cnnTicketIsUrgent.Close
	set cnnTicketIsUrgent = Nothing
	
	TicketIsUrgent = resultTicketIsUrgent

End Function

Function TicketOriginalDispatchDateTime(passedTicketNumber)

	resultTicketOriginalDispatchDateTime = ""
	
	Set cnnTicketOriginalDispatchDateTime = Server.CreateObject("ADODB.Connection")
	cnnTicketOriginalDispatchDateTime.open Session("ClientCnnString")

	SQLTicketOriginalDispatchDateTime = "Select TOP 1 * from FS_ServiceMemosDetail where MemoNumber = '" & passedTicketNumber & "' AND MemoStage = 'Dispatched'"

	Set rsTicketOriginalDispatchDateTime = Server.CreateObject("ADODB.Recordset")
	rsTicketOriginalDispatchDateTime.CursorLocation = 3 
	Set rsTicketOriginalDispatchDateTime = cnnTicketOriginalDispatchDateTime.Execute(SQLTicketOriginalDispatchDateTime)
	
	If not rsTicketOriginalDispatchDateTime.eof then resultTicketOriginalDispatchDateTime = rsTicketOriginalDispatchDateTime("OriginalDispatchDateTime")

	
	set rsTicketOriginalDispatchDateTime = Nothing
	cnnTicketOriginalDispatchDateTime.Close
	set cnnTicketOriginalDispatchDateTime = Nothing
	
	TicketOriginalDispatchDateTime = resultTicketOriginalDispatchDateTime

End Function

Function NumberOfServiceTicketsDispatchedToTech(passedServiceTechNum)

	resultNumberOfServiceTicketsDispatchedToTech = 0 

	SQLNumberOfServiceTicketsDispatchedToTech = "SELECT Distinct Memonumber From FS_ServiceMemosDetail Where UserNoOfServiceTech ='" & passedServiceTechNum & "' AND ClosedorCancelled <> 1"
	SQLNumberOfServiceTicketsDispatchedToTech = SQLNumberOfServiceTicketsDispatchedToTech & " AND (MemoStage  = 'Dispatched'"
	SQLNumberOfServiceTicketsDispatchedToTech = SQLNumberOfServiceTicketsDispatchedToTech & " OR MemoStage = 'Dispatch Acknowledged'"
	SQLNumberOfServiceTicketsDispatchedToTech = SQLNumberOfServiceTicketsDispatchedToTech & " OR MemoStage = 'En Route'"
	SQLNumberOfServiceTicketsDispatchedToTech = SQLNumberOfServiceTicketsDispatchedToTech & " OR MemoStage = 'On Site')"
	SQLNumberOfServiceTicketsDispatchedToTech = SQLNumberOfServiceTicketsDispatchedToTech & " AND MemoNumber Not In (Select MemoNumber from FS_ServiceMemosRedispatch)"
	
	
	Set cnnNumberOfServiceTicketsDispatchedToTech = Server.CreateObject("ADODB.Connection")
	cnnNumberOfServiceTicketsDispatchedToTech.open (Session("ClientCnnString"))
	Set rNumberOfServiceTicketsDispatchedToTech = Server.CreateObject("ADODB.Recordset")
	rNumberOfServiceTicketsDispatchedToTech.CursorLocation = 3 
	Set rNumberOfServiceTicketsDispatchedToTech = cnnNumberOfServiceTicketsDispatchedToTech.Execute(SQLNumberOfServiceTicketsDispatchedToTech)

	If not rNumberOfServiceTicketsDispatchedToTech.EOF Then
		Do While Not rNumberOfServiceTicketsDispatchedToTech.EOF
			If LastTechUserNo(rNumberOfServiceTicketsDispatchedToTech("MemoNumber")) = passedServiceTechNum Then ' If we are not the latest tech, it was reassigned & isnt ours anymore
				resultNumberOfServiceTicketsDispatchedToTech = resultNumberOfServiceTicketsDispatchedToTech + 1
			End If
			rNumberOfServiceTicketsDispatchedToTech.MoveNext
		Loop	
	End IF	
	cnnNumberOfServiceTicketsDispatchedToTech.close
	set rNumberOfServiceTicketsDispatchedToTech = nothing
	set cnnNumberOfServiceTicketsDispatchedToTech= nothing	
	
	NumberOfServiceTicketsDispatchedToTech = resultNumberOfServiceTicketsDispatchedToTech 
	
End Function



Function NumberOfServiceTicketsAcknowledgedByTech(passedServiceTechNum)

	resultNumberOfServiceTicketsAcknowledgedByTech = 0 

	SQLNumberOfServiceTicketsAcknowledgedByTech = "SELECT Distinct Memonumber From FS_ServiceMemosDetail Where UserNoOfServiceTech ='" & passedServiceTechNum & "' AND ClosedorCancelled <> 1"
	SQLNumberOfServiceTicketsAcknowledgedByTech = SQLNumberOfServiceTicketsAcknowledgedByTech & " AND (MemoStage = 'Dispatch Acknowledged'"
	SQLNumberOfServiceTicketsAcknowledgedByTech = SQLNumberOfServiceTicketsAcknowledgedByTech & " OR MemoStage = 'En Route'"
	SQLNumberOfServiceTicketsAcknowledgedByTech = SQLNumberOfServiceTicketsAcknowledgedByTech & " OR MemoStage = 'On Site')"
	SQLNumberOfServiceTicketsAcknowledgedByTech = SQLNumberOfServiceTicketsAcknowledgedByTech & " AND MemoNumber Not In (Select MemoNumber from FS_ServiceMemosRedispatch)"
	
	
	Set cnnNumberOfServiceTicketsAcknowledgedByTech = Server.CreateObject("ADODB.Connection")
	cnnNumberOfServiceTicketsAcknowledgedByTech.open (Session("ClientCnnString"))
	Set rNumberOfServiceTicketsAcknowledgedByTech = Server.CreateObject("ADODB.Recordset")
	rNumberOfServiceTicketsAcknowledgedByTech.CursorLocation = 3 
	Set rNumberOfServiceTicketsAcknowledgedByTech = cnnNumberOfServiceTicketsAcknowledgedByTech.Execute(SQLNumberOfServiceTicketsAcknowledgedByTech)

	If not rNumberOfServiceTicketsAcknowledgedByTech.EOF Then
		Do While Not rNumberOfServiceTicketsAcknowledgedByTech.EOF
			If LastTechUserNo(rNumberOfServiceTicketsAcknowledgedByTech("MemoNumber")) = passedServiceTechNum Then ' If we are not the latest tech, it was reassigned & isnt ours anymore
				resultNumberOfServiceTicketsAcknowledgedByTech = resultNumberOfServiceTicketsAcknowledgedByTech + 1
			End If
			rNumberOfServiceTicketsAcknowledgedByTech.MoveNext
		Loop	
	End IF	
	cnnNumberOfServiceTicketsAcknowledgedByTech.close
	set rNumberOfServiceTicketsAcknowledgedByTech = nothing
	set cnnNumberOfServiceTicketsAcknowledgedByTech= nothing	
	
	NumberOfServiceTicketsAcknowledgedByTech = resultNumberOfServiceTicketsAcknowledgedByTech 
	
End Function




Function NumberOfServiceTicketsAwaitingACKFromTech(passedServiceTechNum)

	resultNumberOfServiceTicketsAwaitingACKFromTech= 0 

	SQLNumberOfServiceTicketsDispatchedToTech = "SELECT Distinct MemoNumber From FS_ServiceMemosDetail Where UserNoOfServiceTech ='" 
	SQLNumberOfServiceTicketsDispatchedToTech = SQLNumberOfServiceTicketsDispatchedToTech & passedServiceTechNum & "' AND MemoStage  = 'Dispatched' AND MemoNumber IN "
	SQLNumberOfServiceTicketsDispatchedToTech = SQLNumberOfServiceTicketsDispatchedToTech & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus='OPEN')"
	
	Set cnnNumberOfServiceTicketsDispatchedToTech = Server.CreateObject("ADODB.Connection")
	cnnNumberOfServiceTicketsDispatchedToTech.open (Session("ClientCnnString"))
	Set rNumberOfServiceTicketsDispatchedToTech = Server.CreateObject("ADODB.Recordset")
	rNumberOfServiceTicketsDispatchedToTech.CursorLocation = 3 
	Set rNumberOfServiceTicketsDispatchedToTech = cnnNumberOfServiceTicketsDispatchedToTech.Execute(SQLNumberOfServiceTicketsDispatchedToTech)

	If not rNumberOfServiceTicketsDispatchedToTech.EOF Then 
		Do While Not rNumberOfServiceTicketsDispatchedToTech.eof
			If GetServiceTicketCurrentStage(rNumberOfServiceTicketsDispatchedToTech("MemoNumber")) = "Dispatched" AND LastTechUserNo(rNumberOfServiceTicketsDispatchedToTech("MemoNumber")) = passedServiceTechNum Then resultNumberOfServiceTicketsAwaitingACKFromTech = resultNumberOfServiceTicketsAwaitingACKFromTech + 1
			rNumberOfServiceTicketsDispatchedToTech.movenext
		Loop
	End IF	
	
	cnnNumberOfServiceTicketsDispatchedToTech.close
	set rNumberOfServiceTicketsDispatchedToTech = nothing
	set cnnNumberOfServiceTicketsDispatchedToTech= nothing	
	
	NumberOfServiceTicketsAwaitingACKFromTech = resultNumberOfServiceTicketsAwaitingACKFromTech
	
End Function







Function NumberOfServiceTicketsClosedOrRedoByTech(passedServiceTechNum)

	resultNumberOfServiceTicketsClosedOrRedoByTech = 0 
	
	SQLNumberOfServiceTicketsClosedOrRedoByTech = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE "
	
	SQLNumberOfServiceTicketsClosedOrRedoByTech = SQLNumberOfServiceTicketsClosedOrRedoByTech & "((CurrentStatus='CLOSE' AND RecordSubType = 'CLOSE') OR (CurrentStatus='CANCEL' AND RecordSubType = 'CANCEL')) "
	
	
	SQLNumberOfServiceTicketsClosedOrRedoByTech = SQLNumberOfServiceTicketsClosedOrRedoByTech & " AND Month(RecordCreateDateTime) = Month(getdate()) "
	
	SQLNumberOfServiceTicketsClosedOrRedoByTech = SQLNumberOfServiceTicketsClosedOrRedoByTech & " AND Day(RecordCreateDateTime) = Day(getdate()) "
	
	SQLNumberOfServiceTicketsClosedOrRedoByTech = SQLNumberOfServiceTicketsClosedOrRedoByTech & " AND Year(RecordCreateDateTime) = Year(getdate()) "	
		 
	SQLNumberOfServiceTicketsClosedOrRedoByTech = SQLNumberOfServiceTicketsClosedOrRedoByTech & " AND UserNoOfServiceTech = " & passedServiceTechNum & " OR "
	
	SQLNumberOfServiceTicketsClosedOrRedoByTech = SQLNumberOfServiceTicketsClosedOrRedoByTech & "MemoNumber In (Select MemoNumber from FS_ServiceMemosRedispatch WHERE MemoNumber IN (SELECT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN'))"

	
	Set cnnNumberOfServiceTicketsClosedOrRedoByTech = Server.CreateObject("ADODB.Connection")
	cnnNumberOfServiceTicketsClosedOrRedoByTech.open (Session("ClientCnnString"))
	Set rNumberOfServiceTicketsClosedOrRedoByTech = Server.CreateObject("ADODB.Recordset")
	rNumberOfServiceTicketsClosedOrRedoByTech.CursorLocation = 3 
	Set rNumberOfServiceTicketsClosedOrRedoByTech = cnnNumberOfServiceTicketsClosedOrRedoByTech.Execute(SQLNumberOfServiceTicketsClosedOrRedoByTech)

	If not rNumberOfServiceTicketsClosedOrRedoByTech.EOF Then
	
		Do While Not rNumberOfServiceTicketsClosedOrRedoByTech.EOF
	
			If LastTechUserNo(rNumberOfServiceTicketsClosedOrRedoByTech("MemoNumber")) = passedServiceTechNum Then ' If we are not the latest tech, it was reassigned & isnt ours anymore
				resultNumberOfServiceTicketsClosedOrRedoByTech = resultNumberOfServiceTicketsClosedOrRedoByTech  + 1 
			End If
			
			rNumberOfServiceTicketsClosedOrRedoByTech.MoveNext
		Loop

	End IF
	cnnNumberOfServiceTicketsClosedOrRedoByTech.close
	set rNumberOfServiceTicketsClosedOrRedoByTech = nothing
	set cnnNumberOfServiceTicketsClosedOrRedoByTech= nothing	
	
	NumberOfServiceTicketsClosedOrRedoByTech = resultNumberOfServiceTicketsClosedOrRedoByTech 
	
End Function





Function advancedDispatchIsOn()

	
	resultadvancedDispatchIsOn = False

	SQLadvancedDispatchIsOn = "SELECT * FROM tblServerInfo where clientKey='"& MUV_READ("ClientID") &"'"
	
	Set ConnectionadvancedDispatchIsOn = Server.CreateObject("ADODB.Connection")
	Set RecordsetadvancedDispatchIsOn = Server.CreateObject("ADODB.Recordset")
	RecordsetadvancedDispatchIsOn.CursorLocation = 3 
	ConnectionadvancedDispatchIsOn.Open "Driver={SQL Server};Server=66.201.99.15;Database=_BIInsight;Uid=biinsight;Pwd=Z32#kje4217;"
	RecordsetadvancedDispatchIsOn.Open SQLadvancedDispatchIsOn,ConnectionadvancedDispatchIsOn,3,3
	
	If RecordsetadvancedDispatchIsOn.EOF Then 
		resultadvancedDispatchIsOn = False
	Else
		If RecordsetadvancedDispatchIsOn("advancedDispatch") = 1 Then resultadvancedDispatchIsOn = True Else resultadvancedDispatchIsOn = False
	End If	
	
	RecordsetadvancedDispatchIsOn.Close
	ConnectionadvancedDispatchIsOn.Close
	Set RecordsetadvancedDispatchIsOn = Nothing
	Set ConnectionadvancedDispatchIsOn = Nothing

	
	advancedDispatchIsOn = resultadvancedDispatchIsOn 

End Function


Function GetServiceTicketCust(passedTicketNumber)

	result = ""
	
	Set cnnCust = Server.CreateObject("ADODB.Connection")
	cnnCust.open Session("ClientCnnString")

	SQLCust = "SELECT AccountNumber FROM FS_ServiceMemos WHERE MemoNumber = '" & passedTicketNumber & "' order by submissionDateTime desc"

	Set rsCust = Server.CreateObject("ADODB.Recordset")
	rsCust.CursorLocation = 3 
	Set rsCust = cnnCust.Execute(SQLCust )
	
	If not rsCust.eof then result = rsCust("AccountNumber")

	set rsCust = Nothing
	set cnnCust= Nothing
	
	GetServiceTicketCust = result

End Function

Function TicketIsFilterChange(passedTicketNumber)

	result = False
	
	Set cnnTicketIsFilterChange = Server.CreateObject("ADODB.Connection")
	cnnTicketIsFilterChange.open Session("ClientCnnString")

	SQLTicketIsFilterChange = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & passedTicketNumber & "'"

	Set rsTicketIsFilterChange = Server.CreateObject("ADODB.Recordset")
	rsTicketIsFilterChange.CursorLocation = 3 
	Set rsTicketIsFilterChange = cnnTicketIsFilterChange.Execute(SQLTicketIsFilterChange)
	
	If not rsTicketIsFilterChange.eof then
		If rsTicketIsFilterChange("FilterChange") = 1 Then result = True
	End If

	set rsTicketIsFilterChange = Nothing
	set cnnTicketIsFilterChange= Nothing
	
	TicketIsFilterChange = result

End Function


Function Redispatch(passedTicketNumber)

	resultRedispatch = 0
	
	Set cnnRedispatch = Server.CreateObject("ADODB.Connection")
	cnnRedispatch.open Session("ClientCnnString")

	SQLRedispatch = "Insert Into FS_ServiceMemosRedispatch (MemoNumber) VALUES ('" & passedTicketNumber & "')"

	Set rsRedispatch = Server.CreateObject("ADODB.Recordset")
	rsRedispatch.CursorLocation = 3 
	Set rsRedispatch = cnnRedispatch.Execute(SQLRedispatch)
	
	set rsRedispatch = Nothing
	set cnnRedispatch= Nothing
	
	Redispatch= resultRedispatch 

End Function

Function AwaitingRedispatch(passedTicketNumber)

	resultAwaitingRedispatch = False
	
	Set cnnAwaitingRedispatch = Server.CreateObject("ADODB.Connection")
	cnnAwaitingRedispatch.open Session("ClientCnnString")

	SQLAwaitingRedispatch = "Select * from FS_ServiceMemosRedispatch Where MemoNumber ='" & passedTicketNumber & "'"
	
	Set rsAwaitingRedispatch = Server.CreateObject("ADODB.Recordset")
	rsAwaitingRedispatch.CursorLocation = 3 
	Set rsAwaitingRedispatch = cnnAwaitingRedispatch.Execute(SQLAwaitingRedispatch)

	If not rsAwaitingRedispatch.Eof Then resultAwaitingRedispatch = True
		
	set rsAwaitingRedispatch = Nothing
	set cnnAwaitingRedispatch= Nothing
	
	AwaitingRedispatch= resultAwaitingRedispatch 

End Function

Function RemoveFromRedispatch(passedTicketNumber)

	resultRemoveFromRedispatch = 0	
	
	Set cnnRemoveFromRedispatch = Server.CreateObject("ADODB.Connection")
	cnnRemoveFromRedispatch.open Session("ClientCnnString")

	SQLRemoveFromRedispatch = "Delete from FS_ServiceMemosRedispatch Where MemoNumber ='" & passedTicketNumber & "'"
	
	Set rsRemoveFromRedispatch = Server.CreateObject("ADODB.Recordset")
	rsRemoveFromRedispatch.CursorLocation = 3 
	Set rsRemoveFromRedispatch = cnnRemoveFromRedispatch.Execute(SQLRemoveFromRedispatch)
	
	set rsRemoveFromRedispatch = Nothing
	set cnnRemoveFromRedispatch= Nothing
	
	RemoveFromRedispatch = resultRemoveFromRedispatch 

End Function

Function GetChainDescByChainNum(passedChainNum)

	resultGetChainDescByChainNum = ""
	
	Set cnnGetChainDescByChainNum = Server.CreateObject("ADODB.Connection")
	cnnGetChainDescByChainNum.open Session("ClientCnnString")

	SQLGetChainDescByChainNum = "Select description from Chain where ChainSequence = " & passedChainNum
	 
	Set rsGetChainDescByChainNum = Server.CreateObject("ADODB.Recordset")
	rsGetChainDescByChainNum.CursorLocation = 3 
	Set rsGetChainDescByChainNum= cnnGetChainDescByChainNum.Execute(SQLGetChainDescByChainNum)
	
	If not rsGetChainDescByChainNum.eof then resultGetChainDescByChainNum = rsGetChainDescByChainNum("description")
	
	set rsGetChainDescByChainNum = Nothing
	set cnnGetChainDescByChainNum = Nothing
	
	GetChainDescByChainNum = resultGetChainDescByChainNum

End Function

Function FilterChangeSubmitted(passedassetNumber,passedfilterchangedate)

	resultFilterChangeSubmitted = False
	
	Set cnnFilterChangeSubmitted = Server.CreateObject("ADODB.Connection")
	cnnFilterChangeSubmitted.open Session("ClientCnnString")

	SQLFilterChangeSubmitted = "Select * from tblAssetPMSubmitted Where assetnumber ='" & passedassetNumber& "' AND PMdate ='" & passedfilterchangedate & "' AND filterOrPM='F'"
	
	Set rsFilterChangeSubmitted = Server.CreateObject("ADODB.Recordset")
	rsFilterChangeSubmitted.CursorLocation = 3 
	Set rsFilterChangeSubmitted = cnnFilterChangeSubmitted.Execute(SQLFilterChangeSubmitted)

	If not rsFilterChangeSubmitted.Eof Then resultFilterChangeSubmitted = True
		
	set rsFilterChangeSubmitted = Nothing
	set cnnFilterChangeSubmitted= Nothing
	
	FilterChangeSubmitted= resultFilterChangeSubmitted 

End Function

Function LastTechUserNo(passedTicketNumber)

	'Use only when advanced dispatch module is on

	resultLastTechUserNo = "Received"
	
	Set cnnLastTechUserNo = Server.CreateObject("ADODB.Connection")
	cnnLastTechUserNo.open Session("ClientCnnString")

	SQLLastTechUserNo = "Select TOP 1 * from FS_ServiceMemosDetail where MemoNumber = '" & passedTicketNumber & "' Order By SubmissionDateTime Desc"

	Set rsLastTechUserNo = Server.CreateObject("ADODB.Recordset")
	rsLastTechUserNo.CursorLocation = 3 
	Set rsLastTechUserNo = cnnLastTechUserNo.Execute(SQLLastTechUserNo)
	
	If not rsLastTechUserNo.eof then 
		resultLastTechUserNo = rsLastTechUserNo("UserNoOfServiceTech")
	End IF	
	
	set rsLastTechUserNo = Nothing
	cnnLastTechUserNo.Close
	set cnnLastTechUserNo = Nothing
	
	LastTechUserNo = resultLastTechUserNo
	

End Function



Function filterChangeModuleOn()

	
	resultfilterChangeModuleOn = False

	SQLfilterChangeModuleOn = "SELECT * FROM tblServerInfo where clientKey='"& MUV_READ("ClientID")  &"'"
	
	Set ConnectionfilterChangeModuleOn = Server.CreateObject("ADODB.Connection")
	Set RecordsetfilterChangeModuleOn = Server.CreateObject("ADODB.Recordset")
	RecordsetfilterChangeModuleOn.CursorLocation = 3 
	ConnectionfilterChangeModuleOn.Open "Driver={SQL Server};Server=66.201.99.15;Database=_BIInsight;Uid=biinsight;Pwd=Z32#kje4217;"
	Set RecordsetfilterChangeModuleOn = ConnectionfilterChangeModuleOn.Execute(SQLfilterChangeModuleOn)
	
	If RecordsetfilterChangeModuleOn.EOF Then 
		resultfilterChangeModuleOn = False
	Else
		If RecordsetfilterChangeModuleOn("filterChangeModule") = 1 Then resultfilterChangeModuleOn = True Else resultfilterChangeModuleOn = False
	End If	
	
	RecordsetfilterChangeModuleOn.Close
	ConnectionfilterChangeModuleOn.Close
	Set RecordsetfilterChangeModuleOn = Nothing
	Set ConnectionfilterChangeModuleOn = Nothing

	'Cant have filter changes without advanced dispatch
	If advancedDispatchIsOn() <> True Then resultfilterChangeModuleOn = False
	
	filterChangeModuleOn = resultfilterChangeModuleOn 

End Function

Function GetCustRouteNum(passedCustID)

	resultGetCustRouteNum = 0

	Set cnnGetCustRouteNum = Server.CreateObject("ADODB.Connection")
	cnnGetCustRouteNum.open Session("ClientCnnString")
			
	SQLGetCustRouteNum = "Select * from AR_Customer where CustNum = '" & passedCustID & "'"
	 
	Set rsGetCustRouteNum = Server.CreateObject("ADODB.Recordset")
	rsGetCustRouteNum.CursorLocation = 3 
	Set rsGetCustRouteNum= cnnGetCustRouteNum.Execute(SQLGetCustRouteNum)
		
	If not rsGetCustRouteNum.eof then resultGetCustRouteNum = rsGetCustRouteNum("RouteNum")
	
	set rsGetCustRouteNum= Nothing
	cnnGetCustRouteNum.Close	
	set cnnGetCustRouteNum= Nothing
	
	GetCustRouteNum = resultGetCustRouteNum
	
End Function

Function GetMyFilterRoutes(passedUserNo)

	resultGetMyFilterRoutes = 0
	
	Set cnnGetMyFilterRoutes = Server.CreateObject("ADODB.Connection")
	cnnGetMyFilterRoutes.open Session("ClientCnnString")

	SQLGetMyFilterRoutes = "Select * from tblUsers where UserNo = " & passedUserNo
	Set rsGetMyFilterRoutes = Server.CreateObject("ADODB.Recordset")
	rsGetMyFilterRoutes.CursorLocation = 3 
	Set rsGetMyFilterRoutes = cnnGetMyFilterRoutes.Execute(SQLGetMyFilterRoutes)
	
	resultGetMyFilterRoutes = rsGetMyFilterRoutes("userFilterRoutes")

	set rsGetMyFilterRoutes = Nothing
	cnnGetMyFilterRoutes.Close
	set cnnGetMyFilterRoutes = Nothing
	
	GetMyFilterRoutes = resultGetMyFilterRoutes 

End Function

Function RemoveFilterChangeSubmitted(passedassetNumber,passedfilterchangedate)

	resultRemoveFilterChangeSubmitted = 0	
	
	Set cnnRemoveFilterChangeSubmitted = Server.CreateObject("ADODB.Connection")
	cnnRemoveFilterChangeSubmitted.open Session("ClientCnnString")

	SQLRemoveFilterChangeSubmitted = "Delete from tblAssetPMSubmitted Where assetnumber ='" & passedassetNumber& "' AND PMdate ='" & passedfilterchangedate & "' AND filterOrPM='F'"
	
	Set rsRemoveFilterChangeSubmitted = Server.CreateObject("ADODB.Recordset")
	rsRemoveFilterChangeSubmitted.CursorLocation = 3 
	Set rsRemoveFilterChangeSubmitted = cnnRemoveFilterChangeSubmitted.Execute(SQLRemoveFilterChangeSubmitted)
	
	set rsRemoveFilterChangeSubmitted = Nothing
	set cnnRemoveFilterChangeSubmitted= Nothing
	
	RemoveFilterChangeSubmitted = resultRemoveFilterChangeSubmitted 

End Function

Function GetCustNumberByInvoiceNum(passedInvoiceNumber)

	Set cnnGetCustNumberByInvoiceNum = Server.CreateObject("ADODB.Connection")
	cnnGetCustNumberByInvoiceNum.open Session("ClientCnnString")

	resultGetCustNumberByInvoiceNum = ""
		
	SQLGetCustNumberByInvoiceNum = "Select * from " & Session("SQL_Owner") & ".InvoiceHistory where IvsNum = " & passedInvoiceNumber
	 
	Set rsGetCustNumberByInvoiceNum = Server.CreateObject("ADODB.Recordset")
	rsGetCustNumberByInvoiceNum.CursorLocation = 3 
	Set rsGetCustNumberByInvoiceNum= cnnGetCustNumberByInvoiceNum.Execute(SQLGetCustNumberByInvoiceNum)
	
	
	If not rsGetCustNumberByInvoiceNum.eof then resultGetCustNumberByInvoiceNum = rsGetCustNumberByInvoiceNum("CustNum")
	
	set rsGetCustNumberByInvoiceNum= Nothing
	set cnnGetCustNumberByInvoiceNum= Nothing
	
	GetCustNumberByInvoiceNum = resultGetCustNumberByInvoiceNum
	
End Function


Sub Check_HOLD_Alerts

	'Now does all it's own lookup and runs through all tickets that are on hold
	
	'Only check alerts during business hours
	If AlertsDuringBusinessHoursOnly = vbTrue Then
		If WeekDay(Date()) = 1 or WeekDay(Date()) = 7 Then Exit Sub ' Sat or Sun
		If Hour(Now()) < Hour(BusinessDayStartorEndTime("S")) or (Hour(Now()) = Hour(BusinessDayStartorEndTime("S")) AND Minute(Now()) < Minute(BusinessDayStartorEndTime("S"))) Then Exit Sub
		If Hour(Now()) > Hour(BusinessDayStartorEndTime("E")) or (Hour(Now()) = Hour(BusinessDayStartorEndTime("E")) AND Minute(Now()) > Minute(BusinessDayStartorEndTime("E"))) Then Exit Sub
	End If
	
	Set cnn99 = Server.CreateObject("ADODB.Connection")
	cnn99.open (Session("ClientCnnString"))
	Set rs99 = Server.CreateObject("ADODB.Recordset")
	rs99.CursorLocation = 3 
	SQL99 = "SELECT * FROM Settings_EmailService "
	Set rs99 = cnn99.Execute(SQL99)
	If not rs99.EOF Then
		AlertsDuringBusinessHoursOnly = rs99("AlertsDuringBizHoursOnly")
		HoldAlertsOn = rs99("HoldAlertsOn")
		SendHoldAlertToFinanceManagers = rs99("SendHoldAlertToFinanceManagers")
		SendHoldAlertToAdditionalEmails = rs99("SendHoldAlertToAdditionalEmails")
		SendHoldAlertHours = rs99("SendHoldAlertHours")
		EscalationAlertHours = rs99("EscalationAlertHours")
		HoldEscalationAlertsOn = rs99("HoldEscalationAlertsOn")
		HoldEscalationAlertToEmails = rs99("HoldEscalationAlertToEmails")
		HoldEscalationAlertHours = rs99("HoldEscalationAlertHours")
	Else
		HoldAlertsOn = vbFalse
		HoldEscalationAlertsOn = vbFalse
	End If
	Set rs99 = Nothing
	cnn99.Close
	Set cnn99 = Nothing

	'Checks to see if any service call alerts need to go out
		
	Set cnn100 = Server.CreateObject("ADODB.Connection")
	cnn100.open Session("ClientCnnString")
		
	SQL100 = "Select * from " & Session("SQL_Owner") & ".FS_ServiceMemos where CurrentStatus = 'HOLD'"
	 
	Set rs100 = Server.CreateObject("ADODB.Recordset")
	rs100.CursorLocation = 3 
	Set rs100= cnn100.Execute(SQL100)

	If not rs100.Eof Then
	
		Do While Not rs100.Eof
		
			elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs100.Fields("MemoNumber")),Now())
			
			'Now need to account for alert periods of less than 1 hour
			If (cdbl(SendHoldAlertHours) >= 1 AND Round((elapsedMinutes /60),2) > cdbl(SendHoldAlertHours) and HoldAlertEmailSent(rs100.Fields("MemoNumber")) <> True)_
			OR (cdbl(SendHoldAlertHours) < 1 AND cdbl(elapsedMinutes) > cdbl(SendHoldAlertHours)*60 and HoldAlertEmailSent(rs100.Fields("MemoNumber")) <> True)Then ' Only if not already sentThen ' Yes, we need to send
		
				Send_To=""
				If SendHoldAlertToFinanceManagers = vbTrue Then
					'Get all the service manager email addresses
					Set cnn_CheckAlerts = Server.CreateObject("ADODB.Connection")
					cnn_CheckAlerts.open (Session("ClientCnnString"))
					Set rs_CheckAlerts = Server.CreateObject("ADODB.Recordset")
					rs_CheckAlerts.CursorLocation = 3 
					SQL_CheckAlerts = "SELECT userEmail FROM tblUsers WHERE userType = 'Finance Manager' and userArchived <> 1"
					Set rs_CheckAlerts = cnn_CheckAlerts.Execute(SQL_CheckAlerts)
					If not rs_CheckAlerts.EOF Then
						Do
							If rs_CheckAlerts("userEmail") <> "" AND Not IsNull(rs_CheckAlerts("userEmail")) Then Send_To = Send_To & rs_CheckAlerts("userEmail") & ";"
							rs_CheckAlerts.MoveNext
						Loop Until rs_CheckAlerts.Eof
					End If
					Set rs_CheckAlerts = Nothing
					cnn_CheckAlerts.Close
					Set cnn_CheckAlerts = Nothing
				End If		
					
					
				'Now see if there any additionals
				If SendHoldAlertToAdditionalEmails <> "" and not IsNull(SendHoldAlertToAdditionalEmails) Then
					tmpSendHoldAlertToAdditionalEmails = trim(SendHoldAlertToAdditionalEmails)		
					If Len(tmpSendHoldAlertToAdditionalEmails) > 1 Then
						If Right(tmpSendHoldAlertToAdditionalEmails,1) <> ";" Then tmpSendHoldAlertToAdditionalEmails = tmpSendHoldAlertToAdditionalEmails & ";"
						Send_To = Send_To & tmpSendHoldAlertToAdditionalEmails
					End If	
				End If
				
				'Got all the addresses so now break them up
				Send_To_Array = Split(Send_To,";")
		
				For x = 0 to Ubound(Send_To_Array) -1
					Send_To = Send_To_Array(x)
						%>
						<!--#include file="../emails/service_hold_alert.asp"-->						
						<%
					'Failsafe for dev
					If Instr(ucase(sURL),"DEV") <> 0 Then Send_To = "rich@ocsaccess.com"
					SendMail "mailsender@" & maildomain,Send_To,emailSubject,emailBody,GetTerm("Service"),"Hold Alert"
					MemNum = rs100.Fields("MemoNumber")
					CreateAuditLogEntry "Service Hold Alert Sent","Service Hold Alert Sent","Minor",0,"Service Hold Alert Sent to " & Send_To & " for ticket #: " & MemNum & " - " & Round((ElapsedMinutes/60),2) & " hours"
				Next 
				Dummy = MarkHoldAlertEmailSent(rs100.Fields("MemoNumber"))
			End If
			
			' Now do the escalations
			
			'Now need to account for alert periods of less than 1 hour
			If (cdbl(HoldEscalationAlertHours) >= 1 AND Round((ElapsedMinutes/60),2) > cdbl(HoldEscalationAlertHours) and HoldEscalationAlertEmailSent(rs100.Fields("MemoNumber")) <> True)_
			OR (cdbl(HoldEscalationAlertHours) < 1 AND cdbl(ElapsedMinutes) > cdbl(HoldEscalationAlertHours)*60 and HoldEscalationAlertEmailSent(rs100.Fields("MemoNumber")) <> True)Then ' Only if not already sentThen ' Yes, we need to send
		
				Send_To=""
		
					
				'Now see if there any additionals
				If HoldEscalationAlertToEmails <> "" and not IsNull(HoldEscalationAlertToEmails) Then
					tmpHoldEscalationAlertToEmails = trim(HoldEscalationAlertToEmails)		
					If Len(tmpHoldEscalationAlertToEmails) > 1 Then
						If Right(tmpHoldEscalationAlertToEmails,1) <> ";" Then tmpHoldEscalationAlertToEmails = tmpHoldEscalationAlertToEmails & ";"
						Send_To = Send_To & tmpHoldEscalationAlertToEmails
					End If	
				End If
				
				'Got all the addresses so now break them up
				Send_To_Array = Split(Send_To,";")
		
				For x = 0 to Ubound(Send_To_Array) -1
					Send_To = Send_To_Array(x)
					%>
					<!--#include file="../emails/service_hold_alert.asp"-->
					<%	
					'Failsafe for dev
					sURL = Request.ServerVariables("SERVER_NAME")
					If Instr(ucase(sURL),"DEV") <> 0 Then Send_To = "rich@ocsaccess.com"
					SendMail "mailsender@" & maildomain,Send_To,emailSubject,emailBody,GetTerm("Service"),"Hold Escalation"
					MemNum = rs100.Fields("MemoNumber")
					CreateAuditLogEntry "Service Hold Escalation Alert Sent","Service Hold Escalation Alert Sent","Minor",0,"Service Hold Escalation Alert Sent to " & Send_To & " for ticket #: " & MemNum & " - " & Round((ElapsedMinutes/60),2) & " hours"
				Next 
				Dummy = MarkHoldEscalationAlertEmailSent(rs100.Fields("MemoNumber"))
			End If
			
			rs100.movenext
			
		Loop
		
	End If
End Sub



Function CustPendingFilterChangeInfo(passedCustid)

	resultCustPendingFilterChangeInfo = ""
	
	'Remember, it reads the setting FieldServiceDays from tblSetting_Global to determine how many days to use in the evaluation
	Set cnnFChange_NextDate = Server.CreateObject("ADODB.Connection")
	Set rsFChange_NextDate = Server.CreateObject("ADODB.Recordset")
	
	cnnFChange_NextDate.open (Session("ClientCnnString"))
	
	SQLFChange_NextDate = "SELECT * FROM Settings_Global"
		
	rsFChange_NextDate.CursorLocation = 3 
	Set rsFChange_NextDate = cnnFChange_NextDate.Execute(SQLFChange_NextDate)
	
	If not rsFChange_NextDate.EOF Then
		FChangeDays = rsFChange_NextDate("FilterChangeDays")
	Else
		FChangeDays = 15
	End If
	
	
	Set cnnCustPendingFilterChangeInfo = Server.CreateObject("ADODB.Connection")
	cnnCustPendingFilterChangeInfo.open Session("ClientCnnString")

	SQLCustPendingFilterChangeInfo = "SELECT * ,"
	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " CASE WHEN FS_CustomerFilters.FrequencyType='D' THEN DATEADD(day, FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " WHEN FS_CustomerFilters.FrequencyType='M' THEN DATEADD(day, 28*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " WHEN FS_CustomerFilters.FrequencyType='W' THEN DATEADD(day, 7*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " ELSE FS_CustomerFilters.LastChangeDateTime END AS nextdate "
	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " FROM FS_CustomerFilters WHERE "
	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " CustID = '" & passedCustid & "' AND "
	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " CASE WHEN FS_CustomerFilters.FrequencyType='D' THEN DATEADD(day, FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " WHEN FS_CustomerFilters.FrequencyType='M' THEN DATEADD(day, 28*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " WHEN FS_CustomerFilters.FrequencyType='W' THEN DATEADD(day, 7*FS_CustomerFilters.FrequencyTime, FS_CustomerFilters.LastChangeDateTime) "
	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " ELSE FS_CustomerFilters.LastChangeDateTime END"
	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " <= DateAdd(day," & FChangeDays & ",getdate()) "
'	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " AND "
'	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & "("
'	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & " FS_CustomerFilters.FilterIntRecID NOT IN "
'	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & "(SELECT FilterIntRecID FROM FS_ServiceMemosFilterInfo WHERE CustID = '" & passedCustid & "' AND ServiceTicketID IN (Select MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus='OPEN')) "
'	SQLCustPendingFilterChangeInfo = SQLCustPendingFilterChangeInfo & ")"

	Set rsCustPendingFilterChangeInfo = Server.CreateObject("ADODB.Recordset")
	rsCustPendingFilterChangeInfo.CursorLocation = 3 
	Set rsCustPendingFilterChangeInfo = cnnCustPendingFilterChangeInfo.Execute(SQLCustPendingFilterChangeInfo)
	
	If not rsCustPendingFilterChangeInfo.eof then 
		Do While NOT rsCustPendingFilterChangeInfo.EOF
		
			resultCustPendingFilterChangeInfo = resultCustPendingFilterChangeInfo  & FormatDateTime(rsCustPendingFilterChangeInfo("nextDate")) & "   " &  rsCustPendingFilterChangeInfo("Notes") & vbCRLF
			
			rsCustPendingFilterChangeInfo.MoveNext
			
		Loop
	End If	
	
	set rsCustPendingFilterChangeInfo = Nothing
	cnnCustPendingFilterChangeInfo.Close
	set cnnCustPendingFilterChangeInfo = Nothing
	
	CustPendingFilterChangeInfo = resultCustPendingFilterChangeInfo

End Function


Function GetPONumberByInvoiceNum(passedInvoiceNumber)

	Set cnnGetPONumberByInvoiceNum = Server.CreateObject("ADODB.Connection")
	cnnGetPONumberByInvoiceNum.open Session("ClientCnnString")

	resultGetPONumberByInvoiceNum = ""
		
	SQLGetPONumberByInvoiceNum = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory where IvsNum = " & passedInvoiceNumber
	 
	Set rsGetPONumberByInvoiceNum = Server.CreateObject("ADODB.Recordset")
	rsGetPONumberByInvoiceNum.CursorLocation = 3 
	Set rsGetPONumberByInvoiceNum= cnnGetPONumberByInvoiceNum.Execute(SQLGetPONumberByInvoiceNum)
	
	
	If not rsGetPONumberByInvoiceNum.eof then resultGetPONumberByInvoiceNum = rsGetPONumberByInvoiceNum("PurchOrderNum")
	
	set rsGetPONumberByInvoiceNum= Nothing
	set cnnGetPONumberByInvoiceNum= Nothing
	
	GetPONumberByInvoiceNum = resultGetPONumberByInvoiceNum
	
End Function

Function GetRouteNumByInvoiceNum(passedInvoiceNumber)

	Set cnnGetRouteNumByInvoiceNum = Server.CreateObject("ADODB.Connection")
	cnnGetRouteNumByInvoiceNum.open Session("ClientCnnString")

	resultGetRouteNumByInvoiceNum = ""
		
	SQLGetRouteNumByInvoiceNum = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory where IvsNum = " & passedInvoiceNumber
	 
	Set rsGetRouteNumByInvoiceNum = Server.CreateObject("ADODB.Recordset")
	rsGetRouteNumByInvoiceNum.CursorLocation = 3 
	Set rsGetRouteNumByInvoiceNum= cnnGetRouteNumByInvoiceNum.Execute(SQLGetRouteNumByInvoiceNum)
	
	
	If not rsGetRouteNumByInvoiceNum.eof then resultGetRouteNumByInvoiceNum = rsGetRouteNumByInvoiceNum("RouteNum")
	
	set rsGetRouteNumByInvoiceNum= Nothing
	set cnnGetRouteNumByInvoiceNum= Nothing
	
	GetRouteNumByInvoiceNum = resultGetRouteNumByInvoiceNum
	
End Function

Function GetRouteNameByRouteNum(passedRouteNumber)

	Set cnnGetRouteNameByRouteNum = Server.CreateObject("ADODB.Connection")
	cnnGetRouteNameByRouteNum.open Session("ClientCnnString")

	resultGetRouteNameByRouteNum = ""
		
	SQLGetRouteNameByRouteNum = "Select * from " & MUV_Read("SQL_Owner") & ".Routes where RouteSequence = " & passedRouteNumber
	 
	Set rsGetRouteNameByRouteNum = Server.CreateObject("ADODB.Recordset")
	rsGetRouteNameByRouteNum.CursorLocation = 3 
	Set rsGetRouteNameByRouteNum= cnnGetRouteNameByRouteNum.Execute(SQLGetRouteNameByRouteNum)
	
	
	If not rsGetRouteNameByRouteNum.eof then resultGetRouteNameByRouteNum = rsGetRouteNameByRouteNum("RouteName")
	
	set rsGetRouteNameByRouteNum= Nothing
	set cnnGetRouteNameByRouteNum= Nothing
	
	GetRouteNameByRouteNum = resultGetRouteNameByRouteNum
	
End Function

Function GetTermsNumByInvoiceNum(passedInvoiceNumber)

	Set cnnGetTermsNumByInvoiceNum = Server.CreateObject("ADODB.Connection")
	cnnGetTermsNumByInvoiceNum.open Session("ClientCnnString")

	resultGetTermsNumByInvoiceNum = ""
		
	SQLGetTermsNumByInvoiceNum = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory where IvsNum = " & passedInvoiceNumber
	 
	Set rsGetTermsNumByInvoiceNum = Server.CreateObject("ADODB.Recordset")
	rsGetTermsNumByInvoiceNum.CursorLocation = 3 
	Set rsGetTermsNumByInvoiceNum= cnnGetTermsNumByInvoiceNum.Execute(SQLGetTermsNumByInvoiceNum)
	
	
	If not rsGetTermsNumByInvoiceNum.eof then resultGetTermsNumByInvoiceNum = rsGetTermsNumByInvoiceNum("Terms")
	
	set rsGetTermsNumByInvoiceNum= Nothing
	set cnnGetTermsNumByInvoiceNum= Nothing
	
	GetTermsNumByInvoiceNum = resultGetTermsNumByInvoiceNum
	
End Function

Function GetTermsNameByTermsNum(passedTermsNumber)

	Set cnnGetTermsNameByTermsNum = Server.CreateObject("ADODB.Connection")
	cnnGetTermsNameByTermsNum.open Session("ClientCnnString")

	resultGetTermsNameByTermsNum = ""
		
	SQLGetTermsNameByTermsNum = "Select * from " & MUV_Read("SQL_Owner") & ".Terms where TermsSequence = " & passedTermsNumber
	 
	Set rsGetTermsNameByTermsNum = Server.CreateObject("ADODB.Recordset")
	rsGetTermsNameByTermsNum.CursorLocation = 3 
	Set rsGetTermsNameByTermsNum= cnnGetTermsNameByTermsNum.Execute(SQLGetTermsNameByTermsNum)
	
	
	If not rsGetTermsNameByTermsNum.eof then resultGetTermsNameByTermsNum = rsGetTermsNameByTermsNum("Description")
	
	set rsGetTermsNameByTermsNum= Nothing
	set cnnGetTermsNameByTermsNum= Nothing
	
	GetTermsNameByTermsNum = resultGetTermsNameByTermsNum
	
End Function

Function GetPrimarySalesmanByInvoiceNum(passedInvoiceNumber)

	Set cnnGetPrimarySalesmanByInvoiceNum = Server.CreateObject("ADODB.Connection")
	cnnGetPrimarySalesmanByInvoiceNum.open Session("ClientCnnString")

	resultGetPrimarySalesmanByInvoiceNum = ""
		
	SQLGetPrimarySalesmanByInvoiceNum = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory where IvsNum = " & passedInvoiceNumber
	 
	Set rsGetPrimarySalesmanByInvoiceNum = Server.CreateObject("ADODB.Recordset")
	rsGetPrimarySalesmanByInvoiceNum.CursorLocation = 3 
	Set rsGetPrimarySalesmanByInvoiceNum= cnnGetPrimarySalesmanByInvoiceNum.Execute(SQLGetPrimarySalesmanByInvoiceNum)
	
	
	If not rsGetPrimarySalesmanByInvoiceNum.eof then resultGetPrimarySalesmanByInvoiceNum = rsGetPrimarySalesmanByInvoiceNum("PrimarySalesman")
	
	set rsGetPrimarySalesmanByInvoiceNum= Nothing
	set cnnGetPrimarySalesmanByInvoiceNum= Nothing
	
	GetPrimarySalesmanByInvoiceNum = resultGetPrimarySalesmanByInvoiceNum
	
End Function

Function GetSalesmanNameBySalesmanNum(passedSalesmanNumber)

	Set cnnGetSalesmanNameBySalesmanNum = Server.CreateObject("ADODB.Connection")
	cnnGetSalesmanNameBySalesmanNum.open Session("ClientCnnString")

	resultGetSalesmanNameBySalesmanNum = ""
		
	SQLGetSalesmanNameBySalesmanNum = "Select * from " & MUV_Read("SQL_Owner") & ".Salesman where SalesmanSequence = " & passedSalesmanNumber
	 
	Set rsGetSalesmanNameBySalesmanNum = Server.CreateObject("ADODB.Recordset")
	rsGetSalesmanNameBySalesmanNum.CursorLocation = 3 
	Set rsGetSalesmanNameBySalesmanNum= cnnGetSalesmanNameBySalesmanNum.Execute(SQLGetSalesmanNameBySalesmanNum)
	
	
	If not rsGetSalesmanNameBySalesmanNum.eof then resultGetSalesmanNameBySalesmanNum = rsGetSalesmanNameBySalesmanNum("Name")
	
	set rsGetSalesmanNameBySalesmanNum= Nothing
	set cnnGetSalesmanNameBySalesmanNum= Nothing
	
	GetSalesmanNameBySalesmanNum = resultGetSalesmanNameBySalesmanNum
	
End Function


Function GetInvoiceDateByInvoiceNum(passedInvoiceNumber)

	Set cnnGetInvoiceDateByInvoiceNum = Server.CreateObject("ADODB.Connection")
	cnnGetInvoiceDateByInvoiceNum.open Session("ClientCnnString")

	resultGetInvoiceDateByInvoiceNum = ""
		
	SQLGetInvoiceDateByInvoiceNum = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory where IvsNum = " & passedInvoiceNumber
	 
	Set rsGetInvoiceDateByInvoiceNum = Server.CreateObject("ADODB.Recordset")
	rsGetInvoiceDateByInvoiceNum.CursorLocation = 3 
	Set rsGetInvoiceDateByInvoiceNum= cnnGetInvoiceDateByInvoiceNum.Execute(SQLGetInvoiceDateByInvoiceNum)
	
	
	If not rsGetInvoiceDateByInvoiceNum.eof then resultGetInvoiceDateByInvoiceNum = rsGetInvoiceDateByInvoiceNum("IvsDate")
	
	set rsGetInvoiceDateByInvoiceNum= Nothing
	set cnnGetInvoiceDateByInvoiceNum= Nothing
	
	GetInvoiceDateByInvoiceNum = resultGetInvoiceDateByInvoiceNum
	
End Function

Function GetSpecialCommentByCustNum(passedCustomerNumber)

	Set cnnGetSpecialCommentByCustNum = Server.CreateObject("ADODB.Connection")
	cnnGetSpecialCommentByCustNum.open Session("ClientCnnString")

	resultGetSpecialCommentByCustNum = ""
		
	SQLGetSpecialCommentByCustNum = "Select * from " & MUV_Read("SQL_Owner") & ".AR_Customer Where CustNum = " & passedCustomerNumber
	 
	Set rsGetSpecialCommentByCustNum = Server.CreateObject("ADODB.Recordset")
	rsGetSpecialCommentByCustNum.CursorLocation = 3 
	Set rsGetSpecialCommentByCustNum= cnnGetSpecialCommentByCustNum.Execute(SQLGetSpecialCommentByCustNum)
	
	
	If not rsGetSpecialCommentByCustNum.eof then resultGetSpecialCommentByCustNum = rsGetSpecialCommentByCustNum("SpecialComment")
	
	set rsGetSpecialCommentByCustNum= Nothing
	set cnnGetSpecialCommentByCustNum= Nothing
	
	GetSpecialCommentByCustNum = resultGetSpecialCommentByCustNum
	
End Function

Function GetInvoiceSubTotsByIvsNum(passedInvoiceNumber,passedSubtotToGet)

	passedSubtotToGet = Ucase(passedSubtotToGet)

	If passedSubtotToGet <> "MERCH" AND passedSubtotToGet <> "RECYCLE" AND passedSubtotToGet <> "TAX" AND passedSubtotToGet <> "DEPOSIT" AND passedSubtotToGet <> "GST" AND 	passedSubtotToGet <> "GRAND" THEN passedSubtotToGet = "GRAND"

	Set cnnGetInvoiceSubTotsByIvsNum = Server.CreateObject("ADODB.Connection")
	cnnGetInvoiceSubTotsByIvsNum.open Session("ClientCnnString")

	resultGetInvoiceSubTotsByIvsNum = 0

	If passedSubtotToGet  <> "RECYCLE" Then ' Recycle is special and entirely different
			
			SQLGetInvoiceSubTotsByIvsNum = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory Where IvsNum = " & passedInvoiceNumber
			 
			Set rsGetInvoiceSubTotsByIvsNum = Server.CreateObject("ADODB.Recordset")
			rsGetInvoiceSubTotsByIvsNum.CursorLocation = 3 
			Set rsGetInvoiceSubTotsByIvsNum= cnnGetInvoiceSubTotsByIvsNum.Execute(SQLGetInvoiceSubTotsByIvsNum)
				
			If not rsGetInvoiceSubTotsByIvsNum.eof then 
		
				Select Case passedSubtotToGet
					Case "MERCH"
						resultGetInvoiceSubTotsByIvsNum = rsGetInvoiceSubTotsByIvsNum("IvsTotalAmt") - ( rsGetInvoiceSubTotsByIvsNum("IvsSalesTax") + rsGetInvoiceSubTotsByIvsNum("IvsDepositChg") + rsGetInvoiceSubTotsByIvsNum("IvsGstTax"))
					Case "TAX"
						resultGetInvoiceSubTotsByIvsNum = rsGetInvoiceSubTotsByIvsNum("IvsSalesTax")
					Case "DEPOSIT"
						resultGetInvoiceSubTotsByIvsNum = rsGetInvoiceSubTotsByIvsNum("IvsDepositChg")			
					Case "GST"
						resultGetInvoiceSubTotsByIvsNum = rsGetInvoiceSubTotsByIvsNum("IvsGstTax")			
					Case "GRAND"
						resultGetInvoiceSubTotsByIvsNum = rsGetInvoiceSubTotsByIvsNum("IvsTotalAmt")			
				End Select	
		
			End If
			
	Else
		
			SQLGetInvoiceSubTotsByIvsNum = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail Where IvsNum = " & passedInvoiceNumber & " AND partNum = 'CYCLE'"
			 
			Set rsGetInvoiceSubTotsByIvsNum = Server.CreateObject("ADODB.Recordset")
			rsGetInvoiceSubTotsByIvsNum.CursorLocation = 3 
			Set rsGetInvoiceSubTotsByIvsNum= cnnGetInvoiceSubTotsByIvsNum.Execute(SQLGetInvoiceSubTotsByIvsNum)
				
			If not rsGetInvoiceSubTotsByIvsNum.eof then resultGetInvoiceSubTotsByIvsNum = rsGetInvoiceSubTotsByIvsNum("itemPrice")
		
	End If
			
	set rsGetInvoiceSubTotsByIvsNum= Nothing
	set cnnGetInvoiceSubTotsByIvsNum= Nothing
	
	GetInvoiceSubTotsByIvsNum = resultGetInvoiceSubTotsByIvsNum
	
End Function

Function GetTaxableFlagByIvsHistDetSequence(passedIvsHistDetSequence)

	Set cnnGetTaxableFlagByIvsHistDetSequence = Server.CreateObject("ADODB.Connection")
	cnnGetTaxableFlagByIvsHistDetSequence.open Session("ClientCnnString")

	resultGetTaxableFlagByIvsHistDetSequence = ""
		
	SQLGetTaxableFlagByIvsHistDetSequence = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail where IvsHistDetSequence = " & passedIvsHistDetSequence
	 
	Set rsGetTaxableFlagByIvsHistDetSequence = Server.CreateObject("ADODB.Recordset")
	rsGetTaxableFlagByIvsHistDetSequence.CursorLocation = 3 
	Set rsGetTaxableFlagByIvsHistDetSequence= cnnGetTaxableFlagByIvsHistDetSequence.Execute(SQLGetTaxableFlagByIvsHistDetSequence)
	
	
	If not rsGetTaxableFlagByIvsHistDetSequence.eof then resultGetTaxableFlagByIvsHistDetSequence = rsGetTaxableFlagByIvsHistDetSequence("prodTaxable")
	
	set rsGetTaxableFlagByIvsHistDetSequence= Nothing
	set cnnGetTaxableFlagByIvsHistDetSequence= Nothing
	
	GetTaxableFlagByIvsHistDetSequence = resultGetTaxableFlagByIvsHistDetSequence
	
End Function


Function GetInvoiceNumberByIvsSeq(passedIvsSeq)

	Set cnnGetInvoiceNumberByIvsSeq = Server.CreateObject("ADODB.Connection")
	cnnGetInvoiceNumberByIvsSeq.open Session("ClientCnnString")

	resultGetInvoiceNumberByIvsSeq = ""
		
	SQLGetInvoiceNumberByIvsSeq = "Select * from " & Session("SQL_Owner") & ".InvoiceHistory where IvsHistSequence = " & passedIvsSeq 
	 
	Set rsGetInvoiceNumberByIvsSeq = Server.CreateObject("ADODB.Recordset")
	rsGetInvoiceNumberByIvsSeq.CursorLocation = 3 
	Set rsGetInvoiceNumberByIvsSeq= cnnGetInvoiceNumberByIvsSeq.Execute(SQLGetInvoiceNumberByIvsSeq)
	
	
	If not rsGetInvoiceNumberByIvsSeq.eof then resultGetInvoiceNumberByIvsSeq = rsGetInvoiceNumberByIvsSeq("IvsNum")
	
	set rsGetInvoiceNumberByIvsSeq= Nothing
	set cnnGetInvoiceNumberByIvsSeq= Nothing
	
	GetInvoiceNumberByIvsSeq = resultGetInvoiceNumberByIvsSeq
	
End Function

Function GetCustNumberByInvSeq(passedIvsSeq)

	Set cnnGetCustNumberByInvSeq = Server.CreateObject("ADODB.Connection")
	cnnGetCustNumberByInvSeq.open Session("ClientCnnString")

	resultGetCustNumberByInvSeq = ""
		
	SQLGetCustNumberByInvSeq = "Select * from " & Session("SQL_Owner") & ".InvoiceHistory where IvsHistSequence = " & passedIvsSeq
	 
	Set rsGetCustNumberByInvSeq = Server.CreateObject("ADODB.Recordset")
	rsGetCustNumberByInvSeq.CursorLocation = 3 
	Set rsGetCustNumberByInvSeq= cnnGetCustNumberByInvSeq.Execute(SQLGetCustNumberByInvSeq)
	
	
	If not rsGetCustNumberByInvSeq.eof then resultGetCustNumberByInvSeq = rsGetCustNumberByInvSeq("CustNum")
	
	set rsGetCustNumberByInvSeq= Nothing
	set cnnGetCustNumberByInvSeq= Nothing
	
	GetCustNumberByInvSeq = resultGetCustNumberByInvSeq
	
End Function

Function GetPONumberByInvSeq(passedIvsSeq)

	Set cnnGetPONumberByInvSeq = Server.CreateObject("ADODB.Connection")
	cnnGetPONumberByInvSeq.open Session("ClientCnnString")

	resultGetPONumberByInvSeq = ""
		
	SQLGetPONumberByInvSeq = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory where IvsHistSequence = " & passedIvsSeq
	 
	Set rsGetPONumberByInvSeq = Server.CreateObject("ADODB.Recordset")
	rsGetPONumberByInvSeq.CursorLocation = 3 
	Set rsGetPONumberByInvSeq= cnnGetPONumberByInvSeq.Execute(SQLGetPONumberByInvSeq)
	
	
	If not rsGetPONumberByInvSeq.eof then resultGetPONumberByInvSeq = rsGetPONumberByInvSeq("PurchOrderNum")
	
	set rsGetPONumberByInvSeq= Nothing
	set cnnGetPONumberByInvSeq= Nothing
	
	GetPONumberByInvSeq = resultGetPONumberByInvSeq
	
End Function

Function GetRouteNumByInvSeq(passedIvsSeq)

	Set cnnGetRouteNumByInvSeq = Server.CreateObject("ADODB.Connection")
	cnnGetRouteNumByInvSeq.open Session("ClientCnnString")

	resultGetRouteNumByInvSeq = ""
		
	SQLGetRouteNumByInvSeq = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory where IvsHistSequence = " & passedIvsSeq
	 
	Set rsGetRouteNumByInvSeq = Server.CreateObject("ADODB.Recordset")
	rsGetRouteNumByInvSeq.CursorLocation = 3 
	Set rsGetRouteNumByInvSeq= cnnGetRouteNumByInvSeq.Execute(SQLGetRouteNumByInvSeq)
	
	
	If not rsGetRouteNumByInvSeq.eof then resultGetRouteNumByInvSeq = rsGetRouteNumByInvSeq("RouteNum")
	
	set rsGetRouteNumByInvSeq= Nothing
	set cnnGetRouteNumByInvSeq= Nothing
	
	GetRouteNumByInvSeq = resultGetRouteNumByInvSeq
	
End Function

Function GetTermsNumByInvSeq(passedIvsSeq)

	Set cnnGetTermsNumByInvSeq = Server.CreateObject("ADODB.Connection")
	cnnGetTermsNumByInvSeq.open Session("ClientCnnString")

	resultGetTermsNumByInvSeq = ""
		
	SQLGetTermsNumByInvSeq = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory where IvsHistSequence = " & passedIvsSeq
	 
	Set rsGetTermsNumByInvSeq = Server.CreateObject("ADODB.Recordset")
	rsGetTermsNumByInvSeq.CursorLocation = 3 
	Set rsGetTermsNumByInvSeq= cnnGetTermsNumByInvSeq.Execute(SQLGetTermsNumByInvSeq)
	
	
	If not rsGetTermsNumByInvSeq.eof then resultGetTermsNumByInvSeq = rsGetTermsNumByInvSeq("Terms")
	
	set rsGetTermsNumByInvSeq= Nothing
	set cnnGetTermsNumByInvSeq= Nothing
	
	GetTermsNumByInvSeq = resultGetTermsNumByInvSeq
	
End Function

Function GetPrimarySalesmanByInvSeq(passedIvsSeq)

	Set cnnGetPrimarySalesmanByInvSeq = Server.CreateObject("ADODB.Connection")
	cnnGetPrimarySalesmanByInvSeq.open Session("ClientCnnString")

	resultGetPrimarySalesmanByInvSeq = ""
		
	SQLGetPrimarySalesmanByInvSeq = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory where IvsHistSequence = " & passedIvsSeq
	 
	Set rsGetPrimarySalesmanByInvSeq = Server.CreateObject("ADODB.Recordset")
	rsGetPrimarySalesmanByInvSeq.CursorLocation = 3 
	Set rsGetPrimarySalesmanByInvSeq= cnnGetPrimarySalesmanByInvSeq.Execute(SQLGetPrimarySalesmanByInvSeq)
	
	
	If not rsGetPrimarySalesmanByInvSeq.eof then resultGetPrimarySalesmanByInvSeq = rsGetPrimarySalesmanByInvSeq("PrimarySalesman")
	
	set rsGetPrimarySalesmanByInvSeq= Nothing
	set cnnGetPrimarySalesmanByInvSeq= Nothing
	
	GetPrimarySalesmanByInvSeq = resultGetPrimarySalesmanByInvSeq
	
End Function

Function GetInvoiceDateByInvSeq(passedIvsSeq)

	Set cnnGetInvoiceDateByInvSeq = Server.CreateObject("ADODB.Connection")
	cnnGetInvoiceDateByInvSeq.open Session("ClientCnnString")

	resultGetInvoiceDateByInvSeq = ""
		
	SQLGetInvoiceDateByInvSeq = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory where IvsHistSequence = " & passedIvsSeq
	 
	Set rsGetInvoiceDateByInvSeq = Server.CreateObject("ADODB.Recordset")
	rsGetInvoiceDateByInvSeq.CursorLocation = 3 
	Set rsGetInvoiceDateByInvSeq= cnnGetInvoiceDateByInvSeq.Execute(SQLGetInvoiceDateByInvSeq)
	
	
	If not rsGetInvoiceDateByInvSeq.eof then resultGetInvoiceDateByInvSeq = rsGetInvoiceDateByInvSeq("IvsDate")
	
	set rsGetInvoiceDateByInvSeq= Nothing
	set cnnGetInvoiceDateByInvSeq= Nothing
	
	GetInvoiceDateByInvSeq = resultGetInvoiceDateByInvSeq
	
End Function

Function GetInvoiceSubTotsByInvSeq(passedIvsSeq,passedSubtotToGet)

	passedSubtotToGet = Ucase(passedSubtotToGet)

	If passedSubtotToGet <> "MERCH" AND passedSubtotToGet <> "RECYCLE" AND passedSubtotToGet <> "TAX" AND passedSubtotToGet <> "DEPOSIT" AND passedSubtotToGet <> "GST" AND 	passedSubtotToGet <> "GRAND" THEN passedSubtotToGet = "GRAND"

	Set cnnGetInvoiceSubTotsByInvSeq = Server.CreateObject("ADODB.Connection")
	cnnGetInvoiceSubTotsByInvSeq.open Session("ClientCnnString")

	resultGetInvoiceSubTotsByInvSeq = 0

	If passedSubtotToGet  <> "RECYCLE" Then ' Recycle is special and entirely different
			
			SQLGetInvoiceSubTotsByInvSeq = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistory Where IvsHistSequence = " & passedIvsSeq
			 
			Set rsGetInvoiceSubTotsByInvSeq = Server.CreateObject("ADODB.Recordset")
			rsGetInvoiceSubTotsByInvSeq.CursorLocation = 3 
			Set rsGetInvoiceSubTotsByInvSeq= cnnGetInvoiceSubTotsByInvSeq.Execute(SQLGetInvoiceSubTotsByInvSeq)
				
			If not rsGetInvoiceSubTotsByInvSeq.eof then 
		
				Select Case passedSubtotToGet
					Case "MERCH"
						resultGetInvoiceSubTotsByInvSeq = rsGetInvoiceSubTotsByInvSeq("IvsTotalAmt") - ( rsGetInvoiceSubTotsByInvSeq("IvsSalesTax") + rsGetInvoiceSubTotsByInvSeq("IvsDepositChg") + rsGetInvoiceSubTotsByInvSeq("IvsGstTax"))
						
						'If Merch sustoal, must also subtract recycle fe
						SQLGetInvoiceSubTotsByInvSeq = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail Where IvsHistSequence = " & passedIvsSeq & " AND partNum = 'CYCLE'"
			 			Set rsGetInvoiceSubTotsByInvSeq = Server.CreateObject("ADODB.Recordset")
						rsGetInvoiceSubTotsByInvSeq.CursorLocation = 3 
						Set rsGetInvoiceSubTotsByInvSeq= cnnGetInvoiceSubTotsByInvSeq.Execute(SQLGetInvoiceSubTotsByInvSeq)
										
						If not rsGetInvoiceSubTotsByInvSeq.eof then
							resultGetInvoiceSubTotsByInvSeq = resultGetInvoiceSubTotsByInvSeq - rsGetInvoiceSubTotsByInvSeq("itemPrice")
						End If
					Case "TAX"
						resultGetInvoiceSubTotsByInvSeq = rsGetInvoiceSubTotsByInvSeq("IvsSalesTax")
					Case "DEPOSIT"
						resultGetInvoiceSubTotsByInvSeq = rsGetInvoiceSubTotsByInvSeq("IvsDepositChg")			
					Case "GST"
						resultGetInvoiceSubTotsByInvSeq = rsGetInvoiceSubTotsByInvSeq("IvsGstTax")			
					Case "GRAND"
						resultGetInvoiceSubTotsByInvSeq = rsGetInvoiceSubTotsByInvSeq("IvsTotalAmt")			
				End Select	
		
			End If
			
	Else
		
			SQLGetInvoiceSubTotsByInvSeq = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail Where IvsHistSequence = " & passedIvsSeq & " AND partNum = 'CYCLE'"
			 
			Set rsGetInvoiceSubTotsByInvSeq = Server.CreateObject("ADODB.Recordset")
			rsGetInvoiceSubTotsByInvSeq.CursorLocation = 3 
			Set rsGetInvoiceSubTotsByInvSeq= cnnGetInvoiceSubTotsByInvSeq.Execute(SQLGetInvoiceSubTotsByInvSeq)
				
			If not rsGetInvoiceSubTotsByInvSeq.eof then resultGetInvoiceSubTotsByInvSeq = rsGetInvoiceSubTotsByInvSeq("itemPrice")
		
	End If
			
	set rsGetInvoiceSubTotsByInvSeq= Nothing
	set cnnGetInvoiceSubTotsByInvSeq= Nothing
	
	GetInvoiceSubTotsByInvSeq = resultGetInvoiceSubTotsByInvSeq
	
End Function



Function GetAddressElementByCustNum(passedCustNum,passedElement)

	passedElement = Ucase(passedElement)
	resultGetAddressElementByCustNum = ""

	Set cnnGetAddressElementByCustNum = Server.CreateObject("ADODB.Connection")
	cnnGetAddressElementByCustNum.open Session("ClientCnnString")

	SQLGetAddressElementByCustNum = "Select * from " & MUV_Read("SQL_Owner") & ".AR_Customer Where CustNum = '" & passedCustNum & "'"
			 
	Set rsGetAddressElementByCustNum = Server.CreateObject("ADODB.Recordset")
	rsGetAddressElementByCustNum.CursorLocation = 3 
	Set rsGetAddressElementByCustNum= cnnGetAddressElementByCustNum.Execute(SQLGetAddressElementByCustNum)
				
	If not rsGetAddressElementByCustNum.eof then 
		
		Select Case passedElement
				Case "ADDRESS1"
					resultGetAddressElementByCustNum = rsGetAddressElementByCustNum("Addr1")
				Case "ADDRESS2"
					resultGetAddressElementByCustNum = rsGetAddressElementByCustNum("Addr2")
				Case "CSZ"
					resultGetAddressElementByCustNum = rsGetAddressElementByCustNum("CityStateZip")			
				Case "PHONE"
					resultGetAddressElementByCustNum = rsGetAddressElementByCustNum("Phone")			
				Case "CONTACT"
					resultGetAddressElementByCustNum = rsGetAddressElementByCustNum("Contact")			
			End Select	
	End If
			
			
	set rsGetAddressElementByCustNum= Nothing
	set cnnGetAddressElementByCustNum= Nothing
	
	GetAddressElementByCustNum = resultGetAddressElementByCustNum
	
End Function

Function InvoiceProfitDollars(passedInvoiceNumber)
	
	resultInvoiceProfitDollars = 0 
	
	Set cnnInvoiceProfitDollars = Server.CreateObject("ADODB.Connection")
	cnnInvoiceProfitDollars.open Session("ClientCnnString")
		
	SQLInvoiceProfitDollars = "Select * from " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail where IvsNum = " & passedInvoiceNumber
	 
	Set rsInvoiceProfitDollars = Server.CreateObject("ADODB.Recordset")
	rsInvoiceProfitDollars.CursorLocation = 3 
	Set rsInvoiceProfitDollars= cnnInvoiceProfitDollars.Execute(SQLInvoiceProfitDollars)
		
			
	If not rsInvoiceProfitDollars.eof then
	
		Do
		
			resultInvoiceProfitDollars = resultInvoiceProfitDollars + ( (rsInvoiceProfitDollars("itemPrice")-rsInvoiceProfitDollars("itemCost")) * rsInvoiceProfitDollars("itemQuantity") )
		
			rsInvoiceProfitDollars.MoveNext
			
		Loop while not rsInvoiceProfitDollars.eof
		
	End If
	
	set rsInvoiceProfitDollars= Nothing
	set cnnInvoiceProfitDollars= Nothing
	
	InvoiceProfitDollars = resultInvoiceProfitDollars 
	
End Function

Function GetNumberOfLinesByInvoiceNumber(passedInvoiceNumber)
	
	resultGetNumberOfLinesByInvoiceNumber = 0 
	
	Set cnnGetNumberOfLinesByInvoiceNumber = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfLinesByInvoiceNumber.open Session("ClientCnnString")
		
	SQLGetNumberOfLinesByInvoiceNumber = "Select Count (*) as Expr1 from " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail where IvsNum = " & passedInvoiceNumber
	 
	Set rsGetNumberOfLinesByInvoiceNumber = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfLinesByInvoiceNumber.CursorLocation = 3 
	Set rsGetNumberOfLinesByInvoiceNumber= cnnGetNumberOfLinesByInvoiceNumber.Execute(SQLGetNumberOfLinesByInvoiceNumber)
		
			
	If not rsGetNumberOfLinesByInvoiceNumber.eof then resultGetNumberOfLinesByInvoiceNumber = rsGetNumberOfLinesByInvoiceNumber("Expr1")
	
	set rsGetNumberOfLinesByInvoiceNumber= Nothing
	set cnnGetNumberOfLinesByInvoiceNumber= Nothing
	
	GetNumberOfLinesByInvoiceNumber = resultGetNumberOfLinesByInvoiceNumber 
	
End Function


Function GetAlertType (passedAlertNumber)

	resultGetAlertType = ""

	Set cnn_GetAlertType = Server.CreateObject("ADODB.Connection")
	cnn_GetAlertType.open (Session("ClientCnnString"))
	Set rsGetAlertType = Server.CreateObject("ADODB.Recordset")
	rsGetAlertType.CursorLocation = 3 
	
	SQL_GetAlertType = "SELECT * FROM SC_Alerts WHERE InternalAlertRecNumber = " & passedAlertNumber
	Set rsGetAlertType = cnn_GetAlertType.Execute(SQL_GetAlertType)

	If not rsGetAlertType.EOF Then
		resultGetAlertType = rsGetAlertType("AlertType")
	End If

	Set rsGetAlertType = Nothing
	cnn_GetAlertType.Close
	Set cnn_GetAlertType = Nothing

	GetAlertType =	resultGetAlertType 
	
End Function


Function TZNow()

	'Returns Now() adjusted for the company time zone settings
	
	resultTZNow = Now()

	Set cnnTZNow = Server.CreateObject("ADODB.Connection")
	cnnTZNow.open Session("ClientCnnString")
		
	SQLTZNow = "Select * from Settings_CompanyID"
 
	Set rsTZNow = Server.CreateObject("ADODB.Recordset")
	rsTZNow.CursorLocation = 3 
	Set rsTZNow = cnnTZNow.Execute(SQLTZNow)
			 
	If not rsTZNow.EOF Then 
		Select Case rsTZNow("Timezone")
			Case "Eastern"
				resultTZNow = Now() 
			Case "Central"
				resultTZNow = dateadd("h",1,Now())
			Case "Mountain"
				resultTZNow = dateadd("h",2,Now())
			Case "Pacific"
				resultTZNow = dateadd("h",3,Now())
		End Select
	End If
	
	rsTZNow.Close
	set rsTZNow= Nothing
	cnnTZNow.Close	
	set cnnTZNow= Nothing
	
	TZNow = resultTZNow
	
End Function

Function TZTime()

	'Returns Time adjusted for the company time zone settings
	
	resultTZTime = Time

	Set cnnTZTime = Server.CreateObject("ADODB.Connection")
	cnnTZTime.open Session("ClientCnnString")
		
	SQLTZTime = "Select * from Settings_CompanyID"
 
	Set rsTZTime = Server.CreateObject("ADODB.Recordset")
	rsTZTime.CursorLocation = 3 
	Set rsTZTime = cnnTZTime.Execute(SQLTZTime)
			 
	If not rsTZTime.EOF Then 
		Select Case rsTZTime("Timezone")
			Case "Eastern"
				resultTZTime = Time 
			Case "Central"
				resultTZTime = dateadd("h",1,Time)
			Case "Mountain"
				resultTZTime = dateadd("h",2,Time)
			Case "Pacific"
				resultTZTime = dateadd("h",3,Time)
		End Select
	End If
	
	rsTZTime.Close
	set rsTZTime= Nothing
	cnnTZTime.Close	
	set cnnTZTime= Nothing
	
	TZTime = resultTZTime
	
End Function

Function BusinessDayEnd()

	resultBusinessDayEnd = ""

	Set cnnBusinessDayEnd = Server.CreateObject("ADODB.Connection")
	cnnBusinessDayEnd.open Session("ClientCnnString")
		
	SQLBusinessDayEnd = "Select * from Settings_CompanyID"
 
	Set rsBusinessDayEnd = Server.CreateObject("ADODB.Recordset")
	rsBusinessDayEnd.CursorLocation = 3 
	Set rsBusinessDayEnd = cnnBusinessDayEnd.Execute(SQLBusinessDayEnd)
			 
	If not rsBusinessDayEnd.EOF Then resultBusinessDayEnd = rsBusinessDayEnd("BusinessDayEnd")

	rsBusinessDayEnd.Close
	set rsBusinessDayEnd= Nothing
	cnnBusinessDayEnd.Close	
	set cnnBusinessDayEnd= Nothing
	
	BusinessDayEnd = resultBusinessDayEnd
	
End Function

Function BusinessDayStart()

	resultBusinessDayStart = ""

	Set cnnBusinessDayStart = Server.CreateObject("ADODB.Connection")
	cnnBusinessDayStart.open Session("ClientCnnString")
		
	SQLBusinessDayStart = "Select * from Settings_CompanyID"
 
	Set rsBusinessDayStart = Server.CreateObject("ADODB.Recordset")
	rsBusinessDayStart.CursorLocation = 3 
	Set rsBusinessDayStart = cnnBusinessDayStart.Execute(SQLBusinessDayStart)
			 
	If not rsBusinessDayStart.EOF Then resultBusinessDayStart = rsBusinessDayStart("BusinessDayStart")

	rsBusinessDayStart.Close
	set rsBusinessDayStart= Nothing
	cnnBusinessDayStart.Close	
	set cnnBusinessDayStart= Nothing
	
	BusinessDayStart = resultBusinessDayStart
	
End Function


Function GetServiceTicketCompanyName(passedTicketNumber)

	result = ""
	
	Set cnnCust = Server.CreateObject("ADODB.Connection")
	cnnCust.open Session("ClientCnnString")

	SQLCust = "SELECT Company FROM FS_ServiceMemos WHERE MemoNumber = '" & passedTicketNumber & "' order by submissionDateTime desc"

	Set rsCust = Server.CreateObject("ADODB.Recordset")
	rsCust.CursorLocation = 3 
	Set rsCust = cnnCust.Execute(SQLCust )
	
	If not rsCust.eof then result = rsCust("Company")

	set rsCust = Nothing
	set cnnCust= Nothing
	
	GetServiceTicketCompanyName = result

End Function


Function GetExtension(FileName)
  Dim DotPos
  DotPos = InstrRev(FileName, "." )
  If DotPos < Len(FileName) Then
    GetExtension = Mid(FileName, DotPos + 1)
  Else
    GetExtension = ""
  End If
End Function

Function GetLastInvoiceFromWebDate(passedCustID)

	resultGetLastInvoiceFromWebDate = ""

	Set cnnGetLastInvoiceFromWebDate = Server.CreateObject("ADODB.Connection")
	cnnGetLastInvoiceFromWebDate.open Session("ClientCnnString")
		
	SQLGetLastInvoiceFromWebDate = "SELECT IvsDate FROM InvoiceHistory WHERE CustNum = " & passedCustID & " AND LoginName = 'websel'"

	Set rsGetLastInvoiceFromWebDate = Server.CreateObject("ADODB.Recordset")
	rsGetLastInvoiceFromWebDate.CursorLocation = 3 
	Set rsGetLastInvoiceFromWebDate = cnnGetLastInvoiceFromWebDate.Execute(SQLGetLastInvoiceFromWebDate)
			 
	If not rsGetLastInvoiceFromWebDate.EOF Then resultGetLastInvoiceFromWebDate = rsGetLastInvoiceFromWebDate("IvsDate")

	rsGetLastInvoiceFromWebDate.Close
	set rsGetLastInvoiceFromWebDate= Nothing
	cnnGetLastInvoiceFromWebDate.Close	
	set cnnGetLastInvoiceFromWebDate= Nothing
	
	GetLastInvoiceFromWebDate = resultGetLastInvoiceFromWebDate
	
End Function


Function NumberOfWorkDays(passedStartDate, passedEndDate)

	DateFrom = passedStartDate
	DateTo = passedEndDate
	Weekends = 0
	ActualDays = 0
	
	' Step 1: Get the actual days
	ActualDays = DateDiff("d",DateFrom,DateTo)
	
	' Step 2: Find the weekends
	For x = 0 to ActualDays - 1
	    xDate = dateadd("d",x,DateFrom)
	    If weekday(xDate,1)=1 OR weekday(xDate,1)=7 Then
	         Weekends = Weekends + 1
	    End If
	next
	
	' Step 3: Find the number of closed days in the company calendar
	
	closedCompanyCalendarDays = 0

	Set cnnGetClosedDaysInRange = Server.CreateObject("ADODB.Connection")
	cnnGetClosedDaysInRange.open Session("ClientCnnString")
	
	yearPassedStartDate = cInt(Year(passedStartDate))
	yearPassedEndDate = cInt(Year(passedEndDate))
	
	SQLGetClosedDaysInRange = "SELECT * FROM Settings_CompanyCalendar WHERE YearNum >= " & yearPassedStartDate & " AND YearNum <= " & yearPassedEndDate 

	Set rsGetClosedDaysInRange = Server.CreateObject("ADODB.Recordset")
	rsGetClosedDaysInRange.CursorLocation = 3 
	Set rsGetClosedDaysInRange = cnnGetClosedDaysInRange.Execute(SQLGetClosedDaysInRange)
			 
	If NOT rsGetClosedDaysInRange.EOF Then 
					 
		Do While NOT rsGetClosedDaysInRange.EOF
		
			curMonthToCheck = rsGetClosedDaysInRange("MonthNum")
			curDayToCheck = rsGetClosedDaysInRange("DayNum")
			curYearToCheck = rsGetClosedDaysInRange("YearNum")
			
			curDateToCheck = cDate(curMonthToCheck & "/" & curDayToCheck & "/" & curYearToCheck)
'Response.Write("curDateToCheck :" & curDateToCheck & "<br>")
'Response.Write("passedStartDate :" & passedStartDate & "<br>")
'Response.Write("passedEndDate :" & passedEndDate & "<br>")			
			If curDateToCheck >= cDate(passedStartDate) AND curDateToCheck <= cDate(passedEndDate) Then
			
				If rsGetClosedDaysInRange("OpenClosedCloseEarly") = "Closed" Then closedCompanyCalendarDays = closedCompanyCalendarDays + 1
				If rsGetClosedDaysInRange("OpenClosedCloseEarly") = "Close Early" Then closedCompanyCalendarDays = closedCompanyCalendarDays + .5
				
			End If
			
		rsGetClosedDaysInRange.MoveNext
		Loop
				
	End If

	rsGetClosedDaysInRange.Close
	set rsGetClosedDaysInRange= Nothing
	cnnGetClosedDaysInRange.Close	
	set cnnGetClosedDaysInRange= Nothing
'Response.Write("closedCompanyCalendarDays :" & closedCompanyCalendarDays  & "<br>")

	NumberOfWorkDays = ActualDays - Weekends - closedCompanyCalendarDays

End Function

Function GetServiceTicketLastEntryDateTime(passedTicketNumber)

	'Use only when advanced dispatch module is on

	resultGetServiceTicketLastEntryDateTime = ""
	
	Set cnnGetServiceTicketLastEntryDateTime = Server.CreateObject("ADODB.Connection")
	cnnGetServiceTicketLastEntryDateTime.open Session("ClientCnnString")

	SQLGetServiceTicketLastEntryDateTime = "Select TOP 1 * from FS_ServiceMemosDetail where MemoNumber = '" & passedTicketNumber & "' Order By RecordCreatedDateTime Desc"

	Set rsGetServiceTicketLastEntryDateTime = Server.CreateObject("ADODB.Recordset")
	rsGetServiceTicketLastEntryDateTime.CursorLocation = 3 
	Set rsGetServiceTicketLastEntryDateTime = cnnGetServiceTicketLastEntryDateTime.Execute(SQLGetServiceTicketLastEntryDateTime)
	
	If not rsGetServiceTicketLastEntryDateTime.eof then 
		resultGetServiceTicketLastEntryDateTime = rsGetServiceTicketLastEntryDateTime("RecordCreatedDateTime")
	Else
		' No detail so need to get it from the header
		SQLGetServiceTicketLastEntryDateTime = "Select TOP 1 * from FS_ServiceMemos where MemoNumber = '" & passedTicketNumber & "' Order By RecordCreatedateTime Desc"
		Set rsGetServiceTicketLastEntryDateTime = cnnGetServiceTicketLastEntryDateTime.Execute(SQLGetServiceTicketLastEntryDateTime)
		If not rsGetServiceTicketLastEntryDateTime.eof then 
			resultGetServiceTicketLastEntryDateTime = rsGetServiceTicketLastEntryDateTime("RecordCreatedateTime")
		End If
	End IF	
	
	set rsGetServiceTicketLastEntryDateTime = Nothing
	cnnGetServiceTicketLastEntryDateTime.Close
	set cnnGetServiceTicketLastEntryDateTime = Nothing
	
	GetServiceTicketLastEntryDateTime = resultGetServiceTicketLastEntryDateTime
	

End Function

Function GetUserNoBySalesPersonNo(passedSalesPersonNo)

	resultGetUserNoBySalesPersonNo = ""
	
	Set cnnGetUserNoBySalesPersonNo = Server.CreateObject("ADODB.Connection")
	cnnGetUserNoBySalesPersonNo.open Session("ClientCnnString")

	SQLGetUserNoBySalesPersonNo = "SELECT UserNo FROM tblUsers WHERE userSalesPersonNumber = " & passedSalesPersonNo

	Set rsGetUserNoBySalesPersonNo = Server.CreateObject("ADODB.Recordset")
	rsGetUserNoBySalesPersonNo.CursorLocation = 3 
	Set rsGetUserNoBySalesPersonNo = cnnGetUserNoBySalesPersonNo.Execute(SQLGetUserNoBySalesPersonNo)
	
	If not rsGetUserNoBySalesPersonNo.eof then resultGetUserNoBySalesPersonNo = rsGetUserNoBySalesPersonNo("UserNo")

	set rsGetUserNoBySalesPersonNo = Nothing
	cnnGetUserNoBySalesPersonNo.Close
	set cnnGetUserNoBySalesPersonNo = Nothing
	
	GetUserNoBySalesPersonNo = resultGetUserNoBySalesPersonNo

End Function

Function GetSalesPersonNoByUserNo(passedUserNo)

	resultGetSalesPersonNoByUserNo = ""
	
	Set cnnGetSalesPersonNoByUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetSalesPersonNoByUserNo.open Session("ClientCnnString")

	SQLGetSalesPersonNoByUserNo = "SELECT userSalesPersonNumber FROM tblUsers WHERE UserNo = " & passedUserNo

	Set rsGetSalesPersonNoByUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetSalesPersonNoByUserNo.CursorLocation = 3 
	Set rsGetSalesPersonNoByUserNo = cnnGetSalesPersonNoByUserNo.Execute(SQLGetSalesPersonNoByUserNo)
	
	If not rsGetSalesPersonNoByUserNo.eof then resultGetSalesPersonNoByUserNo = rsGetSalesPersonNoByUserNo("userSalesPersonNumber")

	set rsGetSalesPersonNoByUserNo = Nothing
	cnnGetSalesPersonNoByUserNo.Close
	set cnnGetSalesPersonNoByUserNo = Nothing
	
	GetSalesPersonNoByUserNo = resultGetSalesPersonNoByUserNo

End Function

Function NAGMasterON()

	resultNAGMasterON = ""
	
	Set cnnNAGMasterON = Server.CreateObject("ADODB.Connection")
	cnnNAGMasterON.open Session("ClientCnnString")

	SQLNAGMasterON = "SELECT MasterNagMessageONOFF FROM Settings_Global"

	Set rsNAGMasterON = Server.CreateObject("ADODB.Recordset")
	rsNAGMasterON.CursorLocation = 3 
	Set rsNAGMasterON = cnnNAGMasterON.Execute(SQLNAGMasterON)
	
	If not rsNAGMasterON.EOF Then
		If rsNAGMasterON("MasterNagMessageONOFF") = 0 Then
			resultNAGMasterON = False
	    ElseIf rsNAGMasterON("MasterNagMessageONOFF") = 1 Then
	    	resultNAGMasterON = True
	    Else
	    	resultNAGMasterON = False
	    End If
    Else
	    resultNAGMasterON = False
    End If

	set rsNAGMasterON = Nothing
	cnnNAGMasterON.Close
	set cnnNAGMasterON = Nothing
	
	NAGMasterON = resultNAGMasterON

End Function

Function GetPassPhrase(passedClientKey)

	resultGetPassPhrase = ""
		
	SQLGetPassPhrase = "SELECT * FROM tblServerInfo where clientKey='"& passedClientKey &"'"
	Set cnnGetPassPhrase = Server.CreateObject("ADODB.Connection")
	Set rsGetPassPhrase = Server.CreateObject("ADODB.Recordset")
	cnnGetPassPhrase.Open "Driver={SQL Server};Server=66.201.99.15;Database=_BIInsight;Uid=biinsight;Pwd=Z32#kje4217;"

	rsGetPassPhrase.Open SQLGetPassPhrase,cnnGetPassPhrase,3,3

	If NOT rsGetPassPhrase.EOF Then resultGetPassPhrase = rsGetPassPhrase("directLaunchPassphrase")
		
	rsGetPassPhrase.close
	cnnGetPassPhrase.close
	set rsGetPassPhrase = Nothing
	set cnnGetPassPhrase = Nothing
	
	GetPassPhrase = resultGetPassPhrase 
	
End Function



function fmt_mmddyy(passedinput)
	
	resultfmt_resultfmt = ""
		
	passedinput = cDate(passedinput)
		
	dim m: m = month(passedinput)
    dim d: d = day(passedinput)
    if (m < 10) then m = "0" & m
    if (d < 10) then d = "0" & d

    resultfmt_mmddyy = m & "/" & d & "/" & right(year(passedinput), 2)
	
    fmt_mmddyy = resultfmt_mmddyy
    
end function

Function GetPeriodOneThisFiscalYearSeqNum()

		resultGetPeriodOneThisFiscalYearSeqNum = ""
		
		YearPart = GetCurrentPeriodAndYear()
		YearPart = Right(YearPart,Len(YearPart) - Instr(YearPart,"-")-1)
		YearPart = Trim(YearPart)

		Set cnnGetPeriodOneThisFiscalYearSeqNum = Server.CreateObject("ADODB.Connection")
		cnnGetPeriodOneThisFiscalYearSeqNum.open Session("ClientCnnString")
	
		SQLGetPeriodOneThisFiscalYearSeqNum = "SELECT BillPerSequence FROM BillingPeriodHistory WHERE Period = 1 AND Year = " & YearPart 
	
		Set rsGetPeriodOneThisFiscalYearSeqNum = Server.CreateObject("ADODB.Recordset")
		rsGetPeriodOneThisFiscalYearSeqNum.CursorLocation = 3 
		Set rsGetPeriodOneThisFiscalYearSeqNum = cnnGetPeriodOneThisFiscalYearSeqNum.Execute(SQLGetPeriodOneThisFiscalYearSeqNum)

		If Not rsGetPeriodOneThisFiscalYearSeqNum.Eof Then resultGetPeriodOneThisFiscalYearSeqNum =  rsGetPeriodOneThisFiscalYearSeqNum("BillPerSequence") 

		Set rsGetPeriodOneThisFiscalYearSeqNum = Nothing
		cnnGetPeriodOneThisFiscalYearSeqNum.Close
		SET cnnGetPeriodOneThisFiscalYearSeqNum = Nothing
				
		GetPeriodOneThisFiscalYearSeqNum = resultGetPeriodOneThisFiscalYearSeqNum 

End Function


Function Old_NumberofWorkMinutes_DateOpened(passedDate,passedTime,passedNormalBizDayStartTime,passedNormalBizDayEndTime)

debugmsg=1

	If passedTime = "" OR passedTime <= passedNormalBizDayEndTime Then ' Otherwise it is after or before hours so return 0
		If debugmsg=1 then response.write("CCCCCCCCCCCCCCCCCCCCCCCCCCpassedTime" & passedTime & "<br>")
		If debugmsg=1 then response.write("CCCCCCCCCCCCCCCCCCCCCCCCCCEndTime:" & passedNormalBizDayStartTime&"<br>")

			SQL = "SELECT * FROM Settings_CompanyCalendar where Monthnum='" & Month(passedDate) & "' AND DayNum ='" & Day(passedDate) & "' AND YearNum='" & Year(passedDate) & "'"
		
			Set cnn9 = Server.CreateObject("ADODB.Connection")
			cnn9.open (Session("ClientCnnString"))
			Set rs9 = Server.CreateObject("ADODB.Recordset")
			rs9.CursorLocation = 3 
			Set rs9 = cnn9.Execute(SQL)
				
			If not rs9.EOF Then
				Select Case rs9("OpenClosedCloseEarly")
					Case "Closed"
						'Do nothing, they are closed
					Case "Close Early"
						ClosingTime = cdate(rs9("ClosingTime"))
						If passedTime = "" Then
							NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,ClosingTime)
						Else
							NumberofWorkMinutesDate_result = DateDiff("n",passedTime,ClosingTime)
						End IF
					End Select
			Else ' EOF MEANS THEY ARE OPEN
				If passedTime = "" Then
					NumberofWorkMinutesDate_result = DateDiff("n",passedNormalBizDayStartTime,passedNormalBizDayEndTime)
				Else
					NumberofWorkMinutesDate_result = DateDiff("n",passedTime,passedNormalBizDayEndTime)
					If debugmsg=1 then Response.Write( passedMemoNumber & ":passedTime:" & passedTime& "<br>")
					If debugmsg=1 then Response.Write( passedMemoNumber & ":EndTime:" & passedNormalBizDayEndTime& "<br>")
				End IF	
		 	End If
		
			set rs9 = Nothing
			cnn9.close
			set cnn9 = Nothing
			
	Else ' After hours
			If debugmsg=1 then response.write("CCCCCCCCCCCCCCCCCCCCCCCCCC<br>")
			NumberofWorkMinutesDate_result = 0
	End If
		
	Old_NumberofWorkMinutes_DateOpened= NumberofWorkMinutesDate_result
	
End Function

Sub Write_API_AuditLog_Entry(passedIdentity,passedLogEntry,passedMode,passedModule)

	'on error resume next
	'Creates an entry in API_AuditLog
	
	
	Set cnnAudit = Server.CreateObject("ADODB.Connection")
	cnnAudit.open (Session("ClientCnnString"))

	Set rsAudit = Server.CreateObject("ADODB.Recordset")
	rsAudit.CursorLocation = 3 
	
	Set rsAudit = cnnAudit.Execute("Select TOP 1 * from API_AuditLog order by EntryThread desc")
	If Not rsAudit.EOF Then
		If IsNull(rsAudit("EntryThread")) Then EntryThread =1 Else EntryThread = rsAudit("EntryThread") + 1
	Else
		EntryThread = 1
	End If

	
	
	passedLogEntry= replace(passedLogEntry,"'","")
	
	
	UserIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If UserIPAddress = "" Then
		UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
	End If
	
	'************************
	'Log to client's database
	'************************
	SQL = "INSERT INTO API_AuditLog([Identity],LogEntry,Mode,SerNo,apiModule,IPAddress,EntryThread)"
	SQL = SQL &  " VALUES ('" & passedIdentity & "'"
	SQL = SQL & ",'"  & passedLogEntry & "'"
	SQL = SQL & ",'"  & passedMode & "'"
	SQL = SQL & ",'"  & MUV_READ("SERNO") & "'"
	SQL = SQL & ",'"  & passedModule & "'"
	SQL = SQL & ",'"  & UserIPAddress & "'"
	SQL = SQL & ","  & EntryThread & ")"


	Set rsAudit = cnnAudit.Execute(SQL)
	set rsAudit = Nothing
	cnnAudit.Close
	Set cnnAudit = Nothing



	On error goto 0

	
End Sub 



Function padDate(n, totalDigits) 
    if totalDigits > len(n) then 
        padDate = String(totalDigits-len(n),"0") & n 
    else 
        padDate = n 
    end if 
End Function 

Function GetCompanyCountry()

	resultGetCompanyCountry = "United States"
	
	Set cnnGetCompanyCountry = Server.CreateObject("ADODB.Connection")
	cnnGetCompanyCountry.open Session("ClientCnnString")

	SQLGetCompanyCountry = "SELECT STMT_Country FROM Settings_CompanyID"

	Set rsGetCompanyCountry = Server.CreateObject("ADODB.Recordset")
	rsGetCompanyCountry.CursorLocation = 3 
	Set rsGetCompanyCountry = cnnGetCompanyCountry.Execute(SQLGetCompanyCountry)
	
	If not rsGetCompanyCountry.eof then resultGetCompanyCountry = rsGetCompanyCountry("STMT_Country")

	set rsGetCompanyCountry = Nothing
	cnnGetCompanyCountry.Close
	set cnnGetCompanyCountry = Nothing
	
	GetCompanyCountry = resultGetCompanyCountry

End Function

Function GetCustTypeByCustTypeNum(passedCustTypeNum)
	
	GetCustTypeByCustTypeNumResult = ""
	
	Set cnnGetCustTypeByCustTypeNum = Server.CreateObject("ADODB.Connection")
	cnnGetCustTypeByCustTypeNum.open Session("ClientCnnString")

	SQLGetCustTypeByCustTypeNum = "SELECT Description FROM CustomerType WHERE CustTypeSequence = " & passedCustTypeNum
	 
	Set rsGetCustTypeByCustTypeNum = Server.CreateObject("ADODB.Recordset")
	rsGetCustTypeByCustTypeNum.CursorLocation = 3 
	Set rsGetCustTypeByCustTypeNum= cnnGetCustTypeByCustTypeNum.Execute(SQLGetCustTypeByCustTypeNum)
	
	If not rsGetCustTypeByCustTypeNum.eof then 
		GetCustTypeByCustTypeNumResult = rsGetCustTypeByCustTypeNum("Description")
	End IF	
	
	set rsGetCustTypeByCustTypeNum = Nothing
	cnnGetCustTypeByCustTypeNum.Close
	set cnnGetCustTypeByCustTypeNum = Nothing
	
	GetCustTypeByCustTypeNum = GetCustTypeByCustTypeNumResult

End Function

Function CheckScheduler(passedScheduleTable,passedScheduleField)
	
	resultCheckScheduler = "-1,No Results returned from CheckScheduler function"
	ScheduleTable = passedScheduleTable
	ScheduleField = passedScheduleField
	ScheduleString = ""
	
	Set cnnCheckScheduler = Server.CreateObject("ADODB.Connection")
	cnnCheckScheduler.open MUV_READ("ClientCnnString") 
	Set rsCheckScheduler = Server.CreateObject("ADODB.Recordset")
	
	SQLCheckScheduler = "SELECT " & ScheduleField & " AS ScheduleString FROM " & ScheduleTable 
	Set rsCheckScheduler = cnnCheckScheduler.Execute(SQLCheckScheduler)
	
	If NOT rsCheckScheduler.EOF Then
		ScheduleString = rsCheckScheduler("ScheduleString")
	End IF
	
	
	If ScheduleString <> "" Then ' ok, there was something there
	
		ScheduleStringArray = Split(ScheduleString,",")
		
		If cInt(ScheduleStringArray(Weekday(Now())-1)) <> 1 Then ' Is it set Not to run today?
		
			resultCheckScheduler = "0,Not set to run today"
			
		Else
		
			' So far so good
			' See if we are closed or closed early today
			' M=F only
			' But only if the option is set not to run
			If cInt(ScheduleStringArray(14)) = 1 OR cInt(ScheduleStringArray(15)) = 1 Then 'NoReportIfClosed or NoReportIfClosingEarly 
			
				If Weekday(Now()) <> 1 and Weekday(Now()) <> 7 Then
	
					SQLCheckScheduler = "SELECT * FROM Settings_CompanyCalendar WHERE "
					SQLCheckScheduler = SQLCheckScheduler & "Year(getdate()) = YearNum AND "
					SQLCheckScheduler = SQLCheckScheduler & "Month(getdate()) = MonthNum AND "
					SQLCheckScheduler = SQLCheckScheduler & "Day(getdate()) = DayNum"
					
					Set rsCheckScheduler = cnnCheckScheduler.Execute(SQLCheckScheduler)
	
					If NOT rsCheckScheduler.EOF Then
						If rsCheckScheduler("OpenClosedCloseEarly") = "Closed" And cInt(ScheduleStringArray(14)) = 1 Then resultCheckScheduler = "0,Set not to run when closed"
						If rsCheckScheduler("OpenClosedCloseEarly") = "Close Early" And cInt(ScheduleStringArray(15)) = 1 Then resultCheckScheduler = "0,Set not to run when closing early"
					End If
				
				End If
		
			End If
			
		End If		
		
		
		' If we got this far & we are still set to run, see if it is the right time
		If left(resultCheckScheduler,1) <> "0" Then
			
			ArrayPosition = Weekday(Now()) + 6
			ScheduleTimeToRun = ScheduleStringArray(ArrayPosition)
		
			ScheduleTimeToRun = cDate(ScheduleTimeToRun)

'Response.Write("ScheduleTimeToRun: " &  ScheduleTimeToRun & "<br>")
'Response.Write("Time(): " &  Time() & "<br>")
'Response.Write("DateDiff: " & DateDiff("n",Time(),ScheduleTimeToRun) & "<br>")
			
			If DateDiff("n",Time(),ScheduleTimeToRun) > 0 Then ' Not time to run yet
			
				resultCheckScheduler = "0,Not time yet"
			
			End If
		
		End If
	
	
		' Last check, if we are still ok, make sure it hasn't already run today
		If left(resultCheckScheduler,1) <> "0" Then
	
			SQLCheckScheduler = "SELECT * FROM SC_SchedulerLog WHERE "
			SQLCheckScheduler = SQLCheckScheduler & "Year(getdate()) = Year(RecordCreationDateTime) AND "
			SQLCheckScheduler = SQLCheckScheduler & "Month(getdate()) = Month(RecordCreationDateTime) AND "
			SQLCheckScheduler = SQLCheckScheduler & "Day(getdate()) = Day(RecordCreationDateTime) AND "
			SQLCheckScheduler = SQLCheckScheduler & "pageName = '" & Request.ServerVariables("SCRIPT_NAME") & "'"
					
			Set rsCheckScheduler = cnnCheckScheduler.Execute(SQLCheckScheduler)
			
			If NOT rsCheckScheduler.EOF Then
				resultCheckScheduler = "0,Already ran today" ' it already ran today
			End If
	
		End If					
		
	
		' Final step, this is a little different
		'It will insert the record saying the report has now been run
		If left(resultCheckScheduler,1) <> "0" Then
			SQLCheckScheduler = "INSERT INTO SC_SchedulerLog (pageName) VALUES ('" & Request.ServerVariables("SCRIPT_NAME") & "')"
						
			Set rsCheckScheduler = cnnCheckScheduler.Execute(SQLCheckScheduler)
		End If
		
		cnnCheckScheduler.Close
		Set rsCheckScheduler = Nothing
		Set cnnCheckScheduler = Nothing
	
	End If

	If left(resultCheckScheduler,1) <> "0" Then resultCheckScheduler = "1,OK"
	
	CheckScheduler = resultCheckScheduler 

End Function

Function GetCustChainIDByCustID(passedCustID)

	resultGetCustChainIDByCustID = ""

	Set cnnGetCustChainIDByCustID = Server.CreateObject("ADODB.Connection")
	cnnGetCustChainIDByCustID.open Session("ClientCnnString")
			
	SQLGetCustChainIDByCustID = "Select ChainNum from AR_Customer where CustNum = '" & passedCustID & "'"
	 
	Set rsGetCustChainIDByCustID = Server.CreateObject("ADODB.Recordset")
	rsGetCustChainIDByCustID.CursorLocation = 3 
	Set rsGetCustChainIDByCustID= cnnGetCustChainIDByCustID.Execute(SQLGetCustChainIDByCustID)
		
	If not rsGetCustChainIDByCustID.eof then resultGetCustChainIDByCustID = rsGetCustChainIDByCustID("ChainNum")
	
	set rsGetCustChainIDByCustID= Nothing
	cnnGetCustChainIDByCustID.Close	
	set cnnGetCustChainIDByCustID= Nothing
	
	GetCustChainIDByCustID = resultGetCustChainIDByCustID
	
End Function

%>
