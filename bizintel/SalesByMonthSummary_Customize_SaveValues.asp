<%


selMonthYearCombinations = Request.Form("selMonthYearCombinations")

DefaultSelectedCustomerClassesForSalesReport = Request.Form("chkClassCode")

If Right(DefaultSelectedCustomerClassesForSalesReport,1) = "," Then 
	DefaultSelectedCustomerClassesForSalesReport = left(DefaultSelectedCustomerClassesForSalesReport,Len(DefaultSelectedCustomerClassesForSalesReport)-1)
End If

CustomerClassArrayForCustomize = ""
CustomerClassArrayForCustomize = Split(DefaultSelectedCustomerClassesForSalesReport,",")

For z = 0 to UBound(CustomerClassArrayForCustomize)
	If z = 0 Then
		DefaultSelectedCustomerClassesForSalesReport = Trim(CustomerClassArrayForCustomize(z))
	Else
		DefaultSelectedCustomerClassesForSalesReport = DefaultSelectedCustomerClassesForSalesReport & "," & Trim(CustomerClassArrayForCustomize(z))
	End If
Next	


InvoiceTypeBackOrder = Request.Form("chkBackorder")
InvoiceTypeCreditMemo = Request.Form("chkCreditMemo")
InvoiceTypeARDebit = Request.Form("chkARDebit")
InvoiceTypeRental = Request.Form("chkRental")
InvoiceTypeRouteInvoicing = Request.Form("chkRouteInvoicing")
InvoiceTypeInterest = Request.Form("chkInterest")
InvoiceTypeTelselInvoicing = Request.Form("chkTelselInvoicing")


If (InvoiceTypeBackOrder <> "" AND InvoiceTypeBackOrder = "on") Then InvoiceTypeBackOrder = "B" Else InvoiceTypeBackOrder = ""
If (InvoiceTypeCreditMemo <> "" AND InvoiceTypeCreditMemo = "on") Then InvoiceTypeCreditMemo = "C" Else InvoiceTypeCreditMemo = ""
If (InvoiceTypeARDebit <> "" AND InvoiceTypeARDebit = "on") Then InvoiceTypeARDebit = "E" Else InvoiceTypeARDebit = ""
If (InvoiceTypeRental <> "" AND InvoiceTypeRental = "on") Then InvoiceTypeRental = "G" Else InvoiceTypeRental = ""
If (InvoiceTypeRouteInvoicing <> "" AND InvoiceTypeRouteInvoicing = "on") Then InvoiceTypeRouteInvoicing = "I" Else InvoiceTypeRouteInvoicing = ""
If (InvoiceTypeInterest <> "" AND InvoiceTypeInterest = "on") Then InvoiceTypeInterest = "O" Else InvoiceTypeInterest = ""
If (InvoiceTypeTelselInvoicing <> "" AND InvoiceTypeTelselInvoicing = "on") Then InvoiceTypeTelselInvoicing = "T" Else InvoiceTypeTelselInvoicing = ""


SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1500 AND UserNo = " & Session("userNo")

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)

'Rec does not exist yet, make it quick but empty, update it later
If rs.EOF Then
	SQL = "INSERT INTO Settings_Reports (ReportNumber, UserNo) VALUES (1500, " & Session("userNo") & ")"
	rs.Close
	Set rs= cnn8.Execute(SQL)
End If

'Now update the table with the values


SQL = "UPDATE Settings_Reports Set ReportSpecificData1 = '', "
SQL = SQL & "ReportSpecificData2 = '', "
SQL = SQL & "ReportSpecificData3 = '', " 
SQL = SQL & "ReportSpecificData4 = '', " 
SQL = SQL & "ReportSpecificData5 = '" & DefaultSelectedCustomerClassesForSalesReport & "', " 
SQL = SQL & "ReportSpecificData6 = '" & InvoiceTypeBackOrder & "', " 
SQL = SQL & "ReportSpecificData7 = '" & InvoiceTypeCreditMemo & "', "
SQL = SQL & "ReportSpecificData8 = '" & InvoiceTypeARDebit & "', "
SQL = SQL & "ReportSpecificData9 = '" & InvoiceTypeRental & "', "
SQL = SQL & "ReportSpecificData10 = '" & InvoiceTypeRouteInvoicing & "', "
SQL = SQL & "ReportSpecificData11 = '" & InvoiceTypeInterest & "', "
SQL = SQL & "ReportSpecificData12 = '" & InvoiceTypeTelselInvoicing & "', " 
SQL = SQL & "ReportSpecificData13 = '', "
SQL = SQL & "ReportSpecificData14 = '', "
SQL = SQL & "ReportSpecificData15 = '" & selMonthYearCombinations & "' "

SQL = SQL & "WHERE ReportNumber = 1500 AND UserNo = " & Session("userNo")

Response.Write(SQL)

Set rs= cnn8.Execute(SQL)

cnn8.Close

Set rs = Nothing
Set cnn8 = Nothing

Response.Redirect ("SalesByMonthSummary.asp")
%>

 
