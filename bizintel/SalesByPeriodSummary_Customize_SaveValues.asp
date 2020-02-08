<%

Dim arr1()

selBasePeriodsRange = ""
selPeriodsForComparison = "" 

selBasePeriodsRange1 = Request.Form("selBasePeriodsRange0")
If selBasePeriodsRange1 <>	"" Then
	selBasePeriodsRangeArray1 = Split(selBasePeriodsRange1,"*")
	selBasePeriodsRange1 = selBasePeriodsRangeArray1(0)
End If

selBasePeriodsRange2 = Request.Form("selBasePeriodsRange1")
If selBasePeriodsRange2 <>	"" Then
	selBasePeriodsRangeArray2 = Split(selBasePeriodsRange2,"*")
	selBasePeriodsRange2 = selBasePeriodsRangeArray2(0)
End If

selBasePeriodsRange3 = Request.Form("selBasePeriodsRange2")
If selBasePeriodsRange3 <>	"" Then
	selBasePeriodsRangeArray3 = Split(selBasePeriodsRange3,"*")
	selBasePeriodsRange3 = selBasePeriodsRangeArray3(0)
End If

selBasePeriodsRange4 = Request.Form("selBasePeriodsRange3")
If selBasePeriodsRange4 <>	"" Then
	selBasePeriodsRangeArray4 = Split(selBasePeriodsRange4,"*")
	selBasePeriodsRange4 = selBasePeriodsRangeArray4(0)
End If

selBasePeriodsRange5 = Request.Form("selBasePeriodsRange4")
If selBasePeriodsRange5 <>	"" Then
	selBasePeriodsRangeArray5 = Split(selBasePeriodsRange5,"*")
	selBasePeriodsRange5 = selBasePeriodsRangeArray5(0)
End If

selBasePeriodsRange6 = Request.Form("selBasePeriodsRange5")
If selBasePeriodsRange6 <>	"" Then
	selBasePeriodsRangeArray6 = Split(selBasePeriodsRange6,"*")
	selBasePeriodsRange6 = selBasePeriodsRangeArray6(0)
End If

If selBasePeriodsRange1 <> "" Then
	selBasePeriodsRange = selBasePeriodsRange1
End If
If selBasePeriodsRange2 <> "" Then
	selBasePeriodsRange = selBasePeriodsRange & "," & selBasePeriodsRange2
End If
If selBasePeriodsRange3 <> "" Then
	selBasePeriodsRange = selBasePeriodsRange & "," & selBasePeriodsRange3
End If
If selBasePeriodsRange4 <> "" Then
	selBasePeriodsRange = selBasePeriodsRange & "," & selBasePeriodsRange4
End If
If selBasePeriodsRange5 <> "" Then
	selBasePeriodsRange = selBasePeriodsRange & "," & selBasePeriodsRange5
End If
If selBasePeriodsRange6 <> "" Then
	selBasePeriodsRange = selBasePeriodsRange & "," & selBasePeriodsRange6
End If



selPeriodsForComparison1 = Request.Form("selPeriodsForComparison0")
If selPeriodsForComparison1 <> "" Then
	selPeriodsForComparisonArray1 = Split(selPeriodsForComparison1,"*")
	selPeriodsForComparison1 = selPeriodsForComparisonArray1(0)
End If

selPeriodsForComparison2 = Request.Form("selPeriodsForComparison1")
If selPeriodsForComparison2 <> "" Then
	selPeriodsForComparisonArray2 = Split(selPeriodsForComparison2,"*")
	selPeriodsForComparison2 = selPeriodsForComparisonArray2(0)
End If

selPeriodsForComparison3 = Request.Form("selPeriodsForComparison2")
If selPeriodsForComparison3 <> "" Then
	selPeriodsForComparisonArray3 = Split(selPeriodsForComparison3,"*")
	selPeriodsForComparison3 = selPeriodsForComparisonArray3(0)
End If

selPeriodsForComparison4 = Request.Form("selPeriodsForComparison3")
If selPeriodsForComparison4 <> "" Then
	selPeriodsForComparisonArray4 = Split(selPeriodsForComparison4,"*")
	selPeriodsForComparison4 = selPeriodsForComparisonArray4(0)
End If

selPeriodsForComparison5 = Request.Form("selPeriodsForComparison4")
If selPeriodsForComparison5 <> "" Then
	selPeriodsForComparisonArray5 = Split(selPeriodsForComparison5,"*")
	selPeriodsForComparison5 = selPeriodsForComparisonArray5(0)
End If

selPeriodsForComparison6 = Request.Form("selPeriodsForComparison5")
If selPeriodsForComparison6 <> "" Then
	selPeriodsForComparisonArray6 = Split(selPeriodsForComparison6,"*")
	selPeriodsForComparison6 = selPeriodsForComparisonArray6(0)
End If



If selPeriodsForComparison1 <> "" Then
	selPeriodsForComparison = selPeriodsForComparison1
End If
If selPeriodsForComparison2 <> "" Then
	selPeriodsForComparison = selPeriodsForComparison & "," & selPeriodsForComparison2
End If
If selPeriodsForComparison3 <> "" Then
	selPeriodsForComparison = selPeriodsForComparison & "," & selPeriodsForComparison3
End If
If selPeriodsForComparison4 <> "" Then
	selPeriodsForComparison = selPeriodsForComparison & "," & selPeriodsForComparison4
End If
If selPeriodsForComparison5 <> "" Then
	selPeriodsForComparison = selPeriodsForComparison & "," & selPeriodsForComparison5
End If
If selPeriodsForComparison6 <> "" Then
	selPeriodsForComparison = selPeriodsForComparison & "," & selPeriodsForComparison6
End If






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
SQL = SQL & "ReportSpecificData13 = '" & selBasePeriodsRange & "', "
SQL = SQL & "ReportSpecificData14 = '" & selPeriodsForComparison & "' "

SQL = SQL & "WHERE ReportNumber = 1500 AND UserNo = " & Session("userNo")

Response.Write(SQL)

Set rs= cnn8.Execute(SQL)

cnn8.Close

Set rs = Nothing
Set cnn8 = Nothing

Response.Redirect ("SalesByPeriodSummary.asp")
%>

 
