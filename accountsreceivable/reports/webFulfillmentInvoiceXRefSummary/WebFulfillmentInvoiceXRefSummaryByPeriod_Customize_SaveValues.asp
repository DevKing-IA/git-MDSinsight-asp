<%

OCSWebOrderOrMDSInvoice = Request.Form("selWebOrderOrMDSInvoiceFilter")

If OCSWebOrderOrMDSInvoice = "" Then
	DatesOrPeriods = "Periods"
	RangeStartDateCustomize = ""
	RangeEndDateCustomize = ""
	StartPeriodBeingEvaluatedCustomize = ""
	EndPeriodBeingEvaluatedCustomize = ""
Else
	DatesOrPeriods = "Periods"
	StartPeriodBeingEvaluatedCustomize = Request.Form("selPeriodStart")
	EndPeriodBeingEvaluatedCustomize = Request.Form("selPeriodEnd")
	RangeStartDateCustomize = ""
	RangeEndDateCustomize = ""
End If

DefaultSelectedCustomerClassesForInvoiceReport = Request.Form("chkClassCode")

If Right(DefaultSelectedCustomerClassesForInvoiceReport,1) = "," Then 
	DefaultSelectedCustomerClassesForInvoiceReport = left(DefaultSelectedCustomerClassesForInvoiceReport,Len(DefaultSelectedCustomerClassesForInvoiceReport)-1)
End If

CustomerClassArrayForCustomize = ""
CustomerClassArrayForCustomize = Split(DefaultSelectedCustomerClassesForInvoiceReport,",")

For z = 0 to UBound(CustomerClassArrayForCustomize)
	If z = 0 Then
		DefaultSelectedCustomerClassesForInvoiceReport = Trim(CustomerClassArrayForCustomize(z))
	Else
		DefaultSelectedCustomerClassesForInvoiceReport = DefaultSelectedCustomerClassesForInvoiceReport & "," & Trim(CustomerClassArrayForCustomize(z))
	End If
Next	



DefaultSelectedCustomerTypesForInvoiceReport = Request.Form("chkCustomerType")

If Right(DefaultSelectedCustomerTypesForInvoiceReport,1) = "," Then 
	DefaultSelectedCustomerTypesForInvoiceReport = left(DefaultSelectedCustomerTypesForInvoiceReport,Len(DefaultSelectedCustomerTypesForInvoiceReport)-1)
End If

CustomerTypeArrayForCustomize = ""
CustomerTypeArrayForCustomize = Split(DefaultSelectedCustomerTypesForInvoiceReport,",")

For z = 0 to UBound(CustomerTypeArrayForCustomize)
	If z = 0 Then
		DefaultSelectedCustomerTypesForInvoiceReport = Trim(CustomerTypeArrayForCustomize(z))
	Else
		DefaultSelectedCustomerTypesForInvoiceReport = DefaultSelectedCustomerTypesForInvoiceReport & "," & Trim(CustomerTypeArrayForCustomize(z))
	End If
Next	


ShowOrdersWithRemarks = Request.Form("chkShowOrdersWithRemarks")
ShowOrdersWithoutRemarks = Request.Form("chkShowOrdersWithoutRemarks")

If (ShowOrdersWithRemarks <> "" AND ShowOrdersWithRemarks = "on") Then ShowOrdersWithRemarks = "true" Else ShowOrdersWithRemarks = "false"
If (ShowOrdersWithoutRemarks <> "" AND ShowOrdersWithoutRemarks = "on") Then ShowOrdersWithoutRemarks = "true" Else ShowOrdersWithoutRemarks = "false"


ShowOrdersThatAreInvoiced = Request.Form("chkShowOrdersThatAreInvoiced")
ShowOrdersThatAreNotInvoiced = Request.Form("chkShowOrdersThatAreNotInvoiced")

If (ShowOrdersThatAreInvoiced <> "" AND ShowOrdersThatAreInvoiced = "on") Then ShowOrdersThatAreInvoiced = "true" Else ShowOrdersThatAreInvoiced = "false"
If (ShowOrdersThatAreNotInvoiced <> "" AND ShowOrdersThatAreNotInvoiced = "on") Then ShowOrdersThatAreNotInvoiced = "true" Else ShowOrdersThatAreNotInvoiced = "false"


ShowOrdersThatAreHidden = Request.Form("chkShowOrdersThatAreHidden")

If (ShowOrdersThatAreHidden <> "" AND ShowOrdersThatAreHidden = "on") Then ShowOrdersThatAreHidden = "true" Else ShowOrdersThatAreHidden = "false"



SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1900 AND UserNo = " & Session("userNo")

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)

'Rec does not exist yet, make it quick but empty, update it later
If rs.EOF Then
	SQL = "INSERT INTO Settings_Reports (ReportNumber, UserNo) VALUES (1900, " & Session("userNo") & ")"
	rs.Close
	Set rs= cnn8.Execute(SQL)
End If

'Now update the table with the values

SQL = "UPDATE Settings_Reports Set ReportSpecificData1 = '" & OCSWebOrderOrMDSInvoice & "', "
SQL = SQL & "ReportSpecificData2 = '" & DatesOrPeriods & "', "
SQL = SQL & "ReportSpecificData3 = '" & StartPeriodBeingEvaluatedCustomize & "', "
SQL = SQL & "ReportSpecificData4 = '" & EndPeriodBeingEvaluatedCustomize & "', " 
SQL = SQL & "ReportSpecificData5 = '" & RangeStartDateCustomize & "', " 
SQL = SQL & "ReportSpecificData6 = '" & RangeEndDateCustomize & "', " 
SQL = SQL & "ReportSpecificData7 = '" & DefaultSelectedCustomerClassesForInvoiceReport & "', "
SQL = SQL & "ReportSpecificData8 = '" & ShowOrdersWithRemarks & "', "
SQL = SQL & "ReportSpecificData9 = '" & ShowOrdersWithoutRemarks & "', "
SQL = SQL & "ReportSpecificData10 = '" & ShowOrdersThatAreInvoiced & "', "
SQL = SQL & "ReportSpecificData11 = '" & ShowOrdersThatAreNotInvoiced & "', "
SQL = SQL & "ReportSpecificData12 = '" & ShowOrdersThatAreHidden & "', " 
SQL = SQL & "ReportSpecificData13 = '" & DefaultSelectedCustomerTypesForInvoiceReport & "' "
SQL = SQL & "WHERE ReportNumber = 1900 AND UserNo = " & Session("userNo")

Set rs= cnn8.Execute(SQL)

cnn8.Close

Set rs = Nothing
Set cnn8 = Nothing

Response.Redirect ("WebFulfillmentInvoiceXRefSummaryByPeriod.asp")
%>

 
