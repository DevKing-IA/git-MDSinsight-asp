<%
Server.ScriptTimeout = 900000 'Default value
%>
<!--#include file="../../../inc/header.asp"-->
<!--#include file="../../../inc/jquery_table_search.asp"-->
<%
CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Web Fulfillment Invoice Cross Reference By Period"

Response.Write("<div id=""PleaseWaitPanel"">")
Response.Write("<br><br>Processing Web Fulfillment and Invoice Cross Reference By Period Data<br><br>This may take up to a full minute, please wait...<br><br>")
Response.Write("<img src=""../../img/loading.gif"" />")
Response.Write("</div>")
Response.Flush()

%>

<script type="text/javascript">

	$(document).ready(function() {
	
	    $("#PleaseWaitPanel").hide();
	    
	});

</script>

	<style>
	.form-control[disabled], .form-control[readonly], fieldset[disabled] .form-control{
		background-color:#fff;
		border: 1px solid #eee;
	}
	
	.invoicerangedatepicker {
		position: absolute;
		bottom: 25px;
		right: 24px;
		top: auto;
		cursor: pointer;
	}
	
	.activefilter {
	    background: #f0ad4e !important;
	}
		

	.modal-footer {
	    /*padding: 0px !important;*/
	    text-align: right !important;
	    border-top: 0px !important;
	}
	.bs-example-modal-lg-customize .row{
		margin-bottom: 10px;
	 	width: 100%;
		overflow: hidden;
	}
	
	.bs-example-modal-lg -customize.left-column{
		background: #eaeaea;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}
	
	.bs-example-modal-lg-customize .left-column h4{
		margin-top: 0px;
	}
	
	.bs-example-modal-lg-customize .right-column{
		background: #fff;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}


	.bs-example-modal-lg-customize .row{
		margin-bottom: 10px;
	 	width: 100%;
		overflow: hidden;
	}

	.bs-example-modal-lg-customize .left-column{
		background: #eaeaea;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}
	
	.bs-example-modal-lg-customize .left-column h4{
		margin-top: 0px;
	}
	
	.bs-example-modal-lg-customize .right-column{
		background: #fff;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}

	.filter-search-width{
		max-width: 36%;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
	    content: " \25B4\25BE" 
	    
	}
	
	table.sortable thead {
	    color:#222;
	    font-weight: bold;
	    cursor: pointer;
	}
	
	#PleaseWaitPanel{
		position: fixed;
		left: 470px;
		top: 275px;
		width: 975px;
		height: 300px;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	}    
	
	markY {
	    background-color: yellow;
	    color: black;
	} 
	
</style>

<%

Set cnnCompanyPeriods = Server.CreateObject("ADODB.Connection")
cnnCompanyPeriods.open (Session("ClientCnnString"))
Set rsCompanyPeriods = Server.CreateObject("ADODB.Recordset")
rsCompanyPeriods.CursorLocation = 3 

Set cnnMasterWebFulfillment = Server.CreateObject("ADODB.Connection")
cnnMasterWebFulfillment.open (Session("ClientCnnString"))
Set rsMasterWebFulfillment = Server.CreateObject("ADODB.Recordset")
rsMasterWebFulfillment.CursorLocation = 3 

Set cnnPeriodsLoopWebFulfillment = Server.CreateObject("ADODB.Connection")
cnnPeriodsLoopWebFulfillment.open (Session("ClientCnnString"))
Set rsPeriodsLoopWebFulfillment = Server.CreateObject("ADODB.Recordset")
rsPeriodsLoopWebFulfillment.CursorLocation = 3 

Set cnnInvoiceCount = Server.CreateObject("ADODB.Connection")
cnnInvoiceCount.open (Session("ClientCnnString"))
Set rsInvoiceCount = Server.CreateObject("ADODB.Recordset")
rsInvoiceCount.CursorLocation = 3 

Set cnnInvoiceGrandTotals = Server.CreateObject("ADODB.Connection")
cnnInvoiceGrandTotals.open (Session("ClientCnnString"))
Set rsSQLInvoiceGrandTotals = Server.CreateObject("ADODB.Recordset")
rsSQLInvoiceGrandTotals.CursorLocation = 3 


	'**************************************************
	'Read Settings_Reports
	'**************************************************
	SQL = "SELECT * from Settings_Reports where ReportNumber = 1900 AND UserNo = " & Session("userNo")
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs= cnn8.Execute(SQL)
	If NOT rs.EOF Then
	
		UseSettings_Reports = True
		
		OCSWebOrderOrMDSInvoice = rs("ReportSpecificData1")
		DatesOrPeriods = "Dates"
		StartPeriodBeingEvaluatedCustomize = rs("ReportSpecificData3")
		EndPeriodBeingEvaluatedCustomize = rs("ReportSpecificData4")
		RangeStartDateCustomize = ""
		RangeEndDateCustomize = ""
		DefaultSelectedCustomerClassesForInvoiceReport = rs("ReportSpecificData7")
		ShowOrdersWithRemarks = rs("ReportSpecificData8")
		ShowOrdersWithoutRemarks = rs("ReportSpecificData9")
		ShowOrdersThatAreInvoiced = rs("ReportSpecificData10")
		ShowOrdersThatAreNotInvoiced = rs("ReportSpecificData11")
		ShowOrdersThatAreHidden = rs("ReportSpecificData12")
		DefaultSelectedCustomerTypesForInvoiceReport = rs("ReportSpecificData13")
		
		If IsNull(OCSWebOrderOrMDSInvoice) Then OCSWebOrderOrMDSInvoice = "OCS"
		If IsNull(DatesOrPeriods) Then DatesOrPeriods = "Periods"


		If IsNull(StartPeriodBeingEvaluatedCustomize) OR StartPeriodBeingEvaluatedCustomize = "" Then
	
			SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods "
			SQLCompanyPeriods = SQLCompanyPeriods & "WHERE InternalRecordIdentifier = " & GetFirstReportPeriodThisYearIntRecID()
			SQLCompanyPeriods = SQLCompanyPeriods & " ORDER BY [Year] DESC, Period DESC"
			
			Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
			
			If NOT rsCompanyPeriods.EOF Then
				firstPeriodofYearIntRecID = rsCompanyPeriods("InternalRecordIdentifier")
				firstPeriodofYearBeginDateDefault = rsCompanyPeriods("BeginDate")
				firstPeriodEndDateDefault = rsCompanyPeriods("EndDate") 
				StartPeriodBeingEvaluatedCustomize = firstPeriodofYearIntRecID
			End If
		End If
		
		
		If IsNull(EndPeriodBeingEvaluatedCustomize) OR EndPeriodBeingEvaluatedCustomize = "" Then
		
			SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods "
			SQLCompanyPeriods = SQLCompanyPeriods & "WHERE InternalRecordIdentifier = " & GetCurrentReportPeriodIntRecID()
			SQLCompanyPeriods = SQLCompanyPeriods & " ORDER BY [Year] DESC, Period DESC"
			
			Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
			
			If NOT rsCompanyPeriods.EOF Then
				currentPeriodIntRecID = rsCompanyPeriods("InternalRecordIdentifier")
				currentPeriodBeginDateDefault = rsCompanyPeriods("BeginDate")
				currentPeriodEndDateDefault = rsCompanyPeriods("EndDate") 
				EndPeriodBeingEvaluatedCustomize = currentPeriodIntRecID
			End If
			
		End If

		
		If IsNull(DefaultSelectedCustomerClassesForInvoiceReport) Then DefaultSelectedCustomerClassesForInvoiceReport = ""
		If IsNull(DefaultSelectedCustomerTypesForInvoiceReport) Then DefaultSelectedCustomerTypesForInvoiceReport = ""
		
		If IsNull(ShowOrdersWithRemarks) OR ShowOrdersWithRemarks = "false" Then 
			ShowOrdersWithRemarks = 0
		ElseIf ShowOrdersWithRemarks = "true" Then
			ShowOrdersWithRemarks = 1
		End If

		If IsNull(ShowOrdersWithoutRemarks) OR ShowOrdersWithoutRemarks = "false" Then 
			ShowOrdersWithoutRemarks = 0
		ElseIf ShowOrdersWithoutRemarks = "true" Then
			ShowOrdersWithoutRemarks = 1
		End If

		If IsNull(ShowOrdersThatAreInvoiced) OR ShowOrdersThatAreInvoiced = "false" Then 
			ShowOrdersThatAreInvoiced = 0
		ElseIf ShowOrdersThatAreInvoiced = "true" Then
			ShowOrdersThatAreInvoiced = 1
		End If
		
		If IsNull(ShowOrdersThatAreNotInvoiced) OR ShowOrdersThatAreNotInvoiced = "false" Then 
			ShowOrdersThatAreNotInvoiced = 0
		ElseIf ShowOrdersThatAreNotInvoiced = "true" Then
			ShowOrdersThatAreNotInvoiced = 1
		End If

		If IsNull(ShowOrdersThatAreHidden) OR ShowOrdersThatAreHidden = "false" Then 
			ShowOrdersThatAreHidden = 0
		ElseIf ShowOrdersThatAreHidden = "true" Then
			ShowOrdersThatAreHidden = 1
		End If
		
	Else
	
		UseSettings_Reports = False
		OCSWebOrderOrMDSInvoice = "OCS"
		
		'-------------------------------------------------------------------------
		'For this report, we need the first period and current period of this year 
		'if no periods were specified in the customizations. Default view is all
		'periods this year.
		'-------------------------------------------------------------------------
		
		DatesOrPeriods = "Periods"


		SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods "
		SQLCompanyPeriods = SQLCompanyPeriods & "WHERE InternalRecordIdentifier = " & GetFirstReportPeriodThisYearIntRecID()
		SQLCompanyPeriods = SQLCompanyPeriods & " ORDER BY [Year] DESC, Period DESC"
		
		Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
		
		If NOT rsCompanyPeriods.EOF Then
			firstPeriodofYearIntRecID = rsCompanyPeriods("InternalRecordIdentifier")
			firstPeriodofYearBeginDateDefault = rsCompanyPeriods("BeginDate")
			firstPeriodEndDateDefault = rsCompanyPeriods("EndDate") 
			StartPeriodBeingEvaluatedCustomize = firstPeriodofYearIntRecID
		End If
		
		
		SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods "
		SQLCompanyPeriods = SQLCompanyPeriods & "WHERE InternalRecordIdentifier = " & GetCurrentReportPeriodIntRecID()
		SQLCompanyPeriods = SQLCompanyPeriods & " ORDER BY [Year] DESC, Period DESC"
		
		Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
		
		If NOT rsCompanyPeriods.EOF Then
			currentPeriodIntRecID = rsCompanyPeriods("InternalRecordIdentifier")
			currentPeriodBeginDateDefault = rsCompanyPeriods("BeginDate")
			currentPeriodEndDateDefault = rsCompanyPeriods("EndDate") 
			EndPeriodBeingEvaluatedCustomize = currentPeriodIntRecID
		End If

		RangeStartDateCustomize= ""
		RangeEndDateCustomize = ""	
		DefaultSelectedCustomerClassesForInvoiceReport = ""
		
		'-------------------------------------------------------------------------
		'As per Warren, Default View Should Include Orders That Are Hidden,
		'Orders Invoiced and Un-Invoiced, Orders With and Without Remarks
		'-------------------------------------------------------------------------
				
		ShowOrdersWithRemarks = 1
		ShowOrdersWithoutRemarks = 1
		ShowOrdersThatAreInvoiced = 1
		ShowOrdersThatAreNotInvoiced = 1
		ShowOrdersThatAreHidden = 0
		
		Set cnnUpdateReportSettings = Server.CreateObject("ADODB.Connection")
		cnnUpdateReportSettings.open Session("ClientCnnString")
			
		SQLUpdateReportSettings = "UPDATE Settings_Reports Set ReportSpecificData1 = '" & OCSWebOrderOrMDSInvoice & "', "
		SQLUpdateReportSettings = SQLUpdateReportSettings & "ReportSpecificData2 = '" & DatesOrPeriods & "', "
		SQLUpdateReportSettings = SQLUpdateReportSettings & "ReportSpecificData3 = '" & StartPeriodBeingEvaluatedCustomize & "', "
		SQLUpdateReportSettings = SQLUpdateReportSettings & "ReportSpecificData4 = '" & EndPeriodBeingEvaluatedCustomize & "', "
		SQLUpdateReportSettings = SQLUpdateReportSettings & "ReportSpecificData8 = '" & true & "', "		
		SQLUpdateReportSettings = SQLUpdateReportSettings & "ReportSpecificData9 = '" & true & "', "
		SQLUpdateReportSettings = SQLUpdateReportSettings & "ReportSpecificData10 = '" & true & "', "
		SQLUpdateReportSettings = SQLUpdateReportSettings & "ReportSpecificData11 = '" & true & "', "
		SQLUpdateReportSettings = SQLUpdateReportSettings & "ReportSpecificData12 = '" & false & "' " 
		SQLUpdateReportSettings = SQLUpdateReportSettings & "WHERE ReportNumber = 1900 AND UserNo = " & Session("userNo")
				
		Set rsUpdateReportSettings = Server.CreateObject("ADODB.Recordset")
		rsUpdateReportSettings.CursorLocation = 3 
		Set rsUpdateReportSettings= cnnUpdateReportSettings.Execute(SQLUpdateReportSettings)
		
		set rsUpdateReportSettings= Nothing
		
		'-------------------------------------------------------------------------
		'-------------------------------------------------------------------------
		
		
	End If										
	'**************************************************
	'End Read Settings_Reports
	'**************************************************
	
	
	
	
	'**************************************************
	'Start Build Page Header Text
	'**************************************************
	
	'----------------------------------------------------------------------
	'Check for customization by OCSAccessOrderDate (By Period)
	'If no customization, use current periods this year
	'----------------------------------------------------------------------	
	
	'Response.write("<br><br><br>StartPeriodBeingEvaluatedCustomize: " & StartPeriodBeingEvaluatedCustomize & "<br>")
	'Response.write("EndPeriodBeingEvaluatedCustomize: " & EndPeriodBeingEvaluatedCustomize& "<br>")
		
	'Response.write("<br>GetFirstReportPeriodThisYearIntRecID: " & GetFirstReportPeriodThisYearIntRecID() & "<br>")
	'Response.write("GetCurrentReportPeriodIntRecID: " & GetCurrentReportPeriodIntRecID() & "<br>")
	
	PeriodStartNum = GetPeriodByIntRecID(StartPeriodBeingEvaluatedCustomize)
	PeriodEndNum = GetPeriodByIntRecID(EndPeriodBeingEvaluatedCustomize)
	PeriodStartYear = GetPeriodYearByIntRecID(StartPeriodBeingEvaluatedCustomize)
	PeriodEndYear = GetPeriodYearByIntRecID(EndPeriodBeingEvaluatedCustomize)
	PeriodStartDate = GetPeriodBeginDateByIntRecID(StartPeriodBeingEvaluatedCustomize)
	PeriodEndDate = GetPeriodEndDateByIntRecID(EndPeriodBeingEvaluatedCustomize)
	
	If OCSWebOrderOrMDSInvoice = "OCS" Then
	
		If UseSettings_Reports = False Then
			PageHeaderText = "Web Fulfillment and Invoice Cross Reference By Period for All Orders This Year "
			PageHeaderText = PageHeaderText & " (" & FormatDateTime(PeriodStartDate,2) & " - " & FormatDateTime(Date(),2) & ")&nbsp;&nbsp;"
		Else
			PageHeaderText = "Web Fulfillment and Invoice Cross Reference By Period For OCS Web Orders In Period " & PeriodStartNum & " of " & PeriodStartYear & " to Period " & PeriodEndNum & " of " & PeriodEndYear
			PageHeaderText = PageHeaderText & " (" & FormatDateTime(PeriodStartDate,2) & " - " & FormatDateTime(PeriodEndDate,2) & ")&nbsp;&nbsp;"
		End If
		
	ElseIf OCSWebOrderOrMDSInvoice = "MDS" Then
	
		PageHeaderText = "Web Fulfillment and Invoice Cross Reference By Period For MDS Invoices In Period " & PeriodStartNum & " of " & PeriodStartYear & " to Period " & PeriodEndNum & " of " & PeriodEndYear			
		PageHeaderText = PageHeaderText & " (" & FormatDateTime(PeriodStartDate,2) & " - " & FormatDateTime(PeriodEndDate,2) & ")&nbsp;&nbsp;"
		
	End If
	'**************************************************
	'End Build Page Header Text
	'**************************************************
	

	CustomerClassArray = ""
	CustomerClassArray = Split(DefaultSelectedCustomerClassesForInvoiceReport,",")
	
	'**************************************************************************************
	'If Customer Class is empty from report settings, obtain all customer
	'classes from AR_CustomerClass
	'**************************************************************************************

	If UBound(CustomerClassArray) < 0 Then
	
		CustomerClassArrayString = ""
		
		Set cnnGetAllValidCustomerClasses = Server.CreateObject("ADODB.Connection")
		cnnGetAllValidCustomerClasses.open Session("ClientCnnString")
	
		resultGetAllValidCustomerClasses = ""
			
		SQLGetAllValidCustomerClasses = "SELECT DISTINCT(ClassCode) FROM AR_CustomerClass ORDER BY ClassCode"
		 
		Set rsGetAllValidCustomerClasses = Server.CreateObject("ADODB.Recordset")
		rsGetAllValidCustomerClasses.CursorLocation = 3 
		Set rsGetAllValidCustomerClasses= cnnGetAllValidCustomerClasses.Execute(SQLGetAllValidCustomerClasses)
		
		If NOT rsGetAllValidCustomerClasses.EOF Then 
		
			Do While NOT rsGetAllValidCustomerClasses.EOF
				CustomerClassArrayString = CustomerClassArrayString & rsGetAllValidCustomerClasses("ClassCode") & ","
				rsGetAllValidCustomerClasses.MoveNext
			Loop
				
			If Right(CustomerClassArrayString,1) = "," Then 
				CustomerClassArrayString = left(CustomerClassArrayString,Len(CustomerClassArrayString)-1)
			End If
			
			CustomerClassArray = Split(CustomerClassArrayString,",")

		End If
	
		rsGetAllValidCustomerClasses.Close
		set rsGetAllValidCustomerClasses= Nothing
		cnnGetAllValidCustomerClasses.Close	
		set cnnGetAllValidCustomerClasses = Nothing	
	
	End If
	
	'**************************************************************************************
	'End Build Customer Class Array
	'**************************************************************************************


	
	'**************************************************************************************
	'Build WHERE Clause For Customer Class Array
	'**************************************************************************************
	
	WHERE_CLAUSE_CUSTCLASS = ""
	
	For z = 0 to UBound(CustomerClassArray)
		
		If z = 0 Then
			WHERE_CLAUSE_CUSTCLASS = WHERE_CLAUSE_CUSTCLASS & " AND (CustClassCode = '" & CustomerClassArray(z) & "'"
		Else
			WHERE_CLAUSE_CUSTCLASS = WHERE_CLAUSE_CUSTCLASS & " OR CustClassCode = '" & CustomerClassArray(z) & "'"
		End If
	Next	
	
	If WHERE_CLAUSE_CUSTCLASS <> "" Then
		WHERE_CLAUSE_CUSTCLASS = WHERE_CLAUSE_CUSTCLASS & ") "
	End IF
	
	
	CustomerTypeArray = ""
	CustomerTypeArray = Split(DefaultSelectedCustomerTypesForInvoiceReport,",")
	
	'**************************************************************************************
	'If Customer Type is empty from report settings, obtain all customer
	'types from AR_Customer and CustomerType
	'**************************************************************************************

	If UBound(CustomerTypeArray) < 0 Then
	
		CustomerTypeArrayString = ""
		
		Set cnnGetAllValidCustomerTypes = Server.CreateObject("ADODB.Connection")
		cnnGetAllValidCustomerTypes.open Session("ClientCnnString")
	
		resultGetAllValidCustomerTypes = ""
			
		SQLGetAllValidCustomerTypes = "SELECT DISTINCT(CustType) FROM AR_Customer ORDER BY CustType"
		 
		Set rsGetAllValidCustomerTypes = Server.CreateObject("ADODB.Recordset")
		rsGetAllValidCustomerTypes.CursorLocation = 3 
		Set rsGetAllValidCustomerTypes= cnnGetAllValidCustomerTypes.Execute(SQLGetAllValidCustomerTypes)
		
		If NOT rsGetAllValidCustomerTypes.EOF Then 
		
			Do While NOT rsGetAllValidCustomerTypes.EOF
				CustomerTypeArrayString = CustomerTypeArrayString & rsGetAllValidCustomerTypes("CustType") & ","
				rsGetAllValidCustomerTypes.MoveNext
			Loop
				
			If Right(CustomerTypeArrayString,1) = "," Then 
				CustomerTypeArrayString = left(CustomerTypeArrayString,Len(CustomerTypeArrayString)-1)
			End If
			
			CustomerTypeArray = Split(CustomerTypeArrayString,",")

		End If
	
		rsGetAllValidCustomerTypes.Close
		set rsGetAllValidCustomerTypes= Nothing
		cnnGetAllValidCustomerTypes.Close	
		set cnnGetAllValidCustomerTypes = Nothing	
	
	End If
	
	'**************************************************************************************
	'End Build Customer Type Array
	'**************************************************************************************


	
	'**************************************************************************************
	'Build WHERE Clause For Customer Type Array
	'**************************************************************************************
	
	WHERE_CLAUSE_CUSTTYPE = ""
	
	For z = 0 to UBound(CustomerTypeArray)
		
		If z = 0 Then
			WHERE_CLAUSE_CUSTTYPE = WHERE_CLAUSE_CUSTTYPE & " AND (CustTypeNum = " & CustomerTypeArray(z) & " "
		Else
			WHERE_CLAUSE_CUSTTYPE = WHERE_CLAUSE_CUSTTYPE & " OR CustTypeNum = " & CustomerTypeArray(z) & " "
		End If
	Next	
	
	If WHERE_CLAUSE_CUSTTYPE <> "" Then
		WHERE_CLAUSE_CUSTTYPE = WHERE_CLAUSE_CUSTTYPE & ") "
	End IF
	


	'**************************************************************************************
	'Build WHERE Clause For Orders That Are/Are Not Invoiced
	'**************************************************************************************

	WHERE_CLAUSE_INVOICED = ""
	
	If ShowOrdersThatAreInvoiced = "1" AND ShowOrdersThatAreNotInvoiced = "0" Then
		WHERE_CLAUSE_INVOICED = " AND (MDSInvoiceID <> '')"
	ElseIf ShowOrdersThatAreInvoiced = "0" AND ShowOrdersThatAreNotInvoiced = "1" Then 
		WHERE_CLAUSE_INVOICED = " AND (MDSInvoiceID = '')"			
	End If
	
	'**************************************************************************************
	'Build WHERE Clause For Orders That Are Hidden
	'**************************************************************************************

	WHERE_CLAUSE_HIDDEN = ""
	
	If ShowOrdersThatAreHidden = "1" Then
		WHERE_CLAUSE_HIDDEN = " AND (DontIncludeOnReport = 1 OR DontIncludeOnReport = 0)"
	ElseIf ShowOrdersThatAreHidden = "0" Then 
		WHERE_CLAUSE_HIDDEN = " AND (DontIncludeOnReport = 0)"				
	End If


%>
	<h3 class="page-header">
	
		<a href="<%= BaseURL %>accountsreceivable/reports/main.asp"><button type="button" class="btn btn-success"><i class="fa fa-arrow-left" aria-hidden="true"></i> Back To <%= GetTerm("Accounts Receivable") %> Reports</button></a><br><br>
	
		<i class="fa fa-file-text" aria-hidden="true"></i> 
		<%= PageHeaderText %>
	
		<!-- modal button !-->
		<button type="button" class="btn btn-primary" data-toggle="modal" data-target=".bs-example-modal-lg-customize">
		  Customize
		</button>
		
		<% If UseSettings_Reports = True Then%>
			<a href="<%= BaseURL %>accountsreceivable/reports/WebFulfillmentInvoiceXRefSummaryByPeriodCustomize_ClearValues.asp"><button type="button" class="btn btn-primary">Clear Customizations</button></a>
		<% End If %>
	
	</h3>

	<!--#include file="WebFulfillmentInvoiceXRefSummaryByPeriod_Customize.asp"-->	
	 
	
	<h6 class="page-header">
	<table id="table-search" class='table table-striped table-condensed table-hover display'>
	<tr>
		<td width="20%">
			<% If UseSettings_Reports = True Then
				Response.Write("<span class='markY'>" & "Using Saved Customization Values</br>" & "</span>")
			End If %>
			
			<% If ShowOrdersThatAreHidden = 1 Then %>
				Include Hidden Orders is <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersThatAreHidden = 0 Then %>
				Include Hidden Orders is <strong>OFF</strong><br>
			<% End If %>
			
			<% If ShowOrdersThatAreInvoiced = 1 Then %>
				Include Orders That Are Invoiced is <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersThatAreNotInvoiced = 1 Then %>
				Include Orders That Are NOT Invoiced is <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersThatAreInvoiced = 0 AND ShowOrdersThatAreNotInvoiced = 0 Then %>
				Include Both Invoiced and Non-Invoiced Orders is <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersWithRemarks = 1 Then %>
				Include Orders With Remarks is <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersWithoutRemarks = 1 Then %>
				Include Orders Without Remarks <strong>ON</strong><br>
			<% End If %>
			
			<% If ShowOrdersWithRemarks = 0 AND ShowOrdersWithoutRemarks = 0 Then %>
				Include Orders With and Without Remarks is <strong>ON</strong><br>
			<% End If %>
		
		</td>
		<td>
		
			<% If OCSWebOrderOrMDSInvoice = "" Then %>
			
				Filter Orders By Date <strong>OFF</strong><br>
				
			<% ElseIf OCSWebOrderOrMDSInvoice = "OCS" Then %>
		
				Filter Orders By OCS Web Order Date Within a Period Range is <strong>ON</strong><br>
				Showing OCS Web Orders from <strong>Period <%= PeriodStartNum %> of  <%= PeriodStartYear %></strong> to <strong>Period <%= PeriodEndNum %> of <%= PeriodEndYear %></strong><br>
				These periods span the dates: <strong><%= FormatDateTime(PeriodStartDate,2) %> to <%= FormatDateTime(PeriodEndDate,2) %></strong><br>
		
			<% ElseIf OCSWebOrderOrMDSInvoice = "MDS" Then %>
		
				Filter Orders By MDS Invoice Date Within a Period Range is <strong>ON</strong><br>
				Showing MDS Invoiced Orders from <strong>Period <%= PeriodStartNum %> of  <%= PeriodStartYear %></strong> to <strong><%= PeriodEndNum %> of <%= PeriodEndYear %></strong><br>
				These periods span the dates: <strong><%= FormatDateTime(PeriodStartDate,2) %> to <%= FormatDateTime(PeriodEndDate,2) %></strong><br>
		
			<% End If %>
		
		</td>
		
		<td>
			<% For z = 0 to UBound(CustomerClassArray)
					currentClass = cStr(CustomerClassArray(z))
					%>Customer Class <%= currentClass %> - <%= GetCustomerClassNameByID(currentClass) %><br><%
			   Next
			%>
		</td>
				
		<td>
			<% For z = 0 to UBound(CustomerTypeArray)
					If CustomerTypeArray(z) <> "" Then
						currentCustType = CustomerTypeArray(z)
						%>Customer Type <strong><%= GetCustTypeByCustTypeNum(currentCustType) %></strong> - (Cust Type <%= currentCustType %>)<br><%
					End If
			   Next
			%>
		</td>
				
	
	</tr>
	</table>
	</h6>
	
<!-- row !-->
<div class="row">

<!-- responsive tables !-->
<div class="table-responsive">
	
<div class="input-group"> <span class="input-group-addon">Narrow Results</span>

    <input id="filter" type="text" class="form-control filter-search-width" placeholder="Type here...">
</div>

	<input type="hidden" name="txtShowOrdersThatAreHidden" id="txtShowOrdersThatAreHidden" value="<%= ShowOrdersThatAreHidden %>">
    <table id="tableSuperSum" class="food_planner sortable table table-striped table-condensed table-hover">
      <thead>
        <tr>
          <th class="sorttable">Period Dates</th>
          <th class="sorttable">Period/Year</th>
          <th class="sorttable">Class</th> 
          <th class="sorttable numeric"># Orders</th> 
          <th class="sorttable numeric"># Invoices</th>
          <th class="sorttable numeric">Order $</th>
          <th class="sorttable numeric">Invoice $</th> 
          <th class="sorttable numeric">Fulfillment Rate</th>
        </tr>
      </thead>
      
      <tbody class="searchable">
	

<%
'****************************************************************************************************************************************
'Loop through the period internal record identifiers
'StartPeriodBeingEvaluatedCustomize = Customized start period or date of first period of current year, if no customization selected
'EndPeriodBeingEvaluatedCustomize= Customized end period or current date, if no customization selected
'
'Within each period/loop, we will group web fulfillment results by Customer Class Code
'****************************************************************************************************************************************
		
GrandTotNumMDSInvoices = 0

for x = cInt(EndPeriodBeingEvaluatedCustomize) to cInt(StartPeriodBeingEvaluatedCustomize) step -1


	PeriodLoopStartDate = GetPeriodBeginDateByIntRecID(x)
	PeriodLoopEndDate = GetPeriodEndDateByIntRecID(x)
	PeriodNumLoop = GetPeriodByIntRecID(x)
	PeriodYearLoop = GetPeriodYearByIntRecID(x)
	
	'Response.write("<br><br>x: " & x & "<br><br>")

	'**************************************************************************************
	'Begin Build SQL STMT To Select From IN_WebFulfillment
	'**************************************************************************************	
	
	SQLMasterWebFulfillment = "SELECT COUNT(OCSAccessOrderID) AS TotNumWebOrders, SUM(OCSAccessMerchTotal) AS TotWebSales, "
	SQLMasterWebFulfillment = SQLMasterWebFulfillment & " SUM(MDSInvoiceTotal) AS TotMDSInvoiceAmt, "
	SQLMasterWebFulfillment = SQLMasterWebFulfillment & " CustClassCode"
	SQLMasterWebFulfillment = SQLMasterWebFulfillment & " FROM IN_WebFulfillment "
	
	'---------------------------------------------------------------------------------------------
	'Check for Customization by OCSAccessOrderDate (By Period) or by MDSInvoiceDate (By Period)
	'---------------------------------------------------------------------------------------------	
	
	If OCSWebOrderOrMDSInvoice = "OCS" Then
		SQLMasterWebFulfillment = SQLMasterWebFulfillment & " WHERE (OCSAccessOrderDate BETWEEN '" & PeriodLoopStartDate & "' AND '" & PeriodLoopEndDate & "') "
	ElseIf OCSWebOrderOrMDSInvoice = "MDS" Then
		SQLMasterWebFulfillment = SQLMasterWebFulfillment & " WHERE (MDSInvoiceDate BETWEEN '" & PeriodLoopStartDate & "' AND '" & PeriodLoopEndDate & "') "
	Else
		SQLMasterWebFulfillment = SQLMasterWebFulfillment & " WHERE (OCSAccessOrderDate BETWEEN '" & PeriodLoopStartDate & "' AND '" & PeriodLoopEndDate & "') "
	End If

	'---------------------------------------------------------------------------
	'Check for Customization by Customer Class, Remarks, Invoiced, and Hidden
	'---------------------------------------------------------------------------	
	
	If WHERE_CLAUSE_CUSTCLASS <> "" Then SQLMasterWebFulfillment = SQLMasterWebFulfillment & WHERE_CLAUSE_CUSTCLASS
	If WHERE_CLAUSE_CUSTTYPE <> "" Then SQLMasterWebFulfillment = SQLMasterWebFulfillment & WHERE_CLAUSE_CUSTTYPE
	If WHERE_CLAUSE_REMARKS <> "" Then SQLMasterWebFulfillment = SQLMasterWebFulfillment & WHERE_CLAUSE_REMARKS
	If WHERE_CLAUSE_INVOICED <> "" Then SQLMasterWebFulfillment = SQLMasterWebFulfillment & WHERE_CLAUSE_INVOICED
	If WHERE_CLAUSE_HIDDEN <> "" Then SQLMasterWebFulfillment = SQLMasterWebFulfillment & WHERE_CLAUSE_HIDDEN
	
	'---------------------------------------------------------------------------
	'Ending GROUP BY and ORDER BY clause
	'---------------------------------------------------------------------------
	
	SQLMasterWebFulfillment = SQLMasterWebFulfillment & " GROUP BY CustClassCode ORDER BY CustClassCode ASC"
	
	'**************************************************************************************
	'END Build SQL STMT To Select From IN_WebFulfillment
	'**************************************************************************************	
	
	'Response.write("<br><br>SQLMasterWebFulfillment: " & SQLMasterWebFulfillment)
	
	Set rsMasterWebFulfillment = cnnMasterWebFulfillment.Execute(SQLMasterWebFulfillment)
		
		If NOT rsMasterWebFulfillment.EOF Then
		
			Do While Not rsMasterWebFulfillment.EOF
			
				'---------------------------------------------------------------------------
				'Need to obtain the count of MDS Invoices in a separate SQL STMT
				'---------------------------------------------------------------------------		
				SQLInvoiceCount = "SELECT COUNT(MDSInvoiceID) AS TotNumMDSInvoices FROM IN_WebFulfillment "
					
				'------------------------------------------------------------------------------------------------
				'Check for customization by Date Range Customer Class, Remarks, Invoiced, and Hidden
				'------------------------------------------------------------------------------------------------	
				
				If OCSWebOrderOrMDSInvoice = "OCS" Then
					WHERE_CLAUSE_INVOICEDATERANGE = " WHERE (OCSAccessOrderDate BETWEEN '" & PeriodLoopStartDate & "' AND '" & PeriodLoopEndDate & "') "
				ElseIf OCSWebOrderOrMDSInvoice = "MDS" Then
					WHERE_CLAUSE_INVOICEDATERANGE = " WHERE (MDSInvoiceDate BETWEEN '" & PeriodLoopStartDate & "' AND '" & PeriodLoopEndDate & "') "
				Else
					WHERE_CLAUSE_INVOICEDATERANGE = " WHERE (OCSAccessOrderDate BETWEEN '" & PeriodLoopStartDate & "' AND '" & PeriodLoopEndDate & "') "
				End If	
				
				If WHERE_CLAUSE_INVOICEDATERANGE <> "" Then SQLInvoiceCount = SQLInvoiceCount & WHERE_CLAUSE_INVOICEDATERANGE
				If WHERE_CLAUSE_REMARKS <> "" Then SQLInvoiceCount = SQLInvoiceCount & WHERE_CLAUSE_REMARKS
				If WHERE_CLAUSE_HIDDEN <> "" Then SQLInvoiceCount = SQLInvoiceCount & WHERE_CLAUSE_HIDDEN
				
				If ShowOrdersThatAreInvoiced = "1" AND ShowOrdersThatAreNotInvoiced = "0" Then
					SQLInvoiceCount = SQLInvoiceCount & " AND (MDSInvoiceID IS NOT NULL AND MDSInvoiceID <> '') "
				ElseIf ShowOrdersThatAreInvoiced = "0" AND ShowOrdersThatAreNotInvoiced = "1" Then 			
					SQLInvoiceCount = SQLInvoiceCount & " AND (MDSInvoiceID IS NULL OR MDSInvoiceID = '') "
				Else	
					SQLInvoiceCount = SQLInvoiceCount & " AND (MDSInvoiceID IS NOT NULL AND MDSInvoiceID <> '') "	
				End If
				
				SQLInvoiceCount = SQLInvoiceCount & " AND CustClassCode = '" & rsMasterWebFulfillment("CustClassCode") & "' "
		
				'---------------------------------------------------------------------------
				'End ORDER BY clauses
				'---------------------------------------------------------------------------
				
				'Response.Write("<br>SQLInvoiceCount: " & SQLInvoiceCount & "<br>")
				
				Set rsSQLInvoiceCount = cnnInvoiceCount.Execute(SQLInvoiceCount)
				
				If NOT rsSQLInvoiceCount.EOF Then
					TotNumMDSInvoices = rsSQLInvoiceCount("TotNumMDSInvoices") 
					GrandTotNumMDSInvoices = GrandTotNumMDSInvoices + TotNumMDSInvoices
				End If
				'---------------------------------------------------------------------------
				'End MDS Invoice Count
				'---------------------------------------------------------------------------
	
				ClassCode = rsMasterWebFulfillment("CustClassCode")
				CustClassCode = GetCustomerClassNameByID(rsMasterWebFulfillment("CustClassCode")) 
				TotNumWebOrders = rsMasterWebFulfillment("TotNumWebOrders")
				TotWebSales = rsMasterWebFulfillment("TotWebSales")
				TotMDSInvoiceAmt = rsMasterWebFulfillment("TotMDSInvoiceAmt")
							
				If TotWebSales > 0 Then
					If TotMDSInvoiceAmt >= TotWebSales Then
						FulfillmentRate = FormatPercent(1,2)					
					Else
						FulfillmentRate = FormatPercent(TotMDSInvoiceAmt/TotWebSales,2)
					End If
				Else
					FulfillmentRate = FormatPercent(0,2)
				End If
							
				If TotWebSales <> "" Then
					TotWebSales = FormatCurrency(TotWebSales,2)
				End If
							
				If TotMDSInvoiceAmt <> "" AND TotMDSInvoiceAmt <> "---" Then
					TotMDSInvoiceAmt = FormatCurrency(TotMDSInvoiceAmt,2)
				End If
				
					
				%>
					<tr id="<%= x %>-<%= CustClassCode %>"> 
						<td><a href="<%= BaseURL %>accountsreceivable/reports/WebFulfillmentInvoiceXRefSummaryByPeriodDetail.asp?p=<%= x %>&cc=<%= ClassCode %>">
							<%= FormatDateTime(PeriodLoopStartDate,2) %> - <%= FormatDateTime(PeriodLoopEndDate,2) %></a>
						</td>
						<td><a href="<%= BaseURL %>accountsreceivable/reports/WebFulfillmentInvoiceXRefSummaryByPeriodDetail.asp?p=<%= x %>&cc=<%= ClassCode %>">
							Period <%= PeriodNumLoop %> of <%= PeriodYearLoop  %></a>
						</td>
						
						<td><a href="<%= BaseURL %>accountsreceivable/reports/WebFulfillmentInvoiceXRefSummaryByPeriodDetail.asp?p=<%= x %>&cc=<%= ClassCode %>">
							<%= CustClassCode %></a>
						</td>
						<td><%= TotNumWebOrders %></td>
						<td><%= TotNumMDSInvoices %></td>
						<td><%= TotWebSales %></td>
						<td><%= TotMDSInvoiceAmt %></td>
						<td><%= FulfillmentRate %></td>
					</tr>
				<%
				
				rsMasterWebFulfillment.movenext
					
			Loop
			
		Else
					
			%>
				<tr id="<%= x %>"> 
					<td><%= FormatDateTime(PeriodLoopStartDate,2) %> - <%= FormatDateTime(PeriodLoopEndDate,2) %></td>
					<td>Period <%= PeriodNumLoop %> of <%= PeriodYearLoop  %></td>
					<td>---</td>
					<td>No Orders This Period</td>
					<td>---</td>
					<td>---</td>
					<td>---</td>
					<td>---</td>
				</tr>
			<%
			
		End If

Next		
		%>
		<tfoot>
			<tr style="border-top:3px #337ab7 solid; background-color:#D3D3D3">
		   		<td>&nbsp;</td>
		   		<td>&nbsp;</td>
			    <td>&nbsp;</td>                
			    <td><strong>Total Web Orders</strong></td>
			    <td><strong>Total MDS Invoices</strong></td>
			    <td><strong>Total Web Sales $</strong></td> 
			    <td><strong>Total MDS Invoice $</strong></td> 
			    <td><strong>Total Avg Fulfillment %</strong></td>
	        </tr>
	        <%
				'---------------------------------------------------------------------------
				'Need to obtain grand totals
				'---------------------------------------------------------------------------		
				SQLInvoiceGrandTotals = "SELECT COUNT(OCSAccessOrderID) AS GrandTotNumWebOrders, SUM(OCSAccessMerchTotal) AS GrandTotWebSales, "
				SQLInvoiceGrandTotals = SQLInvoiceGrandTotals  & " SUM(MDSInvoiceTotal) AS GrandTotMDSInvoiceAmt FROM IN_WebFulfillment "
					
				'------------------------------------------------------------------------------------------------
				'Check for customization by Date Range Customer Class, Remarks, Invoiced, and Hidden
				'------------------------------------------------------------------------------------------------	
				
				If OCSWebOrderOrMDSInvoice = "OCS" Then
					WHERE_CLAUSE_INVOICEDATERANGE = " WHERE (OCSAccessOrderDate BETWEEN '" & PeriodStartDate & "' AND '" & PeriodEndDate & "') "
				ElseIf OCSWebOrderOrMDSInvoice = "MDS" Then
					WHERE_CLAUSE_INVOICEDATERANGE = " WHERE (MDSInvoiceDate BETWEEN '" & PeriodStartDate & "' AND '" & PeriodEndDate & "') "
				Else
					WHERE_CLAUSE_INVOICEDATERANGE = " WHERE (OCSAccessOrderDate BETWEEN '" & PeriodStartDate & "' AND '" & PeriodEndDate & "') "
				End If	
				
				If WHERE_CLAUSE_INVOICEDATERANGE <> "" Then SQLInvoiceGrandTotals = SQLInvoiceGrandTotals & WHERE_CLAUSE_INVOICEDATERANGE
				If WHERE_CLAUSE_CUSTCLASS <> "" Then SQLInvoiceGrandTotals = SQLInvoiceGrandTotals & WHERE_CLAUSE_CUSTCLASS
				If WHERE_CLAUSE_REMARKS <> "" Then SQLInvoiceGrandTotals = SQLInvoiceGrandTotals & WHERE_CLAUSE_REMARKS
				If WHERE_CLAUSE_HIDDEN <> "" Then SQLInvoiceGrandTotals = SQLInvoiceGrandTotals & WHERE_CLAUSE_HIDDEN
				If WHERE_CLAUSE_INVOICED <> "" Then SQLInvoiceGrandTotals = SQLInvoiceGrandTotals & WHERE_CLAUSE_INVOICED
				
				'---------------------------------------------------------------------------
				'End ORDER BY clauses
				'---------------------------------------------------------------------------
				
				'Response.Write("<br>" & SQLInvoiceGrandTotals & "<br>")
				
				Set rsSQLInvoiceGrandTotals = cnnInvoiceGrandTotals.Execute(SQLInvoiceGrandTotals)
				
				If NOT rsSQLInvoiceGrandTotals.EOF Then
					GrandTotNumWebOrders = rsSQLInvoiceGrandTotals("GrandTotNumWebOrders")
					GrandTotWebSales = rsSQLInvoiceGrandTotals("GrandTotWebSales")
					GrandTotMDSInvoiceAmt = rsSQLInvoiceGrandTotals("GrandTotMDSInvoiceAmt") 
				End If
				
				If GrandTotWebSales <> "" Then
					GrandTotWebSales = FormatCurrency(GrandTotWebSales,2)
				End If
				If GrandTotMDSInvoiceAmt <> "" Then
					GrandTotMDSInvoiceAmt = FormatCurrency(GrandTotMDSInvoiceAmt,2)
				End If
				
				'---------------------------------------------------------------------------
				'End MDS Invoice Count
				'---------------------------------------------------------------------------

				If GrandTotWebSales > 0 Then
					If GrandTotMDSInvoiceAmt >= GrandTotWebSales Then
						GrandTotFulfillmentRate = FormatPercent(1,2)					
					Else
						GrandTotFulfillmentRate = FormatPercent(GrandTotMDSInvoiceAmt/GrandTotWebSales,2)
					End If
				Else
					GrandTotFulfillmentRate = FormatPercent(0,2)
				End If
	        
	        %>
			<tr>
		   		<td>&nbsp;</td>
			    <td>&nbsp;</td> 
			    <td>&nbsp;</td>                 
			    <td><strong><%= GrandTotNumWebOrders %></strong></td>
			    <td><strong><%= GrandTotNumMDSInvoices %></strong></td>
			    <td><strong><%= GrandTotWebSales %></strong></td>
			    <td><strong><%= GrandTotMDSInvoiceAmt %></strong></td>
			    <td><strong><%= GrandTotFulfillmentRate %></strong></td>
		    </tr>
		</tfoot>
		<%
		
		Response.Write("</tbody>")
		Response.Write("</table>")		
		Response.Write("</div>")
%>


    </table>
            
            
</div>
<!-- eof responsive tables !-->



<!-- eof row !-->

<!-- row !-->
<div class="row">
<div class="col-lg-12"><hr></div>
</div>
<!-- eof row !-->

<!-- row !-->
<div class="row">
</div>
<!-- eof row !-->

<!--#include file="../../../inc/footer-main.asp"-->