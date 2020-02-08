<%
Server.ScriptTimeout = 900000 'Default value
Dim ReportNumber 
ReportNumber = 1500


	'****************************************
	'First Day In Current Month Function
	'****************************************
			
		Function CustomDate(dtm)
		    Dim y, m, d, h, n, s
		        y = Year(dtm)
		        m = Right("0" & Month(dtm), 2)
		        d = Right("0" & Day(dtm), 2)
		        h = Right("0" & Hour(dtm), 2)
		        n = Right("0" & Minute(dtm), 2)
		        s = Right("0" & Second(dtm), 2)
		    CustomDate = y & "-" & m & "-" & d
		    If h + n + s > 0 Then
		        CustomDate = CustomDate & " "
		        CustomDate = CustomDate & h & ":" & n & ":" & s
		    End If
		End Function

	'****************************************
	'END First Day In Current Month Function
	'****************************************

%>
<!--#include file="../inc/header.asp"-->

<!--#include file="../inc/jquery_table_search.asp"-->
<!--#include file="../inc/InSightFuncs_BizIntel.asp"-->
 	
	
	<style>
	.beatpicker-clear{
		display: block;
		text-indent:-9999em;
		line-height: 0;
		visibility: hidden;
	}
	
	.form-control[disabled], .form-control[readonly], fieldset[disabled] .form-control{
		background-color:#fff;
		border: 1px solid #eee;
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
	
	.top-grey-table{
		background-color:#f8f9fa;
		border-top: 1px solid #ccc;
	}
	
	.top-grey-table tr,td{
		background-color: transparent;
	 }
	 
	 .top-grey-table>tbody>tr>td{
		 border: 0px;
	 }
	 
	.small-date{
		margin-left: 20px;
	} 
	
	.price-volume-net-proof{
		width: 33%;
	}
	
	.period-difference{
		width: 26%
	}
	
	.td-align{
		text-align: right !important;
		width: 100px !important;
	}
	
	.table-size{
		width: 70%;
	}

	table.agenda {
	  /*font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;*/
	  font-family: "Helvetica Neue",Helvetica,Arial,sans-serif;
	  border-collapse: collapse;
	  width: 100%;
	  overflow:scroll;
	}
	
	table.agenda td,
	table.agenda th {
	  border: 1px solid #fff;
	  padding: 8px;
	  text-align: center;
	}
	
	table.agenda th {
	  padding-top: 12px;
	  padding-bottom: 12px;
	  background-color: rgb(193, 212, 174);
	  color: black;
	}

	table.agenda th.customerclass {
	    background: rgb(72, 151, 54);
	    color: white;
	}	
	
	table.agenda th.date {
	    background: #337ab7;
	    color: white;
	    font-weight:normal;
	    width:100px;
	}	
	
	table.agenda tr:nth-child(even) {background: #eee}
	table.agenda tr:nth-child(odd) {background: #FFF}
	
	.border-right{
		border-right: 1px solid #000 !important;
	}

	.positive {
	        /*color: #5cb85c !important;*/
	        color: #009933 !important;
	}
	.negative {
	        color: #d9534f !important;
	}
	.zero {
	        color: #aaa !important;
	}
	
	</style>
	
	<script type="text/javascript">
	$(document).ready(function() {
	    $("#PleaseWaitPanel").hide();
	});
	
	 
	</script>
	
	<%
	Response.Write("<div id=""PleaseWaitPanel"">")
	Response.Write("<br><br>Sales By Day Data<br><br>This may take up to a full minute, please wait...<br><br>")
	Response.Write("<img src=""../img/loading.gif"" />")
	Response.Write("</div>")
	Response.Flush()


	CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Sales By Period Summary View"


	'************************
	'Read Settings_Reports
	'************************
	
	SQL = "SELECT * from Settings_Reports where ReportNumber = 1500 AND UserNo = " & Session("userNo")
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs= cnn8.Execute(SQL)
	If NOT rs.EOF Then

		DefaultSelectedCustomerClassesForSalesReport = rs("ReportSpecificData5")
		InvoiceTypeBackOrder = rs("ReportSpecificData7")
		InvoiceTypeCreditMemo = rs("ReportSpecificData7")
		InvoiceTypeARDebit = rs("ReportSpecificData8")
		InvoiceTypeRental = rs("ReportSpecificData9")
		InvoiceTypeRouteInvoicing = rs("ReportSpecificData10")
		InvoiceTypeInterest = rs("ReportSpecificData11")
		InvoiceTypeTelselInvoicing = rs("ReportSpecificData12")
		
		BasePeriodsRangeIntRecIDs = rs("ReportSpecificData13")
		PeriodsForComparisonRangeIntRecIDs = rs("ReportSpecificData14") 
		
		
		If IsNull(DefaultSelectedCustomerClassesForSalesReport) Then DefaultSelectedCustomerClassesForSalesReport = ""
		If IsNull(InvoiceTypeBackOrder) Then InvoiceTypeBackOrder = ""
		If IsNull(InvoiceTypeCreditMemo) Then InvoiceTypeCreditMemo = ""
		If IsNull(InvoiceTypeARDebit) Then InvoiceTypeARDebit = ""
		If IsNull(InvoiceTypeRental) Then InvoiceTypeRental = ""
		If IsNull(InvoiceTypeRouteInvoicing) Then InvoiceTypeRouteInvoicing = ""
		If IsNull(InvoiceTypeInterest) Then InvoiceTypeInterest = ""
		If IsNull(InvoiceTypeTelselInvoicing) Then InvoiceTypeTelselInvoicing = ""
		
		If IsNull(BasePeriodsRangeIntRecIDs) Then BasePeriodsRangeIntRecIDs = ""
		If IsNull(PeriodsForComparisonRangeIntRecIDs) Then PeriodsForComparisonRangeIntRecIDs = ""
						
	End If										
	'****************************
	'End Read Settings_Reports
	'****************************
	
	'**************************************************************************************
	'Build Page Header From Custome Period Ranges, If Applicable
	'**************************************************************************************

	BasePeriodsRangeIntRecIDsArray = ""
	BasePeriodsRangeIntRecIDsArray = Split(BasePeriodsRangeIntRecIDs,",")

	ComparisonPeriodsRangeIntRecIDsArray = ""
	ComparisonPeriodsRangeIntRecIDsArray = Split(PeriodsForComparisonRangeIntRecIDs,",")
	
	Set cnnCompanyPeriods = Server.CreateObject("ADODB.Connection")
	cnnCompanyPeriods.open (Session("ClientCnnString"))
	Set rsCompanyPeriods = Server.CreateObject("ADODB.Recordset")
	rsCompanyPeriods.CursorLocation = 3 
		
	If UBound(BasePeriodsRangeIntRecIDsArray) > 0 AND UBound(ComparisonPeriodsRangeIntRecIDsArray) > 0 Then
	
		pageHeaderTextBaseVsCompare = ""
		pageHeaderTextBaseVsCompare = pageHeaderTextBaseVsCompare & "<span style='font-size:16px; color:#337ab7; margin-left:40px; '>Currently Viewing:</span><br>"
		
		For z = 0 to UBound(BasePeriodsRangeIntRecIDsArray)
		
			
				SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods WHERE InternalRecordIdentifier = " & BasePeriodsRangeIntRecIDsArray(z)
				Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
						
				If NOT rsCompanyPeriods.EOF Then
				
					periodBeginDateBase = FormatDateTime(rsCompanyPeriods("BeginDate"),2)
					periodEndDateBase = FormatDateTime(rsCompanyPeriods("EndDate"),2)
					periodNumberBase = rsCompanyPeriods("Period")
					periodYearBase = rsCompanyPeriods("Year")
			
				End If

				SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods WHERE InternalRecordIdentifier = " & ComparisonPeriodsRangeIntRecIDsArray(z)
				Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
						
				If NOT rsCompanyPeriods.EOF Then
				
					periodBeginDateComparison = FormatDateTime(rsCompanyPeriods("BeginDate"),2)
					periodEndDateComparison = FormatDateTime(rsCompanyPeriods("EndDate"),2) 
					periodNumberComparison = rsCompanyPeriods("Period")
					periodYearComparison = rsCompanyPeriods("Year")
			
				End If
			
			
			pageHeaderTextBaseVsCompare = pageHeaderTextBaseVsCompare & "<span style='font-size:16px; color:rgb(72, 151, 54); margin-left:40px;'>"
			pageHeaderTextBaseVsCompare = pageHeaderTextBaseVsCompare & "<strong>Period " & periodNumberBase & " of " & periodYearBase & "</strong> (" & periodBeginDateBase & "-" & periodEndDateBase & ") vs <strong>Period " 
			pageHeaderTextBaseVsCompare = pageHeaderTextBaseVsCompare & periodNumberComparison & " of " & periodYearComparison & "</strong> (" & periodBeginDateComparison & "-" & periodEndDateComparison & "<br>" 
			pageHeaderTextBaseVsCompare = pageHeaderTextBaseVsCompare & "</span>"
		Next	
		
		%>
			<h3 class="page-header"><i class="fa fa-graduation-cap"></i> Sales By Period and Customer Class (Summary) For Periods
			<!-- modal button !-->
			<button type="button" class="btn btn-primary" data-toggle="modal" data-target=".bs-example-modal-lg-customize">
			  Customize
			</button>
			<% If SalesByDayReportTableSet(ReportNumber)Then %>
				<a href="<%= BaseURL %>bizintel/SalesByPeriodSummary_Customize_ClearValues.asp"><button type="button" class="btn btn-primary">Clear Customizations</button></a>
			<% End If %>
				<!-- eof modal button !-->
							<br><br>
				<%= pageHeaderTextBaseVsCompare %>

			</h3>
		<%
	End If
	
	
	'***********************************************************************************************
	'Else If The User Did Not Specify Comparison Periods, Use Dates of Current Period/Week to Date
	'***********************************************************************************************
	
	If UBound(BasePeriodsRangeIntRecIDsArray) <= 0 AND UBound(ComparisonPeriodsRangeIntRecIDsArray) <= 0 Then
	
  	  	SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods "
  	  	SQLCompanyPeriods = SQLCompanyPeriods & "WHERE InternalRecordIdentifier <= " & GetLastClosedReportPeriodIntRecID() - 1
  	  	SQLCompanyPeriods = SQLCompanyPeriods & " ORDER BY [Year] DESC, Period DESC"
  	  	
  	  	Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
	
		If NOT rsCompanyPeriods.EOF Then
			periodBeginDateDefault = rsCompanyPeriods("BeginDate")
			periodEndDateDefault = rsCompanyPeriods("EndDate") 
		End If

  	  	SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods "
  	  	SQLCompanyPeriods = SQLCompanyPeriods & "WHERE InternalRecordIdentifier <= " & GetLastClosedReportPeriodIntRecID() - 2
  	  	SQLCompanyPeriods = SQLCompanyPeriods & " ORDER BY [Year] DESC, Period DESC"
  	  	
  	  	Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
	
		If NOT rsCompanyPeriods.EOF Then
			periodBeginDateCompareDefault = rsCompanyPeriods("BeginDate")
			periodEndDateCompareDefault = rsCompanyPeriods("EndDate") 
		End If
		
		%>
			<h3 class="page-header"><i class="fa fa-graduation-cap"></i> Sales By Period and Customer Class (Summary) For <%= FormatDateTime(periodBeginDateDefault,2) %>-<%= FormatDateTime(periodEndDateDefault,2) %>
			vs <%= FormatDateTime(periodBeginDateCompareDefault,2) %>-<%= FormatDateTime(periodEndDateCompareDefault,2) %>
			&nbsp;&nbsp;
			<!-- modal button !-->
			<button type="button" class="btn btn-primary" data-toggle="modal" data-target=".bs-example-modal-lg-customize">
			  Customize
			</button>
			<% If SalesByDayReportTableSet(ReportNumber)Then %>
				<a href="<%= BaseURL %>bizintel/SalesByPeriodDetail_Customize_ClearValues.asp"><button type="button" class="btn btn-primary">Clear Customizations</button></a>
			<% End If %>
				<!-- eof modal button !-->
			</h3>
		<%
		
	End If
	
	'**************************************************************************************
	'If Customer Class is empty from report settings, obtain all customer
	'classes from AR_CustomerClass
	'**************************************************************************************
	
	CustomerClassArray = ""
	CustomerClassArray = Split(DefaultSelectedCustomerClassesForSalesReport,",")

	If UBound(CustomerClassArray) <= 0 Then
	
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
			WHERE_CLAUSE_CUSTCLASS = WHERE_CLAUSE_CUSTCLASS & " AND (ClassCode = '" & CustomerClassArray(z) & "'"
		Else
			WHERE_CLAUSE_CUSTCLASS = WHERE_CLAUSE_CUSTCLASS & " OR ClassCode = '" & CustomerClassArray(z) & "'"
		End If
	Next	
	
	If WHERE_CLAUSE_CUSTCLASS <> "" Then
		WHERE_CLAUSE_CUSTCLASS = WHERE_CLAUSE_CUSTCLASS & ") "
	End IF
	

	
	'**************************************************************************************
	'Build WHERE Clause For Invoice Type
	'**************************************************************************************
	
	WHERE_CLAUSE_IVSTYPE = ""
	
	If InvoiceTypeBackOrder = "B" Then
		If	WHERE_CLAUSE_IVSTYPE = "" Then
			WHERE_CLAUSE_IVSTYPE = " AND (IvsType = 'B'"
		End If
	End If

	If InvoiceTypeCreditMemo = "C" Then
		If	WHERE_CLAUSE_IVSTYPE = "" Then
			WHERE_CLAUSE_IVSTYPE = " AND (IvsType = 'C'"
		Else
			WHERE_CLAUSE_IVSTYPE = WHERE_CLAUSE_IVSTYPE & " OR IvsType = 'C'"
		End If
	End If

	If InvoiceTypeARDebit = "E" Then
		If	WHERE_CLAUSE_IVSTYPE = "" Then
			WHERE_CLAUSE_IVSTYPE = " AND (IvsType = 'E'"
		Else
			WHERE_CLAUSE_IVSTYPE = WHERE_CLAUSE_IVSTYPE & " OR IvsType = 'E'"
		End If
	End If

	If InvoiceTypeRental = "G" Then
		If	WHERE_CLAUSE_IVSTYPE = "" Then
			WHERE_CLAUSE_IVSTYPE = " AND (IvsType = 'G'"
		Else
			WHERE_CLAUSE_IVSTYPE = WHERE_CLAUSE_IVSTYPE & " OR IvsType = 'G'"
		End If
	End If

	If InvoiceTypeRouteInvoicing = "I" Then
		If	WHERE_CLAUSE_IVSTYPE = "" Then
			WHERE_CLAUSE_IVSTYPE = " AND (IvsType = 'I'"
		Else
			WHERE_CLAUSE_IVSTYPE = WHERE_CLAUSE_IVSTYPE & " OR IvsType = 'I'"
		End If
	End If

	If InvoiceTypeInterest = "O" Then
		If	WHERE_CLAUSE_IVSTYPE = "" Then
			WHERE_CLAUSE_IVSTYPE = " AND (IvsType = 'O'"
		Else
			WHERE_CLAUSE_IVSTYPE = WHERE_CLAUSE_IVSTYPE & " OR IvsType = 'O'"
		End If
	End If

	If InvoiceTypeTelselInvoicing = "T" Then
		If	WHERE_CLAUSE_IVSTYPE = "" Then
			WHERE_CLAUSE_IVSTYPE = " AND (IvsType = 'T'"
		Else
			WHERE_CLAUSE_IVSTYPE = WHERE_CLAUSE_IVSTYPE & " OR IvsType = 'T'"
		End If
	End If

	If WHERE_CLAUSE_IVSTYPE <> "" Then
		WHERE_CLAUSE_IVSTYPE = WHERE_CLAUSE_IVSTYPE & ") "
	End IF
	
	
	%>
	
	<!--#include file="SalesByPeriodSummary_Customize.asp"-->	
	 
	
	<div class="row">
	
	<h6 class="page-header">
	<table id="table-search" class='table table-striped table-condensed table-hover display top-grey-table'>
			<tr>
				<%
				For z = 0 to UBound(CustomerClassArray)
					currentClass = cStr(CustomerClassArray(z))
					%><td>Customer Class <%= currentClass %> - <%= GetCustomerClassNameByID(currentClass) %></td><%
				Next
				
				If InvoiceTypeBackOrder = "B" Then
					%><td>Invoice Type: <strong>BACKORDER</strong><br></td><%
				End If
			
				If InvoiceTypeCreditMemo = "C" Then
					%><td>Invoice Type: <strong>CREDIT MEMO</strong><br></td><%
				End If
			
				If InvoiceTypeARDebit = "E" Then
					%><td>Invoice Type: <strong>AR DEBIT</strong><br></td><%
				End If
			
				If InvoiceTypeRental = "G" Then
					%><td>Invoice Type: <strong>RENTAL</strong><br></td><%
				End If
			
				If InvoiceTypeRouteInvoicing = "I" Then
					%><td>Invoice Type: <strong>ROUTE INVOICING</strong><br></td><%
				End If
			
				If InvoiceTypeInterest = "O" Then
					%><td>Invoice Type: <strong>INTEREST</strong><br></td><%
				End If
			
				If InvoiceTypeTelselInvoicing = "T" Then
					%><td>Invoice Type: <strong>TELSEL INVOICING</strong><br></td><%
				End If
			%></tr>
		</table>
	</h6>
	</div>
	
	 <div class="row">
	 
	<%
	
	'**************************************************************************************
	'Build SQL STMT To Select Date From BI_DailySalesByTypeByClass 
	'**************************************************************************************	
	
	'////////////////////////////////////////////////////////////////////////////////////////////
	'CREATE DEFAULT BASE AND COMPARISON PERIODS IF NO CUSTOMIZATION IS SET
	'////////////////////////////////////////////////////////////////////////////////////////////
	
	If UBound(BasePeriodsRangeIntRecIDsArray) <= 0 AND UBound(ComparisonPeriodsRangeIntRecIDsArray) <= 0 Then
	
      	SQL = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".Settings_CompanyPeriods "
  	  	SQL = SQL & "WHERE InternalRecordIdentifier <= " & GetLastClosedReportPeriodIntRecID() - 1
  	  	SQL = SQL & " ORDER BY [Year] DESC, Period DESC"

		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
		Set rs = cnn8.Execute(SQL)
	
		If not rs.EOF Then
			BaseInternalRecordIdentifier = rs("InternalRecordIdentifier")
			ComparisonInternalRecordIdentifier = BaseInternalRecordIdentifier - 1
		End If

		ReDim BasePeriodsRangeIntRecIDsArray(0)
		ReDim ComparisonPeriodsRangeIntRecIDsArray(0)
		
		BasePeriodsRangeIntRecIDsArray(0) = BaseInternalRecordIdentifier
		ComparisonPeriodsRangeIntRecIDsArray(0) = ComparisonInternalRecordIdentifier
			
	End If

	'////////////////////////////////////////////////////////////////////////////////////////////
	'START LOOPING THROUGH THE ARRAY OF BASE PERIODS
	'FOR EACH BASE PERIOD, CREATE ONE TABLE ROW THAT CONSISTS OF EACH CUSTOMER TYPE AND 
	'COMPARES THAT PERIOD TO THE SAME PERIOD IN THE ARRAY OF COMPARISON PERIODS
	'////////////////////////////////////////////////////////////////////////////////////////////
	
	
	CurrentBaseIntRecID = -1

	For y = 0 to UBound(BasePeriodsRangeIntRecIDsArray)
	
		SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods WHERE InternalRecordIdentifier = " & BasePeriodsRangeIntRecIDsArray(y)
		Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
				
		If NOT rsCompanyPeriods.EOF Then
			periodBeginDateBase = rsCompanyPeriods("BeginDate")
			periodEndDateBase = rsCompanyPeriods("EndDate") 
			periodNumBase = rsCompanyPeriods("Period")
			periodYearBase = rsCompanyPeriods("Year")
		End If

		SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods WHERE InternalRecordIdentifier = " & ComparisonPeriodsRangeIntRecIDsArray(y)
		Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
				
		If NOT rsCompanyPeriods.EOF Then
			periodBeginDateCompare = rsCompanyPeriods("BeginDate")
			periodEndDateCompare = rsCompanyPeriods("EndDate") 
			periodNumCompare = rsCompanyPeriods("Period")
			periodYearCompare = rsCompanyPeriods("Year")
		End If
		
	
		%>
		
		<% If y = 0 Then %>			
			<div class="table-responsive">
				<table class="agenda">
				  <thead>
				    <tr>
				    	<th rowspan="2">Date/Period</th>
				    	<%
						For z = 0 to UBound(CustomerClassArray)
							%><th colspan="3" class="customerclass"><%= CustomerClassArray(z) %> - <%= GetCustomerClassNameByID(CustomerClassArray(z)) %></th><%
						Next
				    	%>
				    </tr>
				    <tr>
				    	<%
						For z = 0 to UBound(CustomerClassArray)
							%>
					        <th>Total Sales $</th>
					        <th>Gross Profit %</th>
					        <th class="border-right">+/- % Comparison Period</th>
							<%
						Next
				    	%>
				    </tr>
				  </thead>
				  <tbody>
			<% End If %>
			<%
				

			'////////////////////////////////////////////////////////////////////////////////////////////
			'DETERMINE WHEN TO START A NEW ROW IN THE TABLE
			'A NEW ROW SHOULD START WHEN BOTH THE INTERNAL RECORD IDENTIFIER FOR THE BASE PERIOD
			'HAS ADVANCED TO THE NEXT "Y" VALUE IN THE FOR/NEXT LOOP
			'////////////////////////////////////////////////////////////////////////////////////////////	
								
			If cInt(CurrentBaseIntRecID) <> cInt(y) Then %>
				<tr>
					<th class="date"><%= FormatDateTime(periodBeginDateBase,2) %>-<%= FormatDateTime(periodEndDateBase,2) %> (P<%= periodNumBase %>-<%= periodYearBase %>) 
									<br>vs<br> <%= FormatDateTime(periodBeginDateCompare,2) %>-<%= FormatDateTime(periodEndDateCompare,2) %>  (P<%= periodNumCompare %>-<%= periodYearCompare %>)</th>
					<%  	
					CurrentBaseIntRecID = y
			End If 
			
			For z = 0 to UBound(CustomerClassArray)	
			

				CurrentRecordClassCode = ""
				SelectedBasePeriod_GrossProfitPct = 0
				SelectedComparePeriod_GrossProfitPct = 0
				PlusMinusComparePeriodPct = 0
				PlusMinusComparePeriodDlrs = 0
			
				CurrentRecordClassCode = CustomerClassArray(z)
		
				'***************************************************************************************************
				'Build SQL STMT To Get Periods Sales Data For The BASE Period's Date Range
				'Note: You Need To Run a SQL Statement To Get The Start and End Dates For The Base Period
				'***************************************************************************************************	

				SQL = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost FROM BI_DailySalesByTypeByClass "
				SQL = SQL & " WHERE (ivsDate BETWEEN '" & periodBeginDateBase & "' AND '" & periodEndDateBase & "') AND "
				SQL = SQL & " ClassCode = '" & CurrentRecordClassCode & "' "
				'SQL = SQL & " GROUP BY IvsDate, ClassCode ORDER BY ivsdate,ClassCode"
				If WHERE_CLAUSE_IVSTYPE <> "" Then SQL = SQL & WHERE_CLAUSE_IVSTYPE
			 	
			
				'Response.Write("<strong>SQL For Base Period Loop:</strong> " & SQL & "<br>")
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3
				rs.Open SQL, Session("ClientCnnString")

				'*****************************************************************************************************************************************
				'*****************************************************************************************************************************************


				If rs("TotCost") <> 0 Then
					SelectedBasePeriod_GrossProfitPct = Round(((rs("TotSales") - rs("TotCost")) / rs("TotSales")) * 100,2)
				ElseIf rs("TotCost") = 0 AND rs("TotSales") = 0 Then
					SelectedBasePeriod_GrossProfitPct = 0
				Else
					SelectedBasePeriod_GrossProfitPct = 100
				End If
				
				Set cnnGetSalesDataComparePeriod = Server.CreateObject("ADODB.Connection")
				cnnGetSalesDataComparePeriod.open Session("ClientCnnString")
		
				'***************************************************************************************************
				'Build SQL STMT To Get Compare Periods Sales Data For The Compare Period's Date Range
				'Note: You Need To Run a SQL Statement To Get The Start and End Dates For The Comparison Period
				'***************************************************************************************************	
				
				
				
				SQLCompanyPeriods2 = "SELECT * FROM Settings_CompanyPeriods WHERE InternalRecordIdentifier = " & ComparisonPeriodsRangeIntRecIDsArray(y)
				Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods2)
						
				If NOT rsCompanyPeriods.EOF Then
					periodBeginDateCompare= rsCompanyPeriods("BeginDate")
					periodEndDateCompare = rsCompanyPeriods("EndDate") 
				End If
		
				SQLGetSalesDataComparePeriod = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost FROM BI_DailySalesByTypeByClass "
				SQLGetSalesDataComparePeriod = SQLGetSalesDataComparePeriod & " WHERE (ivsDate BETWEEN '" & periodBeginDateCompare & "' AND '" & periodEndDateCompare & "') AND "
				SQLGetSalesDataComparePeriod = SQLGetSalesDataComparePeriod & " ClassCode = '" & CurrentRecordClassCode & "' "
				If WHERE_CLAUSE_IVSTYPE <> "" Then SQLGetSalesDataComparePeriod = SQLGetSalesDataComparePeriod & WHERE_CLAUSE_IVSTYPE
					
				'**************************************************************************************
				'END Build SQL STMT To Select Date Range From BI_DailySalesByTypeByClass 
				'**************************************************************************************	
				
				'Response.write("<strong>SQLGetSalesDataComparePeriod </strong>: " & SQLGetSalesDataComparePeriod & "<br>")
										 
				Set rsGetSalesDataComparePeriod = Server.CreateObject("ADODB.Recordset")
				rsGetSalesDataComparePeriod.CursorLocation = 3 
				Set rsGetSalesDataComparePeriod= cnnGetSalesDataComparePeriod.Execute(SQLGetSalesDataComparePeriod)
				
				If NOT rsGetSalesDataComparePeriod.EOF then 
						
					If rsGetSalesDataComparePeriod("TotCost") <> 0 Then
						SelectedComparePeriod_GrossProfitPct = Round(((rsGetSalesDataComparePeriod("TotSales") - rsGetSalesDataComparePeriod("TotCost")) / rsGetSalesDataComparePeriod("TotSales")) * 100,2)
					ElseIf rsGetSalesDataComparePeriod("TotCost") = 0 AND rsGetSalesDataComparePeriod("TotSales") = 0 Then
						SelectedComparePeriod_GrossProfitPct = 0
					Else
						SelectedComparePeriod_GrossProfitPct = 100
					End If
				End If
					
				If SelectedBasePeriod_GrossProfitPct = 0 Then
					PlusMinusComparePeriodPct = SelectedComparePeriod_GrossProfitPct * (-1)
				Else
					PlusMinusComparePeriodPct = SelectedBasePeriod_GrossProfitPct - SelectedComparePeriod_GrossProfitPct
				End If
				
				PlusMinusComparePeriodDlrs = rs("TotSales") - rsGetSalesDataComparePeriod("TotSales")
				
				%>
				
				<% If rs("TotSales") = 0 Then %>
					<td class="td-align zero"><%= FormatCurrency(rs("TotSales"),2,-2,-1)%></td> 
				<% ElseIf rs("TotSales") > 0 Then %>
					<td class="td-align positive"><%= FormatCurrency(rs("TotSales"),2,-2,-1)%></td>
				<% ElseIf rs("TotSales") < 0 Then %>
					<td class="td-align negative"><%= FormatCurrency(rs("TotSales"),2,-2,-1)%></td>
				<% End If %>
								
				<% If SelectedBasePeriod_GrossProfitPct = 0 Then %>
					<td class="td-align zero"><%= FormatNumber(SelectedBasePeriod_GrossProfitPct,2) %>%</td> 
				<% ElseIf SelectedBasePeriod_GrossProfitPct > 0 Then %>
					<td class="td-align positive"><%= FormatNumber(SelectedBasePeriod_GrossProfitPct,2) %>%</td>
				<% ElseIf SelectedBasePeriod_GrossProfitPct < 0 Then %>
					<td class="td-align negative"><%= FormatNumber(SelectedBasePeriod_GrossProfitPct,2) %>%</td>
				<% End If %>
				
				<% If PlusMinusComparePeriodPct = 0 Then %>
					<td class="td-align zero border-right"><%= FormatNumber(PlusMinusComparePeriodPct,2) %>%</td> 
				<% ElseIf PlusMinusComparePeriodPct > 0 Then %>
					<td class="td-align positive border-right"><%= FormatNumber(PlusMinusComparePeriodPct,2) %>%</td>
				<% ElseIf PlusMinusComparePeriodPct < 0 Then %>
					<td class="td-align negative border-right"><%= FormatNumber(PlusMinusComparePeriodPct,2) %>%</td>
				<% End If %>
				

			<%
	 		Next '''''(Customer Class Loop)
		 	
		Next ''''(Base Period Loop)
		
		rsGetSalesDataComparePeriod.Close
		set rsGetSalesDataComparePeriod= Nothing
		cnnGetSalesDataComparePeriod.Close	
		set cnnGetSalesDataComparePeriod = Nothing
		
		'****************************************************************************************************************
		'OUTER LOOPS ARE COMPLETE, NOW PREPARE TO WRITE THE LAST ROW OF THE TABLE THAT GIVES MASTER SUMMARY TOTALS
		'****************************************************************************************************************
				
			%>
				
			</tr>

		    <tr>
		    	<th rowspan="2">&nbsp;</th>
		    	<%
				For z = 0 to UBound(CustomerClassArray)
					%><th colspan="3" class="customerclass"><%= CustomerClassArray(z) %> - <%= GetCustomerClassNameByID(CustomerClassArray(z)) %></th><%
				Next
		    	%>
		    </tr>
		    <tr>
		    	<%
				For z = 0 to UBound(CustomerClassArray)
					%>
				        <th>Total Sales $</th>
				        <th>Gross Profit %</th>
				        <th class="border-right">+/- % Comparison Period</th>
					<%
				Next
		    	%>
		    </tr>
		
			    <tr>
			    	<th class="date">TOTALS</th>
			    	<%
					For z = 0 to UBound(CustomerClassArray)
	
	
			    		Set cnnGetTotalsForMasterSummary = Server.CreateObject("ADODB.Connection")
						cnnGetTotalsForMasterSummary.open Session("ClientCnnString")
					
						'*********************************************************************************************************
						'Get Summary of Total Sales For This Customer Class For Current Base Period
						'*********************************************************************************************************
						
						WHERE_CLAUSE_BASE_PERIOD_MASTER = ""
						WHERE_CLAUSE_COMPARISON_PERIOD_MASTER = ""
						
						'*******************************************************************
						'If Periods Were Selected By The User, Loop and Build WHERE Clause
						'*******************************************************************
					
							For i = 0 to UBound(BasePeriodsRangeIntRecIDsArray)
							
								If i = 0 Then
								
									SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods WHERE InternalRecordIdentifier = " & BasePeriodsRangeIntRecIDsArray(i)
									Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
											
									If NOT rsCompanyPeriods.EOF Then
									
										periodBeginDate = rsCompanyPeriods("BeginDate")
										periodEndDate = rsCompanyPeriods("EndDate") 
										periodNumber = rsCompanyPeriods("Period")
										periodYear = rsCompanyPeriods("Year")
								
										WHERE_CLAUSE_BASE_PERIOD_MASTER = WHERE_CLAUSE_BASE_PERIOD_MASTER & " ((ivsDate BETWEEN '" & periodBeginDate & "' AND '" & periodEndDate & "') "
									End If
								Else
								
									SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods WHERE InternalRecordIdentifier = " & BasePeriodsRangeIntRecIDsArray(i)
									Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
											
									If NOT rsCompanyPeriods.EOF Then
									
										periodBeginDate = rsCompanyPeriods("BeginDate")
										periodEndDate = rsCompanyPeriods("EndDate") 
										periodNumber = rsCompanyPeriods("Period")
										periodYear = rsCompanyPeriods("Year")
								
										WHERE_CLAUSE_BASE_PERIOD_MASTER = WHERE_CLAUSE_BASE_PERIOD_MASTER & " OR (ivsDate BETWEEN '" & periodBeginDate & "' AND '" & periodEndDate & "') "
										
									End If
						
								End If
								
							Next	
					
							IF WHERE_CLAUSE_BASE_PERIOD_MASTER <> "" Then WHERE_CLAUSE_BASE_PERIOD_MASTER =  WHERE_CLAUSE_BASE_PERIOD_MASTER & ") "
	
						
					'*******************************************************************
					'If Periods Were Selected By The User, Loop and Build WHERE Clause
					'*******************************************************************
					
						For i = 0 to UBound(ComparisonPeriodsRangeIntRecIDsArray)
						
							If i = 0 Then
							
								SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods WHERE InternalRecordIdentifier = " & ComparisonPeriodsRangeIntRecIDsArray(i)
								Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
										
								If NOT rsCompanyPeriods.EOF Then
								
									periodBeginDate = rsCompanyPeriods("BeginDate")
									periodEndDate = rsCompanyPeriods("EndDate") 
									periodNumber = rsCompanyPeriods("Period")
									periodYear = rsCompanyPeriods("Year")
							
									WHERE_CLAUSE_COMPARISON_PERIOD_MASTER = WHERE_CLAUSE_COMPARISON_PERIOD_MASTER & " (ivsDate BETWEEN '" & periodBeginDate & "' AND '" & periodEndDate & "') "
									
								End If
							Else
							
								SQLCompanyPeriods = "SELECT * FROM Settings_CompanyPeriods WHERE InternalRecordIdentifier = " & ComparisonPeriodsRangeIntRecIDsArray(i)
								Set rsCompanyPeriods = cnnCompanyPeriods.Execute(SQLCompanyPeriods)
										
								If NOT rsCompanyPeriods.EOF Then
								
									periodBeginDate = rsCompanyPeriods("BeginDate")
									periodEndDate = rsCompanyPeriods("EndDate") 
									periodNumber = rsCompanyPeriods("Period")
									periodYear = rsCompanyPeriods("Year")
							
									WHERE_CLAUSE_COMPARISON_PERIOD_MASTER = WHERE_CLAUSE_COMPARISON_PERIOD_MASTER & " OR (ivsDate BETWEEN '" & periodBeginDate & "' AND '" & periodEndDate & "') "
									
								End If
					
							End If
							

						Next	

						
						'******************************
						'BUILD SQL STMT
						'******************************
						
						SQLGetTotalsForMasterSummary = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost FROM BI_DailySalesByTypeByClass WHERE "

						If WHERE_CLAUSE_BASE_PERIOD_MASTER <> "" Then SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & WHERE_CLAUSE_BASE_PERIOD_MASTER
						
						If WHERE_CLAUSE_IVSTYPE <> "" Then SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & WHERE_CLAUSE_IVSTYPE
						
					 	SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " AND ClassCode = '" & CustomerClassArray(z) & "'"
						
						'Response.write("<strong>SQLGetTotalsForMasterSummary Base Period</strong> : " & SQLGetTotalsForMasterSummary & "<br>")
						'******************************
						'END BUILD SQL STMT
						'******************************
						
			 
						Set rsGetTotalsForMasterSummary = Server.CreateObject("ADODB.Recordset")
						rsGetTotalsForMasterSummary.CursorLocation = 3 
						Set rsGetTotalsForMasterSummary= cnnGetTotalsForMasterSummary.Execute(SQLGetTotalsForMasterSummary)
						
						
						If NOT rsGetTotalsForMasterSummary.EOF Then 
						
							TotalSalesBasePeriod_GrandTotal = rsGetTotalsForMasterSummary("TotSales")
							TotalCostBasePeriod_GrandTotal = rsGetTotalsForMasterSummary("TotCost")
		
							If TotalCostBasePeriod_GrandTotal <> 0 Then
								GrossProfitBasePeriod_GrandTotal = Round(((TotalSalesBasePeriod_GrandTotal - TotalCostBasePeriod_GrandTotal) / TotalSalesBasePeriod_GrandTotal) * 100,2)
							ElseIf TotalCostBasePeriod_GrandTotal = 0 AND TotalSalesBasePeriod_GrandTotal = 0 Then
								GrossProfitBasePeriod_GrandTotal = 0
							Else
								GrossProfitBasePeriod_GrandTotal = 100
							End If
	
						End If
							
						'*********************************************************************************************************
						'Get Summary of Total Sales For This Customer Class For Comparison Period
						'*********************************************************************************************************
	
						'******************************
						'BUILD SQL STMT
						'******************************
													
						SQLGetTotalsForMasterSummary = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost FROM BI_DailySalesByTypeByClass WHERE "
						
						If WHERE_CLAUSE_COMPARISON_PERIOD_MASTER <> "" Then SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & WHERE_CLAUSE_COMPARISON_PERIOD_MASTER
						
						If WHERE_CLAUSE_IVSTYPE <> "" Then SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & WHERE_CLAUSE_IVSTYPE
						
					 	SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " AND ClassCode = '" & CustomerClassArray(z) & "'"
					 	
					 	'Response.write("<strong>SQLGetTotalsForMasterSummary Compare Period</strong> : " & SQLGetTotalsForMasterSummary & "<br>")
					
						'******************************
						'END BUILD SQL STMT
						'******************************
							
						Set rsGetTotalsForMasterSummary= cnnGetTotalsForMasterSummary.Execute(SQLGetTotalsForMasterSummary)
						
						
						If NOT rsGetTotalsForMasterSummary.EOF Then 
						
							TotalSalesComparePeriod_GrandTotal = rsGetTotalsForMasterSummary("TotSales")
							TotalCostComparePeriod_GrandTotal = rsGetTotalsForMasterSummary("TotCost")
		
							If TotalCostComparePeriod_GrandTotal <> 0 Then
								GrossProfitComparePeriod_GrandTotal = Round(((TotalSalesComparePeriod_GrandTotal - TotalCostComparePeriod_GrandTotal) / TotalSalesComparePeriod_GrandTotal) * 100,2)
							ElseIf TotalCostComparePeriod_GrandTotal = 0 AND TotalSalesComparePeriod_GrandTotal = 0 Then
								GrossProfitComparePeriod_GrandTotal = 0
							Else
								GrossProfitComparePeriod_GrandTotal = 100
							End If
	
						End If
						
							
						rsGetTotalsForMasterSummary.Close
						set rsGetTotalsForMasterSummary= Nothing
						cnnGetTotalsForMasterSummary.Close	
						set cnnGetTotalsForMasterSummary = Nothing
						
				
						PlusMinusComparePeriodDlrs_GrandTotal = TotalSalesBasePeriod_GrandTotal - TotalSalesComparePeriod_GrandTotal
						
						If GrossProfitBasePeriod_GrandTotal = 0 Then
							PlusMinusComparePeriodPct_GrandTotal = GrossProfitComparePeriod_GrandTotal * (-1)
						Else
							PlusMinusComparePeriodPct_GrandTotal = GrossProfitBasePeriod_GrandTotal - GrossProfitComparePeriod_GrandTotal
						End If
								
	
			    	
				    	%>
				    	
						<% If TotalSalesBasePeriod_GrandTotal = 0 Then %>
							<td class="td-align zero"><strong><%= FormatCurrency(TotalSalesBasePeriod_GrandTotal,2,-2,-1)%></strong></td> 
						<% ElseIf TotalSalesBasePeriod_GrandTotal > 0 Then %>
							<td class="td-align positive"><strong><%= FormatCurrency(TotalSalesBasePeriod_GrandTotal,2,-2,-1)%></strong></td>
						<% ElseIf TotalSalesBasePeriod_GrandTotal < 0 Then %>
							<td class="td-align negative"><strong><%= FormatCurrency(TotalSalesBasePeriod_GrandTotal,2,-2,-1)%></strong></td>
						<% End If %>
						
						
						<% If GrossProfitBasePeriod_GrandTotal = 0 Then %>
							<td class="td-align zero"><strong><%= FormatNumber(GrossProfitBasePeriod_GrandTotal,2)%>%</strong></td> 
						<% ElseIf GrossProfitBasePeriod_GrandTotal > 0 Then %>
							<td class="td-align positive"><strong><%= FormatNumber(GrossProfitBasePeriod_GrandTotal,2)%>%</strong></td>
						<% ElseIf GrossProfitBasePeriod_GrandTotal < 0 Then %>
							<td class="td-align negative"><strong><%= FormatNumber(GrossProfitBasePeriod_GrandTotal,2)%>%</strong></td>
						<% End If %>
				    	
						<% If PlusMinusComparePeriodPct_GrandTotal = 0 Then %>
							<td class="td-align zero border-right"><strong><%= FormatNumber(PlusMinusComparePeriodPct_GrandTotal,2)%>%</strong></td> 
						<% ElseIf PlusMinusComparePeriodPct_GrandTotal > 0 Then %>
							<td class="td-align positive border-right"><strong><%= FormatNumber(PlusMinusComparePeriodPct_GrandTotal,2)%>%</strong></td>
						<% ElseIf PlusMinusComparePeriodPct_GrandTotal < 0 Then %>
							<td class="td-align negative border-right"><strong><%= FormatNumber(PlusMinusComparePeriodPct_GrandTotal,2)%>%</strong></td>
						<% End If %>
						
						<%
				Next
		    	%>
			    
			    </tbody>
			</table>
		</div>
		<%

				
	rsCompanyPeriods.Close
	set rsCompanyPeriods = Nothing
	cnnCompanyPeriods.Close	
	set cnnCompanyPeriods = Nothing	
	
		
%>
					
          
</div>
          

<!--#include file="../inc/footer-main.asp"-->