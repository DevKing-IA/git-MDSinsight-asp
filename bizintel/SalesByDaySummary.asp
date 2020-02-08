<%
Server.ScriptTimeout = 900000 'Default value
Dim ReportNumber 
ReportNumber = 1500
%>
<!--#include file="../inc/header.asp"-->

<!--#include file="../inc/jquery_table_search.asp"-->
<!--#include file="../inc/InSightFuncs_BizIntel.asp"-->
 
<%


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
	
	FirstDayOfThisMonth = CustomDate(DateSerial(Year(Date), Month(Date), 1))
	FirstDayOfThisMonthLastYear = CustomDate(DateSerial(Year(DateAdd("yyyy",-1,Date)), Month(DateAdd("yyyy",-1,Date)), 1))


	CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Sales By Day Summary View"

	'************************
	'Read Settings_Reports
	'************************
	SQL = "SELECT * from Settings_Reports where ReportNumber = 1500 AND UserNo = " & Session("userNo")
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs= cnn8.Execute(SQL)
	If NOT rs.EOF Then
		DatesOrPeriods = rs("ReportSpecificData1")
		PeriodBeingEvaluatedCustomize = rs("ReportSpecificData2")
		RangeStartDateCustomize = rs("ReportSpecificData3")
		RangeEndDateCustomize = rs("ReportSpecificData4")
		DefaultSelectedCustomerClassesForSalesReport = rs("ReportSpecificData5")
		InvoiceTypeBackOrder = rs("ReportSpecificData7")
		InvoiceTypeCreditMemo = rs("ReportSpecificData7")
		InvoiceTypeARDebit = rs("ReportSpecificData8")
		InvoiceTypeRental = rs("ReportSpecificData9")
		InvoiceTypeRouteInvoicing = rs("ReportSpecificData10")
		InvoiceTypeInterest = rs("ReportSpecificData11")
		InvoiceTypeTelselInvoicing = rs("ReportSpecificData12")
		
		If IsNull(PeriodBeingEvaluatedCustomize) Then PeriodBeingEvaluatedCustomize = ""
		If IsNull(RangeStartDateCustomize) OR RangeStartDateCustomize = "" Then RangeStartDateCustomize = Now()
		If IsNull(RangeEndDateCustomize) OR RangeEndDateCustomize = ""  Then RangeEndDateCustomize = Now()
		If IsNull(DefaultSelectedCustomerClassesForSalesReport) Then DefaultSelectedCustomerClassesForSalesReport = ""
		If IsNull(InvoiceTypeBackOrder) Then InvoiceTypeBackOrder = 0
		If IsNull(InvoiceTypeCreditMemo) Then InvoiceTypeCreditMemo = 0
		If IsNull(InvoiceTypeARDebit) Then InvoiceTypeARDebit = 0
		If IsNull(InvoiceTypeRental) Then InvoiceTypeRental = 0
		If IsNull(InvoiceTypeRouteInvoicing) Then InvoiceTypeRouteInvoicing = 0
		If IsNull(InvoiceTypeInterest) Then InvoiceTypeInterest = 0
		If IsNull(InvoiceTypeTelselInvoicing) Then InvoiceTypeTelselInvoicing = 0
	Else
		PeriodBeingEvaluatedCustomize = ""
		RangeStartDateCustomize = Now()
		RangeEndDateCustomize = Now()
		DefaultSelectedCustomerClassesForSalesReport = ""
		InvoiceTypeBackOrder = 0
		InvoiceTypeCreditMemo = 0
		InvoiceTypeARDebit = 0
		InvoiceTypeRental = 0
		InvoiceTypeRouteInvoicing = 0
		InvoiceTypeInterest = 0
		InvoiceTypeTelselInvoicing = 0
	End If										
	'****************************
	'End Read Settings_Reports
	'****************************

	CustomerClassArray = ""
	CustomerClassArray = Split(DefaultSelectedCustomerClassesForSalesReport,",")
	
	'**************************************************************************************
	'If Customer Class is empty from report settings, obtain all customer
	'classes from AR_CustomerClass
	'**************************************************************************************

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
	
	<!-- date picker !-->
		<link rel="stylesheet" href="<%= baseURL %>css/datepicker/BeatPicker.min.css"/>
		<script src="<%= baseURL %>js/datepicker/BeatPicker.min.js"></script>
	<!-- eof date picker !-->
	
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

	
	If SalesByDayReportTableSet(ReportNumber)Then
	
		If SalesByDayGetDatesOrPeriods(ReportNumber)= "Dates" Then 
			PeriodBeingEvaluated = FormatDateTime(SalesByDayGetRangeStart(ReportNumber),2) & " - " & FormatDateTime(SalesByDayGetRangeEnd(ReportNumber),2)
		ElseIf SalesByDayGetDatesOrPeriods(ReportNumber)= "Periods" Then
		
			PeriodYear = GetPeriodYearByIntRecID(PeriodBeingEvaluatedCustomize)
			PeriodNum = GetPeriodByIntRecID(PeriodBeingEvaluatedCustomize)
			PeriodStartDate = FormatDateTime(GetPeriodBeginDateByIntRecID(PeriodBeingEvaluatedCustomize),2)
			PeriodEndDate = FormatDateTime(GetPeriodEndDateByIntRecID(PeriodBeingEvaluatedCustomize),2)
			
			If cInt(PeriodYear) = 0 AND cInt(Period) = 0 Then
				PeriodBeingEvaluated = FormatDateTime(FirstDayOfThisMonth,2) & " in " & FormatDateTime(Now(),2)
			Else
				PeriodBeingEvaluated = "Period " & PeriodNum & " in " & PeriodYear & " (" & PeriodStartDate & "-" & PeriodEndDate & ")"
			End If
			
		End If 
			
	Else
		PeriodBeingEvaluated = FormatDateTime(FirstDayOfThisMonth,2) & " - " & FormatDateTime(Now(),2)

	End If
	
	%>
	
	<h3 class="page-header"><i class="fa fa-graduation-cap"></i> Sales By Day and Customer Class (Summary) For <%= PeriodBeingEvaluated %>
	&nbsp;&nbsp;
	
	<!-- modal button !-->
	<button type="button" class="btn btn-primary" data-toggle="modal" data-target=".bs-example-modal-lg-customize">
	  Customize
	</button>
	<% If SalesByDayReportTableSet(ReportNumber)Then %>
		<a href="<%= BaseURL %>bizintel/SalesByDaySummary_Customize_ClearValues.asp"><button type="button" class="btn btn-primary">Clear Customizations</button></a>
	<% End If %>
		<!-- eof modal button !-->
	</h3>
	
	<!--#include file="SalesByDaySummary_Customize.asp"-->	
	 
	
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
	
	SQL = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost, IvsDate, ClassCode FROM BI_DailySalesByTypeByClass "
	
	
	If SalesByDayGetDatesOrPeriods(ReportNumber) = "Periods" Then

		PeriodYear = GetPeriodYearByIntRecID(PeriodBeingEvaluatedCustomize)
		Period = GetPeriodByIntRecID(PeriodBeingEvaluatedCustomize)
		
		If cInt(PeriodYear) = 0 AND cInt(Period) = 0 Then
			SQL = SQL & " WHERE ivsDate BETWEEN '" & FirstDayOfThisMonth & "' AND '" & Now() & "' "
		Else
			SQL = SQL & " WHERE (PeriodYear = " & PeriodYear & ") AND (Period = " & Period & ")"
		End If
			
	ElseIf SalesByDayGetDatesOrPeriods(ReportNumber) = "Dates" Then
		SQL = SQL & " WHERE ivsDate BETWEEN '" & SalesByDayGetRangeStart(ReportNumber) & "' AND '" & SalesByDayGetRangeEnd(ReportNumber) & "' "
	Else
		SQL = SQL & " WHERE ivsDate BETWEEN '" & FirstDayOfThisMonth & "' AND '" & Now() & "' "
	End If
	
	
	If WHERE_CLAUSE_CUSTCLASS <> "" Then SQL = SQL & WHERE_CLAUSE_CUSTCLASS
	If WHERE_CLAUSE_IVSTYPE <> "" Then SQL = SQL & WHERE_CLAUSE_IVSTYPE
	
	
 	SQL = SQL & " GROUP BY IvsDate, ClassCode ORDER BY ivsdate,ClassCode"
 	
	'**************************************************************************************
	'END Build SQL STMT To Select Date From BI_DailySalesByTypeByClass 
	'**************************************************************************************	
	
	'Response.Write(SQL & "<br>")
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open SQL, Session("ClientCnnString")

	If NOT rs.EOF Then
	
		CurrentDate = ""
		%>			
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
				        <th class="border-right">+/- % Same Period Last Year</th>
						<%
					Next
			    	%>
			    </tr>
			  </thead>
			  <tbody>
		<%
				
		Do While NOT rs.EOF
					
			%> 
			
			<% If CurrentDate <> rs("ivsdate") Then %>
				<tr>
					<th class="date"><%= rs("ivsdate") %></th>
			<%  CurrentDate = rs("ivsdate")
			End If 	
			
			
			
			Do While CurrentDate = rs("ivsdate") 
			
				LastYear = ""
				ThisDateLastYear = ""
				CurrentRecordClassCode = ""
				SelectedPeriod_GrossProfitPct = 0
				SelectedPeriodLastYear_GrossProfitPct = 0
				PlusMinusSamePeriodLastYearPct = 0
				PlusMinusSamePeriodLastYearDlrs = 0
				
				
				If  CurrentDate = rs("ivsdate") Then
				
					If rs("TotCost") <> 0 Then
						SelectedPeriod_GrossProfitPct = Round(((rs("TotSales") - rs("TotCost")) / rs("TotSales")) * 100,2)
					ElseIf rs("TotCost") = 0 AND rs("TotSales") = 0 Then
						SelectedPeriod_GrossProfitPct = 0
					Else
						SelectedPeriod_GrossProfitPct = 100
					End If
					
					Set cnnGetSalesDataLastYear = Server.CreateObject("ADODB.Connection")
					cnnGetSalesDataLastYear.open Session("ClientCnnString")
				
					'**************************************************************************************
					'Build SQL STMT To Get Last Year's Sales Data For This Date
					'**************************************************************************************	
					
					LastYear = DateAdd("yyyy", -1, rs("ivsdate"))
					ThisDateLastYear = CustomDate(DateSerial(Year(LastYear), Month(rs("ivsdate")), Day(rs("ivsdate"))))
					CurrentRecordClassCode = rs("ClassCode")
					
					SQLGetSalesDataLastYear = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost FROM BI_DailySalesByTypeByClass "
					SQLGetSalesDataLastYear = SQLGetSalesDataLastYear & " WHERE ivsDate = '" & ThisDateLastYear & "' AND "
					SQLGetSalesDataLastYear = SQLGetSalesDataLastYear & " ClassCode = '" & CurrentRecordClassCode & "'"
					
					'**************************************************************************************
					'END Build SQL STMT To Select Date From BI_DailySalesByTypeByClass 
					'**************************************************************************************	
					
					'Response.write("<br>" & SQLGetSalesDataLastYear & "<br>")
											 
					Set rsGetSalesDataLastYear = Server.CreateObject("ADODB.Recordset")
					rsGetSalesDataLastYear.CursorLocation = 3 
					Set rsGetSalesDataLastYear= cnnGetSalesDataLastYear.Execute(SQLGetSalesDataLastYear)
					
					If NOT rsGetSalesDataLastYear.EOF then 
						
						If rsGetSalesDataLastYear("TotCost") <> 0 Then
							SelectedPeriodLastYear_GrossProfitPct = Round(((rsGetSalesDataLastYear("TotSales") - rsGetSalesDataLastYear("TotCost")) / rsGetSalesDataLastYear("TotSales")) * 100,2)
						ElseIf rsGetSalesDataLastYear("TotCost") = 0 AND rsGetSalesDataLastYear("TotSales") = 0 Then
							SelectedPeriodLastYear_GrossProfitPct = 0
						Else
							SelectedPeriodLastYear_GrossProfitPct = 100
						End If
					End If
					
					If SelectedPeriod_GrossProfitPct = 0 Then
						PlusMinusSamePeriodLastYearPct = SelectedPeriodLastYear_GrossProfitPct * (-1)
					Else
						PlusMinusSamePeriodLastYearPct = SelectedPeriod_GrossProfitPct - SelectedPeriodLastYear_GrossProfitPct
					End If
					
					PlusMinusSamePeriodLastYearDlrs = rs("TotSales") - rsGetSalesDataLastYear("TotSales")
					
					%>
					
					<% If rs("TotSales") = 0 Then %>
						<td class="td-align zero"><%= FormatCurrency(rs("TotSales"),2,-2,-1)%></td> 
					<% ElseIf rs("TotSales") > 0 Then %>
						<td class="td-align positive"><%= FormatCurrency(rs("TotSales"),2,-2,-1)%></td>
					<% ElseIf rs("TotSales") < 0 Then %>
						<td class="td-align negative"><%= FormatCurrency(rs("TotSales"),2,-2,-1)%></td>
					<% End If %>
										
					<% If SelectedPeriod_GrossProfitPct = 0 Then %>
						<td class="td-align zero"><%= FormatNumber(SelectedPeriod_GrossProfitPct,2) %>%</td> 
					<% ElseIf SelectedPeriod_GrossProfitPct > 0 Then %>
						<td class="td-align positive"><%= FormatNumber(SelectedPeriod_GrossProfitPct,2) %>%</td>
					<% ElseIf SelectedPeriod_GrossProfitPct < 0 Then %>
						<td class="td-align negative"><%= FormatNumber(SelectedPeriod_GrossProfitPct,2) %>%</td>
					<% End If %>
					
					<% If PlusMinusSamePeriodLastYearPct = 0 Then %>
						<td class="td-align zero border-right"><%= FormatNumber(PlusMinusSamePeriodLastYearPct,2) %>%</td> 
					<% ElseIf PlusMinusSamePeriodLastYearPct > 0 Then %>
						<td class="td-align positive border-right"><%= FormatNumber(PlusMinusSamePeriodLastYearPct,2) %>%</td>
					<% ElseIf PlusMinusSamePeriodLastYearPct < 0 Then %>
						<td class="td-align negative border-right"><%= FormatNumber(PlusMinusSamePeriodLastYearPct,2) %>%</td>
					<% End If %>
					

					<%
					
					rsGetSalesDataLastYear.Close
					set rsGetSalesDataLastYear= Nothing
					cnnGetSalesDataLastYear.Close	
					set cnnGetSalesDataLastYear = Nothing
					
				End If

					
				rs.MoveNext
				If rs.EOF Then EXIT DO
					
			Loop
			%>
				
			</tr>
				
		<%               
		Loop
		
		%>
	
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
				        <th class="border-right">+/- % Same Period Last Year</th>
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
					'Get Summary of Total Sales For This Customer Class For Current Period
					'*********************************************************************************************************
					
					
					'******************************
					'BUILD SQL STMT
					'******************************
						
					SQLGetTotalsForMasterSummary = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost FROM BI_DailySalesByTypeByClass "

					If SalesByDayGetDatesOrPeriods(ReportNumber) = "Periods" Then
						PeriodYear = GetPeriodYearByIntRecID(PeriodBeingEvaluatedCustomize)
						PeriodNum = GetPeriodByIntRecID(PeriodBeingEvaluatedCustomize)
						SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " WHERE (PeriodYear = " & PeriodYear & ") AND (Period = " & PeriodNum & ")"
					ElseIf SalesByDayGetDatesOrPeriods(ReportNumber) = "Dates" Then
						SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " WHERE ivsDate BETWEEN '" & SalesByDayGetRangeStart(ReportNumber) & "' AND '" & SalesByDayGetRangeEnd(ReportNumber) & "' "
					Else
						SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " WHERE ivsDate BETWEEN '" & FirstDayOfThisMonth & "' AND '" & Now() & "' "
					End If
					
					If WHERE_CLAUSE_IVSTYPE <> "" Then SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & WHERE_CLAUSE_IVSTYPE
	
					SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " AND ClassCode = '" & CustomerClassArray(z) & "'"
					
					'******************************
					'END BUILD SQL STMT
					'******************************
					
		 
					Set rsGetTotalsForMasterSummary = Server.CreateObject("ADODB.Recordset")
					rsGetTotalsForMasterSummary.CursorLocation = 3 
					Set rsGetTotalsForMasterSummary= cnnGetTotalsForMasterSummary.Execute(SQLGetTotalsForMasterSummary)
					
					
					If NOT rsGetTotalsForMasterSummary.EOF Then 
					
						TotalSalesThisYear_GrandTotal = rsGetTotalsForMasterSummary("TotSales")
						TotalCostThisYear_GrandTotal = rsGetTotalsForMasterSummary("TotCost")
	
						If TotalCostThisYear_GrandTotal <> 0 Then
							GrossProfitThisYear_GrandTotal = Round(((TotalSalesThisYear_GrandTotal - TotalCostThisYear_GrandTotal) / TotalSalesThisYear_GrandTotal) * 100,2)
						ElseIf TotalCostThisYear_GrandTotal = 0 AND TotalSalesThisYear_GrandTotal = 0 Then
							GrossProfitThisYear_GrandTotal = 0
						Else
							GrossProfitThisYear_GrandTotal = 100
						End If

					End If
						
					'*********************************************************************************************************
					'Get Summary of Total Sales For This Customer Class For This Period LAST YEAR
					'*********************************************************************************************************

					'******************************
					'BUILD SQL STMT
					'******************************
					
					SQLGetTotalsForMasterSummary = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost FROM BI_DailySalesByTypeByClass "

					If SalesByDayGetDatesOrPeriods(ReportNumber) = "Periods" Then
						PeriodYear = GetPeriodYearByIntRecID(PeriodBeingEvaluatedCustomize)
						PeriodNum = GetPeriodByIntRecID(PeriodBeingEvaluatedCustomize)
						SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " WHERE (PeriodYear = " & PeriodYear & ") AND (Period = " & PeriodNum & ")"
					ElseIf SalesByDayGetDatesOrPeriods(ReportNumber) = "Dates" Then
						SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " WHERE ivsDate BETWEEN '" & DateAdd("yyyy", -1,SalesByDayGetRangeStart(ReportNumber)) & "' AND '" & DateAdd("yyyy",-1,SalesByDayGetRangeEnd(ReportNumber)) & "' "
					Else
						SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " WHERE ivsDate BETWEEN '" & FirstDayOfThisMonthLastYear & "' AND '" & DateAdd("yyyy",-1,Now()) & "' "
					End If
					
					If WHERE_CLAUSE_IVSTYPE <> "" Then SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & WHERE_CLAUSE_IVSTYPE
	
					SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " AND ClassCode = '" & CustomerClassArray(z) & "'"
					
					'******************************
					'END BUILD SQL STMT
					'******************************
						
					Set rsGetTotalsForMasterSummary= cnnGetTotalsForMasterSummary.Execute(SQLGetTotalsForMasterSummary)
					
					
					If NOT rsGetTotalsForMasterSummary.EOF Then 
					
						TotalSalesLastYear_GrandTotal = rsGetTotalsForMasterSummary("TotSales")
						TotalCostLastYear_GrandTotal = rsGetTotalsForMasterSummary("TotCost")
	
						If TotalCostLastYear_GrandTotal <> 0 Then
							GrossProfitLastYear_GrandTotal = Round(((TotalSalesLastYear_GrandTotal - TotalCostLastYear_GrandTotal) / TotalSalesLastYear_GrandTotal) * 100,2)
						ElseIf TotalCostLastYear_GrandTotal = 0 AND TotalSalesLastYear_GrandTotal = 0 Then
							GrossProfitLastYear_GrandTotal = 0
						Else
							GrossProfitLastYear_GrandTotal = 100
						End If

					End If
					
						
					rsGetTotalsForMasterSummary.Close
					set rsGetTotalsForMasterSummary= Nothing
					cnnGetTotalsForMasterSummary.Close	
					set cnnGetTotalsForMasterSummary = Nothing
					
					
					
			
					PlusMinusSamePeriodLastYearDlrs_GrandTotal = TotalSalesThisYear_GrandTotal - TotalSalesLastYear_GrandTotal
					
					If GrossProfitThisYear_GrandTotal = 0 Then
						PlusMinusSamePeriodLastYearPct_GrandTotal = GrossProfitLastYear_GrandTotal * (-1)
					Else
						PlusMinusSamePeriodLastYearPct_GrandTotal = GrossProfitThisYear_GrandTotal - GrossProfitLastYear_GrandTotal
					End If
							

		    	
			    	%>
			    	
					<% If TotalSalesThisYear_GrandTotal = 0 Then %>
						<td class="td-align zero"><strong><%= FormatCurrency(TotalSalesThisYear_GrandTotal,2,-2,-1)%></strong></td> 
					<% ElseIf TotalSalesThisYear_GrandTotal > 0 Then %>
						<td class="td-align positive"><strong><%= FormatCurrency(TotalSalesThisYear_GrandTotal,2,-2,-1)%></strong></td>
					<% ElseIf TotalSalesThisYear_GrandTotal < 0 Then %>
						<td class="td-align negative"><strong><%= FormatCurrency(TotalSalesThisYear_GrandTotal,2,-2,-1)%></strong></td>
					<% End If %>
					
					
					<% If GrossProfitThisYear_GrandTotal = 0 Then %>
						<td class="td-align zero"><strong><%= FormatNumber(GrossProfitThisYear_GrandTotal,2)%>%</strong></td> 
					<% ElseIf GrossProfitThisYear_GrandTotal > 0 Then %>
						<td class="td-align positive"><strong><%= FormatNumber(GrossProfitThisYear_GrandTotal,2)%>%</strong></td>
					<% ElseIf GrossProfitThisYear_GrandTotal < 0 Then %>
						<td class="td-align negative"><strong><%= FormatNumber(GrossProfitThisYear_GrandTotal,2)%>%</strong></td>
					<% End If %>
			    	
					<% If PlusMinusSamePeriodLastYearPct_GrandTotal = 0 Then %>
						<td class="td-align zero border-right"><strong><%= FormatNumber(PlusMinusSamePeriodLastYearPct_GrandTotal,2)%>%</strong></td> 
					<% ElseIf PlusMinusSamePeriodLastYearPct_GrandTotal > 0 Then %>
						<td class="td-align positive border-right"><strong><%= FormatNumber(PlusMinusSamePeriodLastYearPct_GrandTotal,2)%>%</strong></td>
					<% ElseIf PlusMinusSamePeriodLastYearPct_GrandTotal < 0 Then %>
						<td class="td-align negative border-right"><strong><%= FormatNumber(PlusMinusSamePeriodLastYearPct_GrandTotal,2)%>%</strong></td>
					<% End If %>
					
					<%
			Next
	    	%>
		    
		    </tbody>
		</table>
	</div>
	<%
	End If
	rs.Close
	%>
				
          
</div>
          

<!--#include file="../inc/footer-main.asp"-->