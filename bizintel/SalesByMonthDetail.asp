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


	'****************************************
	'Get Number of Days In Month Function
	'****************************************
	
		Function DaysInMonth(passedMonth,passedYear)
		    
		    Select Case passedMonth
		        'January, March, May, July, August, October, December
		        Case 1, 3, 5, 7, 8, 10, 12
		        DaysInMonth = 31
		        
		        'February
		        Case 2
		         If (passedYear Mod 4) = 0 Then
		                DaysInMonth = 29
		            Else:
		                DaysInMonth = 28
		            End If
		            
		        'April, June, September, November
		        Case 4, 6, 9, 11
		        DaysInMonth = 30
		        
		    End Select
		    
		End Function 	

	'****************************************
	'END Get Number of Days In Month Function
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
	
	.Month-difference{
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


	CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Sales By Month Summary View"


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
		
		selMonthYearCombinations = rs("ReportSpecificData15")

		If IsNull(DefaultSelectedCustomerClassesForSalesReport) Then DefaultSelectedCustomerClassesForSalesReport = ""
		If IsNull(InvoiceTypeBackOrder) Then InvoiceTypeBackOrder = ""
		If IsNull(InvoiceTypeCreditMemo) Then InvoiceTypeCreditMemo = ""
		If IsNull(InvoiceTypeARDebit) Then InvoiceTypeARDebit = ""
		If IsNull(InvoiceTypeRental) Then InvoiceTypeRental = ""
		If IsNull(InvoiceTypeRouteInvoicing) Then InvoiceTypeRouteInvoicing = ""
		If IsNull(InvoiceTypeInterest) Then InvoiceTypeInterest = ""
		If IsNull(InvoiceTypeTelselInvoicing) Then InvoiceTypeTelselInvoicing = ""
		
		If IsNull(selMonthYearCombinations) Then selMonthYearCombinations = ""
						
	End If										
	'****************************
	'End Read Settings_Reports
	'****************************
	
	'**************************************************************************************
	'Build Page Header From Custome Month Ranges, If Applicable
	'**************************************************************************************

	MonthYearCombinationsArray = ""

	
	'////////////////////////////////////////////////////////////////////////////////////////////
	'CREATE DEFAULT MONTH AND YEAR COMBINATION IF NO CUSTOMIZATION IS SET
	'////////////////////////////////////////////////////////////////////////////////////////////
	
	If selMonthYearCombinations = "" Then
	
		ReDim MonthYearCombinationsArray(0)
		
		CurrentMonth = Month(Date())
		CurrentYear = Year(Date()) 
		
		MonthYearCombinationsArray(0) = CurrentMonth & "*" & CurrentYear
			
	Else
		MonthYearCombinationsArray = Split(selMonthYearCombinations,",")
	End If
	
	
	If UBound(MonthYearCombinationsArray) >= 0 Then
	
		pageHeaderTextBaseVsCompare = ""
		pageHeaderTextBaseVsCompare = pageHeaderTextBaseVsCompare & "<span style='font-size:16px; color:#337ab7; margin-left:40px; '>Currently Viewing:</span><br>"
		
		For z = 0 to UBound(MonthYearCombinationsArray)
			
			currentMonthYearCombo = split(MonthYearCombinationsArray(z),"*")
			
			CurrentMonth = currentMonthYearCombo(0)
			CurrentYear = currentMonthYearCombo(1)
			numDaysInThisMonth = DaysInMonth(CurrentMonth,CurrentYear)
			
			lastYear = currentYear - 1
			numDaysInThisMonthLastYear = DaysInMonth(CurrentMonth,LastYear)
			thisMonthLastYear = currentMonth
			
			defaultLoopBaseStartDate = cDate(CurrentMonth & "/1/" & CurrentYear)
			defaultLoopBaseEndDate = cDate(CurrentMonth & "/" & numDaysInThisMonth & "/" & CurrentYear) 
			
			defaultLoopCompareStartDate = cDate(thisMonthLastYear & "/1/" & lastYear)
			defaultLoopCompareEndDate = cDate(thisMonthLastYear & "/" & numDaysInThisMonthLastYear & "/" & lastYear)
						
			pageHeaderTextBaseVsCompare = pageHeaderTextBaseVsCompare & "<span style='font-size:16px; color:rgb(72, 151, 54); margin-left:40px;'>"
			pageHeaderTextBaseVsCompare = pageHeaderTextBaseVsCompare & "(" & defaultLoopBaseStartDate & "-" & defaultLoopBaseEndDate & ") vs " 
			pageHeaderTextBaseVsCompare = pageHeaderTextBaseVsCompare & "(" & defaultLoopCompareStartDate & "-" & defaultLoopCompareEndDate & ")<br>" 
			pageHeaderTextBaseVsCompare = pageHeaderTextBaseVsCompare & "</span>"
			
		Next	
		
		%>
			<h3 class="page-header"><i class="fa fa-graduation-cap"></i> Sales By Month and Customer Class (Summary) For Months
			<!-- modal button !-->
			<button type="button" class="btn btn-primary" data-toggle="modal" data-target=".bs-example-modal-lg-customize">
			  Customize
			</button>
			<% If SalesByDayReportTableSet(ReportNumber)Then %>
				<a href="<%= BaseURL %>bizintel/SalesByMonthDetail_Customize_ClearValues.asp"><button type="button" class="btn btn-primary">Clear Customizations</button></a>
			<% End If %>
				<!-- eof modal button !-->
							<br><br>
				<%= pageHeaderTextBaseVsCompare %>

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
	
	<!--#include file="SalesByMonthDetail_Customize.asp"-->	
	 
	
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
	'START LOOPING THROUGH THE ARRAY OF MONTH AND YEAR COMBINATIONS
	'FOR EACH MONTH/YEAR, CREATE ONE TABLE ROW THAT CONSISTS OF EACH CUSTOMER TYPE AND 
	'COMPARES THAT MONTH/YEAR TO THE SAME MONTH IN THE PREVIOUS YEAR
	'////////////////////////////////////////////////////////////////////////////////////////////
	
	
	CurrentMonthYearArrayIntRecID = -1

	For y = 0 to UBound(MonthYearCombinationsArray)
	
		currentMonthYearCombo = split(MonthYearCombinationsArray(y),"*")
		
		CurrentMonth = currentMonthYearCombo(0)
		CurrentYear = currentMonthYearCombo(1)
		numDaysInThisMonth = DaysInMonth(CurrentMonth,CurrentYear)
		
		lastYear = currentYear - 1
		numDaysInThisMonthLastYear = DaysInMonth(CurrentMonth,LastYear)
		thisMonthLastYear = currentMonth
		
		currentLoopBaseStartDate = cDate(CurrentMonth & "/1/" & CurrentYear)
		currentLoopBaseEndDate = cDate(CurrentMonth & "/" & numDaysInThisMonth & "/" & CurrentYear) 
		
		currentLoopCompareStartDate = cDate(thisMonthLastYear & "/1/" & lastYear)
		currentLoopCompareEndDate = cDate(thisMonthLastYear & "/" & numDaysInThisMonthLastYear & "/" & lastYear) 
		
		%>
		
		<% If y = 0 Then %>			
			<div class="table-responsive">
				<table class="agenda">
				  <thead>
				    <tr>
				    	<th rowspan="2">Date/Month</th>
				    	<%
						For z = 0 to UBound(CustomerClassArray)
							%><th colspan="8" class="customerclass"><%= CustomerClassArray(z) %> - <%= GetCustomerClassNameByID(CustomerClassArray(z)) %></th><%
						Next
				    	%>
				    </tr>
				    <tr>
				    	<%
						For z = 0 to UBound(CustomerClassArray)
							%>
						        <th>Base Month Total Sales $</th>
						        <th>Base Month Total Cost $</th>
						        <th>Base Month Gross Profit %</th>
						        <th>Comparison Month Sales $</th>
						        <th>Comparison Month Cost $</th>
						        <th>Comparison Month Gross Profit %</th>
						        <th>+/- $ for Comparison Month</th>
						        <th class="border-right">+/- % Comparison Month</th>
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
			'A NEW ROW SHOULD START WHEN BOTH THE INTERNAL RECORD IDENTIFIER FOR THE BASE Month
			'HAS ADVANCED TO THE NEXT "Y" VALUE IN THE FOR/NEXT LOOP
			'////////////////////////////////////////////////////////////////////////////////////////////	
								
			If cInt(CurrentMonthYearArrayIntRecID) <> cInt(y) Then %>
				<tr>
					<th class="date"><%= FormatDateTime(currentLoopBaseStartDate,2) %>-<%= FormatDateTime(currentLoopBaseEndDate,2) %> 
									<br>vs<br> <%= FormatDateTime(currentLoopCompareStartDate,2) %>-<%= FormatDateTime(currentLoopCompareEndDate,2) %></th>
					<%  	
					CurrentMonthYearArrayIntRecID = y
			End If 
			
			For z = 0 to UBound(CustomerClassArray)	
			

				CurrentRecordClassCode = ""
				SelectedBaseMonth_GrossProfitPct = 0
				SelectedCompareMonth_GrossProfitPct = 0
				PlusMinusCompareMonthPct = 0
				PlusMinusCompareMonthDlrs = 0
			
				CurrentRecordClassCode = CustomerClassArray(z)
		
				'***************************************************************************************************
				'Build SQL STMT To Get Months Sales Data For The BASE Month Date Range
				'Note: You Need To Run a SQL Statement To Get The Start and End Dates For The Base Month
				'***************************************************************************************************	

				SQL = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost FROM BI_DailySalesByTypeByClass "
				SQL = SQL & " WHERE (ivsDate BETWEEN '" & currentLoopBaseStartDate & "' AND '" & currentLoopBaseEndDate & "') AND "
				SQL = SQL & " ClassCode = '" & CurrentRecordClassCode & "' "
				'SQL = SQL & " GROUP BY IvsDate, ClassCode ORDER BY ivsdate,ClassCode"
				If WHERE_CLAUSE_IVSTYPE <> "" Then SQL = SQL & WHERE_CLAUSE_IVSTYPE
			 	
			
				'Response.Write("<strong>SQL For Base Month Loop:</strong> " & SQL & "<br>")
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3
				rs.Open SQL, Session("ClientCnnString")

				'*****************************************************************************************************************************************
				'*****************************************************************************************************************************************


				If rs("TotCost") <> 0 Then
					SelectedBaseMonth_GrossProfitPct = Round(((rs("TotSales") - rs("TotCost")) / rs("TotSales")) * 100,2)
				ElseIf rs("TotCost") = 0 AND rs("TotSales") = 0 Then
					SelectedBaseMonth_GrossProfitPct = 0
				Else
					SelectedBaseMonth_GrossProfitPct = 100
				End If
				
		
				SQLGetSalesDataCompareMonth = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost FROM BI_DailySalesByTypeByClass "
				SQLGetSalesDataCompareMonth = SQLGetSalesDataCompareMonth & " WHERE (ivsDate BETWEEN '" & currentLoopCompareStartDate & "' AND '" & currentLoopCompareEndDate & "') AND "
				SQLGetSalesDataCompareMonth = SQLGetSalesDataCompareMonth & " ClassCode = '" & CurrentRecordClassCode & "' "
				If WHERE_CLAUSE_IVSTYPE <> "" Then SQLGetSalesDataCompareMonth = SQLGetSalesDataCompareMonth & WHERE_CLAUSE_IVSTYPE
					
				'**************************************************************************************
				'END Build SQL STMT To Select Date Range From BI_DailySalesByTypeByClass 
				'**************************************************************************************	
				
				'Response.write("<strong>SQLGetSalesDataCompareMonth </strong>: " & SQLGetSalesDataCompareMonth & "<br>")
										 
				Set rsGetSalesDataCompareMonth = Server.CreateObject("ADODB.Recordset")
				rsGetSalesDataCompareMonth.CursorLocation = 3 
				Set cnnGetSalesDataCompareMonth = Server.CreateObject("ADODB.Connection")
				cnnGetSalesDataCompareMonth.open Session("ClientCnnString")
						
				Set rsGetSalesDataCompareMonth= cnnGetSalesDataCompareMonth.Execute(SQLGetSalesDataCompareMonth)
				
				If NOT rsGetSalesDataCompareMonth.EOF then 
						
					If rsGetSalesDataCompareMonth("TotCost") <> 0 Then
						SelectedCompareMonth_GrossProfitPct = Round(((rsGetSalesDataCompareMonth("TotSales") - rsGetSalesDataCompareMonth("TotCost")) / rsGetSalesDataCompareMonth("TotSales")) * 100,2)
					ElseIf rsGetSalesDataCompareMonth("TotCost") = 0 AND rsGetSalesDataCompareMonth("TotSales") = 0 Then
						SelectedCompareMonth_GrossProfitPct = 0
					Else
						SelectedCompareMonth_GrossProfitPct = 100
					End If
				End If
					
				If SelectedBaseMonth_GrossProfitPct = 0 Then
					PlusMinusCompareMonthPct = SelectedCompareMonth_GrossProfitPct * (-1)
				Else
					PlusMinusCompareMonthPct = SelectedBaseMonth_GrossProfitPct - SelectedCompareMonth_GrossProfitPct
				End If
				
				PlusMinusCompareMonthDlrs = rs("TotSales") - rsGetSalesDataCompareMonth("TotSales")
				
				%>
				
				
				<% If rs("TotSales") = 0 Then %>
					<td class="td-align zero"><%= FormatCurrency(rs("TotSales"),2,-2,-1)%></td> 
				<% ElseIf rs("TotSales") > 0 Then %>
					<td class="td-align positive"><%= FormatCurrency(rs("TotSales"),2,-2,-1)%></td>
				<% ElseIf rs("TotSales") < 0 Then %>
					<td class="td-align negative"><%= FormatCurrency(rs("TotSales"),2,-2,-1)%></td>
				<% End If %>
				
				<% If rs("TotCost") = 0 Then %>
					<td class="td-align zero"><%= FormatCurrency(rs("TotCost"),2,-2,-1)%></td> 
				<% ElseIf rs("TotCost") > 0 Then %>
					<td class="td-align positive"><%= FormatCurrency(rs("TotCost"),2,-2,-1)%></td>
				<% ElseIf rs("TotCost") < 0 Then %>
					<td class="td-align negative"><%= FormatCurrency(rs("TotCost"),2,-2,-1)%></td>
				<% End If %>
								
				<% If SelectedBaseMonth_GrossProfitPct = 0 Then %>
					<td class="td-align zero"><%= FormatNumber(SelectedBaseMonth_GrossProfitPct,2) %>%</td> 
				<% ElseIf SelectedBaseMonth_GrossProfitPct > 0 Then %>
					<td class="td-align positive"><%= FormatNumber(SelectedBaseMonth_GrossProfitPct,2) %>%</td>
				<% ElseIf SelectedBaseMonth_GrossProfitPct < 0 Then %>
					<td class="td-align negative"><%= FormatNumber(SelectedBaseMonth_GrossProfitPct,2) %>%</td>
				<% End If %>
				
				<% If rsGetSalesDataCompareMonth("TotSales") = 0 Then %>
					<td class="td-align zero"><%= FormatCurrency(rsGetSalesDataCompareMonth("TotSales"),2,-2,-1)%></td> 
				<% ElseIf rsGetSalesDataCompareMonth("TotSales") > 0 Then %>
					<td class="td-align positive"><%= FormatCurrency(rsGetSalesDataCompareMonth("TotSales"),2,-2,-1)%></td>
				<% ElseIf rsGetSalesDataCompareMonth("TotSales") < 0 Then %>
					<td class="td-align negative"><%= FormatCurrency(rsGetSalesDataCompareMonth("TotSales"),2,-2,-1)%></td>
				<% End If %>
				
				<% If rsGetSalesDataCompareMonth("TotCost") = 0 Then %>
					<td class="td-align zero"><%= FormatCurrency(rsGetSalesDataCompareMonth("TotCost"),2,-2,-1)%></td> 
				<% ElseIf rsGetSalesDataCompareMonth("TotCost") > 0 Then %>
					<td class="td-align positive"><%= FormatCurrency(rsGetSalesDataCompareMonth("TotCost"),2,-2,-1)%></td>
				<% ElseIf rsGetSalesDataCompareMonth("TotCost") < 0 Then %>
					<td class="td-align negative"><%= FormatCurrency(rsGetSalesDataCompareMonth("TotCost"),2,-2,-1)%></td>
				<% End If %>
				
				<% If SelectedCompareMonth_GrossProfitPct = 0 Then %>
					<td class="td-align zero"><%= FormatNumber(SelectedCompareMonth_GrossProfitPct,2) %>%</td> 
				<% ElseIf SelectedCompareMonth_GrossProfitPct > 0 Then %>
					<td class="td-align positive"><%= FormatNumber(SelectedCompareMonth_GrossProfitPct,2) %>%</td>
				<% ElseIf SelectedCompareMonth_GrossProfitPct < 0 Then %>
					<td class="td-align negative"><%= FormatNumber(SelectedCompareMonth_GrossProfitPct,2) %>%</td>
				<% End If %>
				
				
				<% If PlusMinusCompareMonthDlrs = 0 Then %>
					<td class="td-align zero"><%= FormatCurrency(PlusMinusCompareMonthDlrs,2)%></td> 
				<% ElseIf PlusMinusCompareMonthDlrs > 0 Then %>
					<td class="td-align positive"><%= FormatCurrency(PlusMinusCompareMonthDlrs,2)%></td>
				<% ElseIf PlusMinusCompareMonthDlrs < 0 Then %>
					<td class="td-align negative"><%= FormatCurrency(PlusMinusCompareMonthDlrs,2)%></td>
				<% End If %>
				
				<% If PlusMinusCompareMonthPct = 0 Then %>
					<td class="td-align zero border-right"><%= FormatNumber(PlusMinusCompareMonthPct,2) %>%</td> 
				<% ElseIf PlusMinusCompareMonthPct > 0 Then %>
					<td class="td-align positive border-right"><%= FormatNumber(PlusMinusCompareMonthPct,2) %>%</td>
				<% ElseIf PlusMinusCompareMonthPct < 0 Then %>
					<td class="td-align negative border-right"><%= FormatNumber(PlusMinusCompareMonthPct,2) %>%</td>
				<% End If %>
				

			<%
	 		Next '''''(Customer Class Loop)
	 		
		 	
		Next ''''(Month Loop)
		
		'****************************************************************************************************************
		'OUTER LOOPS ARE COMPLETE, NOW PREPARE TO WRITE THE LAST ROW OF THE TABLE THAT GIVES MASTER SUMMARY TOTALS
		'****************************************************************************************************************
				
			%>
				
			</tr>

		    <tr>
		    	<th rowspan="2">&nbsp;</th>
		    	<%
				For z = 0 to UBound(CustomerClassArray)
					%><th colspan="8" class="customerclass"><%= CustomerClassArray(z) %> - <%= GetCustomerClassNameByID(CustomerClassArray(z)) %></th><%
				Next
		    	%>
		    </tr>
		    <tr>
		    	<%
				For z = 0 to UBound(CustomerClassArray)
					%>
			        <th>Base Month Total Sales $</th>
			        <th>Base Month Total Cost $</th>
			        <th>Base Month Gross Profit %</th>
			        <th>Comparison Month Sales $</th>
			        <th>Comparison Month Cost $</th>
			        <th>Comparison Month Gross Month %</th>
			        <th>+/- $ for Comparison Month</th>
			        <th class="border-right">+/- % Comparison Month</th>
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
						'Get Summary of Total Sales For This Customer Class For Current Month
						'*********************************************************************************************************
						
						WHERE_CLAUSE_BASE_MONTH_MASTER = ""
						WHERE_CLAUSE_COMPARISON_MONTH_MASTER = ""
						
											
						'*******************************************************************
						'If Months Were Selected By The User, Loop and Build WHERE Clause
						'*******************************************************************
					
							For i = 0 to UBound(MonthYearCombinationsArray)
							
								currentMonthYearCombo = split(MonthYearCombinationsArray(i),"*")
								
								CurrentMonth = currentMonthYearCombo(0)
								CurrentYear = currentMonthYearCombo(1)
								numDaysInThisMonth = DaysInMonth(CurrentMonth,CurrentYear)
																
								currentLoopBaseStartDate = cDate(CurrentMonth & "/1/" & CurrentYear)
								currentLoopBaseEndDate = cDate(CurrentMonth & "/" & numDaysInThisMonth & "/" & CurrentYear) 
								
								If i = 0 Then
									WHERE_CLAUSE_BASE_MONTH_MASTER = WHERE_CLAUSE_BASE_MONTH_MASTER & " (ivsDate BETWEEN '" & currentLoopBaseStartDate & "' AND '" & currentLoopBaseEndDate & "') "
								Else
									WHERE_CLAUSE_BASE_MONTH_MASTER = WHERE_CLAUSE_BASE_MONTH_MASTER & " OR (ivsDate BETWEEN '" & currentLoopBaseStartDate & "' AND '" & currentLoopBaseEndDate & "') "	
								End If
								
							Next	
					
						
					'*******************************************************************
					'If Months Were Selected By The User, Loop and Build WHERE Clause
					'*******************************************************************
					
						For i = 0 to UBound(MonthYearCombinationsArray)
						
						
							currentMonthYearCombo = split(MonthYearCombinationsArray(i),"*")
							
							CurrentMonth = currentMonthYearCombo(0)
							CurrentYear = currentMonthYearCombo(1)
							numDaysInThisMonth = DaysInMonth(CurrentMonth,CurrentYear)
							
							lastYear = currentYear - 1
							numDaysInThisMonthLastYear = DaysInMonth(CurrentMonth,LastYear)
							thisMonthLastYear = currentMonth
							
							currentLoopBaseStartDate = cDate(CurrentMonth & "/1/" & CurrentYear)
							currentLoopBaseEndDate = cDate(CurrentMonth & "/" & numDaysInThisMonth & "/" & CurrentYear) 
							
							currentLoopCompareStartDate = cDate(thisMonthLastYear & "/1/" & lastYear)
							currentLoopCompareEndDate = cDate(thisMonthLastYear & "/" & numDaysInThisMonthLastYear & "/" & lastYear) 
												
						
							If i = 0 Then						
								WHERE_CLAUSE_COMPARISON_MONTH_MASTER = WHERE_CLAUSE_COMPARISON_MONTH_MASTER & " (ivsDate BETWEEN '" & currentLoopCompareStartDate & "' AND '" & currentLoopCompareEndDate & "') "
							Else
								WHERE_CLAUSE_COMPARISON_MONTH_MASTER = WHERE_CLAUSE_COMPARISON_MONTH_MASTER & " OR (ivsDate BETWEEN '" & currentLoopCompareStartDate & "' AND '" & currentLoopCompareEndDate & "') "
							End If
							

						Next						


						
						'******************************
						'BUILD SQL STMT
						'******************************
						
						SQLGetTotalsForMasterSummary = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost FROM BI_DailySalesByTypeByClass WHERE "

						If WHERE_CLAUSE_BASE_MONTH_MASTER <> "" Then SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & "(" & WHERE_CLAUSE_BASE_MONTH_MASTER & ")"
						
						If WHERE_CLAUSE_IVSTYPE <> "" Then SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & WHERE_CLAUSE_IVSTYPE
						
						If WHERE_CLAUSE_BASE_MONTH_MASTER <> "" OR WHERE_CLAUSE_IVSTYPE <> "" THEN
					 		SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " AND ClassCode = '" & CustomerClassArray(z) & "'"
					 	Else
					 		SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " ClassCode = '" & CustomerClassArray(z) & "'"
					 	End If
						
						'Response.write("<strong>SQLGetTotalsForMasterSummary Base Month</strong> : " & SQLGetTotalsForMasterSummary & "<br>")
						'******************************
						'END BUILD SQL STMT
						'******************************
						
			 
						Set rsGetTotalsForMasterSummary = Server.CreateObject("ADODB.Recordset")
						rsGetTotalsForMasterSummary.CursorLocation = 3 
						Set rsGetTotalsForMasterSummary= cnnGetTotalsForMasterSummary.Execute(SQLGetTotalsForMasterSummary)
						
						
						If NOT rsGetTotalsForMasterSummary.EOF Then 
						
							TotalSalesBaseMonth_GrandTotal = rsGetTotalsForMasterSummary("TotSales")
							TotalCostBaseMonth_GrandTotal = rsGetTotalsForMasterSummary("TotCost")
		
							If TotalCostBaseMonth_GrandTotal <> 0 Then
								GrossProfitBaseMonth_GrandTotal = Round(((TotalSalesBaseMonth_GrandTotal - TotalCostBaseMonth_GrandTotal) / TotalSalesBaseMonth_GrandTotal) * 100,2)
							ElseIf TotalCostBaseMonth_GrandTotal = 0 AND TotalSalesBaseMonth_GrandTotal = 0 Then
								GrossProfitBaseMonth_GrandTotal = 0
							Else
								GrossProfitBaseMonth_GrandTotal = 100
							End If
	
						End If
							
						'*********************************************************************************************************
						'Get Summary of Total Sales For This Customer Class For Comparison Month
						'*********************************************************************************************************
	
						'******************************
						'BUILD SQL STMT
						'******************************
													
						SQLGetTotalsForMasterSummary = "SELECT SUM(TotNumOrders) AS TotNumOrders, SUM(TotSales) AS TotSales, SUM(TotCost) AS TotCost FROM BI_DailySalesByTypeByClass WHERE "
						
						If WHERE_CLAUSE_COMPARISON_MONTH_MASTER <> "" Then SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & "(" & WHERE_CLAUSE_COMPARISON_MONTH_MASTER & ")"
						
						If WHERE_CLAUSE_IVSTYPE <> "" Then SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & WHERE_CLAUSE_IVSTYPE
						
						If WHERE_CLAUSE_COMPARISON_MONTH_MASTER <> "" OR WHERE_CLAUSE_IVSTYPE <> "" THEN
					 		SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " AND ClassCode = '" & CustomerClassArray(z) & "'"
					 	Else
					 		SQLGetTotalsForMasterSummary = SQLGetTotalsForMasterSummary & " ClassCode = '" & CustomerClassArray(z) & "'"
					 	End If
						
					 	'Response.write("<strong>SQLGetTotalsForMasterSummary Compare Month</strong> : " & SQLGetTotalsForMasterSummary & "<br>")
					
						'******************************
						'END BUILD SQL STMT
						'******************************
							
						Set rsGetTotalsForMasterSummary= cnnGetTotalsForMasterSummary.Execute(SQLGetTotalsForMasterSummary)
						
						
						If NOT rsGetTotalsForMasterSummary.EOF Then 
						
							TotalSalesCompareMonth_GrandTotal = rsGetTotalsForMasterSummary("TotSales")
							TotalCostCompareMonth_GrandTotal = rsGetTotalsForMasterSummary("TotCost")
		
							If TotalCostCompareMonth_GrandTotal <> 0 Then
								GrossProfitCompareMonth_GrandTotal = Round(((TotalSalesCompareMonth_GrandTotal - TotalCostCompareMonth_GrandTotal) / TotalSalesCompareMonth_GrandTotal) * 100,2)
							ElseIf TotalCostCompareMonth_GrandTotal = 0 AND TotalSalesCompareMonth_GrandTotal = 0 Then
								GrossProfitCompareMonth_GrandTotal = 0
							Else
								GrossProfitCompareMonth_GrandTotal = 100
							End If
	
						End If
						
						PlusMinusCompareMonthDlrs_GrandTotal = TotalSalesBaseMonth_GrandTotal - TotalSalesCompareMonth_GrandTotal
						
						If GrossProfitBaseMonth_GrandTotal = 0 Then
							PlusMinusCompareMonthPct_GrandTotal = GrossProfitCompareMonth_GrandTotal * (-1)
						Else
							PlusMinusCompareMonthPct_GrandTotal = GrossProfitBaseMonth_GrandTotal - GrossProfitCompareMonth_GrandTotal
						End If
								
	
			    	
				    	%>
				    	
						<% If TotalSalesBaseMonth_GrandTotal = 0 Then %>
							<td class="td-align zero"><strong><%= FormatCurrency(TotalSalesBaseMonth_GrandTotal,2,-2,-1)%></strong></td> 
						<% ElseIf TotalSalesBaseMonth_GrandTotal > 0 Then %>
							<td class="td-align positive"><strong><%= FormatCurrency(TotalSalesBaseMonth_GrandTotal,2,-2,-1)%></strong></td>
						<% ElseIf TotalSalesBaseMonth_GrandTotal < 0 Then %>
							<td class="td-align negative"><strong><%= FormatCurrency(TotalSalesBaseMonth_GrandTotal,2,-2,-1)%></strong></td>
						<% End If %>
						
						<% If TotalCostBaseMonth_GrandTotal = 0 Then %>
							<td class="td-align zero"><strong><%= FormatCurrency(TotalCostBaseMonth_GrandTotal,2,-2,-1)%></strong></td> 
						<% ElseIf TotalCostBaseMonth_GrandTotal > 0 Then %>
							<td class="td-align positive"><strong><%= FormatCurrency(TotalCostBaseMonth_GrandTotal,2,-2,-1)%></strong></td>
						<% ElseIf TotalCostBaseMonth_GrandTotal < 0 Then %>
							<td class="td-align negative"><strong><%= FormatCurrency(TotalCostBaseMonth_GrandTotal,2,-2,-1)%></strong></td>
						<% End If %>
					
						<% If GrossProfitBaseMonth_GrandTotal = 0 Then %>
							<td class="td-align zero"><strong><%= FormatNumber(GrossProfitBaseMonth_GrandTotal,2)%>%</strong></td> 
						<% ElseIf GrossProfitBaseMonth_GrandTotal > 0 Then %>
							<td class="td-align positive"><strong><%= FormatNumber(GrossProfitBaseMonth_GrandTotal,2)%>%</strong></td>
						<% ElseIf GrossProfitBaseMonth_GrandTotal < 0 Then %>
							<td class="td-align negative"><strong><%= FormatNumber(GrossProfitBaseMonth_GrandTotal,2)%>%</strong></td>
						<% End If %>
						
	  					<% If TotalSalesCompareMonth_GrandTotal = 0 Then %>
							<td class="td-align zero"><strong><%= FormatCurrency(TotalSalesCompareMonth_GrandTotal,2,-2,-1)%></strong></td> 
						<% ElseIf TotalSalesCompareMonth_GrandTotal > 0 Then %>
							<td class="td-align positive"><strong><%= FormatCurrency(TotalSalesCompareMonth_GrandTotal,2,-2,-1)%></strong></td>
						<% ElseIf TotalSalesCompareMonth_GrandTotal < 0 Then %>
							<td class="td-align negative"><strong><%= FormatCurrency(TotalSalesCompareMonth_GrandTotal,2,-2,-1)%></strong></td>
						<% End If %>
							
						<% If TotalCostCompareMonth_GrandTotal = 0 Then %>
							<td class="td-align zero"><strong><%= FormatCurrency(TotalCostCompareMonth_GrandTotal,2,-2,-1)%></strong></td> 
						<% ElseIf TotalCostCompareMonth_GrandTotal > 0 Then %>
							<td class="td-align positive"><strong><%= FormatCurrency(TotalCostCompareMonth_GrandTotal,2,-2,-1)%></strong></td>
						<% ElseIf TotalCostCompareMonth_GrandTotal < 0 Then %>
							<td class="td-align negative"><strong><%= FormatCurrency(TotalCostCompareMonth_GrandTotal,2,-2,-1)%></strong></td>
						<% End If %>
						
						<% If GrossProfitCompareMonth_GrandTotal = 0 Then %>
							<td class="td-align zero"><strong><%= FormatNumber(GrossProfitCompareMonth_GrandTotal,2)%>%</strong></td> 
						<% ElseIf GrossProfitCompareMonth_GrandTotal > 0 Then %>
							<td class="td-align positive"><strong><%= FormatNumber(GrossProfitCompareMonth_GrandTotal,2)%>%</strong></td>
						<% ElseIf GrossProfitCompareMonth_GrandTotal < 0 Then %>
							<td class="td-align negative"><strong><%= FormatNumber(GrossProfitCompareMonth_GrandTotal,2)%>%</strong></td>
						<% End If %>
						
						<% If PlusMinusCompareMonthDlrs_GrandTotal = 0 Then %>
							<td class="td-align zero"><strong><%= FormatCurrency(PlusMinusCompareMonthDlrs_GrandTotal,2)%></strong></td> 
						<% ElseIf PlusMinusCompareMonthDlrs_GrandTotal > 0 Then %>
							<td class="td-align positive"><strong><%= FormatCurrency(PlusMinusCompareMonthDlrs_GrandTotal,2)%></strong></td>
						<% ElseIf PlusMinusCompareMonthDlrs_GrandTotal < 0 Then %>
							<td class="td-align negative"><strong><%= FormatCurrency(PlusMinusCompareMonthDlrs_GrandTotal,2)%></strong></td>
						<% End If %>
						
						<% If PlusMinusCompareMonthPct_GrandTotal = 0 Then %>
							<td class="td-align zero border-right"><strong><%= FormatNumber(PlusMinusCompareMonthPct_GrandTotal,2)%>%</strong></td> 
						<% ElseIf PlusMinusCompareMonthPct_GrandTotal > 0 Then %>
							<td class="td-align positive border-right"><strong><%= FormatNumber(PlusMinusCompareMonthPct_GrandTotal,2)%>%</strong></td>
						<% ElseIf PlusMinusCompareMonthPct_GrandTotal < 0 Then %>
							<td class="td-align negative border-right"><strong><%= FormatNumber(PlusMinusCompareMonthPct_GrandTotal,2)%>%</strong></td>
						<% End If %>
						
						
						<%
				Next
		    	%>
			    
			    </tbody>
			</table>
		</div>
			
          
</div>
          

<!--#include file="../inc/footer-main.asp"-->