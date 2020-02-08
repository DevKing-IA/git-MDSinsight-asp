<%
Server.ScriptTimeout = 900000 'Default value
Dim ReportNumber 
ReportNumber = 1100
%>
<!--#include file="../inc/InSightFuncs_BizIntel.asp"-->
<%
'************************
'Does not need ot read the seetings table, gets
'DefaultSelectedCategoriesForVPandVPC from the global settings table
'no other exclusions on panel version
'reads DefaultSelectedCategoriesForVPandVPC from insightglobalvars.asp
'************************
Dim CatArray(22)
	
AllOff = True
For x = 0 to 21
	CatArray(x)="off"
Next 

CategoryArrayDefault = ""
CategoryArrayDefault = Split(DefaultSelectedCategoriesForVPandVPC,",")
For z = 0 to UBound(CategoryArrayDefault)
	CatArray(CategoryArrayDefault(z)) = "on"
	AllOff = False
Next
	
'If they are all off, turn them all on, can't have them all off
If AllOFf = True Then
	For x = 0 to 21
		CatArray(x) = "on" 
	Next
End If
	
WHERE_CLAUSE_ADDITIONAL = ""
	
For x = 0 to 21
	If CatArray(x) = "off" then WHERE_CLAUSE_ADDITIONAL = WHERE_CLAUSE_ADDITIONAL & " AND prodCategory <>" & x & " "
Next
%>
		 
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
	
	.td-left{
		text-align: right;
		width: 50%;
	}
	
	.td-right{
 		width: 50%;
	}
	
	.table-size{
		width: 70%;
	}
	
	.table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
		border:0px;
	}
	
	.date-range{
 		margin:0px;
		font-size: 11px;
	}
</style>
	

<%
'Figure out what periods to show

LastClosedPeriod_FiscalYear = 	trim(right(GetLastClosedPeriodAndYear(),Instr(GetLastClosedPeriodAndYear(),"-")+1))
LastClosedPeriod_PeriodNumber = trim(left(GetLastClosedPeriodAndYear(),Instr(GetLastClosedPeriodAndYear(),"-")-1))

Set cnnVPCGetFirstRangeStart = Server.CreateObject("ADODB.Connection")
cnnVPCGetFirstRangeStart.open (Session("ClientCnnString"))
Set rsVPCGetFirstRangeStart = Server.CreateObject("ADODB.Recordset")

SQLVPCGetFirstRangeStart = "SELECT BeginDate from BillingPeriodHistory where Year = " & LastClosedPeriod_FiscalYear -1 & " And Period = 1"
Set rsVPCGetFirstRangeStart = cnnVPCGetFirstRangeStart.Execute(SQLVPCGetFirstRangeStart)
If Not rsVPCGetFirstRangeStart.Eof Then Period1Start = rsVPCGetFirstRangeStart("BeginDate")

SQLVPCGetFirstRangeStart = "SELECT EndDate from BillingPeriodHistory where Year = " & LastClosedPeriod_FiscalYear -1 & " And Period = " & LastClosedPeriod_PeriodNumber
Set rsVPCGetFirstRangeStart = cnnVPCGetFirstRangeStart.Execute(SQLVPCGetFirstRangeStart)
If Not rsVPCGetFirstRangeStart.Eof Then Period1End = rsVPCGetFirstRangeStart("EndDate")

SQLVPCGetFirstRangeStart = "SELECT BeginDate from BillingPeriodHistory where Year = " & LastClosedPeriod_FiscalYear & " And Period = 1"
Set rsVPCGetFirstRangeStart = cnnVPCGetFirstRangeStart.Execute(SQLVPCGetFirstRangeStart)
If Not rsVPCGetFirstRangeStart.Eof Then Period2Start = rsVPCGetFirstRangeStart("BeginDate")

SQLVPCGetFirstRangeStart = "SELECT EndDate from BillingPeriodHistory where Year = " & LastClosedPeriod_FiscalYear & " And Period = " & LastClosedPeriod_PeriodNumber
Set rsVPCGetFirstRangeStart = cnnVPCGetFirstRangeStart.Execute(SQLVPCGetFirstRangeStart)
If Not rsVPCGetFirstRangeStart.Eof Then Period2End = rsVPCGetFirstRangeStart("EndDate")

Set rsVPCGetFirstRangeStart = Nothing
cnnVPCGetFirstRangeStart.Close
Set cnnVPCGetFirstRangeStart = Nothing

NotEnoughFound = False
'****************************************
'Get info for first period being reported
'****************************************
FirstPeriod_TotalSales = 0
FirstPeriod_TotalCases = 0
SQL = "SELECT SUM(itemQuantity * itemPrice) AS TotSales, SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"

SQL = SQL & " WHERE ivsDate BETWEEN '" & Period1Start  & "' AND '" & Period1End  & "' "

If WHERE_CLAUSE_ADDITIONAL <> "" Then SQL = SQL & WHERE_CLAUSE_ADDITIONAL


Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Open SQL, Session("ClientCnnString")
If not rs.eof Then
	FirstPeriod_TotalSales = rs("TotSales")
	FirstPeriod_TotalCases = rs("TotCases")
Else
	NotEnoughFound = True	
End If
rs.Close
'*****************************************
'Get info for second period being reported
'*****************************************
SecondPeriod_TotalSales = 0
SecondPeriod_TotalCases = 0
SQL = "SELECT SUM(itemQuantity * itemPrice) AS TotSales, SUM(NumberOfCases) AS TotCases FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistoryDetail"

SQL = SQL & " WHERE ivsDate BETWEEN '" & Period2Start  &  "' AND '" & Period2End  & "' "

If WHERE_CLAUSE_ADDITIONAL <> "" Then SQL = SQL & WHERE_CLAUSE_ADDITIONAL


Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Open SQL, Session("ClientCnnString")
If not rs.eof Then
	SecondPeriod_TotalSales = rs("TotSales")
	SecondPeriod_TotalCases = rs("TotCases")
Else
	NotEnoughFound = True	
End If
rs.Close

If FirstPeriod_TotalSales = "" or IsNull(FirstPeriod_TotalSales) or FirstPeriod_TotalCases = ""  or IsNull(FirstPeriod_TotalCases)_
or SecondPeriod_TotalSales = ""  or IsNull(SecondPeriod_TotalSales) or SecondPeriod_TotalCases = ""  or IsNull(SecondPeriod_TotalCases) then NotEnoughFound = True
	
If NotEnoughFound <> True Then

	'**************************************
	' Do all the calcs we will need up here
	'**************************************
	FirstPeriod_AVGSellPrice = FirstPeriod_TotalSales /  FirstPeriod_TotalCases
	SecondPeriod_AVGSellPrice = SecondPeriod_TotalSales /  SecondPeriod_TotalCases
	AVGSellPriceDifference =  SecondPeriod_AVGSellPrice - FirstPeriod_AVGSellPrice
	PriceChange = SecondPeriod_TotalCases * AVGSellPriceDifference
	VolumeChange = ((SecondPeriod_TotalCases - FirstPeriod_TotalCases) * FirstPeriod_AVGSellPrice)
	
	
	%> 
<h6><center>&nbsp;<u>Volume And Price Change</u><br><br>
Same periods last FY .vs P1 thru last closed period (P<%=LastClosedPeriod_PeriodNumber%>/<%=LastClosedPeriod_FiscalYear%>)</center></h6>


<div class="row">
	<div class="table-responsive">
		<table class="table">
			<tr>
				<td class="period-difference td-left" ><b>Difference</b></td>
				<td class="td-right"><%= FormatCurrency(SecondPeriod_TotalSales - FirstPeriod_TotalSales,2,-2,-1)%></td>              
			</tr>
			<tr>
				<td class="td-left"><b>Price Change</b></td>
				<td class="td-right"><%= FormatCurrency(PriceChange,2,-2,-1)  %></td>              
			</tr>
			<tr>
				<td  class="td-left"><b>Volume Change</b></td>
				<td class="td-right"><%= FormatCurrency(VolumeChange,2,-2,-1) %></td>              
			</tr>
			<tr>
				<td  class="td-left"><b>Net Difference</b></td>
				<td class="td-right">
					<%If PriceChange + VolumeChange > 0 Then 
						Response.Write("<font color='green'>"& FormatCurrency(PriceChange + VolumeChange,2,-2,-1) &"</font>")
					Else
						Response.Write("<font color='red'>"& FormatCurrency(PriceChange + VolumeChange,2,-2,-1) &"</font>")
					End If%>
 				</td>              
			</tr>
			 
		</table>
	 
	</div>
	
 		<div class="col-lg-12">
			<p class="date-range"><% 
 					Response.Write("Date Range 1: " & FormatDateTime(Period1Start,2) & " - " & FormatDateTime(Period1End,2)& "&nbsp;")
					Response.Write("&nbsp;[" & DATEDIFF("d", FormatDateTime(Period1Start), FormatDateTime(Period1End)) + 1 & " days]")
					%>
			</p>
			
			<p class="date-range"><% 
 					Response.Write("Date Range 2: " & FormatDateTime(Period2Start,2) & " - " & FormatDateTime(Period2End,2)& "&nbsp;")
					Response.Write("&nbsp;[" & DATEDIFF("d", FormatDateTime(Period2Start), FormatDateTime(Period2End)) + 1 & " days]")
					%>
			</p>
 	</div>
	<% End If %>
</div>
          



