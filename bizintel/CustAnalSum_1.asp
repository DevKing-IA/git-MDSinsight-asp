<%
Server.ScriptTimeout = 900000 'Default value

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))


%>
<!--#include file="../inc/header.asp"-->
<!--#include file="../inc/jquery_table_search.asp"-->
<!--#include file="../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../inc/InSightFuncs_Equipment.asp"--> 
<style>
markY {
    background-color: yellow;
    color: black;
} 
</style>

<%
CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Customer Analysis Summary 1"


'Special code for when they are brought here by the automated email
'in this case, it just resets everything to default values and
'runs the page just for the salesperson who logged in
'it does this by writing to the Settings_reports table so the 
'rest of the code can just run normally from that point


If Request.QueryString("qlSls") <> "" Then

	quickloginSalesPerson = Request.QueryString("qlSls")
	
	SQLqlSls = "SELECT * from Settings_Reports where ReportNumber = 2100 AND UserNo = " & Session("userNo")
	
	Set cnnqlSls = Server.CreateObject("ADODB.Connection")
	cnnqlSls.open (Session("ClientCnnString"))
	Set rsqlSls = Server.CreateObject("ADODB.Recordset")
	Set rsqlSls= cnnqlSls.Execute(SQLqlSls)
	
	'Rec does not exist yet, make it quick but empty, update it later
	If rsqlSls.EOF Then
		SQLqlSls = "Insert into Settings_Reports (ReportNumber, UserNo) Values (2100 , " & Session("userNo") & ")"
		rsqlSls.Close
		Set rsqlSls= cnnqlSls.Execute(SQLqlSls)
	End If
	
	'Now update the table with the values
	SQLqlSls = "Update Settings_Reports Set ReportSpecificData1 = '" & quickloginSalesPerson & "', "
	SQLqlSls = SQLqlSls & "ReportSpecificData2 = 'All', " 
	SQLqlSls = SQLqlSls & "ReportSpecificData3 = 'All', "  
	SQLqlSls = SQLqlSls & "ReportSpecificData4 = 'All', "  
	SQLqlSls = SQLqlSls & "ReportSpecificData5 = '100', " 
	SQLqlSls = SQLqlSls & "ReportSpecificData6 = '10'"
	SQLqlSls = SQLqlSls & " WHERE ReportNumber = 2100 AND UserNo = " & Session("userNo")
	Set rsqlSls= cnnqlSls.Execute(SQLqlSls)
	cnnqlSls.Close
	
	Set rsqlSls = Nothing
	Set cnnqlSls = Nothing
	
End If

If Request.QueryString("qlSls2") <> "" Then

	quickloginSalesPerson = Request.QueryString("qlSls2")
	
	SQLqlSls = "SELECT * from Settings_Reports where ReportNumber = 2100 AND UserNo = " & Session("userNo")
	
	Set cnnqlSls = Server.CreateObject("ADODB.Connection")
	cnnqlSls.open (Session("ClientCnnString"))
	Set rsqlSls = Server.CreateObject("ADODB.Recordset")
	Set rsqlSls= cnnqlSls.Execute(SQLqlSls)
	
	'Rec does not exist yet, make it quick but empty, update it later
	If rsqlSls.EOF Then
		SQLqlSls = "Insert into Settings_Reports (ReportNumber, UserNo) Values (2100 , " & Session("userNo") & ")"
		rsqlSls.Close
		Set rsqlSls= cnnqlSls.Execute(SQLqlSls)
	End If
	
	'Now update the table with the values
	SQLqlSls = "Update Settings_Reports Set ReportSpecificData2 = '" & quickloginSalesPerson & "', "
	SQLqlSls = SQLqlSls & "ReportSpecificData1 = 'All', " 
	SQLqlSls = SQLqlSls & "ReportSpecificData3 = 'All', "  
	SQLqlSls = SQLqlSls & "ReportSpecificData4 = 'All', "  
	SQLqlSls = SQLqlSls & "ReportSpecificData5 = '100', " 
	SQLqlSls = SQLqlSls & "ReportSpecificData6 = '10'"
	SQLqlSls = SQLqlSls & " WHERE ReportNumber = 2100 AND UserNo = " & Session("userNo")
	Set rsqlSls= cnnqlSls.Execute(SQLqlSls)
	cnnqlSls.Close
	
	Set rsqlSls = Nothing
	Set cnnqlSls = Nothing
	
End If

PeriodBeingEvaluated = GetLastClosedPeriodAndYear()
PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum()

WorkDaysIn3PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -3), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1
WorkDaysIn12PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -12), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1)) + 1 
WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated )) + 1 
WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated +1)) + 1 
WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1),Date()) + 1

'************************
'Read Settings_Reports
'************************
SQL = "SELECT * from Settings_Reports where ReportNumber = 2100 AND UserNo = " & Session("userNo")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)
UseSettings_Reports = False
If NOT rs.EOF Then
	UseSettings_Reports = True
	FilterSlsmn1 = rs("ReportSpecificData1")
	FilterSlsmn2 = rs("ReportSpecificData2")
	FilterReferral = rs("ReportSpecificData3")
	If FilterSlsmn1 <> "All" Then FilterSlsmn1 = CInt(FilterSlsmn1)
	If FilterSlsmn2 <> "All" Then FilterSlsmn2 = CInt(FilterSlsmn2)
	If FilterReferral <> "All" Then FilterReferral = CInt(FilterReferral)
	FilterSalesDollars = rs("ReportSpecificData5")
	FilterPercentage = rs("ReportSpecificData6")
	If FilterSalesDollars = "" Then FilterSalesDollars = 100
	If FilterPercentage = "" Then FilterPercentage = 10
End If
'****************************
'End Read Settings_Reports
'****************************

%>

  
<style>

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

	.vpc-variance-header{
		background: #D43F3A;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}
	
	.vpc-3pavg-header{
		background: #F0AD4E;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}
	
	.vpc-lcp-header{
		background: #337AB7;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.vpc-current-header{
		background: #5CB85C;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.gen-info-header{
		background: #3B579D;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.negative{
		font-weight:bold;
		color:red;	
	}

	.neutral{
		font-weight:bold;
		color:black;
	}

	.smaller-header{
		font-size: 0.8em;
		vertical-align: top !important;
		text-align: center;
	}	

	.smaller-detail-line{
		font-size: 0.8em;
	}	

</style>
<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript">

$(document).ready(function() {
    $("#PleaseWaitPanel").hide();
    $('#tableSuperSum').DataTable({
        scrollY: 500,
        scrollCollapse: true,
        paging: false,
        order: [ 2, 'asc' ]
    }
    );
});
</script>


<%
Response.Write("<div id=""PleaseWaitPanel"" class=""container"">")
Response.Write("<br><br>Creating Customer Analysis Summary 1 <br><br>This may take up to a full minute, please wait...<br><br>")
Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
Response.Write("</div>")
Response.Flush()

%>





<h3 class="page-header"><i class="fa fa-graduation-cap"></i> Customer Analysis Summary 1 For Period <%=PeriodBeingEvaluated %>
&nbsp;&nbsp;
<!-- modal button !-->
<button type="button" class="btn btn-primary" data-toggle="modal" data-target=".bs-example-modal-lg-customize">
  Customize
</button>
<% If UseSettings_Reports = True Then%>
<a href="<%= BaseURL %>bizintel/CustAnalSum_1_Customize_ClearValues.asp"><button type="button" class="btn btn-primary">Clear Customizations</button></a>
<% End If %>
	<!-- eof modal button !-->
</h3>

<!--#include file="CustAnalSum_1_Customize.asp"-->	
 


<h6 class="page-header">
	<table id="table-search" class='table table-striped table-condensed table-hover display'>
		<tr>
		
			<td>
				<%= GetTerm("Primary Salesman") %>: <b><% If FilterSlsmn1 = "" or FilterSlsmn1 = "All" Then %>All <%Else Response.Write(GetSalesmanNameBySlsmnSequence(FilterSlsmn1)) End If%></b><br>
				<%= GetTerm("Secondary Salesman") %>: <b><% If FilterSlsmn2 = "" or FilterSlsmn2 = "All"  Then %>All <%Else Response.Write(GetSalesmanNameBySlsmnSequence(FilterSlsmn2)) End If%></b><br>
				Referral:  <b><% If FilterReferral = "" or FilterReferral = "All"  Then %>All <%Else Response.Write(GetReferralNameByCode(FilterReferral)) End If%></b><br>
			</td>
		
			<td>Last closed period sales dollars is less than the prior three period average sales dollars by at least:&nbsp;<b><%=FormatCurrency(FilterSalesDollars,0)%></b><br><br>
				The difference between the last closed period sales vs the prior three preriods average sales represents at least:&nbsp;<b><%=FilterPercentage%>%</b><br>
			</td>
		
			<td>
				<div class="alert alert-info">
					<strong><u>Selection Criteria</u></strong><br>
					<strong>1. </strong> If Current (adjusted for days) >= 3Pavg or 12Pavg - Don't Show.<br>
					<strong>2. </strong> If LCP >= 12pAVG - Don't Show<br>
					<strong>3. </strong> If LCP >= SPLY - Don't Show<br>
					<strong>4. </strong> If 3PROI > 10 - Override Anything Else and Show
				</div>		
			</td>
		
		</tr>
	</table>
</h6>


<!-- row !-->
<div class="row">


<%
SQL = "SELECT Distinct CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
'SQL = SQL & ", Total3PPAvgAllCats - LCPTotSalesAllCats As ExprOrder "
SQL = SQL & " FROM CustCatPeriodSales "
SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
SQL = SQL & " AND LCPTotSalesAllCats < Total3PPAvgAllCats "
SQL = SQL & " AND Total3PPAvgAllCats - LCPTotSalesAllCats > " & FilterSalesDollars 
SQL = SQL & " AND (CASE WHEN Total3PPAvgAllCats <> 0 THEN (((LCPTotSalesAllCats  - Total3PPAvgAllCats ) / Total3PPAvgAllCats) * 100) * -1 END) >= " & FilterPercentage 

'Response.Write(now()&"<br>")
'Response.write(SQL)



Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
Set rs = cnn8.Execute(SQL)
'Response.Write(now()&"<br>")
%>

 

<!-- responsive tables !-->

<!--	
<div class="input-group"> <span class="input-group-addon">Narrow Results</span>
    <input id="filter" type="text" class="form-control filter-search-width" placeholder="Type here...">
</div><br>
!-->
<div class="container-fluid">
    <div class="row">
           <table id="tableSuperSum" class="display  compact" style="width:100%;">
              <thead>
                  <tr>	
						<th rowspan="2"  class="sorttable numeric smaller-header"><br>Acct</th>
                		<th rowspan="2"  class="sorttable numeric smaller-header"><br>Client</th>
						<th class="td-align1 vpc-variance-header" colspan="6" style="border-right: 2px solid #555 !important;">Variances</th>
						<th class="td-align1 vpc-3pavg-header" colspan="5" style="border-right: 2px solid #555 !important;">Sales</th>
						<th class="td-align1 vpc-lcp-header" colspan="5" style="border-right: 2px solid #555 !important;">MCS / MES</th>
						<th class="td-align1 vpc-current-header" colspan="2" style="border-right: 2px solid #555 !important;">ROI</th>
						<th class="td-align1 gen-info-header" colspan="3" style="border-right: 2px solid #555 !important;">General</th>

				</tr>
                <tr>
                  
                  
                  <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">LCP vs<br>3P avg $</th> 
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">LCP vs<br>3P avg %</th> 
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Day<br>Impact</th>  
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>ADS</th> 
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">LCP vs<br>12P avg $</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">LCP vs<br>12P avg %</th>
                  
                  
                  <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>LCP $</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>3P avg $</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>12P avg $</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Current $</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>SPLY $</th> 


                  <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>MCS/MES $</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">LCP vs<br> MxS</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">3P avg vs<br> MxS</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">12P avg vs<br> MxS</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Current vs<br> MxS</th>
                  
                  
                  <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">LCP<br>ROI</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">3P avg<br>ROI</th>
                  
                  <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">Primary<br> Slsmn</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Secondary<br> Slsmn</th>
                  <th class="td-align sorttable_numeric smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>Referral</th>
					<!--<th class="td-align sorttable_numeric smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>Snooze</th>-->

                </tr>
              </thead>
              
			

<%		
		'Response.Write("<tbody class='searchable'>")
        Response.Write("<tbody>")
		
		Do While Not rs.EOF

			ShowThisRecord = True

				
			If ShowThisRecord <> False Then			
			
				PrimarySalesMan =  ""
				SecondarySalesMan =  ""
				ReferralCode =  ""
				CustomerType =  ""
				SelectedCustomerID = rs("CustNum")
				CustName = GetCustNameByCustNum(SelectedCustomerID)	
				
				'Extra Fields for Filtering
				SQL4 = "SELECT * FROM AR_Customer WHERE CustNum = '" & SelectedCustomerID & "'"
				Set rs4 = Server.CreateObject("ADODB.Recordset")
				rs4.CursorLocation = 3
				Set rs4= cnn8.Execute(SQL4 )

				If Not rs4.Eof Then

					If rs4("AcctStatus") <> "A" Then ShowThisRecord = False

					PrimarySalesMan = rs4("Salesman")
					SecondarySalesMan = rs4("SecondarySalesman")
					ReferralCode = rs4("ReferalCode")
					CustomerType = rs4("CustType")

					'Decide if this record meets the filter criteria
					If FilterSlsmn1 <> "" And FilterSlsmn1 <> "All" Then
						If CInt(FilterSlsmn1) <> Cint(rs4("Salesman")) Then ShowThisRecord = False
					End If
					If FilterSlsmn2 <> "" And FilterSlsmn2 <> "All" Then
						If CInt(FilterSlsmn2) <> Cint(rs4("SecondarySalesman")) Then ShowThisRecord = False
					End If
					If FilterReferral <> "" And FilterReferral <> "All" Then
						If CInt(FilterReferral) <> Cint(rs4("ReferalCode")) Then ShowThisRecord = False
					End If
				
					Cust_MGP = rs4("ProjGpPerMonth")
					If rs4("ProjSalesPerMonth") <> "" Then Cust_MGPSales = FormatCurrency(rs4("ProjSalesPerMonth"),0) Else Cust_MGPSales =""
								
					MGPTerm = "" 
					' Determine what CCS is going to call it
					If Cust_MGPSales > 0 Then
						If cint(Cust_MGP) = 1 Then
							MGPTerm = "E" 
						Else
							MGPTerm = "C" 
						End If
					Else
						MGPTerm = ""
					End If
					If rs4("ProjSalesPerMonth") <> "" Then Cust_MGPSales = rs4("ProjSalesPerMonth")
					
				Else
					' Customer not found un AR_Customer
					ShowThisRecord = False
				End If

			End If
			
			
			If ShowThisRecord <> False Then
			
				'TotalCustsReported = TotalCustsReported + 1
				
				'Get everything we need for the report data
				SQLReportData = "SELECT SUM(TotalSales) AS LCPSales, SUM([3PriorPeriodsTotalSales]) AS ThreePPSales "
				SQLReportData = SQLReportData & ", SUM(PriorPeriod1Sales+PriorPeriod2Sales+PriorPeriod3Sales+PriorPeriod4Sales+PriorPeriod5Sales+PriorPeriod6Sales+ "
				SQLReportData = SQLReportData & " PriorPeriod7Sales+PriorPeriod8Sales+PriorPeriod9Sales+PriorPeriod10Sales+PriorPeriod11Sales+PriorPeriod12Sales) As TwelvePPSales "
				SQLReportData = SQLReportData & " FROM CustCatPeriodSales_ReportData WHERE CustNum = '" & SelectedCustomerID & "' AND ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated
				

				LCPSales = rs("LCPSales")
				If Not IsNumeric(LCPSales) Then LCPSales = 0
				ThreePPSales = rs("ThreePPSales")
				TwelvePPSales = rs("TwelvePPSales")
				CurrentPSales = GetCurrent_PostedTotal_ByCust(SelectedCustomerID,PeriodSeqBeingEvaluated) + GetCurrent_UnPostedTotal_ByCust(SelectedCustomerID,PeriodSeqBeingEvaluated)
				LCPvs3PAvgSales = LCPSales - (ThreePPSales/3)
				ImpactDays = (WorkDaysIn3PeriodBasis/3)- WorkDaysInLastClosedPeriod
				DayImpact = ImpactDays  * (LCPSales/WorkDaysInLastClosedPeriod)
				DayImpact = Round(DayImpact,2)
				ADS_LastClosed = (LCPSales/WorkDaysInLastClosedPeriod)
				ADS_3PA = ThreePPSales / (WorkDaysIn3PeriodBasis /3)
				ADS_Variance = ADS_LastClosed -  ADS_3PA 
				LCPvs12PAvgSales = LCPSales - (TwelvePPSales/12)
				If LCPvs12PAvgSales <> 0 Then LCPvs12PAvgPercent = ((LCPSales - LCPvs12PAvgSales) / LCPvs12PAvgSales)  * 100 Else LCPvs12PAvgPercent = 0
				SamePLYSales = TotalTPLYAllCats(PeriodSeqBeingEvaluated,SelectedCustomerID)
				If Not IsNumeric(SamePLYSales) Then SamePLYSales = 0
				ThreePPAvgSales = ThreePPSales / 3
				TwelvePPAvgSales = TwelvePPSales / 12
				If ThreePPAvgSales <> 0 Then LCPvs3PAvgPercent = ((LCPSales - ThreePPAvgSales ) / ThreePPAvgSales )  * 100  Else LCPvs3PAvgPercent = 0
				If MGPTerm <> "" Then LCPvsMxS = LCPSales - Cust_MGPSales
				If MGPTerm <> "" Then ThreePAvgVMxS = ThreePPAvgSales - Cust_MGPSales
				If MGPTerm <> "" Then TwelvePAvgVMxS = TwelvePPAvgSales - Cust_MGPSales
				If MGPTerm <> "" Then CurrentVMxS = CurrentPSales - Cust_MGPSales
				'ROI***********
				TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(SelectedCustomerID)
				'If CustHasEquipment(SelectedCustomerID) Then
				If TotalEquipmentValue > 0 Then	
					'LCPGP = LCPSales - TotalCostByPeriodSeq(PeriodSeqBeingEvaluated,SelectedCustomerID)
					LCPGP = LCPSales - rs("TotalCostAllCats")
					ThreePAvgGP = ThreePPAvgSales - ( TotalCostByPeriodSeqPrior3P(PeriodSeqBeingEvaluated,SelectedCustomerID) / 3 )
					If LCPGP <> 0 Then ROI = TotalEquipmentValue/LCPGP Else ROI = ""
					If ThreePAvgGP <> 0 Then ROI3P = TotalEquipmentValue/ThreePAvgGP Else ROI3P = ""
				End If

				' HERE ARE THE RULES
				'1. If Current (adjusted for days) => 3Pavg or 12Pavg - Dont show
				If CurrentPSales >= ThreePPAvgSales OR CurrentPSales >= TwelvePPAvgSales Then  'If the current is already greater or equal, we don't need to adjust for days, we're already ok
					ShowThisRecord = False
				Else  'We need to adjust for days & fiure it out
					ForecastedCurrent = (CurrentPSales / WorkDaysSoFar) * WorkDaysInCurrentPeriod 
					If ForecastedCurrent >= ThreePPAvgSales OR ForecastedCurrent >= TwelvePPAvgSales Then ShowThisRecord = False
				End If
					
				'2. If LCP => 12pAVG - Dont Show
				If LCPSales >= TwelvePPAvgSales Then ShowThisRecord = False
				
				'3. If LCP => SPLY - Dont Show
				If LCPSales >= SamePLYSales Then ShowThisRecord = False
				
				'4. If 3PROI > 10 - Override anything else and Show
				If Not Isnull(ROI3P) Then
					If IsNumeric(ROI3P) Then
						If ROI3P > 10 Then ShowThisRecord = True
					End If
				End If

				If ShowThisRecord <> False Then
				
					TotalCustsReported = TotalCustsReported + 1
					
					Response.Write("<tr>")
				    Response.Write("<td class='smaller-detail-line'><a href='tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>"& SelectedCustomerID  & "</a></td>")
				    Response.Write("<td class='smaller-detail-line'><a href='tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>"& CustName & "</a></td>")
		   		    Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(LCPvs3PAvgSales,0,-2,0) & "</td>")
				    Response.Write("<td align='right' class='smaller-detail-line'>" & FormatNumber(LCPvs3PAvgPercent,0,-2,0)  & "%</td>")
				    Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(DayImpact,0,-2,0) & "</td>")
				    Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ADS_Variance,0,-2,0) & "</td>")
		   		    Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(LCPvs12PAvgSales,0,-2,0) & "</td>")
				    Response.Write("<td align='right' class='smaller-detail-line'>" & FormatNumber(LCPvs12PAvgPercent,0,-2,0)  & "%</td>")
					Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(LCPSales,0,-2,0) & "</td>")
					Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ThreePPAvgSales,0,-2,0) & "</td>")
					Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(TwelvePPAvgSales,0,-2,0) & "</td>")
				   	Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(CurrentPSales,0,-2,0) & "</td>")
				   	Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(SamePLYSales,0,-2,0) & "</td>")
					'Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(CurrentPSales,0,-2,0) & " // " & FormatCurrency(ForecastedCurrent ,0,-2,0) & " // " & WorkDaysSoFar & " // " & WorkDaysInCurrentPeriod &  "</td>")	
					'Response.Write("<td align='right' class='smaller-detail-line'>" & MGPTerm & "</td>")
					If MGPTerm = "" Then 
						Response.Write("<td>&nbsp;</td>")
					Else
						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(Cust_MGPSales,0,-2,0) & "</td>")
					End If
					If MGPTerm = "" Then 
						Response.Write("<td>&nbsp;</td>")
					Else
						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(LCPvsMxS,0,-2,0) & "</td>")
					End If
					If MGPTerm = "" Then 
						Response.Write("<td>&nbsp;</td>")
					Else
						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ThreePAvgVMxS,0,-2,0) & "</td>")
					End If
					
					If MGPTerm = "" Then 
						Response.Write("<td>&nbsp;</td>")
					Else
						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ThreePAvgVMxS,0,-2,0) & "</td>")
					End If
	
	
					If MGPTerm = "" Then 
						Response.Write("<td>&nbsp;</td>")
					Else
						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(CurrentVMxS,0,-2,0) & "</td>")
					End If
	
					If TotalEquipmentValue > 0 Then	
						If IsNumeric(ROI) and IsNumeric(ROI3P) Then
							If ROI >=10 and ROI3P >= 10 Then ' If both over 10 use red
								Response.Write("<td align='right' class='negative smaller-detail-line'>" & FormatNumber(ROI,1)  & "</td>")
								Response.Write("<td align='right' class='negative smaller-detail-line'>" & FormatNumber(ROI3P,1)  & "</td>")
							Else
								If ROI <> "" Then
									Response.Write("<td align='right' class='smaller-detail-line'>" & FormatNumber(ROI,1)  & "</td>")
								Else
									Response.Write("<td align='right' class='smaller-detail-line'>No Sales</td>")
								End If
								Response.Write("<td align='right' class='smaller-detail-line'>" & FormatNumber(ROI3P,1)  & "</td>")
							End If
						Else
							If IsNumeric(ROI) Then
								Response.Write("<td align='right' class='smaller-detail-line'>" & FormatNumber(ROI,1)  & "</td>")
							Else
								Response.Write("<td align='right' class='smaller-detail-line'>No Sales</td>")
							End If
							If IsNumeric(ROI3P) Then
								Response.Write("<td align='right' class='smaller-detail-line'>" & FormatNumber(ROI3P,1)  & "</td>")
							Else
								Response.Write("<td>&nbsp;</td>")
							End If
						End If
					Else
						Response.Write("<td align='right' class='smaller-detail-line'>No</td>")
						Response.Write("<td align='right' class='smaller-detail-line'>Equipment</td>")
					End If
	
	
					' General info
					PrimarySalesPerson = GetSalesmanNameBySlsmnSequence(PrimarySalesMan)
				    SecondarySalesPerson = GetSalesmanNameBySlsmnSequence(SecondarySalesman)
				    If Instr(PrimarySalesPerson ," ") <> 0 Then
						Response.Write("<td class='smaller-detail-line'>" & Left(PrimarySalesPerson,Instr(PrimarySalesPerson ," ")+1) & "</td>")
					Else
						Response.Write("<td class='smaller-detail-line'>" & PrimarySalesPerson & "</td>")
					End If
					If Instr(SecondarySalesPerson," ") <> 0 Then
						Response.Write("<td class='smaller-detail-line'>" & Left(SecondarySalesPerson,Instr(SecondarySalesPerson," ")+1) & "</td>")
					Else
						Response.Write("<td class='smaller-detail-line'>" & SecondarySalesPerson & "</td>")
					End If

					
					
					
					Response.Write("<td class='smaller-detail-line'>" & GetReferralNameByCode(ReferralCode)  & "</td>")
	
					btncolor = "#fff;"
					'Response.Write("<td class='smaller-detail-line'>")
					'Response.Write "<button type=""button"" class=""" & btncolor & """ id=""btn" & SelectedCustomerID & """ data-toggle=""modal"" data-target=""#modalGeneralNotesGroupM"" >Snooze</button>"
					'Response.Write("</td>")
	                
				    Response.Write("</tr>")
			    
			    End If

			End If
			
			rs.movenext
				
		Loop
		
		Response.Write("</tbody>")
		Response.Write("</table>")		
		Response.Write("</div>")

'		Response.Write(now()&"<br>")
%>


            </table>
    </div>
         
          </div>
<!-- eof responsive tables !-->



<!-- eof row !-->

<!-- row !-->
<div class="row">
   <%
'Response.Write("<div class='col-lg-12'><h3>" & "Total Customers Listed:" & TotalCustsReported  & "</h3></div>")
%>

<div class="col-lg-12"><hr></div>
</div>
<!-- eof row !-->

<!-- row !-->
<div class="row">


<% 	rs.Close %>


</div>
<!-- eof row !-->
<!--#include file="../inc/footer-main.asp"-->