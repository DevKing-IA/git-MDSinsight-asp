<!--#include file="../inc/header.asp"-->
<!--#include file="../inc/jquery_table_search.asp"-->
<!--#include file="../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../inc/InSightFuncs_Equipment.asp"--> 
<!--#include file="../inc/InSightFuncs.asp"-->

<%
Server.ScriptTimeout = 900000 'Default value

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Leakage Overview"

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	If Request.Form("chkLimitSelection") = "on" Then LimitSelection = 1 Else LimitSelection = 0
Else
	'Default values
	LimitSelection = 0
	FilterSalesDollars = 100
	FilterPercentage = 10
End If

If LimitSelection = 1 Then
	FilterSalesDollars = 100
	FilterPercentage = 10
Else
	FilterSalesDollars = 0
	FilterPercentage = 0
End If

PeriodBeingEvaluated = GetLastClosedPeriodAndYear()
PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum()

WorkDaysIn3PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -3), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1
WorkDaysIn12PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -12), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1)) + 1 
WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated )) + 1 
WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated +1)) + 1 
WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1),Date()) + 1

WorkDaysInProjectionBasis =  (NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -2), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1) + WorkDaysInLastClosedPeriod 
%>

  
<style>

	
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


	.referral-header{
		background: #D43F3A;
		color:#fff;
		text-align:center;
		width:60%;
	}
	
	.cust-type-header{
		background: #F0AD4E;
		color:#fff;
		text-align:center;
		font-weight:bold;
		width:60%;
	}
	

	.primary-slsmn-header{
		background: #337AB7;
		color:#fff;
		text-align:center;
		font-weight:bold;
		width:60%;
	}

	.secondary-slsmn-header{
		background: #5CB85C;
		color:#fff;
		text-align:center;
		font-weight:bold;
		width:60%;
	}

	.gen-info-header{
		background: #3B579D;
		color:#fff;
		text-align:center;
		font-weight:bold;
		width:60%;
	}
	
	.dollar-amount-header{
		background: #808080;
		color:#fff;
		text-align:center;
		width:20%;
		font-size: 0.8em;
	}


	.percent-header{
		background: #808080;
		color:#fff;
		text-align:center;
		width:20%;
		font-size: 0.8em;
	}
		
	.dataTables_wrapper .dataTables_filter input {
	    margin-left: 0.5em;
	    box-shadow: none;
  		border-radius:6px;
		-webkit-border-radius: 6px;
		-moz-border-radius: 6px; 
	    padding: 3px;
	    border: solid 1px #E4E4E4;
	    background-color: #fff;		
	    margin-top:10px;
	    margin-bottom:10px;
	}	
	
	.negative{
		font-weight:normal;
		color:red;	
	}

	.neutral{
		font-weight:bold;
		color:black;
	}


	.smaller-detail-line{
		font-size: 0.8em;
	}	

	.fixed-col-header{
		color:#000;
		font-size: 1.1em;
		text-align:left;
		font-weight:bold;
		margin-bottom:10px;
		margin-left: -10px;
    	margin-right: 63px;
    	/*width: 350px;*/
	}

	.fixed-col {
	    height: 100%;
	    background-color: #fff;
	    text-align: center;
	    margin-right: 20px;
	    /*overflow-y: scroll;*/
	    border: solid 1px #000;
	    width: 430px;
	}
		
	.headerText {
		display: inline-block;
		text-align:left;
		vertical-align:middle;
		margin-left:10px;
	}

	.smaller-header{
		font-size: 0.8em;
		vertical-align: top !important;
		text-align: center;
	}	
	
	#chartdivSls2,
	#cchartdivSls1,
	#cchartdivCustType,
	#cchartdivRef{
		width: 100%;
		height: 100%;
	}	
	
</style>




<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.18/sorting/currency.js"></script>
<script type="text/javascript">

$(document).ready(function() {

		$("#chkLimitSelection").change(function() {
			$("#frmOverview").submit();
		});

	    $("#PleaseWaitPanel").hide();
	    
	   
	    $('#tableSuperSumPrimarySlsmn').DataTable({
 		    searching: false,
			scrollY: 500,
		    scrollCollapse: true,
		    paging: false,
		    info: false,		    
	        order: [[ 3, 'desc' ],[ 1, 'desc' ]],
   			columnDefs: [
			        { type: 'currency', targets: 1}
			    ]		        
	    });
	    
	    $('#tableSuperSumSecondarySlsmn').DataTable({
			searching: false,
		  	scrollY: 500,
		    scrollCollapse: true,
		    paging: false,
		    info: false,		    
	        order: [[ 3, 'desc' ],[ 1, 'desc' ]],
   			columnDefs: [
			        { type: 'currency', targets: 1}
			    ]		        
	    });

	    $('#tableSuperSumCustType').DataTable({
			searching: false,
		  	scrollY: 500,
		    scrollCollapse: true,
		    paging: false,
		    info: false,		    
	        order: [[ 3, 'desc' ],[ 1, 'desc' ]],
   			columnDefs: [
			        { type: 'currency', targets: 1}
			    ]		        
	    });

	    $('#tableSuperSumReferral').DataTable({
			searching: false,
		  	scrollY: 500,
		    scrollCollapse: true,
		    paging: false,
		    info: false,		    
	        order: [[ 3, 'desc' ],[ 1, 'desc' ]],
   			columnDefs: [
			        { type: 'currency', targets: 1}
			    ]		        
	    });
	    
    
});
</script>


<%
Response.Write("<div id=""PleaseWaitPanel"" class=""container"">")
Response.Write("<br><br>Creating Leakage Overview <br><br>Please wait...<br><br>")
Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
Response.Write("</div>")
Response.Flush()
%>

<h3 class="page-header"><i class="fa fa-pie-chart"></i> Leakage Overview For Period <%=PeriodBeingEvaluated %> vs. 3 Prior Periods</h3>

	

<%
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 

%>

<!-- row !-->

<div class="row">

<div class="container-fluid">
	
		<div class="col-sm-12">

			<div class="col-lg-3" style="padding-right:0px !important; padding-left:0px !important; ">	

				<div id="chartdivRef" style="width: 100%; height: 350px; margin: 0 auto"></div>

			</div>

			<div class="col-lg-3" style="padding-right:0px !important; padding-left:0px !important; ">	

				<div id="chartdivCustType" style="width: 100%; height: 350px; margin: 0 auto"></div>

			</div>
		
			<div class="col-lg-3" style="padding-right:0px !important; padding-left:0px !important; ">	

				<div id="chartdivSls1" style="width: 100%; height: 350px; margin: 0 auto"></div>

			</div>
			
			<div class="col-lg-3" style="padding-right:0px !important; padding-left:0px !important; ">	

				<div id="chartdivSls2" style="width: 100%; height: 350px; margin: 0 auto"></div>

			</div>

</div>
</div>
</div>
<br><br><br>
<div class="row">
<div class="container-fluid">
<div class="col-sm-12">
			<div class="col-sm-3 fixed-col">	
					
		      	<% 
		     	 	'Get all Referral Codes

					TotSalesRef = 0
					Tot3PAvgRef = 0
					TotDollarDiff =0
					TotalNegDiff = 0
					
					'***************************************************************
					' This is the one part that is different in the referral section
					'We get these  numbers here but all other sections will use them
					'There is no need to get the same number over & over again
					'***************************************************************					
					
					SQL = "SELECT SUM(TotalSales) AS PP1Sales FROM CustCatPeriodSales_ReportData WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated-1  
'Response.Write(SQL & "<br>")
					Set rs = cnn8.Execute(SQL)
					PP1Sales = rs("PP1Sales")
					SQL = "SELECT SUM(TotalSales) AS PP2Sales FROM CustCatPeriodSales_ReportData WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated-2  
'Response.Write(SQL & "<br>")
					Set rs = cnn8.Execute(SQL)
					PP2Sales = rs("PP2Sales")
					SQL = "SELECT SUM(TotalSales) AS PP3Sales FROM CustCatPeriodSales WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated-3 
'Response.Write(SQL & "<br>")
					Set rs = cnn8.Execute(SQL)
					PP3Sales = rs("PP3Sales")


		      		SQL = "SELECT SUM(TotalSales) AS TotSales "
		      		SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg"
					SQL = SQL & " ,Referal.Description2 As ReferralDesc2"
					SQL = SQL & " ,SUM(TotalSales+PriorPeriod1Sales+PriorPeriod2Sales) / 3 AS ProjectionBasis "
					SQL = SQL & " FROM CustCatPeriodSales_ReportData "
					SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
					SQL = SQL & " INNER JOIN Referal ON Referal.ReferalCode = AR_Customer.ReferalCode"
					SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
					If LimitSelection = 1 Then
						SQL = SQL & " AND TotalSales < [3PriorPeriodsAeverage] "
						SQL = SQL & " AND [3PriorPeriodsAeverage] - TotalSales > " & FilterSalesDollars 
						SQL = SQL & " AND (CASE WHEN [3PriorPeriodsAeverage] <> 0 THEN (((TotalSales - [3PriorPeriodsAeverage] ) / [3PriorPeriodsAeverage]) * 100) * -1 END) >= " & FilterPercentage 
					End If
					SQL = SQL & " GROUP BY Referal.Description2"
					SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"



					Set rs = cnn8.Execute(SQL)
					
					If not rs.EOF Then
					
						'Need to get totals first
						Do
							If rs("TotSales") - rs("Tot3PPAvg") < 0 Then TotalNegDiff = TotalNegDiff + (rs("TotSales") - rs("Tot3PPAvg"))
					
							TotDollarDiff = TotDollarDiff + ( rs("TotSales") - rs("Tot3PPAvg")) 
							TotSalesRef = TotSalesRef + rs("TotSales")
							Tot3PAvgRef = Tot3PAvgRef + rs("Tot3PPAvg")



							rs.MoveNext
						Loop While Not rs.Eof

		
						%>	
						<br>
						<table id="tableSuperSumReferral" class="display compact" style="width:100%;">
						
							<thead>
							  	<tr>	
									<th class="td-align1 referral-header" style="border-right: 2px solid #555 !important;">Referral Code</th>
									<th class="td-align1 dollar-amount-header" style="border-right: 2px solid #555 !important;">LCP vs 3Pavg</th>
									<th class="td-align1 dollar-amount-header" style="border-right: 2px solid #555 !important;">Projected P<%=GetLastClosedPeriod()+1%> vs 3Pavg</th>
									<th class="td-align1 percent-header" style="border-right: 2px solid #555 !important;">% of<br>Referral</th>
								</tr>
							</thead>
							
							<tbody>

								<%
								rs.MoveFirst
								
								ChartElementNumber = 1
								ChartDataReferral = ""
								ChartRemainder = 100
								amChartDataReferral = ""
								RemainderDollarDiff = 0
								NextPeriodProj = 0
								
								Do
								
									Response.Write("<tr>")
									
								    Response.Write("<td align='left' class='smaller-detail-line'><a href='dashboard_segment_referral.asp?r=" & rs("ReferralDesc2") & "' target='_blank'>"& rs("ReferralDesc2") & "</a></td>")									
									
									DollarDiff = rs("TotSales") - rs("Tot3PPAvg")
									
									If DollarDiff > 0 Then
										Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(DollarDiff,0,-1,-2) & "</td>")
									Else
										Response.Write("<td align='right' class='smaller-detail-line negative'>" & FormatCurrency(DollarDiff,0,-1,-2) & "</td>")
									End If
									
									'*****************
									' PROJECTION LOGIC
									'*****************
									P3PADS = rs("ProjectionBasis") / WorkDaysInProjectionBasis
									
									P3PSoFar = P3PADS  * WorkDaysSoFar 

									CurrentDollars = GetCurrent_PostedTotal_ByReferralDesc2(rs("ReferralDesc2")) + GetCurrent_UnPostedTotal_ByReferralDesc2(rs("ReferralDesc2"))
									
									CurrentADS = CurrentDollars / WorkDaysSoFar 
									
									ProjByRef = rs("Tot3PPAvg") -(CurrentADS * WorkDaysInCurrentPeriod)
									
									CurrentDiff = CurrentDollars - P3PSoFar 
									
									NextPeriodProj = NextPeriodProj + (CurrentADS * WorkDaysInCurrentPeriod)
									
									'*********************
									' END PROJECTION LOGIC
									'*********************
									'Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(CurrentDiff ,0,-1,-2))
									If ProjByRef >=0 Then
										Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ProjByRef ,0,-1,-2))
									Else
										Response.Write("<td align='right' class='smaller-detail-line negative'>" & FormatCurrency(ProjByRef ,0,-1,-2))
									End If
									Response.Write("</td>")

									If TotalNegDiff <> 0 Then ContributionPercent = (DollarDiff / TotalNegDiff ) * 100 Else ContributionPercent = 0 * 100
									
									Response.Write("<td align='right' class='smaller-detail-line'>" & Round(ContributionPercent,0) & "%</td>")
									Response.Write("</tr>")
									
									'Now handle the part for the chart (Hah! "The part for the chart")
									If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
										ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
										'am Charts
										amChartDataReferral = amChartDataReferral & "{'referral': '" & rs("ReferralDesc2") & "',"
										amChartDataReferral = amChartDataReferral &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
										amChartDataReferral = amChartDataReferral &  "'contribDollars': " & Round(DollarDiff ,0) & "}," 
										
									Else
									
											RemainderDollarDiff = RemainderDollarDiff + DollarDiff	
									End If
									
									ChartElementNumber = ChartElementNumber + 1
									rs.movenext
								Loop until rs.eof
								
								'am Charts
								amChartDataReferral = amChartDataReferral & "{'referral': 'Other',"
								amChartDataReferral = amChartDataReferral &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
								amChartDataReferral = amChartDataReferral &  "'contribDollars': " & Round((RemainderDollarDiff * -1) ,0) & "}" 
								
							End If
							
							
				      	%>
      		
					
					</tbody>
				
				</table><br>	

				<%Response.Write("<table>")
				Response.Write("<tr><td width='25%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='25%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='15%' class='smaller-detail-line'><b><u>Total</u></b></td>")
				Response.Write("<td width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='15%' class='smaller-detail-line'><b><u>ADS</u></b></td>")
				Response.Write("<td width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='5%' class='smaller-detail-line'><b><u>Days</u></b></td>")
				Response.Write("</tr>")
				
				Tot_P3PADS = Tot3PAvgRef / (WorkDaysIn3PeriodBasis / 3)
				WD_P3PADS = WorkDaysIn3PeriodBasis / 3
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Three prior periods avg total sales:&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'> " & FormatCurrency(Tot3PAvgRef ,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(Tot_P3PADS,0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WD_P3PADS,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Last closed period total sales:&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(TotSalesRef,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((TotSalesRef/WorkDaysInLastClosedPeriod),0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInLastClosedPeriod,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Last closed period .vs P3P avg :&nbsp;</td>")
				If TotDollarDiff >= 0 Then
					Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(TotDollarDiff,0,0) & "</td>")
				Else
					Response.Write("<td align='right' width='15%' class='smaller-detail-line negative'>" & FormatCurrency(TotDollarDiff,0,0) & "</td>")
				End If
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				If (TotSalesRef/WorkDaysInLastClosedPeriod) - Tot_P3PADS >= 0 Then
					Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((TotSalesRef/WorkDaysInLastClosedPeriod) - Tot_P3PADS,0,-1,-1) & "</td>")
				Else
					Response.Write("<td align='right' width='15%' class='smaller-detail-line negative'>" & FormatCurrency((TotSalesRef/WorkDaysInLastClosedPeriod) - Tot_P3PADS,0,-1,-1) & "</td>")				
				End If
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				If Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) >= 0 Then
					Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) & "</td>")
				Else
					Response.Write("<td align='right' width='5%' class='smaller-detail-line negative'>" & Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) & "</td>")				
				End If
				Response.Write("</tr>")
				Response.Write("<tr>")
				Response.Write("<tr><td align='left' colspan='2' width='50%' class='smaller-detail-line'>Projected Current Period :&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(NextPeriodProj,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")			
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((NextPeriodProj/WorkDaysInCurrentPeriod),0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInCurrentPeriod,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("</table>")%>


				
			</div>	
			


			
			<div class="col-sm-3 fixed-col">	
			
		      	<% 
		     	 	'Get all Custmoer Types
		     	 	
					TotSalesTyp = 0
					Tot3PAvgTyp = 0
					TotDollarDiff =0
					TotalNegDiff = 0


		      		SQL = "SELECT SUM(TotalSales) AS TotSales "
		      		SQL = SQL & ",SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3) As Tot3PPAvg"
		      		SQL = SQL & ",SUM(PriorPeriod1Sales) As TotPP1Sales"
		      		SQL = SQL & ",SUM(PriorPeriod2Sales) As TotPP2Sales"
		      		SQL = SQL & ",SUM(PriorPeriod3Sales) As TotPP3Sales"
					SQL = SQL & " ,SUM(TotalSales+PriorPeriod1Sales+PriorPeriod2Sales) / 3 AS ProjectionBasis "
					SQL = SQL & ",CustType  "
					SQL = SQL & " FROM CustCatPeriodSales_ReportData "
					SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
					SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
					If LimitSelection = 1 Then
						SQL = SQL & " AND TotalSales < [3PriorPeriodsAeverage] "
						SQL = SQL & " AND [3PriorPeriodsAeverage] - TotalSales > " & FilterSalesDollars 
						SQL = SQL & " AND (CASE WHEN [3PriorPeriodsAeverage] <> 0 THEN (((TotalSales - [3PriorPeriodsAeverage] ) / [3PriorPeriodsAeverage]) * 100) * -1 END) >= " & FilterPercentage 
					End If
					SQL = SQL & " GROUP BY AR_Customer.CustType"
					SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"


					Set rs = cnn8.Execute(SQL)
					
					If not rs.EOF Then

						'Need to get totals first
						Do
							If rs("TotSales") - rs("Tot3PPAvg") < 0 Then TotalNegDiff = TotalNegDiff + (rs("TotSales") - rs("Tot3PPAvg"))
					
							TotDollarDiff = TotDollarDiff + ( rs("TotSales") - rs("Tot3PPAvg")) 
							TotSalesTyp = TotSalesTyp + rs("TotSales")
							Tot3PAvgTyp = Tot3PAvgTyp + rs("Tot3PPAvg")

							rs.MoveNext
						Loop While Not rs.Eof
						%>		
						<br>
						<table id="tableSuperSumCustType" class="display compact" style="width:100%;">
						
							<thead>
							  	<tr>
									<th class="td-align1 cust-type-header" style="border-right: 2px solid #555 !important;">Customer Type</th>
									<th class="td-align1 dollar-amount-header" style="border-right: 2px solid #555 !important;">LCP vs 3P avg</th>
									<th class="td-align1 dollar-amount-header" style="border-right: 2px solid #555 !important;">Projected P<%=GetLastClosedPeriod()+1%> vs 3Pavg</th>									
									<th class="td-align1 percent-header" style="border-right: 2px solid #555 !important;">% of<br>Type</th>						
								</tr>
							</thead>
							
							<tbody>
								<%													
								
								rs.MoveFirst

								ChartElementNumber = 1
								ChartDataCustType = ""
								ChartRemainder = 100
								NextPeriodProj = 0
								
								Do
								
									Response.Write("<tr>")
								    Response.Write("<td align='left' class='smaller-detail-line'><a href='dashboard_segment_customertype.asp?t=" & rs("CustType") & "' target='_blank'>"& rs("CustType") & " - " & GetCustTypeByCode(rs("CustType")) & "</a></td>")									
																	    
									DollarDiff = rs("TotSales") - rs("Tot3PPAvg")

									If DollarDiff > 0 Then
										Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(DollarDiff,0,-1,-2) & "</td>")
									Else
										Response.Write("<td align='right' class='smaller-detail-line negative'>" & FormatCurrency(DollarDiff,0,-1,-2) & "</td>")
									End If
									
									'*****************
									' PROJECTION LOGIC
									'*****************

									P3PADS = rs("ProjectionBasis") / WorkDaysInProjectionBasis
									P3PSoFar = P3PADS  * WorkDaysSoFar 

									CurrentDollars = GetCurrent_PostedTotal_ByCustType(rs("CustType")) + GetCurrent_UnPostedTotal_ByCustType(rs("CustType"))
									
									CurrentADS = CurrentDollars / WorkDaysSoFar 
													
									ProjByTyp = rs("Tot3PPAvg") -(CurrentADS * WorkDaysInCurrentPeriod)				
									
									CurrentDiff = CurrentDollars - P3PSoFar 

									NextPeriodProj = NextPeriodProj + (CurrentADS * WorkDaysInCurrentPeriod)
									
									'*********************
									' END PROJECTION LOGIC
									'*********************
									
									'Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(CurrentDiff ,0,-1,-2)& "</td>")
									If ProjByTyp >= 0 Then
										Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ProjByTyp,0,-1,-2)& "</td>")
									Else
										Response.Write("<td align='right' class='smaller-detail-line negative'>" & FormatCurrency(ProjByTyp,0,-1,-2)& "</td>")
									End If
									
									If TotalNegDiff <> 0 Then ContributionPercent = (DollarDiff / TotalNegDiff ) * 100 Else ContributionPercent = 0 * 100
									Response.Write("<td align='right' class='smaller-detail-line'>" & Round(ContributionPercent,0) & "%</td>")
									Response.Write("</tr>")

									'Now handle the part for the chart (Hah! "The part for the chart")
									If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
										ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
										'am Charts
										amChartDataCustType  = amChartDataCustType  & "{'custtype': '" & GetCustTypeByCode(rs("CustType")) & "',"
										amChartDataCustType  = amChartDataCustType  &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
										amChartDataCustType  = amChartDataCustType  &  "'contribDollars': " & Round(DollarDiff ,0) & "}," 

										
									End If
									
									ChartElementNumber = ChartElementNumber + 1
									
									rs.movenext
								Loop until rs.eof
								
								'am Charts
								amChartDataCustType  = amChartDataCustType  & "{'custtype': 'Other',"
								amChartDataCustType  = amChartDataCustType  &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
								amChartDataCustType  = amChartDataCustType  &  "'contribDollars': " & Round((RemainderDollarDiff * -1) ,0) & "}" 
								
							End If

							
				      	%>
																														       		
					
					</tbody>
				
				</table><br>	
						
				<%Response.Write("<table>")
				Response.Write("<tr><td width='25%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='25%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='15%' class='smaller-detail-line'><b><u>Total</u></b></td>")
				Response.Write("<td width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='15%' class='smaller-detail-line'><b><u>ADS</u></b></td>")
				Response.Write("<td width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='5%' class='smaller-detail-line'><b><u>Days</u></b></td>")
				Response.Write("</tr>")
				
				Tot_P3PADS = Tot3PAvgTyp / (WorkDaysIn3PeriodBasis / 3)
				WD_P3PADS = WorkDaysIn3PeriodBasis / 3
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Three prior periods avg total sales:&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'> " & FormatCurrency(Tot3PAvgTyp ,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(Tot_P3PADS,0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WD_P3PADS,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Last closed period total sales:&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(TotSalesTyp,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((TotSalesTyp/WorkDaysInLastClosedPeriod),0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInLastClosedPeriod,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Last closed period .vs P3P avg :&nbsp;</td>")
				If TotDollarDiff >= 0 Then
					Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(TotDollarDiff,0,0) & "</td>")
				Else
					Response.Write("<td align='right' width='15%' class='smaller-detail-line negative'>" & FormatCurrency(TotDollarDiff,0,0) & "</td>")
				End If
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				If (TotSalesTyp/WorkDaysInLastClosedPeriod) - Tot_P3PADS >= 0 Then
					Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((TotSalesTyp/WorkDaysInLastClosedPeriod) - Tot_P3PADS,0,-1,-2) & "</td>")
				Else
					Response.Write("<td align='right' width='15%' class='smaller-detail-line negative'>" & FormatCurrency((TotSalesTyp/WorkDaysInLastClosedPeriod) - Tot_P3PADS,0,-1,-2) & "</td>")				
				End If
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				If Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) >= 0 Then
					Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) & "</td>")
				Else
					Response.Write("<td align='right' width='5%' class='smaller-detail-line negative'>" & Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) & "</td>")				
				End If
				Response.Write("</tr>")
				Response.Write("<tr>")
				Response.Write("<tr><td align='left' colspan='2' width='50%' class='smaller-detail-line'>Projected Current Period :&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(NextPeriodProj,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")			
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((NextPeriodProj/WorkDaysInCurrentPeriod),0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInCurrentPeriod,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("</table>")%>
		
			</div>	
	

		
			<div class="col-sm-3 fixed-col">	
					
		      	<% 
	     	 	'Get all Slsmn 1
	     	 	
				TotSalesSls1 = 0
				Tot3PAvgSls1 = 0
				TotDollarDiff =0
				TotalNegDiff = 0				     	 		
				
	      		SQL = "SELECT SUM(TotalSales) AS TotSales "
	      		SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg"
				SQL = SQL & " ,SUM(TotalSales+PriorPeriod1Sales+PriorPeriod2Sales) / 3 AS ProjectionBasis "
				SQL = SQL & ",Salesman "
				SQL = SQL & " FROM CustCatPeriodSales_ReportData "
				SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
				SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				If LimitSelection = 1 Then							
					SQL = SQL & " AND TotalSales < [3PriorPeriodsAeverage] "
					SQL = SQL & " AND [3PriorPeriodsAeverage] - TotalSales > " & FilterSalesDollars 
					SQL = SQL & " AND (CASE WHEN [3PriorPeriodsAeverage] <> 0 THEN (((TotalSales - [3PriorPeriodsAeverage] ) / [3PriorPeriodsAeverage]) * 100) * -1 END) >= " & FilterPercentage 
				End If
				SQL = SQL & " GROUP BY AR_Customer.Salesman"
				SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"

'Response.Write(SQL&"<br>")	
				Set rs = cnn8.Execute(SQL)
				
				If not rs.EOF Then

					'Need to get totals first
					Do
						If rs("TotSales") - rs("Tot3PPAvg") < 0 Then TotalNegDiff = TotalNegDiff + (rs("TotSales") - rs("Tot3PPAvg"))
				
						TotDollarDiff = TotDollarDiff + ( rs("TotSales") - rs("Tot3PPAvg")) 
						TotSalesSls1 = TotSalesSls1 + rs("TotSales")
						Tot3PAvgSls1 = Tot3PAvgSls1 + rs("Tot3PPAvg")

						rs.MoveNext
						
					Loop While Not rs.Eof
					%>
					<br>			
					<table id="tableSuperSumPrimarySlsmn" class="display compact" style="width:100%;">
					
						<thead>
						  	<tr>
								<th class="td-align1 primary-slsmn-header" style="border-right: 2px solid #555 !important;"><%= GetTerm("Primary Salesman") %></th>
								<th class="td-align1 dollar-amount-header" style="border-right: 2px solid #555 !important;">LCP vs 3P avg</th>
									<th class="td-align1 dollar-amount-header" style="border-right: 2px solid #555 !important;">Projected P<%=GetLastClosedPeriod()+1%> vs 3Pavg</th>								
								<th class="td-align1 percent-header" style="border-right: 2px solid #555 !important;">% of<br>Primary</th>						
							</tr>
						</thead>
						
						<tbody>
								
						<%								

								ChartElementNumber = 1
								ChartDataSls1 = ""
								ChartRemainder = 100
								NextPeriodProj = 0
								
								rs.MoveFirst
								
								Do
								
									Response.Write("<tr>")
								    Response.Write("<td align='left' class='smaller-detail-line'><a href='dashboard_segment_primarysalesman.asp?p=" & rs("Salesman") & "' target='_blank'>"& rs("Salesman") & " - " & GetSalesmanNameBySlsmnSequence(rs("Salesman")) & "</a></td>")									
									
									DollarDiff = rs("TotSales") - rs("Tot3PPAvg")

									If DollarDiff > 0 Then
										Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(DollarDiff,0,-1,-2) & "</td>")
									Else
										Response.Write("<td align='right' class='smaller-detail-line negative'>" & FormatCurrency(DollarDiff,0,-1,-2) & "</td>")
									End If
									
									
									'*****************
									' PROJECTION LOGIC
									'*****************

									P3PADS = rs("ProjectionBasis") / WorkDaysInProjectionBasis
									P3PSoFar = P3PADS  * WorkDaysSoFar 

									CurrentDollars = GetCurrent_PostedTotal_ByPrimary(rs("Salesman")) + GetCurrent_UnPostedTotal_ByPrimary(rs("Salesman"))

									CurrentADS = CurrentDollars / WorkDaysSoFar 									
									
									ProjBySls1 = rs("Tot3PPAvg") -(CurrentADS * WorkDaysInCurrentPeriod)	
									
									CurrentDiff = CurrentDollars - P3PSoFar 

									NextPeriodProj = NextPeriodProj + (CurrentADS * WorkDaysInCurrentPeriod)
									
									'*********************
									' END PROJECTION LOGIC
									'*********************

									
									'Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(CurrentDiff ,0,-1,-2)& "</td>")
									If ProjBySls1 >= 0 Then
										Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ProjBySls1 ,0,-1,-2)& "</td>")
									Else
										Response.Write("<td align='right' class='smaller-detail-line negative'>" & FormatCurrency(ProjBySls1 ,0,-1,-2)& "</td>")									
									End If
									
									
									If TotalNegDiff <> 0 Then ContributionPercent = (DollarDiff / TotalNegDiff ) * 100 Else ContributionPercent = 0 * 100
									Response.Write("<td align='right' class='smaller-detail-line'>" & Round(ContributionPercent,0) & "%</td>")
									Response.Write("</tr>")
									
									'Now handle the part for the chart (Hah! "The part for the chart")
									If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
										ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
										'am Charts
										If Instr(GetSalesmanNameBySlsmnSequence(rs("Salesman"))," ") <> 0 Then 
											amChartDataSls1  = amChartDataSls1  & "{'primary': '" & Left(GetSalesmanNameBySlsmnSequence(rs("Salesman")),Instr(GetSalesmanNameBySlsmnSequence(rs("Salesman"))," ")+1)  & "',"
										Else
											amChartDataSls1  = amChartDataSls1  & "{'primary': '" & GetSalesmanNameBySlsmnSequence(rs("Salesman")) & "',"										
										End If
										amChartDataSls1  = amChartDataSls1  &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
										amChartDataSls1  = amChartDataSls1  &  "'contribDollars': " & Round(DollarDiff ,0) & "}," 
									End If
									
									ChartElementNumber = ChartElementNumber + 1
									
									rs.movenext
								Loop until rs.eof
								
								'am Charts
								amChartDataSls1  = amChartDataSls1  & "{'primary': 'Other',"
								amChartDataSls1  = amChartDataSls1  &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
								amChartDataSls1  = amChartDataSls1  &  "'contribDollars': " & Round((RemainderDollarDiff * -1) ,0) & "}" 

							End If
								
				      	%>
										        		
					
					</tbody>
				
				</table><br>	


				<%Response.Write("<table>")
				Response.Write("<tr><td width='25%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='25%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='15%' class='smaller-detail-line'><b><u>Total</u></b></td>")
				Response.Write("<td width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='15%' class='smaller-detail-line'><b><u>ADS</u></b></td>")
				Response.Write("<td width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='5%' class='smaller-detail-line'><b><u>Days</u></b></td>")
				Response.Write("</tr>")
				
				Tot_P3PADS = Tot3PAvgSls1 / (WorkDaysIn3PeriodBasis / 3)
				WD_P3PADS = WorkDaysIn3PeriodBasis / 3
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Three prior periods avg total sales:&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'> " & FormatCurrency(Tot3PAvgSls1 ,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(Tot_P3PADS,0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WD_P3PADS,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Last closed period total sales:&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(TotSalesSls1,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((TotSalesSls1/WorkDaysInLastClosedPeriod),0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInLastClosedPeriod,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Last closed period .vs P3P avg :&nbsp;</td>")
				If TotDollarDiff >= 0 Then
					Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(TotDollarDiff,0,0) & "</td>")
				Else
					Response.Write("<td align='right' width='15%' class='smaller-detail-line negative'>" & FormatCurrency(TotDollarDiff,0,0) & "</td>")
				End If
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				If (TotSalesRef/WorkDaysInLastClosedPeriod) - Tot_P3PADS >= 0 Then
					Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((TotSalesRef/WorkDaysInLastClosedPeriod) - Tot_P3PADS,0,-1,-2) & "</td>")
				Else
					Response.Write("<td align='right' width='15%' class='smaller-detail-line negative'>" & FormatCurrency((TotSalesRef/WorkDaysInLastClosedPeriod) - Tot_P3PADS,0,-1,-2) & "</td>")				
				End If
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				If Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) >= 0 Then
					Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) & "</td>")
				Else
					Response.Write("<td align='right' width='5%' class='smaller-detail-line negative'>" & Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) & "</td>")				
				End If
				Response.Write("</tr>")
				Response.Write("<tr>")
				Response.Write("<tr><td align='left' colspan='2' width='50%' class='smaller-detail-line'>Projected Current Period :&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(NextPeriodProj,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")			
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((NextPeriodProj/WorkDaysInCurrentPeriod),0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInCurrentPeriod,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("</table>")%>

		
			</div>	
		
		
		
		
			<div class="col-sm-3 fixed-col">	
		      	<% 
	     	 	'Get all Slsmn 2
	     	 	
				TotSalesSls2 = 0
				Tot3PAvgSls2 = 0
				TotDollarDiff =0
				TotalNegDiff = 0

	      		SQL = "SELECT SUM(TotalSales) AS TotSales "
	      		SQL = SQL & ",SUM([3PriorPeriodsAeverage]) As Tot3PPAvg"
				SQL = SQL & " ,SUM(TotalSales+PriorPeriod1Sales+PriorPeriod2Sales) / 3 AS ProjectionBasis "
				SQL = SQL & ",SecondarySalesman "
				SQL = SQL & " FROM CustCatPeriodSales_ReportData "
				SQL = SQL & " INNER JOIN AR_Customer ON AR_Customer.CustNum = CustCatPeriodSales_ReportData.CustNum"
				SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				If LimitSelection = 1 Then
					SQL = SQL & " AND TotalSales < [3PriorPeriodsAeverage] "
					SQL = SQL & " AND [3PriorPeriodsAeverage] - TotalSales > " & FilterSalesDollars 
					SQL = SQL & " AND (CASE WHEN [3PriorPeriodsAeverage] <> 0 THEN (((TotalSales - [3PriorPeriodsAeverage] ) / [3PriorPeriodsAeverage]) * 100) * -1 END) >= " & FilterPercentage 
				End If
				SQL = SQL & " GROUP BY AR_Customer.SecondarySalesman "
				SQL = SQL & " ORDER BY (SUM(TotalSales)- SUM(([PriorPeriod1Sales]+[PriorPeriod2Sales]+[PriorPeriod3Sales])/3))"

'Response.Write(SQL&"<br>")	
				Set rs = cnn8.Execute(SQL)
				
				If not rs.EOF Then

					'Need to get totals first
					Do
						If rs("TotSales") - rs("Tot3PPAvg") < 0 Then TotalNegDiff = TotalNegDiff + (rs("TotSales") - rs("Tot3PPAvg"))
				
						TotDollarDiff = TotDollarDiff + ( rs("TotSales") - rs("Tot3PPAvg")) 
						TotSalesSls2 = TotSalesSls2 + rs("TotSales")
						Tot3PAvgSls2 = Tot3PAvgSls2 + rs("Tot3PPAvg")
						
						rs.MoveNext
						
					Loop While Not rs.Eof

					%>
					
					<br>								
					<table id="tableSuperSumSecondarySlsmn" class="display compact" style="width:100%;">
				
					<thead>
					  	<tr>
							<th class="td-align1 secondary-slsmn-header" style="border-right: 2px solid #555 !important;"><%= GetTerm("Secondary Salesman") %></th>
							<th class="td-align1 dollar-amount-header" style="border-right: 2px solid #555 !important;">LCP vs 3P avg</th>
							<th class="td-align1 dollar-amount-header" style="border-right: 2px solid #555 !important;">Projected P<%=GetLastClosedPeriod()+1%> vs 3Pavg</th>						
							<th class="td-align1 percent-header" style="border-right: 2px solid #555 !important;">% of<br>Secondary</th>							
						</tr>
					</thead>
					
					<tbody>
					

								
								<%
						
								ChartElementNumber = 1
								ChartDataSls2 = ""
								ChartRemainder = 100
								NextPeriodProj = 0

								rs.MoveFirst
															
								Do
								
									Response.Write("<tr>")
									Response.Write("<td align='left' class='smaller-detail-line'><a href='dashboard_segment_secondarysalesman.asp?p=" & rs("SecondarySalesman") & "' target='_blank'>"& rs("SecondarySalesman") & " - " & GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman")) & "</a></td>")									
									
									DollarDiff = rs("TotSales") - rs("Tot3PPAvg")

									If DollarDiff > 0 Then
										Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(DollarDiff,0,-1,-2) & "</td>")
									Else
										Response.Write("<td align='right' class='smaller-detail-line negative'>" & FormatCurrency(DollarDiff,0,-1,-2) & "</td>")
									End If
									
									
									
									'*****************
									' PROJECTION LOGIC
									'*****************

									P3PADS = rs("ProjectionBasis") / WorkDaysInProjectionBasis
									P3PSoFar = P3PADS  * WorkDaysSoFar 

									CurrentDollars = GetCurrent_PostedTotal_BySecondary(rs("SecondarySalesman")) + GetCurrent_UnPostedTotal_BySecondary(rs("SecondarySalesman"))

									CurrentADS = CurrentDollars / WorkDaysSoFar 									
									
									ProjBySls2 = rs("Tot3PPAvg") -(CurrentADS * WorkDaysInCurrentPeriod)	
																		
									CurrentDiff = CurrentDollars - P3PSoFar 

									NextPeriodProj = NextPeriodProj + (CurrentADS * WorkDaysInCurrentPeriod)
									
									'*********************
									' END PROJECTION LOGIC
									'*********************

									
									'Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(CurrentDiff ,0,-1,-2)& "</td>")
									If ProjBySls2 >=0 Then
										Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ProjBySls2 ,0,-1,-2)& "</td>")
									Else
										Response.Write("<td align='right' class='smaller-detail-line negative'>" & FormatCurrency(ProjBySls2 ,0,-1,-2)& "</td>")									
									End If
									
									
									If TotalNegDiff <> 0 Then ContributionPercent = (DollarDiff / TotalNegDiff ) * 100 Else ContributionPercent = 0 * 100
									Response.Write("<td align='right' class='smaller-detail-line'>" & Round(ContributionPercent,0) & "%</td>")
									Response.Write("</tr>")
									
									
									'Now handle the part for the chart (Hah! "The part for the chart")
									If ChartElementNumber < 6 and Round(ContributionPercent) > 9.99 Then 
										ChartRemainder = Round(ChartRemainder - ContributionPercent ,0)
										'am Charts
										If Instr(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman"))," ") <> 0 Then 										
											amChartDataSls2  = amChartDataSls2  & "{'secondary': '" & Left(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman")),Instr(GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman"))," ")+1) & "',"
										Else
											amChartDataSls2  = amChartDataSls2  & "{'secondary': '" & GetSalesmanNameBySlsmnSequence(rs("SecondarySalesman")) & "',"										
										End If
										amChartDataSls2  = amChartDataSls2  &  "'contribPercent': " & Round(ContributionPercent ,0) & "," 
										amChartDataSls2  = amChartDataSls2  &  "'contribDollars': " & Round(DollarDiff ,0) & "}," 
										
									End If
									
									ChartElementNumber = ChartElementNumber + 1


									rs.movenext
								Loop until rs.eof
								'am Charts
								amChartDataSls2  = amChartDataSls2  & "{'secondary': 'Other',"
								amChartDataSls2  = amChartDataSls2  &  "'contribPercent': " & Round(ChartRemainder ,0) & ", " 
								amChartDataSls2  = amChartDataSls2  &  "'contribDollars': " & Round((RemainderDollarDiff * -1) ,0) & "}" 

								
							End If

					      	%>
         		
					
					</tbody>
				
				</table><br>
					

				<%Response.Write("<table>")
				Response.Write("<tr><td width='25%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='25%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='15%' class='smaller-detail-line'><b><u>Total</u></b></td>")
				Response.Write("<td width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='15%' class='smaller-detail-line'><b><u>ADS</u></b></td>")
				Response.Write("<td width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td width='5%' class='smaller-detail-line'><b><u>Days</u></b></td>")
				Response.Write("</tr>")
				
				Tot_P3PADS = Tot3PAvgSls2 / (WorkDaysIn3PeriodBasis / 3)
				WD_P3PADS = WorkDaysIn3PeriodBasis / 3
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Three prior periods avg total sales:&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'> " & FormatCurrency(Tot3PAvgSls2 ,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(Tot_P3PADS,0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WD_P3PADS,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Last closed period total sales:&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(TotSalesSls2,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((TotSalesSls2/WorkDaysInLastClosedPeriod),0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInLastClosedPeriod,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("<tr>")
				Response.Write("<td align='left' colspan='2' width='50%' class='smaller-detail-line'>Last closed period .vs P3P avg :&nbsp;</td>")
				If TotDollarDiff >= 0 Then
					Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(TotDollarDiff,0,0) & "</td>")
				Else
					Response.Write("<td align='right' width='15%' class='smaller-detail-line negative'>" & FormatCurrency(TotDollarDiff,0,0) & "</td>")
				End If
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")
				If (TotSalesRef/WorkDaysInLastClosedPeriod) - Tot_P3PADS >= 0 Then
					Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((TotSalesRef/WorkDaysInLastClosedPeriod) - Tot_P3PADS,0,-1,-2) & "</td>")
				Else
					Response.Write("<td align='right' width='15%' class='smaller-detail-line negative'>" & FormatCurrency((TotSalesRef/WorkDaysInLastClosedPeriod) - Tot_P3PADS,0,-1,-2) & "</td>")				
				End If
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				If Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) >= 0 Then
					Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) & "</td>")
				Else
					Response.Write("<td align='right' width='5%' class='smaller-detail-line negative'>" & Round(WorkDaysInLastClosedPeriod,0) - Round(WD_P3PADS,0) & "</td>")				
				End If
				Response.Write("</tr>")
				Response.Write("<tr>")
				Response.Write("<tr><td align='left' colspan='2' width='50%' class='smaller-detail-line'>Projected Current Period :&nbsp;</td>")
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency(NextPeriodProj,0,0) & "</td>")
				Response.Write("<td align='right' width='3%' class='smaller-detail-line'>&nbsp;</td>")			
				Response.Write("<td align='right' width='15%' class='smaller-detail-line'>" & FormatCurrency((NextPeriodProj/WorkDaysInCurrentPeriod),0,0) & "</td>")
				Response.Write("<td align='right' width='2%' class='smaller-detail-line'>&nbsp;</td>")
				Response.Write("<td align='right' width='5%' class='smaller-detail-line'>" & Round(WorkDaysInCurrentPeriod,0) & "</td>")
				Response.Write("</tr>")
				
				Response.Write("</table>")%>

	
			</div>	
		
		
		
		
	</div>
	</div>
	         
</div>





<%


Function GetCurrent_PostedTotal_ByReferralDesc2(passedReferralDesc2)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	
	resultGetCurrent_PostedTotal_ByReferralDesc2 = 0

	Set cnnGetCurrent_PostedTotal_ByReferralDesc2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByReferralDesc2.open Session("ClientCnnString")
		

	SQLGetCurrent_PostedTotal_ByReferralDesc2 = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_PostedTotal_ByReferralDesc2 = SQLGetCurrent_PostedTotal_ByReferralDesc2 & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_PostedTotal_ByReferralDesc2 = SQLGetCurrent_PostedTotal_ByReferralDesc2 & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_PostedTotal_ByReferralDesc2 = SQLGetCurrent_PostedTotal_ByReferralDesc2 & " (SELECT CustNum FROM AR_Customer WHERE ReferalCode IN (SELECT ReferalCode FROM Referal WHERE Description2 = '" & passedReferralDesc2 & "'))"


	Set rsGetCurrent_PostedTotal_ByReferralDesc2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByReferralDesc2.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByReferralDesc2 = cnnGetCurrent_PostedTotal_ByReferralDesc2.Execute(SQLGetCurrent_PostedTotal_ByReferralDesc2)

	If not rsGetCurrent_PostedTotal_ByReferralDesc2.EOF Then resultGetCurrent_PostedTotal_ByReferralDesc2 = rsGetCurrent_PostedTotal_ByReferralDesc2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByReferralDesc2) Then resultGetCurrent_PostedTotal_ByReferralDesc2 = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByReferralDesc2.Close
	set rsGetCurrent_PostedTotal_ByReferralDesc2= Nothing
	cnnGetCurrent_PostedTotal_ByReferralDesc2.Close	
	set cnnGetCurrent_PostedTotal_ByReferralDesc2= Nothing

	
	GetCurrent_PostedTotal_ByReferralDesc2 = resultGetCurrent_PostedTotal_ByReferralDesc2

End Function

Function GetCurrent_UnPostedTotal_ByReferralDesc2(passedReferralDesc2)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	
	resultGetCurrent_UnPostedTotal_ByReferralDesc2 = 0

	Set cnnGetCurrent_UnPostedTotal_ByReferralDesc2 = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnPostedTotal_ByReferralDesc2.open Session("ClientCnnString")
		

	SQLGetCurrent_UnPostedTotal_ByReferralDesc2 = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_UnPostedTotal_ByReferralDesc2 = SQLGetCurrent_UnPostedTotal_ByReferralDesc2 & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_UnPostedTotal_ByReferralDesc2 = SQLGetCurrent_UnPostedTotal_ByReferralDesc2 & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_UnPostedTotal_ByReferralDesc2 = SQLGetCurrent_UnPostedTotal_ByReferralDesc2 & " (SELECT CustNum FROM AR_Customer WHERE ReferalCode IN (SELECT ReferalCode FROM Referal WHERE Description2 = '" & passedReferralDesc2 & "'))"

	Set rsGetCurrent_UnPostedTotal_ByReferralDesc2 = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnPostedTotal_ByReferralDesc2.CursorLocation = 3 
	Set rsGetCurrent_UnPostedTotal_ByReferralDesc2 = cnnGetCurrent_UnPostedTotal_ByReferralDesc2.Execute(SQLGetCurrent_UnPostedTotal_ByReferralDesc2)

	If not rsGetCurrent_UnPostedTotal_ByReferralDesc2.EOF Then resultGetCurrent_UnPostedTotal_ByReferralDesc2 = rsGetCurrent_UnPostedTotal_ByReferralDesc2("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnPostedTotal_ByReferralDesc2) Then resultGetCurrent_UnPostedTotal_ByReferralDesc2 = 0 ' In case there are no results
	
	rsGetCurrent_UnPostedTotal_ByReferralDesc2.Close
	set rsGetCurrent_UnPostedTotal_ByReferralDesc2= Nothing
	cnnGetCurrent_UnPostedTotal_ByReferralDesc2.Close	
	set cnnGetCurrent_UnPostedTotal_ByReferralDesc2= Nothing

	
	GetCurrent_UnPostedTotal_ByReferralDesc2 = resultGetCurrent_UnPostedTotal_ByReferralDesc2

End Function


Function GetCurrent_PostedTotal_ByCustType(passedCustType)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	
	resultGetCurrent_PostedTotal_ByCustType = 0

	Set cnnGetCurrent_PostedTotal_ByCustType = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByCustType.open Session("ClientCnnString")
		

	SQLGetCurrent_PostedTotal_ByCustType = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_PostedTotal_ByCustType = SQLGetCurrent_PostedTotal_ByCustType & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_PostedTotal_ByCustType = SQLGetCurrent_PostedTotal_ByCustType & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_PostedTotal_ByCustType = SQLGetCurrent_PostedTotal_ByCustType & " (SELECT CustNum FROM AR_Customer WHERE CustType ='" & passedCustType & "')"


	Set rsGetCurrent_PostedTotal_ByCustType = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByCustType.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByCustType = cnnGetCurrent_PostedTotal_ByCustType.Execute(SQLGetCurrent_PostedTotal_ByCustType)

	If not rsGetCurrent_PostedTotal_ByCustType.EOF Then resultGetCurrent_PostedTotal_ByCustType = rsGetCurrent_PostedTotal_ByCustType("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByCustType) Then resultGetCurrent_PostedTotal_ByCustType = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByCustType.Close
	set rsGetCurrent_PostedTotal_ByCustType= Nothing
	cnnGetCurrent_PostedTotal_ByCustType.Close	
	set cnnGetCurrent_PostedTotal_ByCustType= Nothing

	
	GetCurrent_PostedTotal_ByCustType = resultGetCurrent_PostedTotal_ByCustType

End Function

Function GetCurrent_UnPostedTotal_ByCustType(passedCustType)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	resultGetCurrent_UnPostedTotal_ByCustType = 0

	Set cnnGetCurrent_UnPostedTotal_ByCustType = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnPostedTotal_ByCustType.open Session("ClientCnnString")
		

	SQLGetCurrent_UnPostedTotal_ByCustType = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_UnPostedTotal_ByCustType = SQLGetCurrent_UnPostedTotal_ByCustType & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_UnPostedTotal_ByCustType = SQLGetCurrent_UnPostedTotal_ByCustType & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_UnPostedTotal_ByCustType = SQLGetCurrent_UnPostedTotal_ByCustType & " (SELECT CustNum FROM AR_Customer WHERE CustType ='" & passedCustType & "')"

	Set rsGetCurrent_UnPostedTotal_ByCustType = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnPostedTotal_ByCustType.CursorLocation = 3 
	Set rsGetCurrent_UnPostedTotal_ByCustType = cnnGetCurrent_UnPostedTotal_ByCustType.Execute(SQLGetCurrent_UnPostedTotal_ByCustType)

	If not rsGetCurrent_UnPostedTotal_ByCustType.EOF Then resultGetCurrent_UnPostedTotal_ByCustType = rsGetCurrent_UnPostedTotal_ByCustType("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnPostedTotal_ByCustType) Then resultGetCurrent_UnPostedTotal_ByCustType = 0 ' In case there are no results
	
	rsGetCurrent_UnPostedTotal_ByCustType.Close
	set rsGetCurrent_UnPostedTotal_ByCustType= Nothing
	cnnGetCurrent_UnPostedTotal_ByCustType.Close	
	set cnnGetCurrent_UnPostedTotal_ByCustType= Nothing

	
	GetCurrent_UnPostedTotal_ByCustType = resultGetCurrent_UnPostedTotal_ByCustType

End Function


Function GetCurrent_PostedTotal_ByPrimary(passedPrimary)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	
	resultGetCurrent_PostedTotal_ByPrimary = 0

	Set cnnGetCurrent_PostedTotal_ByPrimary = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByPrimary.open Session("ClientCnnString")
		

	SQLGetCurrent_PostedTotal_ByPrimary = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_PostedTotal_ByPrimary = SQLGetCurrent_PostedTotal_ByPrimary & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_PostedTotal_ByPrimary = SQLGetCurrent_PostedTotal_ByPrimary & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_PostedTotal_ByPrimary = SQLGetCurrent_PostedTotal_ByPrimary & " (SELECT CustNum FROM AR_Customer WHERE Salesman  ='" & passedPrimary & "')"


	Set rsGetCurrent_PostedTotal_ByPrimary = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByPrimary.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByPrimary = cnnGetCurrent_PostedTotal_ByPrimary.Execute(SQLGetCurrent_PostedTotal_ByPrimary)

	If not rsGetCurrent_PostedTotal_ByPrimary.EOF Then resultGetCurrent_PostedTotal_ByPrimary = rsGetCurrent_PostedTotal_ByPrimary("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByPrimary) Then resultGetCurrent_PostedTotal_ByPrimary = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByPrimary.Close
	set rsGetCurrent_PostedTotal_ByPrimary= Nothing
	cnnGetCurrent_PostedTotal_ByPrimary.Close	
	set cnnGetCurrent_PostedTotal_ByPrimary= Nothing

	
	GetCurrent_PostedTotal_ByPrimary = resultGetCurrent_PostedTotal_ByPrimary

End Function

Function GetCurrent_UnPostedTotal_ByPrimary(passedPrimary)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	resultGetCurrent_UnPostedTotal_ByPrimary = 0

	Set cnnGetCurrent_UnPostedTotal_ByPrimary = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnPostedTotal_ByPrimary.open Session("ClientCnnString")
		

	SQLGetCurrent_UnPostedTotal_ByPrimary = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_UnPostedTotal_ByPrimary = SQLGetCurrent_UnPostedTotal_ByPrimary & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_UnPostedTotal_ByPrimary = SQLGetCurrent_UnPostedTotal_ByPrimary & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_UnPostedTotal_ByPrimary = SQLGetCurrent_UnPostedTotal_ByPrimary & " (SELECT CustNum FROM AR_Customer WHERE Salesman  ='" & passedPrimary & "')"

	Set rsGetCurrent_UnPostedTotal_ByPrimary = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnPostedTotal_ByPrimary.CursorLocation = 3 
	Set rsGetCurrent_UnPostedTotal_ByPrimary = cnnGetCurrent_UnPostedTotal_ByPrimary.Execute(SQLGetCurrent_UnPostedTotal_ByPrimary)

	If not rsGetCurrent_UnPostedTotal_ByPrimary.EOF Then resultGetCurrent_UnPostedTotal_ByPrimary = rsGetCurrent_UnPostedTotal_ByPrimary("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnPostedTotal_ByPrimary) Then resultGetCurrent_UnPostedTotal_ByPrimary = 0 ' In case there are no results
	
	rsGetCurrent_UnPostedTotal_ByPrimary.Close
	set rsGetCurrent_UnPostedTotal_ByPrimary= Nothing
	cnnGetCurrent_UnPostedTotal_ByPrimary.Close	
	set cnnGetCurrent_UnPostedTotal_ByPrimary= Nothing

	
	GetCurrent_UnPostedTotal_ByPrimary = resultGetCurrent_UnPostedTotal_ByPrimary

End Function



Function GetCurrent_PostedTotal_BySecondary(passedSecondary)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	
	resultGetCurrent_PostedTotal_BySecondary = 0

	Set cnnGetCurrent_PostedTotal_BySecondary = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_BySecondary.open Session("ClientCnnString")
		

	SQLGetCurrent_PostedTotal_BySecondary = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_PostedTotal_BySecondary = SQLGetCurrent_PostedTotal_BySecondary & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_PostedTotal_BySecondary = SQLGetCurrent_PostedTotal_BySecondary & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_PostedTotal_BySecondary = SQLGetCurrent_PostedTotal_BySecondary & " (SELECT CustNum FROM AR_Customer WHERE SecondarySalesman  ='" & passedSecondary & "')"


	Set rsGetCurrent_PostedTotal_BySecondary = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_BySecondary.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_BySecondary = cnnGetCurrent_PostedTotal_BySecondary.Execute(SQLGetCurrent_PostedTotal_BySecondary)

	If not rsGetCurrent_PostedTotal_BySecondary.EOF Then resultGetCurrent_PostedTotal_BySecondary = rsGetCurrent_PostedTotal_BySecondary("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_BySecondary) Then resultGetCurrent_PostedTotal_BySecondary = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_BySecondary.Close
	set rsGetCurrent_PostedTotal_BySecondary= Nothing
	cnnGetCurrent_PostedTotal_BySecondary.Close	
	set cnnGetCurrent_PostedTotal_BySecondary= Nothing

	
	GetCurrent_PostedTotal_BySecondary = resultGetCurrent_PostedTotal_BySecondary

End Function

Function GetCurrent_UnPostedTotal_BySecondary(passedSecondary)

	LCPvar = GetLastClosedPeriodSeqNum() + 1 ' To get to the current period
	resultGetCurrent_UnPostedTotal_BySecondary = 0

	Set cnnGetCurrent_UnPostedTotal_BySecondary = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnPostedTotal_BySecondary.open Session("ClientCnnString")
		

	SQLGetCurrent_UnPostedTotal_BySecondary = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE "
	SQLGetCurrent_UnPostedTotal_BySecondary = SQLGetCurrent_UnPostedTotal_BySecondary & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & LCPvar 
	SQLGetCurrent_UnPostedTotal_BySecondary = SQLGetCurrent_UnPostedTotal_BySecondary & " AND BI_PostedUnpostedByCustCatPeriod.CustID IN "
	SQLGetCurrent_UnPostedTotal_BySecondary = SQLGetCurrent_UnPostedTotal_BySecondary & " (SELECT CustNum FROM AR_Customer WHERE SecondarySalesman  ='" & passedSecondary & "')"

	Set rsGetCurrent_UnPostedTotal_BySecondary = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnPostedTotal_BySecondary.CursorLocation = 3 
	Set rsGetCurrent_UnPostedTotal_BySecondary = cnnGetCurrent_UnPostedTotal_BySecondary.Execute(SQLGetCurrent_UnPostedTotal_BySecondary)

	If not rsGetCurrent_UnPostedTotal_BySecondary.EOF Then resultGetCurrent_UnPostedTotal_BySecondary = rsGetCurrent_UnPostedTotal_BySecondary("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnPostedTotal_BySecondary) Then resultGetCurrent_UnPostedTotal_BySecondary = 0 ' In case there are no results
	
	rsGetCurrent_UnPostedTotal_BySecondary.Close
	set rsGetCurrent_UnPostedTotal_BySecondary= Nothing
	cnnGetCurrent_UnPostedTotal_BySecondary.Close	
	set cnnGetCurrent_UnPostedTotal_BySecondary= Nothing

	
	GetCurrent_UnPostedTotal_BySecondary = resultGetCurrent_UnPostedTotal_BySecondary

End Function

%>
 
<!-- chart js !-->

	<script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/8.3/highlight.min.js"></script>		

    <!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
    <script src="<%= BaseURL %>js/ie10-viewport-bug-workaround.js"></script>
    

<!-- tooltip JS !-->
<script type="text/javascript">
$(function () {
  $('[data-toggle="tooltip"]').tooltip()
})
 </script>

<!-- eof chart js !-->

<!-- am chart js !-->
<!-- Styles -->
<style>
#chartdiv {
  width: 100%;
  height: 100%;
}
</style>

<!-- Resources -->
<script src="https://www.amcharts.com/lib/3/amcharts.js"></script>
<script src="https://www.amcharts.com/lib/3/pie.js"></script>
<script src="https://www.amcharts.com/lib/3/plugins/export/export.min.js"></script>
<link rel="stylesheet" href="https://www.amcharts.com/lib/3/plugins/export/export.css" type="text/css" media="all" />
<script src="https://www.amcharts.com/lib/3/themes/light.js"></script>

<%
 balloon = "35,000"
%>

<!-- Chart code -->
<script>
	
	var chart = AmCharts.makeChart( "chartdivRef", {
	
		"titles": [
				{
					"text": "Referral Code",
					"size": 18,
					"bold": "false"
				},
				{
					"text": "Leakage",
					"size": 18,
					"bold": "false"
				}
				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[referral]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataReferral%> ],
		"valueField": "contribPercent",
		"titleField": "referral",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",
		"pullOutRadius": 0,
		"balloonText": "[[referral]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});
	
	var chart = AmCharts.makeChart( "chartdivCustType", {
	
		"titles": [
				{
					"text": "Cust Type",
					"size": 18,
					"bold": "false"
				},
				{
					"text": "Leakage",
					"size": 18,
					"bold": "false"
				}
				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[custtype]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataCustType%> ],
		"valueField": "contribPercent",
		"titleField": "custtype",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",		
		"pullOutRadius": 0,
		"balloonText": "[[referral]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});
	var chart = AmCharts.makeChart( "chartdivSls1", {
	
		"titles": [
				{
					"text": "Primary",
					"size": 18,
					"bold": "false"
				},
				{
					"text": "Leakage",
					"size": 18,
					"bold": "false"
				}
				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[primary]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataSls1%> ],
		"valueField": "contribPercent",
		"titleField": "primary",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",		
		"pullOutRadius": 0,
		"balloonText": "[[referral]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});
	var chart = AmCharts.makeChart( "chartdivSls2", {
	
		"titles": [
				{
					"text": "Secondary",
					"size": 18,
					"bold": "false"
				},
				{
					"text": "Leakage",
					"size": 18,
					"bold": "false"
				}
				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[secondary]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataSls2%> ],
		"valueField": "contribPercent",
		"titleField": "secondary",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",		
		"pullOutRadius": 0,
		"balloonText": "[[referral]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});

</script>
<!-- am chart js !-->  
