<!--#include file="../inc/header.asp"-->
<!--#include file="../inc/InsightFuncs_AR_AP.asp"-->

<%


'**********************************************************************************
'HELPER FUNCTION TO GET LAST DAY OF MONTH
'**********************************************************************************
Function GetLastDayofMonth(aDate)
    dim intMonth
    dim dteFirstDayNextMonth

    dtefirstdaynextmonth = dateserial(year(adate),month(adate) + 1, 1)
    GetLastDayofMonth = Day(DateAdd ("d", -1, dteFirstDayNextMonth))
End Function

'**********************************************************************************

%>
<style>
.well {
    box-shadow: 0 0 4px 0 rgba(0,0,0,.08),0 2px 4px 0 rgba(0,0,0,.12);
    border-radius: 4px;

}
.container {
    width: 100%;
}
</style>

<h1 class="page-header"><i class="fa fa-dollar"></i> <%= GetTerm("Accounts Receivable") %> Dashboard</h1>

<!-- Begin Cotainer Top Row -->
<div class="container">

	<!-- Begin Top Row -->
	<div class="row">
		
		<div class="col-md-6">
		    <div class="panel" style="background-color:#F3FAFE">
	          <div class="panel-heading">
	            <div class="row">
	            	<!--#include file="dashboard/graph_InvoiceHistoryCounts.asp"-->
	            </div>
	          </div>
	        </div>
		</div>
		
		
		<div class="col-md-6">
		    <div class="panel" style="background-color:#F3F6F6">
	          <div class="panel-heading">
	            <div class="row">
	            	<!--#include file="dashboard/graph_InvoiceHistoryDollars.asp"-->
	            </div>
	          </div>
	        </div>
		</div>
		
	</div><!-- End Top Row -->
	
	<!-- Begin Second Row -->
	<div class="row">
	    
		<div class="col-md-12">
		    <div class="panel">
	          <div class="panel-heading">
	            <div class="row">
	            	<!--#include file="dashboard/graph_ARCustomerCounts.asp"-->
	            </div>
	          </div>
	        </div>
		</div>
	</div>
	<!-- End Second Row -->	
		
	
</div><!-- End Container For Top Row -->
	
	

<!--#include file="dashboard/graph_dashboard_footer_js_code.asp"-->
<!--#include file="../inc/footer-main.asp"-->