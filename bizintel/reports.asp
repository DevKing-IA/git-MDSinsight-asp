<!--#include file="../inc/header.asp"-->
<h1 class="page-header"><i class="fa fa-graduation-cap"></i> Business Intelligence Reports</h1>

<!-- row !-->
<div class="row">

    <!-- Customer Category Leakage Summary 
   	<div class="col-md-6 reports-box">
    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
    	<p><a href="#" class="title">1. Customer Category Leakage Summary</a></p>
        <p>This report shows "leakage" based on changes in customers' sales examined on several key period comparisons and the Acceptable Variance Threshold.</p>
        <p align="right"><a href="<%= BaseURL %>bizintel/CustomerLeakageSummary.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
    </div> 
    <!-- eof Customer Category Leakage Summary !-->

	<!-- Volume And Price Change By Category 
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">2. Volume And Price Change By Category</a></p>
    	<p>This report calculates the total sales for two specified time periods and shows the Volume and Pricing changes that constitute the difference in sales between the two periods, broken down by category.</p>
		<p align="right"><a href="<%= BaseURL %>bizintel/VolumeAndPricingSummaryByCategory.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- eof Volume And Price Change By Category !-->

	<!-- Volume And Price Change 
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">3. Volume And Price Change</a></p>
    	<p>This report calculates the total sales for two specified time periods and shows the Volume and Pricing changes that constitute the difference in sales between the two periods.</p>
		<p align="right"><a href="<%= BaseURL %>bizintel/VolumeAndPricingSummary.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- eof Volume And Price Change !-->

	<!-- Volume And Price Change By Category - Single Account 
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">4. Volume And Price Change By Category - Single Account</a></p>
    	<p>This report calculates the total sales for two specified time periods and shows the Volume and Pricing changes that constitute the difference in sales between the two periods, broken down by category. This report is run for a single account, specified at the time it is run.</p>
    	<p align="right"><a href="<%= BaseURL %>bizintel/VolumeAndPricingSummaryByCategoryOneCustomer.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- eof Volume And Price Change By Category - Single Account !-->
    <!-- Customer Analysis 1 Summary !-->

   	<div class="col-md-6 reports-box">
    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
    	<p><a href="#" class="title">1. Customer Analysis Summary 1</a></p>
        <p>Analyzes customer</p>
        <p align="right"><a href="<%= BaseURL %>bizintel/CustAnalSum_1.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
    </div> 
    <!-- eof Customer Analysis 1 Summary !-->

   	<div class="col-md-6 reports-box">
    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
    	<p><a href="#" class="title">2. Leakage Overview</a></p>
        <p>Analyzes customer</p>
        <p align="right"><a href="<%= BaseURL %>bizintel/Leakage_Overview.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
    </div> 
    <!-- eof Customer Analysis 1 Summary !-->


	<!-- Sale By Day And Customer Class !-->
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">5. Sales By Day and Customer Class (Summary)</a></p>
    	<p>This report calculates the daily total sales and gross profit between two specified time periods/date ranges, broken down by customer type. This report is run for a specified # of days, specified at the time it is run.</p>
    	<p align="right"><a href="<%= BaseURL %>bizintel/SalesByDaySummary.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- EOF Sale By Day And Customer Class !-->

	<!-- Sale By Day And Customer Class !-->
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">6. Sales By Day and Customer Class (Detailed View)</a></p>
    	<p>This report calculates a detailed view of the daily total sales, cost and gross profit between two specified time periods/date ranges, broken down by customer type. This report is run for a specified # of days, specified at the time it is run.</p>
    	<p align="right"><a href="<%= BaseURL %>bizintel/SalesByDayDetail.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- EOF Sale By Day And Customer Class !-->
	

	<!-- Sale By Period And Customer Class !-->
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">7. Sales By Period and Customer Class (Summary)</a></p>
    	<p>This report calculates the period total sales and gross profit between two specified period ranges, broken down by customer type. This report is run for a specified range of periods, specified at the time it is run.</p>
    	<p align="right"><a href="<%= BaseURL %>bizintel/SalesByPeriodSummary.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- EOF Sale By Period And Customer Class !-->

	<!-- Sale By Period And Customer Class !-->
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">8. Sales By Period and Customer Class (Detailed View)</a></p>
    	<p>This report calculates a detailed view of the period total sales, cost and gross profit between two specified period ranges, broken down by customer type. This report is run for a specified range of periods, specified at the time it is run.</p>
    	<p align="right"><a href="<%= BaseURL %>bizintel/SalesByPeriodDetail.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- EOF Sale By Period And Customer Class !-->
	

	<!-- Sale By Period And Customer Class !-->
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">9. Sales By Month and Customer Class (Summary)</a></p>
    	<p>This report calculates the monthly total sales and gross profit for specified months (and compared to previous year's same month), broken down by customer type. This report is run for a specified range of months, specified at the time it is run.</p>
    	<p align="right"><a href="<%= BaseURL %>bizintel/SalesByMonthSummary.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- EOF Sale By Period And Customer Class !-->
	
	
	<!-- Sale By Period And Customer Class !-->
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">10. Sales By Month and Customer Class (Detailed View)</a></p>
    	<p>This report calculates a detailed view of the monthly total sales and gross profit for specified months (and compared to previous year's same month), broken down by customer type. This report is run for a specified range of months, specified at the time it is run.</p>
    	<p align="right"><a href="<%= BaseURL %>bizintel/SalesByMonthDetail.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- EOF Sale By Period And Customer Class !-->


</div>
<!--#include file="../inc/footer-main.asp"-->