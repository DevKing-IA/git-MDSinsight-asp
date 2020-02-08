<!--#include file="../inc/header.asp"-->

<style>

	.well {
	    box-shadow: 0 0 4px 0 rgba(0,0,0,.08),0 2px 4px 0 rgba(0,0,0,.12);
	    border-radius: 4px;
	
	}
	.container {
	    width: 100%;
	}

	#chartdiv {
	  width: 100%;
	  height: 100%;
	}
	
	h3, h4 {
		text-align:center;
	}

</style>

<h1 class="page-header"><i class="fa fa-plug"></i> API Dashboard</h1>

<!-- Begin Cotainer First Row -->
<div class="container">
	
  <!-- Begin First Row -->
  <div class="row">
  
  	<!-- Begin First Row First Column -->
    <div class ="col-md-5">
    
  		<div class="well">
  			<!--#include file="dashboard/daily_api_activity_summary_chart_data.asp"-->
			<h3><strong>Orders API Activity</strong></h3>
			<h4><%= DateRangeTitleForAPIGraph %></h4>
			<% If Session("AdminPrivelages") = True Then %>
				<div id="OrdersAPIActivityDiv" style="width:100%; height: 350px; margin: 0 auto"></div>
			<% End If %>

		</div> <!-- well ends -->
				
    </div> <!-- End First Row First Column -->
    
  </div> <!-- End First Row -->
  
</div><!-- End Container First Row -->

<!--#include file="dashboard/dashboard_footer_code.asp"-->

<!--#include file="../inc/footer-main.asp"-->