<!--#include file="../inc/header.asp"-->
<!--#include file="../inc/InSightFuncs_Routing.asp"-->
<h1 class="page-header"><i class="fa fa-truck"></i> <%=GetTerm("Routing") %> Reports</h1>

<!-- row !-->
<div class="row">
	
	<!-- Today's Delivery Board 
	<div class="col-md-4 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">Today's Delivery Board</a></p>
		<p>Lists all the deliveries currently listed on the delivery board grouped by driver.</p>
		<p align="right"><a href="<%= BaseURL %>routing/reports/TodaysDeliveryBoard.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- eof Today's Delivery Board  !-->
	
	<!-- Today's Delivery Board Driver Summary!-->
	<div class="col-md-4 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">Today's Deliveries By Driver</a></p>
		<p>Shows one line per driver with # stops, # invoices and total $ value; Report Shows Today Only. </p>
		<p align="right"><a href="<%= BaseURL %>routing/reports/TodaysDeliveriesByDriver.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- eof Today's Delivery Board Driver Summary!-->
	
	
	<!-- Today's Delivery Board Driver Summary!-->
	<div class="col-md-4 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">Historical Deliveries By Driver</a></p>
		<p>Shows one line per driver with # stops, # invoices and total $ value; Report Shows Past Delivery Days Only. </p>
		<p align="right"><a href="<%= BaseURL %>routing/reports/HistoricalDeliveriesByDriver.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- eof Today's Delivery Board Driver Summary!-->
	
	<% If DelBoardDontUseStopSequencing() = False Then %>
		<!-- Historical Delivery Board Driver Summary!-->
		<div class="col-md-4 reports-box">
			<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
			<p><a href="#" class="title">Driver Sequence Report</a></p>
			<p>Past deliveries made out of sequence by driver.</p>
			<p align="right"><a href="<%= BaseURL %>routing/reports/OutOfSequenceDeliveriesByDriver.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
		</div> 
		<!-- eof Historical Delivery Board Driver Summary!-->
	<% End If %>

</div>
<!--#include file="../inc/footer-main.asp"-->