<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Routing.asp"-->

<h1 class="page-header"><i class="fa fa-fw fa-truck"></i> <%=GetTerm("Routing") %> Reports</h1>

<!-- Export to Google Earth !-->
<div class="col-md-6 reports-box">
	<img src="<%= BaseURL %>img/GoogleEarthPro.jpg" class="imgleft">
	<p><a href="#" class="title">Export Deliveries To Google Earth</a></p>
	<p>Creates a .KML* or .csv file of the selected deliveries for importing into Google Earth. Users can select the date to export as well as the route(s).
	 The resulting file can be imported into Google Earth Pro version 7 or higher. 
	 *KML is a file format used to display geographic data in an Earth browser such as Google Earth.</p>
	<p align="right"><a href="<%= BaseURL %>routing/reports/GoogleEarthProExport.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
</div> 
<!-- Export to Google Earth !-->


<!-- Today's Delivery Board Driver Summary!-->
<div class="col-md-6 reports-box">
	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
	<p><a href="#" class="title">Today's Deliveries By Driver</a></p>
	<p>Shows one line per driver with # stops, # invoices and total $ value; Report Shows Today Only. </p>
	<p align="right"><a href="<%= BaseURL %>routing/reports/TodaysDeliveriesByDriver.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
</div> 
<!-- eof Today's Delivery Board Driver Summary!-->


<!-- Today's Delivery Board Driver Summary!-->
<div class="col-md-6 reports-box">
	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
	<p><a href="#" class="title">Historical Deliveries By Driver</a></p>
	<p>Shows one line per driver with # stops, # invoices and total $ value; Report Shows Past Delivery Days Only. </p>
	<p align="right"><a href="<%= BaseURL %>routing/reports/HistoricalDeliveriesByDriver.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
</div> 
<!-- eof Today's Delivery Board Driver Summary!-->

<% If DelBoardDontUseStopSequencing() = False Then %>
	<!-- Historical Delivery Board Driver Summary!-->
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">Driver Sequence Report</a></p>
		<p>Past deliveries made out of sequence by driver.</p>
		<p align="right"><a href="<%= BaseURL %>routing/reports/OutOfSequenceDeliveriesByDriver.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- eof Historical Delivery Board Driver Summary!-->
<% End If %><!-- eof Audit Trail - One Line Per User !--><!--#include file="../../inc/footer-main.asp"-->