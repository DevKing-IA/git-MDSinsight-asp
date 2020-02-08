<!--#include file="../inc/header.asp"-->
<h1 class="page-header"><i class="fas fa-forklift"></i> <%= GetTerm("Inventory Control") %> Reports</h1>

<!-- row !-->
<div class="row">

	<% If userIsCSR(Session("userNo")) or userIsCSRManager(Session("userNo")) or userIsAdmin(Session("userNo")) or userIsInsideSales(Session("userNo")) or userIsInsideSalesManager(Session("userNo")) Then %>

		<!-- Today's Audit Trail (Full) !-->
		<div class="col-md-6 reports-box">
			<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
			<p><a href="#" class="title">Product UPC Report</a></p>
			<p>Product report that shows Unit and Case UPC codes as well as inventory & picking flags. </p>
			<p align="right"><a href="<%= BaseURL %>inventorycontrol/reports/ProductInventoryReport.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
		</div> 
		<!-- eof Today's Audit Trail (Full) !-->
		
	<% End If %>


</div>
<!--#include file="../inc/footer-main.asp"-->