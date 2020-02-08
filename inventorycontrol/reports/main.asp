<!--#include file="../../inc/header.asp"-->

<h1 class="page-header"><i class="fa fa-file-text-o"></i>&nbsp;<%= GetTerm("Inventory Control") %> Reports</h1>

<!-- Today's Audit Trail (Full) !-->
<div class="col-md-6 reports-box">
	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
	<p><a href="#" class="title">Product UPC Report</a></p>
	<p>Product report that displays basic product information, inventory, unit UPC's and case UPC's.</p>
	<p align="right"><a href="<%= BaseURL %>inventorycontrol/reports/ProductInventoryReport.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
</div> 
<!-- eof Today's Audit Trail (Full) !-->


<!--#include file="../../inc/footer-main.asp"-->