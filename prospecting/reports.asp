<!--#include file="../inc/header.asp"-->
<h1 class="page-header"><i class="fa fa-asterisk"></i> <%= GetTerm("Prospecting") %> Reports</h1>

<!-- row !-->
<div class="row">

	<% If userIsAdmin(Session("userNo")) or (userIsOutsideSalesManager(Session("userNo")) or userIsInsideSalesManager(Session("userNo")) AND (GetCRMPermissionLevel(Session("userNo")) <> "NONE")) Then%>
	   	
	   	<!-- Start Propspecting Analysis 1 Summary !-->
	   	<div class="col-md-6 reports-box">
	    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
	    	<p><a href="<%= BaseURL %>prospecting/dashboard_graph_view.asp" class="title">1. <%= GetTerm("Prospecting") %> Graph View</a></p>
	        <p>Shows prospects created, appointments completed, new client converted to customers, qualified prospects and unqualified prospects by reason. All reports are shown by both sales rep and lead source.</p>
	        <p align="right"><a href="<%= BaseURL %>prospecting/reports/dashboard_graph_view.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	    </div> 
	    <!-- eof Propspecting Analysis 1 Summary !-->
	
		<!-- Start Propspecting Analysis 2 Summary !-->
	   	<div class="col-md-6 reports-box">
	    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
	    	<p><a href="<%= BaseURL %>prospecting/dashboard_text_view.asp" class="title">2. <%= GetTerm("Prospecting") %> Text View</a></p>
	        <p>Shows prospects created, appointments completed, new client converted to customers, qualified prospects and unqualified prospects by reason. All reports are shown by both sales rep and lead source.</p>
	        <p align="right"><a href="<%= BaseURL %>prospecting/reports/dashboard_text_view.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	    </div> 
	    <!-- eof Propspecting Analysis 2 Summary !-->
    
    <% End If %>

</div>
<!--#include file="../inc/footer-main.asp"-->