<!--#include file="../../inc/header-accounts-receivable.asp"-->
<h1 class="page-header"><i class="fa fa-file-text"></i>  Invoices / Statements</h1>

<!-- row !-->
<div class="row">

    <!-- START !-->
   	<div class="col-md-6 reports-box">
    <img src="<%= BaseURL %>img/general/statement.png" class="imgleft">
    
    	<p><a href="#" class="title">1. Consolidated Invoice</a></p>
        
        <p>Creates a consolidated invoice PDF for a given <%=GetTerm("account")%> or chain, combining all invoices from the specified date range. Line items are not shown.</p>
        
        <p align="right"><a href="consolidatedInv/consolidatedStatement.asp"><button type="button" class="btn btn-primary">Run</button></a></p>

    </div> 
    <!-- END !-->

    <!-- START !-->
   	<div class="col-md-6 reports-box">
    <img src="<%= BaseURL %>img/general/statement.png" class="imgleft">
    
    	<p><a href="#" class="title">2. Consolidated Invoice (Unpaid Only)</a></p>
        
        <p>Creates a consolidated invoice PDF for a given <%=GetTerm("account")%> or chain, combining all UNPAID invoices from the specified date range. Line items are not shown.</p>
        
        <p align="right"><a href="consolidatedInvUnpaid/consolidatedStatement.asp"><button type="button" class="btn btn-primary">Run</button></a></p>

    </div> 
    <!-- END !-->
    

    <!-- START !-->
   	<div class="col-md-6 reports-box">
    <img src="<%= BaseURL %>img/general/statement.png" class="imgleft">
    
    	<p><a href="#" class="title">3. Detailed Consolidated Invoice</a></p>
        
        <p>Creates a consolidated invoice PDF for a given <%=GetTerm("account")%> or chain, combining all invoices from the specified date range. This report list each invoice including the line items.</p>
        
        <p align="right"><a href="consInvDetail/consolidatedInvDetail.asp"><button type="button" class="btn btn-primary">Run</button></a></p>

    </div> 
    <!-- END !-->
    
    

    <!-- START !-->
   	<div class="col-md-6 reports-box">
    <img src="<%= BaseURL %>img/general/statement.png" class="imgleft">
    
    	<p><a href="#" class="title">4. Consolidated Invoice By Location</a></p>
        
        <p>Creates a consolidated invoice PDF for a given <%=GetTerm("account")%> or chain, combining all invoices, for all locations from the specified date range. This report list each invoice including the line items.</p>
        
        <p align="right"><a href="consolidatedInvByLocation/consolidatedStatement.asp"><button type="button" class="btn btn-primary">Run</button></a></p>

    </div> 
    <!-- END !-->
    
    
    <!-- Web Fulfillment and Invoice Cross Reference Report !-->
    <% If MUV_Read("webFulfillmentModuleOn") = "Enabled" Then %>
    
	   	<div class="col-md-6 reports-box">
	    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
	    	<p><a href="#" class="title">5. Web Fulfillment and Invoice Cross Reference Summary Report</a></p>
	        <p>This report tracks web orders and invoiced orders to help identify missing orders and track fulfillment of each order.</p>
	        <p align="right"><a href="<%= BaseURL %>accountsreceivable/reports/webFulfillmentInvoiceXRefSummary/WebFulfillmentInvoiceXRefSummary.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	    </div> 
	    
	   	<div class="col-md-6 reports-box">
	    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
	    	<p><a href="#" class="title">6. Web Fulfillment and Invoice Cross Reference Summary Report (By Customer)</a></p>
	        <p>This report tracks web orders and invoiced orders to help identify missing orders and track fulfillment of each order, grouped by customer.</p>
	        <p align="right"><a href="<%= BaseURL %>accountsreceivable/reports/webFulfillmentInvoiceXRefSummary/WebFulfillmentInvoiceXRefSummaryByCust.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	    </div> 
	    
	   	<div class="col-md-6 reports-box">
	    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
	    	<p><a href="#" class="title">7. Web Fulfillment and Invoice Cross Reference Summary Report (By Period)</a></p>
	        <p>This report tracks web orders and invoiced orders to help identify missing orders and track fulfillment of each order, grouped by period.</p>
	        <p align="right"><a href="<%= BaseURL %>accountsreceivable/reports/webFulfillmentInvoiceXRefSummary/WebFulfillmentInvoiceXRefSummaryByPeriod.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	    </div> 
	    
    <% End If %>
    <!-- eof Web Fulfillment and Invoice Cross Reference Report  !-->
    

    <!-- START !-->
   	<div class="col-md-6 reports-box">
    <img src="<%= BaseURL %>img/general/peoplesoft.png" class="imgleft">
    
    	<p><a href="#" class="title">8. Export Invoices To Peoplesoft</a></p>
        
        <p>Creates an export file file for the selected <%=GetTerm("account")%> or chain, compliant with Peoplesoft's AP system.</p>
        
        <p align="right"><a href="peoplesoft/main.asp"><button type="button" class="btn btn-primary">Run</button></a></p>

    </div> 
    <!-- END !-->

</div>
<!-- eof row !-->    

<!--#include file="../../inc/footer-main.asp"-->