<!--#include file="../inc/header.asp"-->
<h1 class="page-header"><i class="fa fa-plug"></i> API Reports</h1>

<!-- row !-->
<div class="row">


   	<div class="col-md-6 reports-box">
    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
    	<p><a href="#" class="title">1. Daily API Activity by Partner (Text View)</a></p>
        <p>Shows orders, invoices, return authorization, credit memos and summary invoices by partner.</p>
        <p align="right"><a href="<%= BaseURL %>api/reports/daily_api_activity_by_partner_text_view.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
    </div> 


   	<div class="col-md-6 reports-box">
    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
    	<p><a href="#" class="title">2. Daily API Activity by Partner (Detailed View)</a></p>
        <p>Shows orders, invoices, return authorization, credit memos and summary invoices by partner.</p>
        <p align="right"><a href="<%= BaseURL %>api/reports/daily_api_activity_by_partner_text_view_detail.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
    </div> 

</div>
<!--#include file="../inc/footer-main.asp"-->