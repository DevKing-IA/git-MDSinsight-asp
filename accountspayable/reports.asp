<!--#include file="../inc/header.asp"-->
<h1 class="page-header"><i class="fas fa-envelope-open-dollar"></i> <%= GetTerm("Accounts Payable") %> Reports</h1>

<!-- row !-->
<div class="row">

    <!-- AP 1 Summary !-->
   	<div class="col-md-6 reports-box">
    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
    	<p><a href="#" class="title">1. AP Summary 1</a></p>
        <p>Analyzes customer</p>
        <p align="right"><a href="<%= BaseURL %>accountspayable/#########"><button type="button" class="btn btn-primary">Run</button></a></p>
    </div> 
    <!-- eof AP 1 Summary !-->

	<!-- AP 2 Summary !-->
   	<div class="col-md-6 reports-box">
    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
    	<p><a href="#" class="title">2. AP Overview</a></p>
        <p>Analyzes customer</p>
        <p align="right"><a href="<%= BaseURL %>accountspayable//#########"><"><button type="button" class="btn btn-primary">Run</button></a></p>
    </div> 
    <!-- eof AP 1 Summary !-->


</div>
<!--#include file="../inc/footer-main.asp"-->