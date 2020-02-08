<!--#include file="../inc/header.asp"-->
<h1 class="page-header"><i class="fa fa-wrench"></i> <%= GetTerm("Service") %> Reports</h1>

<!-- row !-->
<div class="row">

    <!-- Customer Analysis 1 Summary !-->
   	<div class="col-md-6 reports-box">
    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
    	<p><a href="#" class="title">1. Sample Report 1</a></p>
        <p>Analyzes <%= GetTerm("Service") %> feature 1</p>
        <p align="right"><a href="<%= BaseURL %>service/reports/#####"><button type="button" class="btn btn-primary">Run</button></a></p>
    </div> 
    <!-- eof Customer Analysis 1 Summary !-->

	<!-- Customer Analysis 2 Summary !-->
   	<div class="col-md-6 reports-box">
    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
    	<p><a href="#" class="title">2. Sample Report 2</a></p>
        <p>Analyzes <%= GetTerm("Service") %> feature 2</p>
        <p align="right"><a href="<%= BaseURL %>service/reports/#####"><button type="button" class="btn btn-primary">Run</button></a></p>
    </div> 
    <!-- eof Customer Analysis 2 Summary !-->



</div>
<!--#include file="../inc/footer-main.asp"-->