<!--#include file="../../../../inc/header.asp"-->

<style type="text/css">

	.tile-container {
	  font-family: "Open Sans", "Segoe UI", Frutiger, "Frutiger Linotype", "Dejavu Sans", "Helvetica Neue", Arial, sans-serif;
	  font-size: 14px;
	  line-height: 1.5em;
	  font-weight: 400;
	  width: 100%;
	}
	
	.tile-container p, .tile-container span, .tile-container a, .tile-container ul, .tile-container li, .tile-container button {
	  font-family: inherit;
	  font-size: inherit;
	  font-weight: inherit;
	  line-height: inherit;
	}
	
	.tile-container strong {
	  font-weight: 600;
	}
	
	.tile-container h1, .tile-container h2, .tile-container h3, .tile-container h4, .tile-container h5, .tile-containerh6 {
	  font-family: "Open Sans", "Segoe UI", Frutiger, "Frutiger Linotype", "Dejavu Sans", "Helvetica Neue", Arial, sans-serif;
	  line-height: 1.5em;
	  font-weight: 300;
	}
	
	.tile-container strong {
	  font-weight: 400;
	}
	
	.tile-container .col-sm-2 {
	    width: 20%;
	}
	
	.tile {
	  width: 100%;
	  display: inline-block;
	  box-sizing: border-box;
	  background: #fff;
	  padding: 20px;
	  margin-bottom: 30px;
	  min-height: 160px;
	  color:#FFF;
	}
	
	.tile .title {
	  margin-top: 0px;
	}
	.tile.purple, .tile.blue, .tile.red, .tile.orange, .tile.green {
	  color: #fff;
	}
	.tile.purple {
	  background: #5133AB;
	}
	.tile.purple:hover {
	  background: #3e2784;
	}
	.tile.red {
	  background: #AC193D;
	}
	.tile.red:hover {
	  background: #7f132d;
	}
	.tile.green {
	  background: #00A600;
	}
	.tile.green:hover {
	  background: #007300;
	}
	.tile.blue {
	  background: #2672EC;
	}
	.tile.blue:hover {
	  background: #125acd;
	}
	.tile.orange {
	  background: #DC572E;
	}
	.tile.orange:hover {
	  background: #b8431f;
	}
	
	.tile.aqua {
	  background: #03997e;
	}
	.tile.aqua:hover {
	  background: #07806a;
	}
	.tile.brown {
	  background: #716c4c;
	}
	.tile.brown:hover {
	  background: #59553a;
	}
	.tile.gold {
	  background: #eacf46;
	}
	.tile.gold:hover {
	  background: #cab238;
	}
	
	.tile.brightred {
	  background: #d42c2c;
	}
	.tile.brightred:hover {
	  background: #ba2424;
	}
	
	.tile.mediumblue {
	  background: #3b579d;
	}
	.tile.mediumblue:hover {
	  background: #2f4782;
	}
	.tile.darkgray {
	  background: #2d2d2d;
	}
	.tile.darkgray:hover {
	  background: #1a1919;
	}
	.tile.kellygreen {
	  background: #2e8a57;
	}
	.tile.kellygreen:hover {
	  background: #206941;
	}
	.tile.pink {
	  background: #ff6eb4;
	}
	.tile.pink:hover {
	  background: #eb5ba1;
	}

	.tile.teal {
	  background: #2EBDAA;
	}
	.tile.teal:hover {
	  background: #2CB1A0;
	}
	.tile.lightorange {
	  background: #FF8C00;
	}
	.tile.lightorange:hover {
	  background: #e07d05;
	}

	
	.tile .disabled{
	    opacity: 0.5;
	}
	
	#PleaseWaitPanel{
		position: fixed;
		left: 470px;
		top: 275px;
		width: 975px;
		height: 300px;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	}   

	.btn-huge{
	    padding: 18px 28px;
	    font-size: 22px;	    
	}
</style>

<h1 class="page-header"><i class="fa fa-external-link"></i>&nbsp;API Settings
	<a href="<%= BaseURL %>admin/global/main.asp"><button class="btn btn-small btn-secondary pull-right" style="margin-left:20px"><i class="fas fa-arrow-alt-left"></i>&nbsp;<i class="fas fa-globe"></i>&nbsp;GLOBAL SETTINGS MAIN</button></a>
</h1>

<div class="container tile-container">

  <!--<div class="row">
    <div class="col-md-12">
      <h1><strong>Bootstrap - Microsoft Metro Tiles</strong></h1>
    </div>
  </div>-->
  
  <div class="row">
  	
    <% If MUV_READ("orderAPIModuleOn") <> "Hidden" Then %>
	    <a href="<%= BaseURL %>admin/global/tiles/api/order-api.asp"><div class="col-sm-2">
	      <div class="tile pink">
	        <h3 class="title"><i class="fa fa-link"></i>&nbsp;Order API</h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a> 
	<% Else %>
	    <a href="#"><div class="col-sm-2">
	      <div class="tile pink disabled">
	        <h3 class="title"><i class="fa fa-link"></i>&nbsp;Order API</h3>
	        <p>This module is currently disabled.</p>
	      </div>
	    </div></a> 	
	<% End If %>

    <% If cint(MUV_Read("arModuleOn")) = 1 Then %>
	    <a href="<%= BaseURL %>admin/global/tiles/api/accounts-receivable.asp"><div class="col-sm-2">
	      <div class="tile kellygreen">
	        <h3 class="title"><i class="fas fa-file-invoice-dollar"></i>&nbsp;<%= GetTerm("Accounts Receivable") %> API</h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a> 
	<% Else %>
	    <a href="#"><div class="col-sm-2">
	      <div class="tile kellygreen disabled">
	        <h3 class="title"><i class="fas fa-file-invoice-dollar"></i>&nbsp;<%= GetTerm("Accounts Receivable") %> API</h3>
	        <p>This module is currently disabled.</p>
	      </div>
	    </div></a> 	
	<% End If %>
	
	
   
  </div>
  
  
</div>

<!--#include file="../../../../inc/footer-main.asp"-->