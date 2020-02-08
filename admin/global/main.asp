<!--#include file="../../inc/header.asp"-->

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

</style>

<h1 class="page-header"><i class="fa fa-globe"></i>&nbsp;Global Settings</h1>

<div class="container tile-container">
  <!--<div class="row">
    <div class="col-md-12">
      <h1><strong>Bootstrap - Microsoft Metro Tiles</strong></h1>
    </div>
  </div>-->
  <div class="row">
  	
    <!--<<a href="<%= BaseURL %>admin/global/tiles/leakage-reports.asp"><div class="col-sm-2">
      <div class="tile purple">
        <h3 class="title"><i class="fa fa-usd"></i>&nbsp;Leakage Reports</h3>
        <p>&nbsp;</p>
      </div>
    </div></a>-->
    
    <!--<a href="<%= BaseURL %>admin/global/tiles/default-categories-vpc.asp"><div class="col-sm-2">
      <div class="tile red">
        <h3 class="title"><i class="fa fa-list-ul"></i>&nbsp;Default Categories For VPC Reports</h3>
        <p>&nbsp;</p>
      </div>
    </div></a>>-->
    
    <a href="<%= BaseURL %>admin/global/tiles/post-settings.asp"><div class="col-sm-2">
      <div class="tile orange">
        <h3 class="title"><i class="fa fa-arrow-circle-up"></i>&nbsp;POST Settings</h3>
        <p>&nbsp;</p>
      </div>
    </div></a>
    
    <% If MUV_Read("FilterTrax") <> "1" Then %>
	    <a href="<%= BaseURL %>admin/global/tiles/category-groupings.asp"><div class="col-sm-2">
	      <div class="tile green">
	        <h3 class="title"><i class="fa fa-list-alt"></i>&nbsp;Category Groupings</h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a>
    <% End If %>
    
    <a href="<%= BaseURL %>admin/global/tiles/customize-terminology.asp"><div class="col-sm-2">
      <div class="tile blue">
        <h3 class="title"><i class="fa fa-quote-left"></i>&nbsp;Customize Terminology</h3>
        <p>&nbsp;</p>
      </div>
    </div></a>
    
    <% If MUV_Read("custServiceOn") = "1" Then %>
	    <a href="<%= BaseURL %>admin/global/tiles/client-care-settings.asp"><div class="col-sm-2">
	      <div class="tile gold">
	        <h3 class="title"><i class="fa fa-users"></i>&nbsp;<%= GetTerm("Customer Service") %> Settings</h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a>
    <% End If %>
    
    <% If filterChangeModuleOn() Then %>
	    <a href="<%= BaseURL %>admin/global/tiles/filter-changes.asp"><div class="col-sm-2">
	      <div class="tile brown">
	        <h3 class="title"><i class="fa fa-filter"></i>&nbsp;Filter Changes</h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a>
	<% Else %>
	    <a href="#"><div class="col-sm-2">
	      <div class="tile brown disabled">
	        <h3 class="title"><i class="fa fa-filter"></i>&nbsp;Filter Changes</h3>
	        <p>This module is currently disabled.</p>
	      </div>
	    </div></a>
	<% End If %>
	
	
    
    
    <a href="<%= BaseURL %>admin/global/tiles/texting-settings.asp"><div class="col-sm-2">
      <div class="tile brightred">
        <h3 class="title"><i class="fa fa-mobile"></i>&nbsp;Texting Settings</h3>
        <p>&nbsp;</p>
      </div>
    </div></a>
    
    <% If MUV_READ("routingModuleOn") <> "Hidden" Then %>
	    <a href="<%= BaseURL %>admin/global/tiles/delivery-board.asp"><div class="col-sm-2">
	      <div class="tile darkgray">
	        <h3 class="title"><i class="fa fa-truck"></i>&nbsp;Routing</h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a>  
	<% Else %>
	    <a href="#"><div class="col-sm-2">
	      <div class="tile darkgray disabled">
	        <h3 class="title"><i class="fa fa-truck"></i>&nbsp;Routing</h3>
	        <p>This module is currently disabled.</p>
	      </div>
	    </div></a>	
	<% End If %>    
	
	
	<% If MUV_READ("serviceModuleOn") <> "Hidden" Then %>
	    <a href="<%= BaseURL %>admin/global/tiles/field-service.asp"><div class="col-sm-2">
	      <div class="tile kellygreen">
	        <h3 class="title"><i class="fa fa-wrench"></i>&nbsp;<%= GetTerm("Field Service") %></h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a> 
	<% Else %>
	    <a href="#"><div class="col-sm-2">
	      <div class="tile kellygreen disabled">
	        <h3 class="title"><i class="fa fa-wrench"></i>&nbsp;<%= GetTerm("Field Service") %></h3>
	        <p>This module is currently disabled.</p>
	      </div>
	    </div></a> 	
	<% End If %>
    
    
    <% If MUV_READ("prospectingModuleOn") <> "Hidden" Then %>
	    <a href="<%= BaseURL %>admin/global/tiles/prospecting-settings.asp"><div class="col-sm-2">
	      <div class="tile mediumblue">
	        <h3 class="title"><i class="fa fa-address-book"></i>&nbsp;<%= GetTerm("Prospecting") %><br>Settings</h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a> 
    <% Else %>
	    <a href="#"><div class="col-sm-2">
	      <div class="tile mediumblue disabled">
	        <h3 class="title"><i class="fa fa-address-book"></i>&nbsp;<%= GetTerm("Prospecting") %><br>Settings</h3>
	        <p>This module is currently disabled.</p>
	      </div>
	    </div></a>     
    <% End If %>
    
    
    <% If MUV_READ("orderAPIModuleOn") <> "Hidden" OR cint(MUV_Read("arModuleOn")) = 1 Then %>
	    <a href="<%= BaseURL %>admin/global/tiles/api/main.asp"><div class="col-sm-2">
	      <div class="tile pink">
	        <h3 class="title"><i class="fa fa-external-link"></i>&nbsp;APIs</h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a> 
	<% Else %>
	    <a href="#"><div class="col-sm-2">
	      <div class="tile pink disabled">
	        <h3 class="title"><i class="fa fa-external-link"></i>&nbsp;APIs</h3>
	        <p>This module is currently disabled.</p>
	      </div>
	    </div></a> 	
	<% End If %>

    <% If MUV_Read("equipmentModuleOn")  <> "Hidden" Then %>
	    <!--<a href="<%= BaseURL %>admin/global/tiles/equipment.asp"><div class="col-sm-2">
	      <div class="tile teal">
	        <h3 class="title"><i class="fa fa-tint"></i><i class="fa fa-coffee"></i>&nbsp;Equipment</h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a> -->
	<% Else %>
	     <!--<a href="#"><div class="col-sm-2">
	      <div class="tile teal disabled">
	        <h3 class="title"><i class="fa fa-tint"></i><i class="fa fa-coffee"></i>&nbsp;Equipment</h3>
	        <p>This module is currently disabled.</p>
	      </div>
	    </div></a> 	 -->
	<% End If %>


     <% If MUV_Read("InventoryControlModuleOn")  <> "Hidden" Then %>
	    <a href="<%= BaseURL %>admin/global/tiles/inventory.asp"><div class="col-sm-2">
	      <div class="tile lightorange">
	        <h3 class="title"><i class="fas fa-forklift"></i>&nbsp;Inventory</h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a> 
	<% Else %>
	    <a href="#"><div class="col-sm-2">
	      <div class="tile lightorange disabled">
	        <h3 class="title"><i class="fas fa-forklift"></i>&nbsp;Inventory</h3>
	        <p>This module is currently disabled.</p>
	      </div>
	    </div></a> 	
	<% End If %>

     <% If MUV_Read("biModuleOn")  <> "Hidden" Then %>
	    <a href="<%= BaseURL %>admin/global/tiles/bizintel.asp"><div class="col-sm-2">
	      <div class="tile purple">
	        <h3 class="title"><i class="fa fa-graduation-cap"></i>&nbsp;Business Intelligence</h3>
	        <p>&nbsp;</p>
	      </div>
	    </div></a> 
	<% Else %>
	    <a href="#"><div class="col-sm-2">
	      <div class="tile purple disabled">
	        <h3 class="title"><i class="fa fa-graduation-cap"></i>&nbsp;Business Intelligence</h3>
	        <p>This module is currently disabled.</p>
	      </div>
	    </div></a> 	
	<% End If %>

    <a href="<%= BaseURL %>admin/global/tiles/needtoknow/main.asp"><div class="col-sm-2">
      <div class="tile blue">
        <h3 class="title"><i class="fas fa-lightbulb-on"></i>&nbsp;Need To Know Reports</h3>
        <p>&nbsp;</p>
      </div>
    </div></a>
	
  
  </div>
  
  
</div>

<!--#include file="../../inc/footer-main.asp"-->