<!--#include file="../../../inc/header-inventory-upc.asp"-->
<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InsightFuncs_InventoryControl.asp"-->

<% 
ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")
%>

<style type="text/css">

	body{
		margin:0;
		padding: 0;
	}

	input:focus {
	  background: pink;
	}	
	
	.general-button{
		font-size: 24px;
		border-bottom: 4px solid #3c8f3c;
		margin-bottom: 15px;
		border-radius: 0px !important;
	}
	
	.general-image{
		max-width: 100%;
		height: auto;
	}

	.magnifier{
		max-height: 30px;
	}

	.left-arrow{
		color: #fff;
		margin-top: -2px;
	}

	.btn-go{
		width: 100%;
		text-align: center;
	}

	.pull-left{
		margin-left: 5px;
	}

	.red{
		color: red;
	}

	.green{
		color: green;
	}
	
	.row-line{
		margin-bottom: 25px;
	}

	.row-info{
		margin-bottom: 15px;
	}

	/* mobile only css */

	@media (max-width: 768px) {

		.mobile-col{
			padding-left: 2px;
			padding-right: 2px;
 		}

 		.mobile-col .label{
 			width: 100%;
 			display: block;
 			font-size: 16px;
 			font-weight: bold;
 			margin-top: 5px;
 			padding: 0px !important;
 			white-space: normal !important;
 		}

 		.mobile-image{
 			max-height: 150px;
 			width: auto;
 		}
        
 		 
	}
 		/* eof mobile only css */


	
	.tt-menu,
	.gist {
	  text-align: left;
	  width: 100%;
	}
	
	.typeahead,
	.tt-query,
	.tt-hint {
	 width: 100% !important;
	  height: 50px;
	  padding: 8px 12px;
	  font-size: 16px;
	  line-height: 30px;
	  -webkit-border-radius: 8px;
	     -moz-border-radius: 8px;
	          border-radius: 8px;
	  outline: none;
	  
		border: 1px solid #ccc;
	    border-radius: 4px;
	    -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
	    box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
	    -webkit-transition: border-color ease-in-out .15s,-webkit-box-shadow ease-in-out .15s;
	    -o-transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;
	    transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;	  
	}
	
	.typeahead {
	  background-color: #fff;
	}
	
	.typeahead:focus {
	  border: 2px solid #0097cf;
	}
	
	.tt-query {
	  -webkit-box-shadow: inset 0 1px 1px rgba(0, 0, 0, 0.075);
	     -moz-box-shadow: inset 0 1px 1px rgba(0, 0, 0, 0.075);
	          box-shadow: inset 0 1px 1px rgba(0, 0, 0, 0.075);
	}
	
	.tt-hint {
	  color: #999
	}
	
	.tt-menu {
	  width: 100%;
	  margin: 12px 0;
	  padding: 8px 0;
	  background-color: #fff;
	  border: 1px solid #ccc;
	  border: 1px solid rgba(0, 0, 0, 0.2);
	  -webkit-border-radius: 8px;
	     -moz-border-radius: 8px;
	          border-radius: 8px;
	  -webkit-box-shadow: 0 5px 10px rgba(0,0,0,.2);
	     -moz-box-shadow: 0 5px 10px rgba(0,0,0,.2);
	          box-shadow: 0 5px 10px rgba(0,0,0,.2);
	}
	
	.tt-suggestion {
	  padding: 3px 20px;
	  font-size: 16px;
	  line-height: 18px;
	}
	
	.tt-suggestion:hover {
	  cursor: pointer;
	  color: #fff;
	  background-color: #0097cf;
	}
	
	.tt-suggestion.tt-cursor {
	  color: #fff;
	  background-color: #0097cf;
	
	}
	
	.tt-suggestion p {
	  margin: 0;
	}
		
	/* scrollable dropdown specific styles */
	/* ----------------------- */
	
	#scrollable-dropdown-menu .empty-message {
	  padding: 5px 10px;
	 text-align: center;
	}
		
	
	#scrollable-dropdown-menu .tt-menu {
	   max-height: 150px;
	   overflow-y: auto;
	 }
 
	/** Added tp make typeahead 100% screen width */
	.twitter-typeahead{
	     width: 98%;
	}
	.tt-dropdown-menu{
	    width: 102%;
	}
	input.typeahead.tt-query{ /* This is optional */
	    width: 300px !important;
	}	
	

</style>       


<SCRIPT LANGUAGE="JavaScript">

	
	$(document).ready(function() { 
		
		var productList = new Bloodhound({
		  datumTokenizer: Bloodhound.tokenizers.obj.whitespace(['value','description']),
		  queryTokenizer: Bloodhound.tokenizers.whitespace,
		  prefetch: "../../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/product_list_mobile_<%= ClientKeyForFileNames %>.json",
		});
		
		productList.initialize();
		productList.clearPrefetchCache();		
	
		$('#scrollable-dropdown-menu .typeahead').typeahead(null, {
		  name: 'product-list',
		  limit: 10,
		  display: 'display',
		  source: productList,
		  hint: false,
		  highlight: true,
		  minLength: 1,	  
		  templates: {
		    empty: [
		      '<div class="empty-message">',
		        'unable to find any products that match the current query',
		      '</div>'
		    ].join('\n'),
		    suggestion: function(data) {
	    		return '<p><strong>' + data.value + '</strong> – ' + data.description + '</p>';
			}
		  }
		  
		}).on('typeahead:selected', function (obj, datum) {
		
		    var prodSKU = datum.value;
		    $("#txtProdSKUSelected").val(prodSKU);
		    
			 if (prodSKU != "") {
			 
			 	$.ajax({
					type:"POST",
					url: "../../../inc/InsightFuncs_AjaxForInventoryControl.asp",
					data: "action=DisplaySKULookupInformation&prodSKU="+encodeURIComponent(prodSKU),
						success: function(msg){					        
							$("#divDisplaySKULookupInformation").html(msg);
						}
				}) 
			  }
		});
		
		
		$("#btnClearTypeahead").click(function() {
             $('.typeahead').typeahead('val', '');
             $('.typeahead').focus();
             $("#divDisplaySKULookupInformation").html("");
   		});		
	    
	});   
	

    $(window).on("load", function () {
        event.preventDefault();
        $('.typeahead').focus();
    });

        
    
</SCRIPT>

<h1 class="inventory-upc-heading"><a href="../main_menu.asp" class="left-arrow"><i class="fa fa-arrow-left pull-left" aria-hidden="true"></i></a> Product Lookup</h1>


<!-- driver menu starts here !-->
<div class="container-fluid inventory-upc-container">

	<!-- label -->
	<label>Search for Product by SKU or Description</label>
	<!-- eof label -->
	 
    <div class="row row-line">
        <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
			<div id="scrollable-dropdown-menu">
			  <input class="typeahead" type="text" placeholder="Search By SKU or Description">
			</div>		     
        </div>
     </div>
     
	<div class="row row-line">
		<div class="col-xs-3 pull-right" style="padding:left:0px"><button class="btn btn-info btn-go btn-md" id="btnClearTypeahead">CLEAR</button></div>
		<div class="col-xs-9 pull-right">&nbsp;</div>
	</div>
	
    <input type="hidden" id="txtProdSKUSelected" name="txtProdSKUSelected">

	<div id="divDisplaySKULookupInformation"></div>

</div>

<!--#include file="../../../inc/footer-mobile.asp"-->