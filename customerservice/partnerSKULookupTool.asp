<% If Session("Userno") = "" Then Response.Redirect("../default.asp") %>

<!DOCTYPE html>
<!--[if lt IE 7 ]> <html class="no-js ie6 oldie" lang="en"> <![endif]-->
<!--[if IE 7 ]>    <html class="no-js ie7 oldie" lang="en"> <![endif]-->
<!--[if IE 8 ]>    <html class="no-js ie8 oldie" lang="en"> <![endif]-->
<!--[if IE 9 ]>    <html class="no-js ie9" lang="en"> <![endif]-->
<!--[if (gte IE 9)|!(IE)]><![endif]--><!-->
<html class="no-js" lang="en">
<!--<![endif]-->
<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/protect.asp"-->
<!--#include file="../inc/InsightFuncs.asp"-->
<!--#include file="../inc/InsightFuncs_InventoryControl.asp"-->
<!--#include file="../inc/InSightFuncs_AjaxForInventoryControlModals.asp"-->

<%
	'**************************************************************************
    'Get Company Information
    '**************************************************************************
    
	SQLCustomLogin = "SELECT * FROM tblServerInfo where clientKey='"& MUV_Read("ClientID") &"'"
	Set ConnectionCustomLogin = Server.CreateObject("ADODB.Connection")
	Set RecordsetCustomLogin = Server.CreateObject("ADODB.Recordset")
	ConnectionCustomLogin.Open InsightCnnString

	'Open the recordset object executing the SQL statement and return records
	RecordsetCustomLogin.Open SQLCustomLogin,ConnectionCustomLogin,3,3

	'First lookup the ClientKey in tblServerInfo
	If NOT RecordsetCustomLogin.EOF then
		companyName = RecordsetCustomLogin.Fields("companyName")
		shortCompanyName = RecordsetCustomLogin.Fields("shortCompanyName")
		RecordsetCustomLogin.close
		ConnectionCustomLogin.close	
	End If	
	
	If shortCompanyName = "" Then
		shortCompanyName = companyName
	End If


%>
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta name="description" content="">
    <meta name="author" content="">

    <title>MDS Partner SKU Lookup Tool</title>

    <!-- Bootstrap core CSS -->
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet">
    <!-- End Bootstrap core CSS -->

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
    
    <!-- icons and notification styles !-->
     <script src="https://use.fontawesome.com/3382135cdc.js"></script>
    
    <!-- fonts !-->
    <link href='http://fonts.googleapis.com/css?family=Coda' rel='stylesheet' type='text/css'>
    <link href='http://fonts.googleapis.com/css?family=Oswald:400,300,700' rel='stylesheet' type='text/css'>
	<link href='http://fonts.googleapis.com/css?family=Indie+Flower' rel='stylesheet' type='text/css'>
    
    <!-- eof fonts !-->
	
	<!-- *********************************************************************** -->
	<!-- IMPORTANT - USE OLDER VERSION OF JQUERY FOR SORTABLE PLUGIN             -->
	<!-- *********************************************************************** -->
  	<script src="http://code.jquery.com/jquery-1.11.2.min.js"></script>
	<!--<script src="https://code.jquery.com/jquery-3.1.1.js" integrity="sha256-16cdPddA6VdVInumRGo6IbivbERE8p7CQR3HzTBuELA=" crossorigin="anonymous"></script>  -->	
	<!-- *********************************************************************** -->
		
	<!-- Including jQuery UI CSS & jQuery Dialog UI Here-->
	<link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/themes/ui-darkness/jquery-ui.css" rel="stylesheet">
	<script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/jquery-ui.min.js"></script>
	<!-- End Including jQuery UI CSS & jQuery Dialog UI Here-->
 	
	<!-- Bootstrap core JS - must load after jQuery -->
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
    <!-- End Bootstrap core JS - must load after jQuery -->
	
	<!-- sweet alert jquery modal alerts !-->	
	<script src="<%= BaseURL %>js/sweetalert/sweetalert.min.js"></script>
	<link rel="stylesheet" type="text/css" href="<%= BaseURL %>js/sweetalert/sweetalert.css">
	<!-- end sweet alert jquery modal alerts !-->	
    
	<!-- Easy Autocomplete Files -->
	<!-- JS file -->
	<script src="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/jquery.easy-autocomplete.js"></script> 
	<!-- CSS file -->
	<link rel="stylesheet" href="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete.css"> 
	<!-- Additional CSS Themes file - not required-->
	<link rel="stylesheet" href="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete.themes.css"> 

    <!-- jQuery Cookie Files To Save State In Place of Session Variables -->
    <script src="<%= BaseURL %>js/jquery.cookie.js"></script>
    <!-- End jQuery Cookie -->
    
    <style>
    body {
    font-family: arial;
  }

	h1{
	margin-left:15px;
	margin-right:15px;
	font-size:24px;
	font-weight:bold;
	color:rgb(50, 118, 177);
	
}
  table {
    border: 1px solid #ccc;
    width: 100%;
    margin:0;
    padding:0;
    border-collapse: collapse;
    border-spacing: 0; 
  }

  table tr {
    border: 1px solid #ddd;
    padding: 5px;
  }

  table th, table td {
    padding: 10px;
    text-align: center;
  }

  table th {
    text-transform: uppercase;
    font-size: 14px;
    letter-spacing: 1px;
  }
	/* CSS for Credit Card Payment form */
	.product-sku-box .panel-title {
	    display: inline;
	    font-weight: bold;
	}
	.product-sku-box .form-control.error {
	    border-color: red;
	    outline: 0;
	    box-shadow: inset 0 1px 1px rgba(0,0,0,0.075),0 0 8px rgba(255,0,0,0.6);
	}
	.product-sku-box label.error {
	  font-weight: bold;
	  color: red;
	  padding: 2px 8px;
	  margin-top: 2px;
	}
	.product-sku-box .payment-errors {
	  font-weight: bold;
	  color: red;
	  padding: 2px 8px;
	  margin-top: 2px;
	}
	.product-sku-box label {
	    display: block;
	}
	/* The old "center div vertically" hack */
	.product-sku-box .display-table {
	    display: table;
	}
	.product-sku-box .display-tr {
	    display: table-row;
	}
	.product-sku-box .display-td {
	    display: table-cell;
	    vertical-align: middle;
	    width: 100%;
	    text-align:center;
	}
	/* Just looks nicer */
	.product-sku-box .panel-heading img {
	    min-width: 180px;
	}
	.input-group-addon.primary {
	    color: rgb(255, 255, 255);
	    background-color: rgb(50, 118, 177);
	    border-color: rgb(40, 94, 142);
	}
	.input-group-addon.success {
	    color: rgb(255, 255, 255);
	    background-color: rgb(92, 184, 92);
	    border-color: rgb(76, 174, 76);
	    cursor: pointer;
	}
	.input-group-addon.info {
	    color: rgb(255, 255, 255);
	    background-color: rgb(57, 179, 215);
	    border-color: rgb(38, 154, 188);
	    cursor: pointer;
	}
	.input-group-addon.warning {
	    color: rgb(255, 255, 255);
	    background-color: rgb(240, 173, 78);
	    border-color: rgb(238, 162, 54);
	    cursor: pointer;
	}
	.input-group-addon.danger {
	    color: rgb(255, 255, 255);
	    background-color: rgb(217, 83, 79);
	    border-color: rgb(212, 63, 58);
	    cursor: pointer;
	}	
	.panel-title {
	    margin-top: 0;
	    margin-bottom: 0;
	    font-size: 16px;
	    /*color: #1a6ecc !important;*/
	    color:rgb(217, 83, 79) !important;;
	}	
  @media screen and (max-width: 600px) {

    table {
      border: 0;
    }

    table thead {
      display: none;
    }

    table tr {
      margin-bottom: 10px;
      display: block;
      border-bottom: 2px solid #ddd;
    }

    table td {
      display: block;
      text-align: right;
      font-size: 13px;
      border-bottom: 1px dotted #ccc;
    }

    table td:last-child {
      border-bottom: 0;
    }

    table td:before {
      content: attr(data-label);
      float: left;
      text-transform: uppercase;
      font-weight: bold;
    }
  }
    </style>
    

<script type="text/javascript">
		
	$(document).ready(function() {
	
		$("#frmProductLookupResults").hide();
		$("#frmPartnerProductLookupResults").hide();
		
		
				
		$( "#btnProductLookupSubmit" ).click(function() {

   			var txtPartnerSKU = $("#txtPartnerSKU").val();
   			var txtShortCompanyName = $("#txtShortCompanyName").val();
   		
	   		if (txtPartnerSKU !== ""){
	   		
				$.ajax({
				
					type:"POST",
					url: "../inc/InSightFuncs_AjaxForInventoryControlModals.asp",
					cache: false,
					data: "action=ReturnProductEquivalentSKUs&partnerSKU=" + encodeURIComponent(txtPartnerSKU) + "&shortCompanyName=" + encodeURIComponent(txtShortCompanyName),
					success: function(response)
					{				
						$("#frmProductLookupResults").html(response);
					    $("#frmProductLookupResults").show();	
			        }
				})
			 } 

		});  
		
		
		
		
		$( "#btnPartnerProductLookupSubmit" ).click(function() {

   			var txtCompanySKU = $("#txtCompanySKU").val();
   		
	   		if (txtCompanySKU !== ""){
	   		
				$.ajax({
				
					type:"POST",
					url: "../inc/InSightFuncs_AjaxForInventoryControlModals.asp",
					cache: false,
					data: "action=ReturnPartnerEquivalentSKUs&companySKU=" + encodeURIComponent(txtCompanySKU),
					success: function(response)
					{				
						$("#frmPartnerProductLookupResults").html(response);
					    $("#frmPartnerProductLookupResults").show();	
			        }
				})
			 } 

		}); 
		
		
		
				
		$("#frmProductLookup").on("submit", function(){
	
   			var txtPartnerSKU = $("#txtPartnerSKU").val();
   			var txtShortCompanyName = $("#txtShortCompanyName").val();
   		
	   		if (txtPartnerSKU !== ""){
	   		
				$.ajax({
				
					type:"POST",
					url: "../inc/InSightFuncs_AjaxForInventoryControlModals.asp",
					cache: false,
					data: "action=ReturnProductEquivalentSKUs&partnerSKU=" + encodeURIComponent(txtPartnerSKU) + "&shortCompanyName=" + encodeURIComponent(txtShortCompanyName),
					success: function(response)
					{				
						$("#frmProductLookupResults").html(response);
					    $("#frmProductLookupResults").show();	
			        }
				})
			 } 
			   
			return false;
		});
		 
		
		 
		 
		$("#frmPartnerProductLookup").on("submit", function(){
	
   			var txtCompanySKU = $("#txtCompanySKU").val();
   		
	   		if (txtCompanySKU !== ""){
	   		
				$.ajax({
				
					type:"POST",
					url: "../inc/InSightFuncs_AjaxForInventoryControlModals.asp",
					cache: false,
					data: "action=ReturnPartnerEquivalentSKUs&companySKU=" + encodeURIComponent(txtCompanySKU),
					success: function(response)
					{				
						$("#frmPartnerProductLookupResults").html(response);
					    $("#frmPartnerProductLookupResults").show();	
			        }
				})
			 } 
			   
			return false;
		 });
		   

		   
	});	
	
	
		   	
</script>
    
  </head>

<body>

<h1>Partner SKU Lookup Tool</h1>

<div class="container">
<div class="row">
<div class="col-xs-12 col-md-12">


	<div class="panel panel-default product-sku-box">
		<div class="panel-heading display-table">
			<div class="row display-tr">
				<h3 class="panel-title display-td">Product Lookup (for Client Care)</h3>
			</div>                    
		</div>
		
		<div class="panel-body">
				<form role="form" id="frmProductLookup" name="frmProductLookup">
					<input type="hidden" name="txtShortCompanyName" id="txtShortCompanyName" value="<%= shortCompanyName %>">
					<div class="row">
						<div class="col-xs-12">
							<div class="form-group">
								<label for="cardNumber">Partner Code</label>
						        <div class="input-group">
						            <input type="text" class="form-control" name="txtPartnerSKU" id="txtPartnerSKU" />
						            <span class="input-group-addon success" id="btnProductLookupSubmit">GO</span>
						        </div>								
							</div>                            
						</div>
					</div>
				</form>
		</div> <!-- end Panel Body -->
	</div><!-- end Panel Container -->            


	<div id="frmProductLookupResults"></div>
	
	<div class="panel panel-default product-sku-box" style="margin-top:50px;">
		<div class="panel-heading display-table">
			<div class="row display-tr">
				<h3 class="panel-title display-td">Reverse Lookup (for Purchasing)</h3>
			</div>                    
		</div>
		
		<div class="panel-body">
				<form role="form" id="frmPartnerProductLookup" name="frmPartnerProductLookup">
					<input type="hidden" name="txtShortCompanyName" value="<%= shortCompanyName %>">
					<div class="row">
						<div class="col-xs-12">
							<div class="form-group">
								<label for="cardNumber"><%= shortCompanyName %> Code</label>
						        <div class="input-group">
						            <input type="text" class="form-control" name="txtCompanySKU" id="txtCompanySKU" />
						            <span class="input-group-addon success" id="btnPartnerProductLookupSubmit">GO</span>
						        </div>								
							</div>                            
						</div>
					</div>
				</form>
		</div> <!-- end Panel Body -->
	</div><!-- end Panel Container -->            

	<div id="frmPartnerProductLookupResults"></div>
	
	

</div> <!-- end col-xs-12 col-md-12 -->           
</div> <!-- end class="row" -->
</div> <!-- end class="container" -->


  </body>
</html>