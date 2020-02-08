<!--#include file="../../../inc/header-inventory-upc.asp"-->
<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InsightFuncs_InventoryControl.asp"-->


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


</style>       


<SCRIPT LANGUAGE="JavaScript">

	
	$(document).ready(function() { 
	    $('#txtUPCCode').bind('focusout', function(e) {
	        //e.preventDefault();
	        //$(this).focus();
	    });
	});   
    
    $(window).on("load", function () {
        event.preventDefault();
        $('#txtUPCCode').focus();
        doSeek();                 
      
    });

    $(document).keypress(function (e) {
        if (e.which == 13) {
            event.preventDefault();
            doSeek();
        } 
    });

	function validateUPCForm()
	{
		if (document.frmUPCLookup.txtUPCCode.value == "") {
			swal("UPC code cannot be blank.");
			event.preventDefault(); // Prevent the page from redirecting before ackowledging swal()
			return false;
		}
		return true;
	}

    function doSeek() {
        $(".scan-result").html("");
        $(".scan-result").load("upcload.asp", { code: $('#txtUPCCode').val() }, function (response, status, xhr) {
            if (status == "error") {
                var msg = "Sorry but there was an error: ";
                $("#error").html(msg + xhr.status + " " + xhr.statusText);
             } 
            $('#txtUPCCode').val("");
        });
        $("#txtUPCCode").focus();
        
    }
        
    
</SCRIPT>

<h1 class="inventory-upc-heading"><a href="../main_menu.asp" class="left-arrow"><i class="fa fa-arrow-left pull-left" aria-hidden="true"></i></a> UPC Lookup</h1>


<!-- driver menu starts here !-->
<div class="container-fluid inventory-upc-container">

		<!-- label -->
		<label>Type or Scan UPC Code or Prod ID</label>
		<!-- eof label -->

		<!-- text / button line -->
		<div class="row">

			<!-- textbox -->
			<div class="col-lg-11 col-md-11 col-sm-11 col-xs-9">
				<input type="search" class="form-control" name="txtUPCCode" id="txtUPCCode" AUTOCOMPLETE="off">
			</div>
			<!-- eof textbox -->
			
			<!-- go button -->
			<div class="col-lg-1 col-md-1 col-sm-1 col-xs-3">
				<button class="btn btn-primary btn-go" onClick="javascript:doSeek();">GO</button>		
			</div>
			<!-- eof go button -->
		</div>
		<!-- eof text / button line -->
    	<div class="row">
            <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 scan-result">
            </div>
    	</div>
	
</div>

<!--#include file="../../../inc/footer-mobile.asp"-->