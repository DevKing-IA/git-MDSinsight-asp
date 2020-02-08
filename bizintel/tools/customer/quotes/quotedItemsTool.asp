<!--#include file="../../../../inc/header.asp"-->
<!--#include file="../../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../../inc/InSightFuncs_BizIntel.asp"-->

<style>

#loadingmodal {
    display:    none;
    position:   fixed;
    z-index:    1000;
    top:        0;
    left:       0;
    height:     100%;
    width:      100%;
    background: rgba( 255, 255, 255, .8 )  
                url('../../../../img/preloader.gif') 
                50% 50% 
                no-repeat;
}

#loadingmodal {
    overflow: hidden;   
}
#loadingmodal {
    display: block;
}
</style>


<div id="loadingmodal"><h1>Loading Quoted Items</h1></div>

<script>
 
	$(window).on('load', function (e) {
	    $('#loadingmodal').fadeOut(1000);
	})
	
</script>

    
<!-- Bootstrap DataTables JS -->
	<script src="https://cdn.datatables.net/1.10.13/js/jquery.dataTables.min.js"></script>
	<script src="https://cdn.datatables.net/1.10.13/js/dataTables.bootstrap.min.js"></script>
	<script src="https://cdn.datatables.net/select/1.2.1/js/dataTables.select.min.js"></script>
	<script src="https://cdn.datatables.net/buttons/1.2.4/js/dataTables.buttons.min.js"></script>
<!-- End Bootstrap DataTables JS -->	


<!-- Bootstrap DataTables CSS -->	
	<link href="https://cdn.datatables.net/1.10.13/css/dataTables.bootstrap.min.css" rel="stylesheet">
	<link href="https://cdn.datatables.net/select/1.2.1/css/select.dataTables.min.css" rel="stylesheet">
	<link href="https://cdn.datatables.net/buttons/1.2.4/css/buttons.dataTables.min.css" rel="stylesheet">

<!-- End Bootstrap DataTables CSS -->    	
<!-- datepicker for EXPIRED DATE !-->
	<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
	<link href="<%= baseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet" type="text/css">
	<script src="<%= baseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.js" type="text/javascript"></script>
<!-- end datepicker for EXPIRED DATE !-->


<%
	custID = Request.QueryString("custID")
	
	If custID = "" Then
		custID = Request.Form("txtCustID")
	End If
	
	radHighlightBasedOn = Request.Form("radHighlightBasedOn")
	QuotedAtZero = Request.Form("QuotedAtZero")
	ZeroOrders = Request.Form("ZeroOrders")
	QuotedViaChain = Request.Form("QuotedViaChain")
	selNumMonthsHistoricalImpact = Request.Form("selNumMonthsHistoricalImpact")
	
	If selNumMonthsHistoricalImpact = "" Then selNumMonthsHistoricalImpact = 6
	
	function UDate(oldDate)
    	UDate = DateDiff("s", "01/01/1970 00:00:00", oldDate)
	end function
%>



<script language="JavaScript">

	function round(value, exp) {
	  if (typeof exp === 'undefined' || +exp === 0)
	    return Math.round(value);
	
	  value = +value;
	  exp = +exp;
	
	  if (isNaN(value) || !(typeof exp === 'number' && exp % 1 === 0))
	    return NaN;
	
	  // Shift
	  value = value.toString().split('e');
	  value = Math.round(+(value[0] + 'e' + (value[1] ? (+value[1] + exp) : exp)));
	
	  // Shift back
	  value = value.toString().split('e');
	  return +(value[0] + 'e' + (value[1] ? (+value[1] - exp) : -exp));
	}	

	$(document).ready(function() {
	

		var table = $('#quotedItemsTable').DataTable({
				"scrollY":        "600px",
				"scrollCollapse": false,
				"columnDefs": [
				        {"className": "dt-center", "targets": "_all"}
				      ],		
		        "lengthMenu": [[10, 25, 50, -1], [10, 25, 50, "All"]],
		        "order": [[ 3, "asc" ],[1, "asc" ]],
		        "stateSave": false,
				"select": true,
				"paging": false,	
				"dom": '<"pull-left"f><"pull-right"l>tip',
				"rowCallback": function( row, data, index ) {
				  	if ((data[16] != "---") && (data[16] != "*")) {
				  	
				  		var oldGP = Number(data[12].replace(/[^0-9\.]+/g,""));
				  		var newGP = Number(data[16].replace(/[^0-9\.]+/g,""));
				  			  	
					    if ( newGP < oldGP ) {
						      $('td:eq(16)', row).removeClass("highlight-orange-ish");
						      $('td:eq(17)', row).removeClass("highlight-orange-ish");
						      $('td:eq(16)', row).removeClass("highlight-green");
						      $('td:eq(17)', row).removeClass("highlight-green");
						      $('td:eq(16)', row).addClass("highlight-red");
						      $('td:eq(17)', row).addClass("highlight-red");
					    } else if (newGP > oldGP ) {
						      $('td:eq(16)', row).removeClass("highlight-red");
						      $('td:eq(17)', row).removeClass("highlight-red");
						      $('td:eq(16)', row).removeClass("highlight-orange-ish");
						      $('td:eq(17)', row).removeClass("highlight-orange-ish");
						      $('td:eq(16)', row).addClass("highlight-green");
						      $('td:eq(17)', row).addClass("highlight-green");
					    } else if (newGP == oldGP ) {
						      $('td:eq(16)', row).removeClass("highlight-green");
						      $('td:eq(17)', row).removeClass("highlight-green");
						      $('td:eq(16)', row).removeClass("highlight-red");
						      $('td:eq(17)', row).removeClass("highlight-red");
						      $('td:eq(16)', row).addClass("highlight-orange-ish");
						      $('td:eq(17)', row).addClass("highlight-orange-ish");
					    }
		
				    }
				  }
			
		 });	 
		 

	
		if( $.cookie('bizintel-hist-impact-months') === '6' ){
		
			var cookie_value = '6';
			$('#selNumMonthsHistoricalImpact option[value=cookie_value]').prop('selected',true);
		}
		
		
		if( $.cookie('bizintel-highlight-items-value') === 'IMPACT' ){
		
            $('#radHighlightBasedOnImpact').prop('checked', true);
 
        	var table = $("#quotedItemsTable");

		    table.find('tr').each(function (i) {
		    	
		    	var trID = $(this).attr("id");
		        var $tds = $(this).find('td'),
		            productId = $tds.eq(1).text(),
		            product = $tds.eq(2).text(),
		            numOrders = $tds.eq(6).text();
	              	
		            if (numOrders == 0) {
		            
				      	$('#' + trID + '-16').removeClass("highlight-green");
				      	$('#' + trID + '-17').removeClass("highlight-green");
				      	$('#' + trID + '-16').removeClass("highlight-red");
				      	$('#' + trID + '-17').removeClass("highlight-red");					    
				      	$('#' + trID + '-16').addClass("highlight-orange-ish");
				      	$('#' + trID + '-17').addClass("highlight-orange-ish");
		            
		            }
		    });
		}
         else if( $.cookie('bizintel-highlight-items-value') === 'GPCHANGE' ){
         
            $('#radHighlightBasedOnGPChange').prop('checked', true);
            var table = $('#quotedItemsTable').DataTable();
            table.draw();
         }
         
         
		
        // this code will run after all other $(document).ready() scripts
        // have completely finished, AND all page elements are fully loaded.
	        
	    // Check if alert has been closed
	    if( $.cookie('bizintel-hide-zero-item-orders') === 'hide' ){
	
			$('#chkZeroOrders').prop('checked', true);
			
			$.fn.dataTableExt.afnFiltering.push(
			  function (oSettings, aData, iDataIndex) {
				    var ZeroOrders= $('#chkZeroOrders').is(':checked'); //True
				    var element = $(oSettings.aoData[iDataIndex].nTr);
				    return element.is('.notZeroOrders') || ! ZeroOrders;
			  });
			var table = $('#quotedItemsTable').DataTable();
			table.draw();
	    }
	    
	    
	    if( $.cookie('bizintel-hide-items-quoted-at-zero') === 'hide' ){
	
   			$('#chkQuotedAtZero').prop('checked', true);
   			
   			$.fn.dataTableExt.afnFiltering.push(
			  function (oSettings, aData, iDataIndex) {
			   	    var QuotedViaChain= $('#chkQuotedViaChain').is(':checked'); //True
				    var element2 = $(oSettings.aoData[iDataIndex].nTr);
					return element2.is('.notQuotedViaChain') || ! QuotedViaChain;
			  });
			var table = $('#quotedItemsTable').DataTable();
			table.draw();
	    }


	    if( $.cookie('bizintel-hide-items-quoted-via-chain') === 'hide' ){
	
   			$('#chkQuotedViaChain').prop('checked', true);
   			
   			$.fn.dataTableExt.afnFiltering.push(
			  function (oSettings, aData, iDataIndex) {
		   	    var QuotedViaChain= $('#chkQuotedViaChain').is(':checked'); //True
			    var element2 = $(oSettings.aoData[iDataIndex].nTr);
				return element2.is('.notQuotedViaChain') || ! QuotedViaChain;
			  });
			var table = $('#quotedItemsTable').DataTable();
			table.draw();
 
	    }
		    
			 
		//table.column(0).visible(false);
		
	    $('#quotedItemsTable tbody').on( 'click', 'tr', function () {
	        var table = $('#quotedItemsTable').DataTable();
	        if ( $(this).hasClass('selected') ) {
	            $(this).removeClass('selected');
	        }
	        else {
	            table.$('tr.selected').removeClass('selected');
	            $(this).addClass('selected');
	        }
	    } );

	  	
		    		
		$('#chkZeroOrders').change(function () { 
		
			/* If you just want the cookie for a session don't provide an expires
	        Set the path as root, so the cookie will be valid across the whole site */
	        var ZeroOrdersChecked = $('#chkZeroOrders').is(':checked'); //True
	        if (ZeroOrdersChecked == true) {
	        	$.cookie('bizintel-hide-zero-item-orders', 'hide', { path: '/' });
	        }
	        else {
	        	$.cookie('bizintel-hide-zero-item-orders', 'show', { path: '/' });
	        }

			$.fn.dataTableExt.afnFiltering.push(
			
			  function (oSettings, aData, iDataIndex) {
			  
			    var ZeroOrders= $('#chkZeroOrders').is(':checked'); //True
			    var element = $(oSettings.aoData[iDataIndex].nTr);
			    return element.is('.notZeroOrders') || ! ZeroOrders;
				
			  });
			var table = $('#quotedItemsTable').DataTable();
			table.draw();
		});
		
		
		
   		$('#chkQuotedAtZero').change(function () { 
   		
			/* If you just want the cookie for a session don't provide an expires
	        Set the path as root, so the cookie will be valid across the whole site */
	        var QuotedAtZeroChecked = $('#chkQuotedAtZero').is(':checked'); //True
	        
	        if (QuotedAtZeroChecked == true) {
	        	$.cookie('bizintel-hide-items-quoted-at-zero', 'hide', { path: '/' });
	        }
	        else {
	        	$.cookie('bizintel-hide-items-quoted-at-zero', 'show', { path: '/' });
	        }
   		
	   		
			$.fn.dataTableExt.afnFiltering.push(
			
			  function (oSettings, aData, iDataIndex) {
			  
		   	    var QuotedAtZero= $('#chkQuotedAtZero').is(':checked'); //True
			    var element1 = $(oSettings.aoData[iDataIndex].nTr);
				return element1.is('.notQuotedAtZero') || ! QuotedAtZero;
				
			  });
   			var table = $('#quotedItemsTable').DataTable();
   			table.draw();
   		});
   		
   		
   		$('#chkQuotedViaChain').change(function () { 
   		
   		
			/* If you just want the cookie for a session don't provide an expires
	        Set the path as root, so the cookie will be valid across the whole site */
	        var QuotedViaChainChecked = $('#chkQuotedViaChain').is(':checked'); //True
	        if (QuotedViaChainChecked == true) {
	        	$.cookie('bizintel-hide-items-quoted-via-chain', 'hide', { path: '/' });
	        }
	        else {
	        	//$.removeCookie("bizintel-hide-items-quoted-via-chain");
	        	$.cookie('bizintel-hide-items-quoted-via-chain', 'show', { path: '/' });
	        	
	        }
   		   		
   		
   			$.fn.dataTableExt.afnFiltering.push(
		
			  function (oSettings, aData, iDataIndex) {
			    
		   	    var QuotedViaChain= $('#chkQuotedViaChain').is(':checked'); //True
			    var element2 = $(oSettings.aoData[iDataIndex].nTr);
				return element2.is('.notQuotedViaChain') || ! QuotedViaChain;
				
			  });
			var table = $('#quotedItemsTable').DataTable();
   			table.draw(); 
   		});
   		
   		
		 
		$(".date").each(function(){

		    $(this).datetimepicker({
		    	format: 'MM/DD/YYYY',
		    	useCurrent: false,
		    	minDate:moment(),
		    	maxDate:moment().add(24,'months'),
	        	ignoreReadonly: true,
	        	showClear: true, 	
	    	});
	    	
	    });

		    	    
		$("[id^='datepicker']").on("dp.change", function (e) {
		
	    	selectedDate = $(this).find("input").val();
	    	datepickerID = $(this).find("input").attr('id');
	    	
	    	//var IntRecID = datepickerID.substr(datepickerID.length-1);
	    	var IntRecID = datepickerID.match(/\d+/);
	    		        
			$.ajax({
				type:"POST",
				url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				async:false,
				data: "action=UpdateExpireDateSingleQuotedItem&recid="+encodeURIComponent(IntRecID)+"&expdate="+encodeURIComponent(selectedDate),
				success: function(msg)
				{
					//location.reload();
				}
			}) 
	    });	
	    
		$('#quotedItemsTable').on('draw.dt', function() {

			$(".date").each(function(){
	
			    $(this).datetimepicker({
			    	format: 'MM/DD/YYYY',
			    	useCurrent: false,
			    	minDate:moment(),
			    	maxDate:moment().add(24,'months'),
		        	ignoreReadonly: true,
		        	showClear: true, 	
		    	});
		    	
		    });
	
   
			$("[id^='datepicker']").on("dp.change", function (e) {
			
		    	selectedDate = $(this).find("input").val();
		    	datepickerID = $(this).find("input").attr('id');
		    	
		    	//var IntRecID = datepickerID.substr(datepickerID.length-1);
		    	var IntRecID = datepickerID.match(/\d+/);
		    		        
				$.ajax({
					type:"POST",
					url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
					async:false,
					data: "action=UpdateExpireDateSingleQuotedItem&recid="+encodeURIComponent(IntRecID)+"&expdate="+encodeURIComponent(selectedDate),
					success: function(msg)
					{
						//location.reload();
					}
				}) 
		    });	
	    
		});		
	    
	   // $("#quotedItemsTable_paginate").on("click", "a", function() { alert("clicked") });
				    
	    $('#dateselectExpireDateAllProducts').datetimepicker({
	    	format: 'MM/DD/YYYY',
	    	useCurrent: false,
	    	minDate:moment(),
	    	maxDate:moment().add(48, 'months'),
        	ignoreReadonly: true,
        	showClear: true,

	    });

		
		$("#dateselectExpireDateAllProducts").on("dp.change", function (e) {
	    	selectedDate = $("#dateselectExpireDateAllProducts").find("input").val();
	    });	
	    

	    $('#deleteQuotedItem').click( function () {


				$('tr.selected').each(function(index,item){
				
				    if(parseInt($(item).data('index'))>0){
				    	
				        IntRecID = $(item).data('index');
				        ProdSKU = $(item).data('sku');
				        
						$.ajax({
							type:"POST",
							url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
							async:false,
							data: "action=DeleteQuotedItemFromCustomer&recid="+encodeURIComponent(IntRecID),
							success: function(msg)
							{
						        if (msg.indexOf("CHAIN") >= 0) {
						        	var result = msg.split(',');
						        	var currProdSKU = result[1];
						        	var currUM = result[2];
						            swal("One or more of the selected items is part of a chain quote and cannot be deleted.");
						        }
						        else {
						        	var result = msg.split(',');
						        	var currProdSKU = result[1];
						        	var currUM = result[2];
						        	var currRowID = parseInt(result[3]);
						        	table.row('.account.selected').remove().draw(false);
						        }
							}
						}) 
 
				    }
				});			
	    });	

	    $('#autoQuoteAllAlternateUMS').click( function () {

			custID = $('#txtCustID').val();
			
			$.ajax({
				type:"POST",
				url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				data: "action=AutoQuoteAllAlternateUMSsForCustomer",
				success: function(msg)
				{
			        location.reload();
				},
			    beforeSend: function() {
			        $('#loadingmodal').show();	
			    },	
			    complete: function() {
			        //$('#loadingmodal').hide();
			    }			    			
			}) 
	    });	


	    $(document).on('click','[name="autoQuoteAlternateUM"]',function(){

			var IntRecID = $(this).attr('id').match(/\d+/);
			
			$.ajax({
				type:"POST",
				url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				data: "action=AutoQuoteSingleUMForCustomer&recid="+encodeURIComponent(IntRecID),
				success: function(msg)
				{
		        	location.reload();
				},
			    beforeSend: function() {
			        $('#loadingmodal').show();	
			    },	
			    complete: function() {
			        //$('#loadingmodal').hide();
			    }					
			}) 
	
	    });	


	    $('#undoAllChanges').click( function () {

			$.ajax({
				type:"POST",
				url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				data: "action=UndoQuotedItemChangesForCustomer",
				success: function(msg)
				{
			        location.reload();
				},
			    beforeSend: function() {
			        $('#loadingmodal').show();	
			    },	
			    complete: function() {
			        //$('#loadingmodal').hide();
			    }					
			}) 
	
	    });	
   	    	
	

	    $(document).on('click','[name="btnUndoSingleItemChanges"]',function(){

			var IntRecID = $(this).attr('id').match(/\d+/);
			
			$.ajax({
				type:"POST",
				url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				data: "action=UndoSingleQuotedItemChangeForCustomer&recid="+encodeURIComponent(IntRecID),
				success: function(msg)
				{
		        	$("#newGPDollars" + IntRecID).html("");
					$("#txtNewPrice" + IntRecID).val("");
					$("#txtNewGPPercent" + IntRecID).val("");
					$("#txtProductPriceExpireDate" + IntRecID).val("");
					$("#changeGPPercentTrend" + IntRecID).html("");
				}
			}) 
	
	    });	
		
		$( "#btnExpireDateAllItems" ).click(function() {

		  	selectedDate = $("#txtExpireDateAllProducts").val();
		  	
		  	if (selectedDate != "") {
		  	
				$.ajax({
					type:"POST",
					url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
					async:false,
					data: "action=UpdateExpireDateAllQuotedItems&expdate="+encodeURIComponent(selectedDate),
					success: function(msg)
					{
						location.reload();
					}
				}) 
			}
			
			else {
				swal("Please select a date for all items to expire on.")
			}
		});


        $(document).on('click','[name="txtNewPrice"]',function(){
			var IntRecID = $(this).attr('id').match(/\d+/);
			$("#txtCurrentRecID").val(IntRecID);
        });		

        $(document).on('click','[name="txtNewGPPercent"]',function(){
			var IntRecID = $(this).attr('id').match(/\d+/);
			$("#txtCurrentRecID").val(IntRecID);
        });		

        $(document).on('focus','[name="txtNewPrice"]',function(){
			var IntRecID = $(this).attr('id').match(/\d+/);
			$("#txtCurrentRecID").val(IntRecID);
        });		

        $(document).on('focus','[name="txtNewGPPercent"]',function(){
			var IntRecID = $(this).attr('id').match(/\d+/);
			$("#txtCurrentRecID").val(IntRecID);
        });		
		
		
        $(document).on('focusout','[name="txtNewPrice"]',function(){
			var currentIntRecID = $("#txtCurrentRecID").val();
			var modifiedIntRecID = $(this).attr('id').match(/\d+/);
						
			if ((parseInt(currentIntRecID) == parseInt(modifiedIntRecID)) && $(this).val() != '') 
			{
				var newPrice = $(this).val();
				//alert("Changing price for record " + currentIntRecID + "/" + modifiedIntRecID + " to " + newPrice);

				$.ajax({
					type:"POST",
					url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
					async:false,
					data: "action=UpdateNewPriceSingleQuotedItem&recid="+encodeURIComponent(modifiedIntRecID)+"&newprice="+encodeURIComponent(newPrice),
					success: function(msg)
					{
						//alert(msg);
						
						var fields = msg.split('*');
						var currentCost = fields[0];
						var quotedPrice = fields[1];	
						var changeMessage = fields[2];
						
						if (newPrice == 0) {
							var newGPDollars = round(newPrice-currentCost,2);
							var newGPPercent = 0;
						}
						else {
							var newGPDollars = round(newPrice-currentCost,2);
							var newGPPercent = round((newGPDollars/newPrice)*100,2);
						}
						
					    if (changeMessage == 'DECREASE') {
					   		$('#' + modifiedIntRecID + '-16').removeClass("highlight-orange-ish");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-orange-ish");
					      	$('#' + modifiedIntRecID + '-16').removeClass("highlight-green");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-green");	
					      	$('#' + modifiedIntRecID + '-16').addClass("highlight-red");
					      	$('#' + modifiedIntRecID + '-17').addClass("highlight-red");
					    } else if (changeMessage == 'INCREASE') {
					   		$('#' + modifiedIntRecID + '-16').removeClass("highlight-orange-ish");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-orange-ish");
					      	$('#' + modifiedIntRecID + '-16').removeClass("highlight-red");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-red");	
					      	$('#' + modifiedIntRecID + '-16').addClass("highlight-green");
					      	$('#' + modifiedIntRecID + '-17').addClass("highlight-green");
					    } else if (changeMessage == 'NOCHANGE') {
					      	$('#' + modifiedIntRecID + '-16').removeClass("highlight-green");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-green");
					      	$('#' + modifiedIntRecID + '-16').removeClass("highlight-red");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-red");					    
					      	$('#' + modifiedIntRecID + '-16').addClass("highlight-orange-ish");
					      	$('#' + modifiedIntRecID + '-17').addClass("highlight-orange-ish");
					    }
	
						$("#newGPDollars" + modifiedIntRecID).html("");
						$("#txtNewGPPercent" + modifiedIntRecID).val(newGPPercent);
						$("#newGPDollars" + modifiedIntRecID).html("$" + newGPDollars);
						
						$("#DiffTotRevenueProjected").html("please recalc");
						$("#DiffTotGPDollars").html("please recalc");
						$("#DiffTotGPPercent").html("please recalc");
						
					}
				}) 
			}
        });		


		
	  $(document).on('focusout','[name="txtNewGPPercent"]',function(){
		
			var currentIntRecID = $("#txtCurrentRecID").val();
			var modifiedIntRecID = $(this).attr('id').match(/\d+/);
						
			if ((parseInt(currentIntRecID) == parseInt(modifiedIntRecID)) && $(this).val() != '') 
			{
				var newGPPercent = $(this).val();

				$.ajax({
					type:"POST",
					url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
					async:false,
					data: "action=UpdateNewGPPercentSingleQuotedItem&recid="+encodeURIComponent(modifiedIntRecID)+"&gpp="+encodeURIComponent(newGPPercent),
					success: function(msg)
					{
						//alert(msg);
						
						var fields = msg.split('*');
						var currentCost = fields[0];
						var quotedPrice = fields[1];
						var changeMessage = fields[2];
	
						
						if (newGPPercent == 100) {
							var NewGPPercentDecimal = newGPPercent/100;
							var newPrice = currentCost * 2;
							var newGPDollars = round(newPrice-currentCost,2);
						}
						else {
							var NewGPPercentDecimal = newGPPercent/100;
							var newPrice = currentCost/(1 - NewGPPercentDecimal);
							var newGPDollars = round(newPrice-currentCost,2);
						}
						

					    if (changeMessage == 'DECREASE') {
					   		$('#' + modifiedIntRecID + '-16').removeClass("highlight-orange-ish");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-orange-ish");
					      	$('#' + modifiedIntRecID + '-16').removeClass("highlight-green");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-green");	
					      	$('#' + modifiedIntRecID + '-16').addClass("highlight-red");
					      	$('#' + modifiedIntRecID + '-17').addClass("highlight-red");
					    } else if (changeMessage == 'INCREASE') {
					   		$('#' + modifiedIntRecID + '-16').removeClass("highlight-orange-ish");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-orange-ish");
					      	$('#' + modifiedIntRecID + '-16').removeClass("highlight-red");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-red");	
					      	$('#' + modifiedIntRecID + '-16').addClass("highlight-green");
					      	$('#' + modifiedIntRecID + '-17').addClass("highlight-green");
					    } else if (changeMessage == 'NOCHANGE') {
					      	$('#' + modifiedIntRecID + '-16').removeClass("highlight-green");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-green");
					      	$('#' + modifiedIntRecID + '-16').removeClass("highlight-red");
					      	$('#' + modifiedIntRecID + '-17').removeClass("highlight-red");					    
					      	$('#' + modifiedIntRecID + '-16').addClass("highlight-orange-ish");
					      	$('#' + modifiedIntRecID + '-17').addClass("highlight-orange-ish");
					    }
												
						$("#newGPDollars" + modifiedIntRecID).html("");
						$("#txtNewPrice" + modifiedIntRecID).val(round(newPrice,2));
						$("#newGPDollars" + modifiedIntRecID).html("$" + newGPDollars);
						
						$("#DiffTotRevenueProjected").html("please recalc");
						$("#DiffTotGPDollars").html("please recalc");
						$("#DiffTotGPPercent").html("please recalc");
						


					}
				}) 
			}

			
		});
  
		
		$("input[name='radHighlightBasedOn']").change(function(){
		
            selected_value = $("input[name='radHighlightBasedOn']:checked").val();
 
	        $.cookie('bizintel-highlight-items-value', selected_value, { path: '/' });

            if (selected_value == 'IMPACT'){
            
            	var table = $("#quotedItemsTable");

			    table.find('tr').each(function (i) {
			    	
			    	var trID = $(this).attr("id");
			        var $tds = $(this).find('td'),
			            productId = $tds.eq(1).text(),
			            product = $tds.eq(2).text(),
			            numOrders = $tds.eq(6).text();
		              	
			            if (numOrders == 0) {
			            
					      	$('#' + trID + '-16').removeClass("highlight-green");
					      	$('#' + trID + '-17').removeClass("highlight-green");
					      	$('#' + trID + '-16').removeClass("highlight-red");
					      	$('#' + trID + '-17').removeClass("highlight-red");					    
					      	$('#' + trID + '-16').addClass("highlight-orange-ish");
					      	$('#' + trID + '-17').addClass("highlight-orange-ish");
			            
			            }
			    });
            
            }
            else if (selected_value == 'GPCHANGE') {
            	var table = $('#quotedItemsTable').DataTable();
            	table.draw();
            }
        });
		
		

		$('#selNumMonthsHistoricalImpact').on('change', function(e) {
		
            selected_value = $('#selNumMonthsHistoricalImpact').val();
	        $.cookie('bizintel-hist-impact-months', selected_value, { path: '/' });
	        $('#loadingmodal').show();
			$("#frmQuotedItemsTool").submit();	        

        });
        
        

		$('#addQuotedItemModal').on('show.bs.modal', function(e) {
		
		    
		    var custID = $('input[name="txtCustID"]').val();
		    	    
		    var $modal = $(this);
			$('#loadingOverlayCategories').show();
						
	    
	    	$.ajax({
				type:"POST",
				url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				cache: false,
				data: "action=GetCategoryInformationForAddQuotedItemModal&custID=" + encodeURIComponent(custID),
			    beforeSend: function(){
			        $('#loadingOverlay').show();
					$('#categoryInformationForCustomer').hide();
					$('#productInformationForCustomer').hide();
					$('#editableProductFieldsForAddQuotedItems').hide();
			    },	
			    complete: function(){
			        $('#loadingOverlayCategories').hide();
			    },			    			
				success: function(response)
				 {
	               	 $('#loadingOverlayCategories').hide();
	               	 $('#categoryInformationForCustomer').show();
	               	 $modal.find('#categoryInformationForCustomer').html(response);	               	 
	             },
	            failure: function(response)
				 {
				   $modal.find('#categoryInformationForCustomer').html("Product Load Failed");
	             }
			});
  
		});	
		


		$( "#btnSendQuotedItemsToMetroplex" ).click(function() {

				swal({
				  title: "Are you sure you want to send?",
				  text: "You will not be able to undo changes posted to Metroplex for this cutomer's quoted items!",
				  type: "warning",
				  showCancelButton: true,
				  confirmButtonColor: "#DD6B55",
				  confirmButtonText: "Yes, Send To Metroplex!",
				  closeOnConfirm: false
				},
				function(){
				  window.location.href = 'writeQuotesloading.asp?custid=<%=custID %>';
				});

		});
		


		$( "#btnChangeCustomer" ).click(function() {

			$.cookie('bizintel-hide-zero-item-orders', 'show', { path: '/' })
			$.cookie('bizintel-hide-items-quoted-at-zero', 'show', { path: '/' });
			$.cookie('bizintel-hide-items-quoted-via-chain', 'show', { path: '/' });
			$.cookie('bizintel-highlight-items-value', 'GPCHANGE', { path: '/' });
			$.cookie('bizintel-hist-impact-months','6', { path: '/' });
			window.location.href = 'reports.asp';
		});
		
		
	  
	    $('.wrapper1').on('scroll', function (e) {
	        $('.wrapper2').scrollLeft($('.wrapper1').scrollLeft());
	    }); 
	    $('.wrapper2').on('scroll', function (e) {
	        $('.wrapper1').scrollLeft($('.wrapper2').scrollLeft());
	    });	  
	    
	    
	    
	    
	    
});
	
</script>

<style>

	.form-inline .form-control {
	    display: inline-block;
	    width: 80px;
	    vertical-align: middle;
	    text-align: right;
	}

	.form-inline .input-group {
	    display: inline-table;
	    vertical-align: middle;
	    width: 135px;
	}	
	
	.dollarSignSpan {
	    float: left;
	    margin-left: 5px;
	    margin-top: 8px;
	    position: absolute;
	    z-index: 2;
	    color: green;
	}	
	
	h2.chain {
		color:#337ab7; 
		padding-bottom:10px;
	}
	
	/* Center Align Columns 1, 4-17  */
	th:nth-child(1).dt-center, td:nth-child(1).dt-center { text-align: center; }
	th:nth-child(4).dt-center, td:nth-child(4).dt-center { text-align: center; }
	th:nth-child(5).dt-center, td:nth-child(5).dt-center { text-align: center; }
	th:nth-child(6).dt-center, td:nth-child(6).dt-center { text-align: center; }
	th:nth-child(7).dt-center, td:nth-child(7).dt-center { text-align: center; }
	th:nth-child(8).dt-center, td:nth-child(8).dt-center { text-align: center; }
	th:nth-child(9).dt-center, td:nth-child(9).dt-center { text-align: center; }
	th:nth-child(10).dt-center, td:nth-child(10).dt-center { text-align: center; }
	th:nth-child(11).dt-center, td:nth-child(11).dt-center { text-align: center; }
	th:nth-child(12).dt-center, td:nth-child(12).dt-center { text-align: center; }
	th:nth-child(13).dt-center, td:nth-child(13).dt-center { text-align: center; }
	th:nth-child(14).dt-center, td:nth-child(14).dt-center { text-align: center; }
	th:nth-child(15).dt-center, td:nth-child(15).dt-center { text-align: center; }
	th:nth-child(16).dt-center, td:nth-child(16).dt-center { text-align: center; }
	th:nth-child(17).dt-center, td:nth-child(17).dt-center { text-align: center; }
	th:nth-child(18).dt-center, td:nth-child(18).dt-center { text-align: center; }
	th:nth-child(19).dt-center, td:nth-child(19).dt-center { text-align: center; }

	
	div.dataTables_wrapper div.dataTables_filter input {
	    margin-left: 0.5em;
	    display: inline-block;
	    width: auto;
	   /* margin-top: 20px;*/
	}
	
.the-table thead,tbody{
		font-size: 11px;
	}

.no-arrows::after{
	display: none;
	visibility: hidden;
}

.highlight-green{
	background: #eafcde;
}

.highlight-red{
	background: #fcdede;
}

.highlight-orange-ish{
	background: #FAF4D7;
}


.highlight-multi-colors{
	background: -webkit-linear-gradient(top, #cfeffc, #cfeffc 50%, #eafcde 50%, #eafcde);
}


.row-margin{
	margin-bottom: 15px;
}

.accordion-line{
	width: 100%;
	float: left;
}

.historical-impact{
	margin-top: 30px;
}

.radios-checkboxes{
	margin-top: 35px;
}

.radios-checkboxes ul{
	list-style-type: none;
 }
</style>


<!-- accordion line starts here !-->
<form name="frmQuotedItemsTool" id="frmQuotedItemsTool" action="quotedItemsTool.asp" method="POST">
	

	<div class="row">
		
		<div class="row">

			<div class="col-lg-7">
				<%
				custChainNum = GetCustChainNum(custID)
				If custChainNum <> 0 Then %>
					<div class="alert alert-info alert-dismissable" style="margin-top:15px;margin-bottom:0px;">
						<a href="#" class="close" data-dismiss="alert" aria-label="close">×</a>
						<strong>This <%=GetTerm("account")%> is part of a chain.&nbsp;&nbsp;&nbsp;Chain ID:<%=custChainNum%>&nbsp;&nbsp;&nbsp;Chain Name:<%= GetChainDescByChainNum(custChainNum)%></strong>.<br><br>
					</div>
				<% End If %>

		 	<h3 class="page-header"><i class="fa fa-file-text-o"></i> <%=GetTerm("Customer")%> Quoted Items for Account <%= custID %>, <%= GetCustNameByCustNum(custID) %></h3>
		 	
			<div class="panel-group accordion-line" id="accordion" role="tablist" aria-multiselectable="true">
				<div class="panel panel-default">
					<div class="panel-heading" role="tab" id="headingOne">
						<h4 class="panel-title">
							<a role="button" data-toggle="collapse" data-parent="#accordion" href="#collapseOne" aria-expanded="false" aria-controls="collapseOne">
								<i class="fa fa-wrench" aria-hidden="true"></i> Click for Tools 
							</a>
						</h4>
					</div>
					<div id="collapseOne" class="panel-collapse collapse" role="tabpanel" aria-labelledby="headingOne">
						<div class="panel-body">
		
							<div class="col-lg-4 reports-box">
									<div class="row" style="margin-bottom:5px;">
							
									    <button type="button" class="btn btn-primary btn-lg btn-block" style="margin-bottom:10px;" id="btnChangeCustomer">
									        <i class="fa fa-user"></i>&nbsp;Change <%=GetTerm("Customer")%>
									    </button>
									    <!--
									    <button type="button" class="btn btn-success btn-lg btn-block" style="margin-bottom:10px;" data-toggle="modal" data-target="#addQuotedItemModal" id="addNewQuotedItem">
									        <i class="fa fa-plus"></i>&nbsp;Add New Quoted Item
									    </button>
									    -->
		
									    <button type="button" class="btn btn-disabled btn-lg btn-block" style="margin-bottom:10px;">
									        <i class="fa fa-plus"></i>&nbsp;Add New Quoted Item - Coming Soon
									    </button>
							
									    <button type="button" class="btn btn-success btn-lg btn-block" style="margin-bottom:10px;" id="autoQuoteAllAlternateUMS">
									        <i class="fas fa-envelope-open-dollar"></i>&nbsp;Auto Quote All Alt UM's
									    </button>
									    
									    <button type="button" id="deleteQuotedItem" class="btn btn-danger btn-lg btn-block" style="margin-bottom:10px;">
									        <i class="fas fa-trash-alt"></i>&nbsp;Delete Quoted Item
									    </button>
									    		    
									    <button type="button" id="undoAllChanges" class="btn btn-warning btn-lg btn-block" style="margin-bottom:10px;">
									        <i class="fa fa-undo"></i>&nbsp;Undo All Changes
									    </button>		    
							
									</div>
							</div>
							
							<div class="col-lg-1 reports-box">
							&nbsp;
							</div>
							
							<div class="col-lg-5 reports-box">		
								<div class="row">
									<div class="form-group">
										<label for="txtExpireDateAllProductse"><i class="fa fa-calendar-times-o" aria-hidden="true"></i> Expire Date All Products</label>
										<div class="input-group date" id="dateselectExpireDateAllProducts">
											<input type="text" class="form-control" name="txtExpireDateAllProducts" id="txtExpireDateAllProducts" readonly="readonly">
											<span class="input-group-addon">
											<span class="glyphicon glyphicon-calendar"></span>
											</span>
										</div>
									</div>
								</div>
								
									<div class="row pull-right row-margin">
										<button type="button" class="btn btn-primary" id="btnExpireDateAllItems">Set Date</button>
		 									</div>
		 									
		 									<div class="row">
			 									 <button type="button" class="btn btn-primary btn-lg btn-block pull-left" style="margin-bottom:10px;" id="btnSendQuotedItemsToMetroplex">
									        <i class="fa fa-share-square"></i>&nbsp;Send Quoted Items To <%=GetTerm("Backend")%>
									    </button> 
		 									</div>
								</div>
								
								 	
							</div>
						</div>
					</div>
				</div>		
			</div>	
				<!-- legends !-->
		<!-- checkbox / radios -->
		<div class="col-lg-2 radios-checkboxes">
			<ul>
				<li>
					<input type="radio" id="radHighlightBasedOnImpact" name="radHighlightBasedOn" value="IMPACT" <% If radHighlightBasedOn = "IMPACT" Then Response.Write("Checked='Checked'") %>> Highlight based on impact
				</li>
				<li>
					<input type="radio" id="radHighlightBasedOnGPChange" name="radHighlightBasedOn" value="GPCHANGE" <% If radHighlightBasedOn = "GPCHANGE" Then Response.Write("Checked='Checked'") %>> Highlight based on GP change
				</li>
			</ul>
			<ul>
				<li>
					<input type="checkbox" id="chkQuotedAtZero" name="QuotedAtZero"  <% If QuotedAtZero = "on" Then Response.Write("Checked='Checked'") %>> Hide items quoted at $0
				</li>
			<li>
				<input type="checkbox" id="chkZeroOrders" name="ZeroOrders" <% If ZeroOrders = "on" Then Response.Write("Checked='Checked'") %>> Hide items with 0 orders
			</li>
			<li>
				<input type="checkbox" id="chkQuotedViaChain" name="QuotedViaChain" <% If QuotedViaChain = "on" Then Response.Write("Checked='Checked'") %>> Hide items quoted via chain
			</li>
			</ul>
		</div>
		<!-- eof checkbox / radios -->
		
			<div class="col-lg-3 pull-right historical-impact">
			

				<label for="lblHistLine" style="margin-bottom:10px;">Historical Impact</label>&nbsp;
				
				<div class="col-lg-12" style="margin-left:-30px;">
				
					<div class="col-lg-8" style="margin-bottom:10px;">
						<select name="selNumMonthsHistoricalImpact" id="selNumMonthsHistoricalImpact" class="form-control">
							<option value="3" <% If cInt(selNumMonthsHistoricalImpact) = 3 Then Response.Write("selected") %>>Last 3 mos</option>
							<option value="6" <% If cInt(selNumMonthsHistoricalImpact) = 6 Then Response.Write("selected") %>>Last 6 mos</option>
							<option value="12" <% If cInt(selNumMonthsHistoricalImpact) = 12 Then Response.Write("selected") %>>Last 12 mos</option>
							<option value="18" <% If cInt(selNumMonthsHistoricalImpact) = 18 Then Response.Write("selected") %>>Last 18 mos</option>
							<option value="24" <% If cInt(selNumMonthsHistoricalImpact) = 24 Then Response.Write("selected") %>>Last 24 mos</option>
						</select>
					</div>
				
					<div class="col-lg-4" style="margin-bottom:10px;">
						<button type="submit" class="btn btn-primary" id="btnRecalc">Recalculate</button>
					</div>
					
					<div class="col-lg-12">
						<!--#include file="historical_impact.asp"-->
					</div>
				
				</div>
								 
			</div>
		
		</div> <!-- end row -->

		
		<input type="hidden" name="txtCustID" id="txtCustID" value="<%= custID %>">
		<input type="hidden" name="txtCurrentRecID" id="txtCurrentRecID">
	
		 

<!-- row !-->
<div class="row">
	<div class="col-lg-12">
			    <!-- Content Here -->		
				<table id="quotedItemsTable" class="table table-striped table-bordered the-table" cellspacing="0" width="100%">
		        <thead>
		            <tr>
		            	<th>C/A</th>
		                <th>SKU</th>
		                <th>DESCRIPTION</th>
		                <th>CATEGORY</th>
		                <th>UM</th>
		                <th>LIST<br>FLAG</th>
		                <th># ORDS</th>
		                <th>LAST PURCH</th>
		                <th>DATE QUOTED</th>
		                <th>EXPIRE DATE</th>
		                <th>NEW EXP DATE</th>
		                <th>COST</th>
		                
		                <th>GP $</th>
		                <th>GP %</th>
		                <th>QUOTED PRICE</th>
		                <th>NEW PRICE</th>
		                <th>NEW GP $</th>  
		                <th>NEW GP %</th>
		                <th class="no-arrows">AUTO QUOTE</th>	
		                <th class="no-arrows">RESET</th>
		            </tr>
		        </thead>
		        <tbody>
				<%       
				Set rsQuotedItems = Server.CreateObject("ADODB.Recordset")
				rsQuotedItems.CursorLocation = 3 
			
				SQLQuotedItems = "SELECT * FROM zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " WHERE DeleteFlag <> 1"
				'ORDER BY MUST BE DONE USING DATATABLES.JS
				
				Set rsImpactSummary = Server.CreateObject("ADODB.Recordset")
								
				Set cnnQuotedItems = Server.CreateObject("ADODB.Connection")
				cnnQuotedItems.open (Session("ClientCnnString"))
				Set rsQuotedItems = cnnQuotedItems.Execute(SQLQuotedItems)
				
				LineCounter = 1
				
				If NOT rsQuotedItems.EOF Then
				
					Do While NOT rsQuotedItems.EOF
					
					
			            QuotedToChainOrAccount = rsQuotedItems("QuotedToChainOrAccount")
			            
	                	SQLImpactSummary = "SELECT * FROM zPRC_AccountQuotedItems_Impact_Summary_" & trim(Session("Userno")) & " WHERE prodSKU = '" & rsQuotedItems("ProdSKU") & "' "
	                	SQLImpactSummary = SQLImpactSummary & " AND UM = '" & rsQuotedItems("QuoteType") & "'"
	                	'Response.Write("<BR>"& SQLImpactSummary & "<br>")
	                	Set rsImpactSummary = cnnQuotedItems.Execute(SQLImpactSummary)
	                	
	                	MostRecentOrdDate = ""
	                	NumTimesOrdered = 0
	                	
						If Not rsImpactSummary.EOF Then
							If Not IsNull(rsImpactSummary("MostRecentOrdDate")) Then MostRecentOrdDate = rsImpactSummary("MostRecentOrdDate")
							NumTimesOrdered = rsImpactSummary("NumTimesOrdered")
						End If

			            
			            
			     %>
		                <% If QuotedToChainOrAccount = "C" Then
		                
				               trclass = ""
				               If NumTimesOrdered > 0 Then trclass = trclass & " notZeroOrders "
				               If cDbl(rsQuotedItems("Price")) > 0 Then trclass = trclass & " notQuotedAtZero "	
				               trclass = trclass & " QuotedViaChain "	
				               								
								If trclass <> "" Then
									%><tr class="<%= trclass %>" id="<%= rsQuotedItems("InternalRecordIdentifier") %>" data-index="<%= rsQuotedItems("InternalRecordIdentifier") %>" data-sku="<%= rsQuotedItems("ProdSKU") %>" data-unit="<%= rsQuotedItems("QuoteType") %>"><%
								Else
									%><tr id="<%= rsQuotedItems("InternalRecordIdentifier") %>" data-index="<%= rsQuotedItems("InternalRecordIdentifier") %>" data-sku="<%= rsQuotedItems("ProdSKU") %>" data-unit="<%= rsQuotedItems("QuoteType") %>">><%
								End If
								%>
								

		                		<!--<tr id="<%= rsQuotedItems("InternalRecordIdentifier") %>" data-index="<%= rsQuotedItems("InternalRecordIdentifier") %>" data-sku="<%= rsQuotedItems("ProdSKU") %>" data-unit="<%= rsQuotedItems("QuoteType") %>" style="color:#ccc;">-->
			                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-0"><i class="fa fa-link" aria-hidden="true"></i>&nbsp;<%= QuotedToChainOrAccount %></td>
				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-1"><%= rsQuotedItems("ProdSKU") %></td>
				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-2"><%= rsQuotedItems("Description") %></td>
				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-3" data-order="<%= cint(rsQuotedItems("Category")) %>"><%= cint(rsQuotedItems("Category")) %>&nbsp;<%= GetCategoryByID(rsQuotedItems("Category")) %></td>
				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-4"><%= rsQuotedItems("QuoteType") %></td>
				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-5"><%= rsQuotedItems("ListFlag") %></td>
				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-6"><%= NumTimesOrdered %></td>

								<% If MostRecentOrdDate <> "" Then %>
									<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-7" data-order="<%= UDate(MostRecentOrdDate) %>"><%= FormatDateTime(MostRecentOrdDate,2) %></td>
								<% Else %>
									<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-7" data-order="0">NA</td>
								<% End If %>
										                
								<% If rsQuotedItems("DateQuoted") <> "" Then %>
									<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-8" data-order="<%= UDate(rsQuotedItems("DateQuoted")) %>"><%= rsQuotedItems("DateQuoted") %></td>
								<% Else %>
									<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-8" data-order="0">NA</td>
								<% End If %>
				                
								<% If rsQuotedItems("ExpireDate") <> "" Then %>
									<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-9" data-order="<%= UDate(rsQuotedItems("ExpireDate")) %>"><%= rsQuotedItems("ExpireDate") %></td>
								<% Else %>
									<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-9" data-order="0">NA</td>
								<% End If %>

				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-10"><%= rsQuotedItems("NewExpireDate") %></td> 		
								<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-11"><%= formatCurrency(rsQuotedItems("Cost"),2) %></td>                

				                <% If rsQuotedItems("Price") <> "" AND NOT IsNull(rsQuotedItems("Price")) AND NOT IsEmpty(rsQuotedItems("Price")) AND rsQuotedItems("Price") > 0 AND rsQuotedItems("Cost") <> ""  Then %>
				                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-12"><%= formatCurrency(rsQuotedItems("Price") - rsQuotedItems("Cost"),2) %></td>
				                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-13">
				                	<%
					                GPPerc = ((rsQuotedItems("Price") - rsQuotedItems("Cost"))/rsQuotedItems("Price")) * 100
					                GPPerc = ROund(GPPerc,2)
					                Response.Write(FormatNumber(GPPerc,2))%>%</td>
						        <% Else %>
				                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-14">*</td>
				                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-15">*</td>
				                <% End If %>              
								<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-16"><%= formatCurrency(rsQuotedItems("Price"),2) %></td>
				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-17">*</td>
				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-18">*</td>
				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-19">*</td>
				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-20">*</td>
				                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-21">*</td>
			                </tr>
			                
				                	
			                <% Else 
			                
				                	QuotedPrice = rsQuotedItems("Price")
				                	NewPrice = rsQuotedItems("NewPrice")
				                	NewGPPercent = rsQuotedItems("NewGPPercent")
				                	CurrentCost = rsQuotedItems("Cost")
				                	
				                	If NewGPPercent <> "" AND NOT IsNull(NewGPPercent) Then
				                	
										If NewGPPercent = 100 Then
											NewGPPercentDecimal = NewGPPercent/100		
											NewPrice = round(CurrentCost * 2,2)
						                	NewGPDollars = NewPrice -CurrentCost
											NewGPDollars = Round(NewGPDollars,2)
										Else
											NewGPPercentDecimal = NewGPPercent/100		
											NewPrice = round(CurrentCost/(1 - NewGPPercentDecimal),2)
						                	NewGPDollars = NewPrice -CurrentCost
											NewGPDollars = Round(NewGPDollars,2)
										End If
										
									ElseIf NOT IsNull(NewPrice) AND NewPrice <> "0" Then
									
						                NewGPDollars = NewPrice -CurrentCost
										NewGPPercent = Round((NewGPDollars/NewPrice)*100,2)
										NewGPDollars = Round(NewGPDollars,2)

									ElseIf NOT IsNull(NewPrice) AND NewPrice = "0" Then
									
						                NewGPDollars = NewPrice - CurrentCost
										'NewGPPercent = Round((NewGPDollars/NewPrice)*100,2)
										NewGPPercent = 0
										NewGPDollars = Round(NewGPDollars,2)	
										
									Else
										NewGPDollars = ""
										NewGPPercent = ""
									End If
									
									changeGPPercentTrendVar = "UNCHANGED"
									
					                If NewGPDollars <> "" AND NewGPDollars <> "0" Then
					                	If QuotedPrice > 0 Then
					                		CurrentGPPercent = formatNumber(((QuotedPrice - CurrentCost)/QuotedPrice ) * 100,2) 
					                		
					                		If cDbl(NewGPPercent) > cDbl(CurrentGPPercent) Then 
								                changeGPPercentTrendVar = "UP"
											ElseIf NewGPPercent < CurrentGPPercent Then
								                changeGPPercentTrendVar = "DOWN"
								            Else   
								                changeGPPercentTrendVar = "UNCHANGED"
								            End If
					                	End If
						            End If    	

					                
					                %>

					                <!-- <tr class="account highlight-multi-colors" !-->
					               <%
					               
		                
				               		trclass = ""
				               		If NumTimesOrdered > 0 Then trclass = trclass & " notZeroOrders "
				               		If cDbl(rsQuotedItems("Price")) > 0 Then trclass = trclass & " notQuotedAtZero "	
				               		trclass = trclass & " notQuotedViaChain "

									If trclass <> "" Then
										%><tr class="<%= trclass %>" id="<%= rsQuotedItems("InternalRecordIdentifier") %>" data-index="<%= rsQuotedItems("InternalRecordIdentifier") %>" data-sku="<%= rsQuotedItems("ProdSKU") %>" data-unit="<%= rsQuotedItems("QuoteType") %>"><%
									Else
										%><tr id="<%= rsQuotedItems("InternalRecordIdentifier") %>" data-index="<%= rsQuotedItems("InternalRecordIdentifier") %>" data-sku="<%= rsQuotedItems("ProdSKU") %>" data-unit="<%= rsQuotedItems("QuoteType") %>"><%
									End If
									%>

				                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-0"><i class="fa fa-building" aria-hidden="true"></i>&nbsp;<%= QuotedToChainOrAccount %></td>			                
					                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-1"><%= rsQuotedItems("ProdSKU") %></td>
					                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-2"><%= rsQuotedItems("Description") %></td>
					                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-3" data-order="<%= cint(rsQuotedItems("Category")) %>"><%= cint(rsQuotedItems("Category")) %>&nbsp;<%= GetCategoryByID(rsQuotedItems("Category")) %></td>
					                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-4"><%= rsQuotedItems("QuoteType") %></td>
					                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-5"><%= rsQuotedItems("ListFlag") %></td>
					                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-6"><%= NumTimesOrdered %></td>	
					                
									<% If MostRecentOrdDate <> "" Then %>
										<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-7" data-order="<%= UDate(MostRecentOrdDate) %>"><%= FormatDateTime(MostRecentOrdDate,2) %></td>
									<% Else %>
										<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-7" data-order="0">NA</td>
									<% End If %>
											                
									<% If rsQuotedItems("DateQuoted") <> "" Then %>
										<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-8" data-order="<%= UDate(rsQuotedItems("DateQuoted")) %>"><%= rsQuotedItems("DateQuoted") %></td>
									<% Else %>
										<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-8" data-order="0">NA</td>
									<% End If %>
					                
									<% If rsQuotedItems("ExpireDate") <> "" Then %>
										<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-9" data-order="<%= UDate(rsQuotedItems("ExpireDate")) %>"><%= rsQuotedItems("ExpireDate") %></td>
									<% Else %>
										<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-9" data-order="0">NA</td>
									<% End If %>
					                
					                
					                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-10">
						                <div class="input-group date" id="datepicker<%= rsQuotedItems("InternalRecordIdentifier") %>">
						                    <input type="text" class="form-control" name="txtProductPriceExpireDate" id="txtProductPriceExpireDate<%= rsQuotedItems("InternalRecordIdentifier") %>"  value="<%= rsQuotedItems("NewExpireDate") %>" readonly="readonly">
						                    <span class="input-group-addon">
						                        <span class="glyphicon glyphicon-calendar"></span>
						                    </span>
						                </div>
					                </td>
				                			
									<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-11"><%= formatCurrency(rsQuotedItems("Cost"),2) %></td>                
					                
					                
					                <% If rsQuotedItems("Price") <> "" AND NOT IsNull(rsQuotedItems("Price")) AND NOT IsEmpty(rsQuotedItems("Price")) AND rsQuotedItems("Price") > 0 AND rsQuotedItems("Cost") <> ""  Then %>
					                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-12"><%= formatCurrency(rsQuotedItems("Price") - rsQuotedItems("Cost"),2) %></td>
					                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-13">
					                	<%
					                	GPPerc = ((rsQuotedItems("Price") - rsQuotedItems("Cost"))/rsQuotedItems("Price")) * 100
					                	GPPerc = ROund(GPPerc,2)
					                	Response.Write(FormatNumber(GPPerc,2))%>%</td>
							        <% Else %>
					                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-12">NA</td>
					                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-13">NA</td>
					                <% End If %>              
									<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-14"><%= formatCurrency(rsQuotedItems("Price"),2) %></td>
					                
					                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-15"><span class="fa fa-usd dollarSignSpan"></span><input type="text" id="txtNewPrice<%= rsQuotedItems("InternalRecordIdentifier") %>" name="txtNewPrice" value="<%= NewPrice %>" class="form-control last-run-inputs"></td>
					                
					                
					                <% If NewGPDollars <> "" AND NewGPDollars <> "0" Then %>
					                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-16" id="newGPDollars<%= rsQuotedItems("InternalRecordIdentifier") %>">$<%= NewGPDollars %></td>
					                <% Else %>
					                	<td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-16" id="newGPDollars<%= rsQuotedItems("InternalRecordIdentifier") %>">---</td>
									<% End If %>
									
					                <td id="<%= rsQuotedItems("InternalRecordIdentifier") %>-17"><span class="fa fa-percent dollarSignSpan"></span><input type="text" id="txtNewGPPercent<%= rsQuotedItems("InternalRecordIdentifier") %>" name="txtNewGPPercent" value="<%= NewGPPercent %>" class="form-control last-run-inputs"></td>
					                
					                
					                
					                <%			
					                If rsQuotedItems("AutoGenerated") = -1 Then%>
					                	<td id="autoQuoteAlternateUM">Auto<br>Quoted</td>
									<%Else 						
										ShowQuoteAltUMButton = True
										CaseConversionFactoroZeroOrOne = False
										SkuCountOfOne = False
										QuoteTypeOfN = False
										ProdSKUCheck = rsQuotedItems("ProdSKU")
										
										Set rsAltUMButtonCheck = Server.CreateObject("ADODB.Recordset")
										rsAltUMButtonCheck.CursorLocation = 3 
	
										SQLAltUMButtonCount = "SELECT COUNT (prodSKU) AS skuCount FROM zPRC_AccountQuotedItems_" & trim(Session("Userno"))
										SQLAltUMButtonCount = SQLAltUMButtonCount & " WHERE prodSKU='" & rsQuotedItems("ProdSKU") & "'"
	
										Set rsAltUMButtonCheck = cnnQuotedItems.Execute(SQLAltUMButtonCount)
	
										If NOT rsAltUMButtonCheck.EOF Then
											If rsAltUMButtonCheck("skuCount") > 1 OR rsQuotedItems("QuoteType") = "N" Then 
												ShowQuoteAltUMButton = False
												If rsAltUMButtonCheck("skuCount") > 1 Then
													SkuCountOfOne = True
												End If
												If rsQuotedItems("QuoteType") = "N" Then
													QuoteTypeOfN = True
												End If
											Else
												'Also have to make sure the case conversion factor is not 0
												Set rsProductCheck = Server.CreateObject("ADODB.Recordset")
												rsProductCheck.CursorLocation = 3 
												SQLProductCheck = "SELECT * FROM Product WHERE PartNo = '" & ProdSKUCheck & "'"
												Set rsProductCheck = cnnQuotedItems.Execute(SQLProductCheck)
												If not rsProductCheck.Eof Then
													If rsProductCheck("CaseConversionFactor") = 0 or rsProductCheck("CaseConversionFactor") = 1 Then
														ShowQuoteAltUMButton = False
														CaseConversionFactoroZeroOrOne = True
													End If
												End IF
	
										Set rsAltUMButtonCheck = cnnQuotedItems.Execute(SQLAltUMButtonCount)


											End IF
											skuCount = rsAltUMButtonCheck("skuCount")
										End IF
	
										SET rsAltUMButtonCheck = Nothing
										%>
		
						   			    <% If ShowQuoteAltUMButton = False Then 
							   			    If CaseConversionFactoroZeroOrOne = True Then %>
	   						                	<td id="autoQuoteAlternateUM">N/A when conv<br> factor is: <%= rsProductCheck("CaseConversionFactor") %></td>
	   						                <% ElseIf SkuCountOfOne = True Then %>
	   						                	<td id="autoQuoteAlternateUM">Alt UM<br>quote exists</td>
	   						                <% ElseIf QuoteTypeOfN = True Then %>
	   						                	<td id="autoQuoteAlternateUM">N/A when<br>UM is N</td>
	   						                <% Else %>
	   						                	<td id="autoQuoteAlternateUM">&nbsp;</td>
	   						                <% End If %>
						   			    
						                	
						                <% Else %>
						                	<%  If rsQuotedItems("QuoteType") = "C" then 
						                			alternateProductUM = "U"
						                		ElseIf rsQuotedItems("QuoteType") = "U" then
						                			alternateProductUM = "C"
						                		End If
						                	%>
						                	<td id="autoQuoteAlternateUM"><button type="button" class="btn btn-success btn-sm" id="autoQuoteSinglelUM<%= rsQuotedItems("InternalRecordIdentifier") %>" name="autoQuoteAlternateUM"><i class="fas fa-envelope-open-dollar" aria-hidden="true"></i></button></td>
										<% End If
									End If%>
									
									
									<td><button type="button" class="btn btn-warning btn-sm btnUndoSingleItemChanges" id="btnUndoSingleItemChanges<%= rsQuotedItems("InternalRecordIdentifier") %>"><i class="fa fa-undo" aria-hidden="true"></i></button></td>
	  
														                
					              </tr>
				              <% End If %>
			            
						<%		
						rsQuotedItems.MoveNext
						
						LineCounter = LineCounter + 1
						
						Loop
					End If
					
					set rsQuotedItems = Nothing
					%>		
		        </tbody>
		    </table></div>
	

</div>


</form>
 


	<!-- modal placeholder for add quoted item modal begins here !-->
	 <!-- Modal -->
	 
		<div class="modal fade" id="addQuotedItemModal" tabindex="-1" role="dialog" aria-labelledby="addQuotedItemModalLabel">
		  <div class="modal-dialog" role="document">
		  
			<script>
			
				$(document).ready(function() {

				    $('#datepicker2').datetimepicker({
				    	format: 'MM/DD/YYYY',
				    	useCurrent: false,
				    	defaultDate: moment().add(365, 'days'),
				    	minDate:moment(),
				    	maxDate:moment().add(24, 'months')
				    });
				    
					$("#datepicker2").on("dp.change", function (e) {
				    	selectedDate = $("#datepicker2").find("input").val();
				    });	    
			    
				}); 
					  
			   function quotedItemCategorySelected() {
			   
				    var custID = $('input[name="txtCustID"]').val();
				    var categoryID = $("#selAddQuotedItemCategories option:selected").val();
			    
			    	$.ajax({
						type:"POST",
						url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
						cache: false,
						data: "action=GetProductInformationForAddQuotedItemModal&custID=" + encodeURIComponent(custID) + "&categoryID=" + encodeURIComponent(categoryID),
					    complete: function(){
					        $('#loadingOverlayProducts').hide();
					    },			    			
						success: function(response)
						 {
			               	 $('#loadingOverlayProducts').hide();
						     $('#productInformationForCustomer').show();
					         $('#editableProductFieldsForAddQuotedItems').show();
			               	 $('#productInformationForCustomer').html(response);	               	 
			             },
			            failure: function(response)
						 {
						   $('#productInformationForCustomer').html("Product Load Failed");
			             }
					});
			    }
			    
				function quotedItemSelected() {
				
					var $modal = $(this);
					var custID = $('input[name="txtCustID"]').val();
					$('#editableProductFieldsForAddQuotedItems').show();
					
				}	
		   
		
		  </script>

		  
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="addQuotedItemModalLabel">Add Quoted Item To Account <%= custID %>, <%= GetCustNameByCustNum(custID) %></h4>
		      </div>
		      
				

		      <form name="frmAddQuotedItemToCustomerFromModal" id="frmAddQuotedItemToCustomerFromModal" action="addQuotedItemToCustomerFromModal.asp" method="POST">
		      		
			      <div class="modal-body">   
			      
						<div class="col-lg-12" id="loadingOverlayCategories" style="display:none; text-align:center;">
							<label for="loadingOverlayCategories">Preparing Available Products For This Account, Please Wait..</label>
							<img src="<%= BaseURL %>img/preloader.gif">
						</div>  
						            					  
					  	<div class="col-lg-12" id="categoryInformationForCustomer">
					  	<!-- Content for the current activity in this modal will be generated and written here -->
						<!-- Content generated by Sub GetCategoryInformationForAddQuotedItemModal() in InsightFuncs_AjaxForBizIntelModals.asp -->
					  	</div>
					
						<div class="col-lg-12" id="loadingOverlayProducts" style="display:none; text-align:center;">
							<label for="loadingOverlayProducts">Preparing Available Products For This Category, Please Wait..</label>
							<img src="<%= BaseURL %>img/preloader.gif">
						</div>  
	
					  	<div class="col-lg-12" id="productInformationForCustomer">
					  	<!-- Content for the current activity in this modal will be generated and written here -->
						<!-- Content generated by Sub GetProductInformationForAddQuotedItemModal() in InsightFuncs_AjaxForBizIntelModals.asp -->
					  	</div>
					  	
					  	<div id="editableProductFieldsForAddQuotedItems" style="display:none;">
					  		<div class="col-lg-5" id="cost" style="margin-left:5px; margin-top:15px;">	
								<div class="form-group">
								  <label for="txtAddQuotedItemCost"><i class="fa fa-usd" aria-hidden="true"></i> Cost:</label>
								  <input type="text" id="txtAddQuotedItemCost" name="txtAddQuotedItemCost" value="" class="form-control last-run-inputs">
								</div>
							</div>		
							<div class="col-lg-5" id="listprice" style="margin-left:5px; margin-top:15px;">	
								<div class="form-group">
								  <label for="txtAddQuotedItemListPrice"><i class="fa fa-usd" aria-hidden="true"></i> List Price:</label>
								  <input type="text" id="txtAddQuotedItemListPrice" name="txtAddQuotedItemListPrice" value="" class="form-control last-run-inputs">
								</div>
							</div>
							<div class="col-lg-5" id="gppercent" style="margin-left:5px; margin-top:15px;">	
								<div class="form-group">
								  <label for="txtAddQuotedItemGPPercent">New GP <i class="fa fa-percent" aria-hidden="true"></i>:</label>
								  <input type="text" id="txtAddQuotedItemGPPercent" name="txtAddQuotedItemGPPercent" value="" class="form-control last-run-inputs">
								</div>
							</div>			
							<div class="col-lg-5" id="gpdollars" style="margin-left:5px; margin-top:15px;">	
								<div class="form-group">
								  <label for="txtAddQuotedItemGPDollars">New GP <i class="fa fa-usd" aria-hidden="true"></i>:</label>
								  <div id="txtAddQuotedItemGPDollars">GENERATED HERE</div>
								</div>
							</div>	
							<br clear="all">
							<div class="col-lg-5" id="datequoted" style="margin-left:5px; margin-top:15px;">	
								<div class="form-group">
								  <label for="datequoted"><i class="fa fa-calendar-plus-o" aria-hidden="true"></i> Date Quoted: <%= Date() %></label>
								</div>
							</div>											
							<div class="col-lg-6" id="expireDate" style="margin-left:5px; margin-top:15px;">	
								<div class="form-group">
								  	<label for="txtAddQuotedItemExpireDate"><i class="fa fa-calendar-times-o" aria-hidden="true"></i> Expire Date:</label>
					                <div class="input-group date" id="datepicker2">
					                    <input type="text" class="form-control" name="txtAddQuotedItemExpireDate" id="txtAddQuotedItemExpireDate">
					                    <span class="input-group-addon">
					                        <span class="glyphicon glyphicon-calendar"></span>
					                    </span>
					                </div>
								</div>
							</div>
					  	</div>

					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-primary">Add Quoted Item To <%=GetTerm("Customer")%></button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- add quoted item modal ends here !-->


<!--#include file="../../../../inc/footer-main.asp"-->