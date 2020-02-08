<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"--> 


<%

Server.ScriptTimeout = 900000 'Default value

dummy = MUV_Write("ShowServiceTicketAlertSwalShown","false")

Set cnnCustFilters = Server.CreateObject("ADODB.Connection")
cnnCustFilters.open (Session("ClientCnnString"))


'Special code for when they are brought here by the automated email
'in this case, it just resets everything to default values and
'runs the page just for the salesperson who logged in
'it does this by writing to the Settings_reports table so the 
'rest of the code can just run normally from that point

%>

<div class="waitdiv d-none" style="position: fixed;z-index: 999999999; top: 0px; left: 0px; width: 100%; height:80%; background-color:transparent; text-align: center; padding-top: 20%; filter: alpha(opacity=0); opacity:0; "></div>
	<div id="waitdiv" class="waitdiv d-none small" style="padding-bottom: 90px;text-align: center; vertical-align:middle;padding-top:50px;background-color:#ebebeb;width:300px;height:100px;margin: 0 auto; top:40%; left:40%;position:absolute;-webkit-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); -moz-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); z-index:999999999;">
		<img src="<%= BaseURL %>/img/loading_gray.gif" alt="" /><br /><span id="waitmsg">Loading Customers.</span> <br />Please wait ...
</div>
 
<style>

	.bs-example-modal-lg-customize .row{
		margin-bottom: 10px;
	 	width: 100%;
		overflow: hidden;
	}
	
	.bs-example-modal-lg-customize .left-column{
		background: #eaeaea;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}
	
	.bs-example-modal-lg-customize .left-column h4{
		margin-top: 0px;
	}
	
	.bs-example-modal-lg-customize .right-column{
		background: #fff;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}

	.filter-search-width{
		max-width: 36%;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
	    content: " \25B4\25BE" 
	    
	}
	
	table.sortable thead {
	    color:#222;
	    font-weight: bold;
	    cursor: pointer;
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

	.vpc-variance-header{
		background: #D43F3A;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}
	
	.vpc-3pavg-header{
		background: #F0AD4E;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}
	
	.vpc-lcp-header{
		background: #337AB7;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.pricing-header{
		background: #5cb85c;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.vpc-misc-header{
		background: #db27d8;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.vpc-current-header{
		background: #5CB85C;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.gen-info-header{
		background: #3B579D;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}
	
	.activities-header{
		background: #9C02FE;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}


	.negative{
		font-weight:bold;
		color:red;	
	}

	.negative-thin{
		color:red;	
	}


	.positive-thin{
		color:green;	
	}

	.positive{
		font-weight:bold;
		color:green;	
	}


	.neutral{
		font-weight:bold;
		color:black;
	}

	.smaller-header{
		font-size: 0.8em;
		vertical-align: top !important;
		text-align: center;
	}	

	.smaller-detail-line{
		font-size: 0.8em;
	}	

	.smaller-detail-line-r{
		font-size: 0.8em;
		font-weight:bold;
		color:red;
	}	
	
	.not-as-small-detail-line{
		font-size: 0.9em;
	}
	
	.table-top .table > tbody > tr > td, .table > tbody > tr > th, .table > tfoot > tr > td, .table > tfoot > tr > th, .table > thead > tr > td, .table > thead > tr > th{
		border: 1px solid #ddd !important;
	}	



	.modal.modal-wide .modal-dialog {
	  width: 50%;
	}
	.modal-wide .modal-body {
	  overflow-y: auto;
	}
	
	.modal.modal-xwide .modal-dialog {
	  width: 70%;
	}
	.modal-xwide .modal-body {
	  overflow-y: auto;
	  max-height:600px;
	}
	
	.bs-example-modal-lg-customize .row{
		margin-bottom: 10px;
	 	width: 100%;
		overflow: hidden;
		font-size:11px;
	}
	
	.bs-example-modal-lg-customize .left-column{
		background: #eaeaea;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}
	
	.bs-example-modal-lg-customize .left-column h4{
		margin-top: 0px;
	}
	
	.bs-example-modal-lg-customize .right-column{
		background: #fff;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}

	.custom-table{
	 	font-size: 11px;
	}

	#tableSuperSum .hidden {
   		position: absolute !important;
   		top: -9999px !important;
   		left: -9999px !important;
	    /*display:none; */
	}


	.ole {
		font-family: consolas, courier, monaco, menlo, monospace;
		background: rgb(200,240,255);
		padding: 6px 10px;
		display: inline-block;
		font-size: 1.2em;
		border-radius: 4px;
		border:0;
		cursor: pointer;
		color: #000;
	}
	
	.ole:hover {
		background: dodgerblue;
		color: #fff;
		text-shadow: 1px 1px 1px #000;
		box-shadow: 0 0 0 #555;
	}
	
	/*customizing tooltip color*/
	
	/*left tooltip*/
	.tooltip> .tooltip-arrow {
		border-left-color: dodgerblue;
	}
	
	/*tooltip inner*/
	.tooltip > .tooltip-inner {
		background-color: dodgerblue;
		text-shadow: 0 1px 1px #000;
		font-weight: normal;
	}
	
	.ajaxRowView .visibleRowEdit, .ajaxRowEdit .visibleRowView { display: none; }
	
	.hidden {
		display : none ;
	}
	.small {
		font-size:12px;
	}
	table.dataTable tbody tr.group {background-color: #f0f0f0; border-top:2px solid #000000;}
	table.dataTable.display tbody tr.group td {
    border-top: 2px solid #000000;
}
.red-bkg {background-color:red; font-weight:bold;}
.d-none {display:none;}
.font-bold {
	font-weight:bold;
}
</style>

<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.18/sorting/currency.js"></script>

<script type="text/javascript">

	var datatableWidget;
	

	function doEditCustomer(obj) {
	
		currentCustomerID=$(obj).attr("data-id");
		
		if (currentCustomerID.length>0) {
			window.location.href = "editViewCustomerDetail.asp?customerID="+currentCustomerID;
		}
	}


		
	$(document).ready(function() {
	
	 	var collapsedGroups = {};
		
	    $("#PleaseWaitPanel").hide();
		
	    $("#AddedCustID").val("");
		
	    $("[rel='tooltip']").tooltip('destroy');
		$("[rel='tooltip']").tooltip({ placement: 'left' });
		
	
	    //$('[data-tooltip="tooltip"]').tooltip();
	    var groupColumn = 0;
		datatableIni(1);
		
		$('#ViewMode a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
			datatableWidget.clear();
			datatableWidget.destroy();
			if($("#ViewMode li.active > a").attr("data-grouping")=="1") {
				datatableIni(1);
			}
			else {
				datatableIni(2);
			}
			
		});
		
		function datatableIni(type) {
		
			switch (type) {
				case 1:

				datatableWidget=$('#tableSuperSum').on('preXhr.dt', function ( e, settings, data ) {
				$(".waitdiv").removeClass("d-none");
				})
				.on('xhr.dt', function ( e, settings, json, xhr ) {
					$(".waitdiv").addClass("d-none");
					
				} )
				.DataTable({
		        scrollY: 500,
		        scrollCollapse: true,
		        paging: true,
				ajax: "datatableArCustomerJSON.asp?maxID=1",
				deferRender: true,
				serverSide:true,
				procesing: true,
				lengthMenu: [[10, 25, 50, 75, 100, 500, -1],[10, 25, 50, 75, 100, 500, "All"]],
				pageLength: 100,
				order: [[ 0, "asc" ]],
				columns: [
						{ "data": "id" },
						{ "data": "CustName" },
						{ "data": "City" },
						{ "data": "State" },	
						{ "data": "Zip" },	
						{ "data": "LastPriceChangeDate" },	
						{ "data": "Action" },
					],
					
				columnDefs: [
					{"orderable": true,"targets": [0,1,2,3,4,5] },
					{"orderable": false,"targets": [6] },
					{"className": "dt-center", "targets": [2,3,4,5,6]},
					{"className": "text-left", "targets": [0,1]},
					//{"visible": false, "targets": [13]},
					{"visible": true, "targets": [0,1,2,3,4,5,6]},		
					{"targets": [ 6 ],	
					"createdCell": function (td, cellData, rowData, row, col) {
						 var nameClient=rowData.CustName;
						 var clientID=rowData.id;
						
						  if($("#ViewMode li.active > a").attr("data-grouping")=="1") {
							$(td).html('<div class="dropdown"><button class="btn btn-success dropdown-toggle" type="button" id="dropdownMenu1" data-toggle="dropdown">Action</button><ul class="dropdown-menu dropdown-menu-right" aria-labelledby="dropdownMenu1"><li><a data-id="'+clientID+'" href="#" onclick="javascript:doEditCustomer(this);">Edit Customer</a></li><li><a data-id="'+clientID+'" href="#" onclick="javascript:toExclude(this);">Change to Inactive</a></li></ul></div>');
						  }
						}				
					
					}
					
				],
				initComplete : function() {
					var input = $('.dataTables_filter input').unbind(),
		            self = this.api(),
		            $searchButton = $('<button>')
		                       .text('search')
		                       .click(function() {
		                          self.search(input.val()).draw();
		                       }),
		            $clearButton = $('<button>')
		                       .text('clear')
		                       .click(function() {
		                          input.val('');
		                          $searchButton.click(); 
		                       }) 
		        	$('.dataTables_filter').append("&nbsp;",$searchButton, $clearButton);
	    		}   
				
				});
			
				break;
			
			
			case 2:
			
				datatableWidget=$('#tableSuperSum').on('preXhr.dt', function ( e, settings, data ) {
				$(".waitdiv").removeClass("d-none");
				})
				.on('xhr.dt', function ( e, settings, json, xhr ) {
					$(".waitdiv").addClass("d-none");
					
				} )
				.DataTable({
		        scrollY: 500,
		        scrollCollapse: true,
		        paging: true,
				ajax: "datatableArCustomerJSON.asp?maxID=2",
				deferRender: true,
				serverSide:true,
				procesing: true,
				lengthMenu: [[10, 25, 50, 75, 100, 500, -1],[10, 25, 50, 75, 100, 500, "All"]],
				pageLength: 100,
				order: [[ 0, "asc" ]],
				createdRow: function ( row, data, index ) {
					if($("#ViewMode li.active > a").attr("data-grouping")=="1") {
					$(row).attr("data-child-value",data.id).css("display","none");
					}
				},	
				columns: [
						{ "data": "id" },
						{ "data": "CustName" },
						{ "data": "City" },
						{ "data": "State" },	
						{ "data": "Zip" },		
						{ "data": "LastPriceChangeDate" },
						{ "data": "Action" },
					],
					
				columnDefs: [
					{"orderable": true,"targets": [0,1,2,3,4,5] },
					{"orderable": false,"targets": [6] },
					{"className": "dt-center", "targets": [2,3,4,5,6]},
					{"className": "text-left", "targets": [0,1]},
					//{"visible": false, "targets": [13]},
					{"visible": true, "targets": [0,1,2,3,4,5,6]},		
					{"targets": [ 6 ],	
					"createdCell": function (td, cellData, rowData, row, col) {
						 var nameClient=rowData.CustName;
						 var clientID=rowData.id;
						
						  if($("#ViewMode li.active > a").attr("data-grouping")=="0") {
							$(td).html('<div class="dropdown"><button class="btn btn-success dropdown-toggle" type="button" id="dropdownMenu1" data-toggle="dropdown">Action</button><ul class="dropdown-menu dropdown-menu-right" aria-labelledby="dropdownMenu1"><li><a data-id="'+clientID+'" href="#" onclick="javascript:doEditCustomer(this);">Edit Customer</a></li><li><a data-id="'+clientID+'" href="#" onclick="javascript:toExclude(this);">Change To Inactive</a></li></ul></div>');
						  }
						}				
					
					}
					
				],
				initComplete : function() {
					var input = $('.dataTables_filter input').unbind(),
		            self = this.api(),
		            $searchButton = $('<button>')
		                       .text('search')
		                       .click(function() {
		                          self.search(input.val()).draw();
		                       }),
		            $clearButton = $('<button>')
		                       .text('clear')
		                       .click(function() {
		                          input.val('');
		                          $searchButton.click(); 
		                       }) 
		        	$('.dataTables_filter').append("&nbsp;",$searchButton, $clearButton);
	    		}   
				
				});
			
				break;
			
			}
			
		
		
		}
			
		
	});



	function ajaxRowMode(type, id, mode) {
	
		$('#ajaxRow'+type+'-'+id).attr("class", "ajaxRow"+mode);
		if(id==0){
			$('#ajaxRow'+type+'-' + 0 + '').remove();
		}	
	
		 $(".ajaxRowEdit").find('input[disabled="true"]').each(function () {
		     $(this).removeAttr("disabled");
		 });
		
	}

</script>




<h3 class="page-header"><i class="fad fa-users"></i>&nbsp;Manage Customers&nbsp;&nbsp;
	<a href="<%= BaseURL %>accountsreceivable/customermininfo/addCustomer.asp">
		<button type="button" class="btn btn-success" id="btnAddCustomer">
			<i class="fas fa-users-medical"></i>&nbsp;Add Customer
		</button>
	</a>
</h3>

 


<!-- row !-->
<div class="row">
<!-- Nav tabs -->
  <ul class="nav nav-tabs" id="ViewMode" role="tablist">
    <li role="presentation" class="active"><a href="#bycustomer" role="tab" data-toggle="tab" data-grouping="1">Active Customers</a></li>
    <li role="presentation"><a href="#detailed" aria-controls="profile" role="tab" data-toggle="tab" data-grouping="0">Inactive Customers</a></li>
  </ul>
  

  <div class="container-fluid" style="padding-top:20px;">
		<div class="row">
           <table id="tableSuperSum" class="display compact" style="width:100%;">
              <thead>
                  	<tr>	
						<th class="td-align1 vpc-3pavg-header" colspan="2" style="border-right: 2px solid #555 !important;"><%= GetTerm("Customer") %> Information</th>
						<th class="td-align1 vpc-lcp-header" colspan="3" style="border-right: 2px solid #555 !important;">Location</th>
						<th class="td-align1 pricing-header" style="border-right: 2px solid #555 !important;">Pricing</th>
						<th class="td-align1 activities-header" style="border-right: 2px solid #555 !important;">Activities</th>
					</tr>
				
					<tr>
						<th class="td-align smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;">Acct</th>
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;">Client</th>

						<th class="td-align smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;">City</th>
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;">State</th>
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;">Zip</th>
						
						<th class="td-align smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;">Last Price Change</th>
						
						<th class="td-align smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;">Action</th>
					</tr>
				</thead>
			</table>
		</div>
    </div>





<!-- eof row !-->

<!-- row !-->
<div class="row">

<div class="col-lg-12"><hr></div>
</div>
<!-- eof row !-->

<!-- row !-->
<div class="row">



</div>
<!-- eof row !-->

<!--#include file="customerModals.asp"-->

<!--#include file="../../inc/footer-main.asp"-->