<%
Server.ScriptTimeout = 900000 'Default value

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))


FilterSlsmn1 = Request.QueryString("p")



%>
<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->
<!--#include file="../../css/fa_animation_styles.css"-->
 

<%
CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Leakage Overview Primary Salesman"

ShowPercentageColumns = False

PeriodBeingEvaluated = GetLastClosedPeriodAndYear()
PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum()

WorkDaysIn3PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -3), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1))+1
WorkDaysIn12PeriodBasis =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated -12), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated -1)) + 1 
WorkDaysInLastClosedPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated )) + 1 
WorkDaysInCurrentPeriod = NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1), GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated +1)) + 1 
WorkDaysSoFar =  NumberOfWorkDays(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated +1),Date()) + 1

%>
<script type="text/javascript">
//**********************************************
	$.ajaxSetup ({
	    // Disable caching of AJAX responses
	    cache: false
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
	
	var datatableWidget;
	var datatableWidgetPrimary;
	
	var ruleColVisible;
	$(window).on("load",function() {
	
		var activeTab=$(".filter-tabs li.active");
		console.log(activeTab);
		var order=$(".filter-tabs li.active").attr("data-order");
		var tab=$(".filter-tabs li.active a").attr("href");
		var colorder=$(".filter-tabs li.active").attr("data-col-order");
		 if ($(".filter-tabs li.active a").attr("id") === "tabruledout") ruleColVisible=true;
					 else ruleColVisible=false;
		$('a[data-toggle="tab"]').on("show.bs.tab", function (e) {
			
			
			var selectedTabID=$(e.target).parent().attr("data-source");
			var order=$(e.target).parent().attr("data-order")
			var colorder=$(e.target).parent().attr("data-col-order")
			if ($(e.target).attr("id") === "tabruledout") ruleColVisible=true;
					 else ruleColVisible=false;
			
			datatableIni(selectedTabID,order,colorder,$(e.target).attr("href"));
			
			
		});
		datatableIni($(".filter-tabs li.active").attr("data-source"), order,colorder,tab);
		
	});
	
	function datatableIni(ajaxURL, orderResult, orderColumn, tab) {
		if (tab=="#tab1") {
			if (typeof datatableWidget!="undefined") {
				$("#tableData").off( 'order.dt');
				datatableWidget.clear();
				datatableWidget.destroy();
				}
			datatableWidget=$("#tableData")
		
			.on('preXhr.dt', function ( e, settings, data ) {
					$(".waitdiv").removeClass("d-none");
				})
			.on('xhr.dt', function ( e, settings, json, xhr ) {
				
				$(".waitdiv").addClass("d-none");
				
			
				})
			
			.DataTable({
	    
	        scrollY: 500,
	        scrollCollapse: true,
			ajax: ajaxURL ,
	        paging: false,
			order: [ parseInt(orderColumn), orderResult ],
			footerCallback: function ( row, data, start, end, display ) {
				
				
				
				
				var api = this.api();
				var columns=[2,4,5,6,7,8,9,10,11,12,13,14,15,16,17,20];
				// Remove the formatting to get integer data for summation
				var intVal = function ( i ) {
					return typeof i === 'string' ?
						i.replace(/[\$(),]/g, '')*1 :
						typeof i === 'number' ?
							i : 0;
				};
				// Total over all pages
				for (j=0;j<columns.length;j++) {
				
					total = api
						.column(columns[j])
						.data()
						.reduce( function (a, b) {
														
							return intVal(a) + intVal(b);
						}, 0 );
					
					// Update footer
					strTotal=total.toFixed(1).replace(/\d(?=(\d{3})+\.)/g, '$&,').substring(0, total.toFixed(1).replace(/\d(?=(\d{3})+\.)/g, '$&,').length-2);
					
					if (total<0) {
						$(api.column(columns[j]).footer()).addClass("negative");
						$(api.column(columns[j]).footer()).html(strTotal.substring(0,1)+"$"+strTotal.substring(1));
					}
					else $(api.column(columns[j]).footer()).html("$"+strTotal);
					
				}
				
				
 
           
            
			},
			
			columns: [
					{ "data": "SelectedCustomerID" },
					{ "data": "CustName" },
					{ "data": "LCPvs3PAvgSales" },
					{ "data": "DayImpact" },
					{ "data": "ADS_Variance" },
					{ "data": "LCPvs12PAvgSales" },
					{ "data": "PP1Sales" },
					{ "data": "PP2Sales" },
					{ "data": "LCPSales" },
					{ "data": "ThreePPAvgSales" },
					{ "data": "TwelvePPAvgSales" },
					{ "data": "CurrentPSales" },
					{ "data": "SamePLYSales" },
					{ "data": "MCS" },
					{ "data": "LCPvsMCS" },
					{ "data": "3PavgvsMCS" },
					{ "data": "12PavgvsMCS" },
					{ "data": "CurrentvsMCS" },
					{ "data": "LCP_ROI" },
					{ "data": "PavgROI" },
					{ "data": "TotalEquipmentValue" },
					{ "data": "SecondarySalesPerson" },
					{ "data": "CustomerType" },
					{ "data": "CustomerNotes" },
					{ "data": "rules" }
					
				],
	  		columnDefs: [
				{"className": "smaller-detail-line", "targets": "_all"},
				{"orderable": true,"targets": [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22] },
				{"orderable": false,"targets": [23] },
				
				
 				{
				"targets": [ 0 ],
				"createdCell": function (td, cellData, rowData, row, col) {
					  $(td).html("<a href='tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID="+rowData.SelectedCustomerID+"&ZDC=0&VB=3Periods&oon=new' target='_blank'>"+rowData.SelectedCustomerID+ "</a>");
					}				
				},
				{
				"targets": [ 1 ],
				"createdCell": function (td, cellData, rowData, row, col) {
					  $(td).addClass("clip");
					  $(td).html("<a href='tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID="+rowData.SelectedCustomerID+"&ZDC=0&VB=3Periods&oon=new' target='_blank'>"+rowData.CustName+ "</a>");
					  
					}				
				},
				{
				"targets": [2],
				"createdCell": function (td, cellData,rowData, row, col) {
					var number = Number(cellData.replace(/[^0-9.-]+/g,""));
					
					if (number<0) $(td).addClass("negative");
					
					}
				},
				{
				"targets": [8],
				"createdCell": function (td, cellData,rowData, row, col) {
					$(td).addClass("not-as-small-detail-line").addClass("border-left").addClass("border-right");
				 }
				},
				{
				"targets": [12,13,14,15],
				"createdCell": function (td, cellData,rowData, row, col) {
					var number = Number(cellData.replace(/[^0-9.-]+/g,""));
					
					if (number<0) $(td).addClass("negative");
					
					}
				},
				{
				"targets": [16,17],
				"createdCell": function (td, cellData,rowData, row, col) {
					var PavgROI = Number(rowData.PavgROI.replace(/[^0-9.-]+/g,""));
					var LCP_ROI = Number(rowData.LCP_ROI.replace(/[^0-9.-]+/g,""));
					if (PavgROI>=10 && LCP_ROI>=10) $(td).addClass("negative");
					
					}
				},
				{
				"targets": [20],
				"createdCell": function (td, cellData,rowData, row, col) {
				
					$(td).html("<a data-toggle='modal' data-show='true' href='#' data-cust-id='"+rowData.SelectedCustomerID+"' data-lcp-gp='0' data-target='#modalEquipmentVPC' data-tooltip='true' data-title='View Customer Equipment'>" +cellData+ "</a>");
					
					
					}
				},
				{
				"targets": [23],
				"createdCell": function (td, cellData,rowData, row, col) {
				
		               	if (cellData == "False") {
		               	 	$(td).html("<a data-toggle='modal' data-target='#modalEditCustomerNotes' data-category-id='-2' data-cust-id='"+rowData.SelectedCustomerID+"' class='ole' rel='tooltip' style='cursor:pointer;'><i class='fa fa-file-text-o' aria-hidden='true'></i></a>");
		               	}
		               	else {
		               	 	$(td).html("<a data-toggle='modal' data-target='#modalEditCustomerNotes' data-category-id='-2' data-cust-id='"+rowData.SelectedCustomerID+"' class='ole' rel='tooltip' style='cursor:pointer;'><i class='fa fa-file-text-o faa-pulse animated fa-2x' aria-hidden='true'></i></a>");
		               	 }
				}
				},
				{
				"targets": [24],
				"visible": ruleColVisible
				
				}
				
			]		

	    }
		
	    );
		
		$("#tableData").on( 'order.dt', function () {
				// This will show: "Ordering on column 1 (asc)", for example
				var order = datatableWidget.order();
				
				if (order.length==2) setOrderMemo(order[0],order[1]);
				else setOrderMemo(order[0][0],order[0][1]);
				});
				
		}
		else {
			if (typeof datatableWidgetPrimary!="undefined") {
				$("#tableData2").off( 'order.dt');	
				datatableWidgetPrimary.clear();
				datatableWidgetPrimary.destroy();
			}
			datatableWidgetPrimary=$("#tableData2")
		
			.on('preXhr.dt', function ( e, settings, data ) {
					$(".waitdiv").removeClass("d-none");
				})
			.on('xhr.dt', function ( e, settings, json, xhr ) {
				
				$(".waitdiv").addClass("d-none");
				
			
				})
			
			.DataTable({
	    
	        scrollY: 500,
	        scrollCollapse: true,
			ajax: ajaxURL ,
	        paging: false,
			order: [ parseInt(orderColumn), orderResult ],
			footerCallback: function ( row, data, start, end, display ) {
				var api = this.api();
				var columns=[1,2,3,4,5,6,7,8,9];
				// Remove the formatting to get integer data for summation
				var intVal = function ( i ) {
					return typeof i === 'string' ?
						i.replace(/[\$(),]/g, '')*1 :
						typeof i === 'number' ?
							i : 0;
				};
				// Total over all pages
				for (j=0;j<columns.length;j++) {
				
					total = api
						.column(columns[j])
						.data()
						.reduce( function (a, b) {
														
							return intVal(a) + intVal(b);
						}, 0 );
					
					// Update footer
					strTotal=total.toFixed(1).replace(/\d(?=(\d{3})+\.)/g, '$&,').substring(0, total.toFixed(1).replace(/\d(?=(\d{3})+\.)/g, '$&,').length-2);
					
					if (total<0) {
						$(api.column(columns[j]).footer()).addClass("negative");
						$(api.column(columns[j]).footer()).html(strTotal.substring(0,1)+"$"+strTotal.substring(1));
					}
					else $(api.column(columns[j]).footer()).html("$"+strTotal);
					
				}
				
				
 
           
            
			},
			columns: [
					{ "data": "Category" },
					{ "data": "TotalSales" },
					{ "data": "CurrentSales" },
					{ "data": "ThreePriorPeriodsAeverage" },
					{ "data": "TwelvePriorPeriodsAeverage" },
					{ "data": "ThisPeriodLastYearSales" },
					{ "data": "CurrentSalesVariance" },
					{ "data": "ThreePriorPeriodsAeverageVariance" },
					{ "data": "TwelvePriorPeriodsAeverageVariance" },
					{ "data": "ThisPeriodLastYearSalesVariance" }
					
				],
	  		columnDefs: [
				{"className": "smaller-detail-line", "targets": "_all"},
				{"orderable": true,"targets": [0,1,2,3,4,5,6,7,8,9] },
				{
				"targets": [1,2,3,4,5,6,7,8,9],
				"createdCell": function (td, cellData,rowData, row, col) {
					var number = Number(cellData.replace(/[^0-9.-]+/g,""));
					
					if (number<0) $(td).addClass("negative");
					
					}
				}
				
			]		

	    }
		
	    );
		
		$("#tableData2").on( 'order.dt', function () {
				// This will show: "Ordering on column 1 (asc)", for example
				var order = datatableWidgetSecondary.order();
				
				if (order.length==2) setOrderMemo(order[0],order[1]);
				else setOrderMemo(order[0][0],order[0][1]);
				});
				
			
		
		
		}
	}
	
	function setOrderMemo(column, orderType) {
		//column - ID of column 
		//orderType - asc or desc
		var sortedTitle=$(".tab-pane.active table thead tr").eq(1).find("th").eq(column).text();
		var outTextData="Sort: "+sortedTitle+" "+(orderType=="desc"?"highest to lowest":"lowest to highest");
		$(".tab-pane.active table thead tr").eq(0).find("th").eq(0).html(outTextData);
	}
</script>
  
<style>

.nav-tabs>li>a {
color: #555;
    cursor: default;
    background-color: #fff;
    border: 1px solid #ddd;
    border-bottom-color: transparent;
	cursor:pointer;
}

.nav-tabs>li.active>a,  .nav-tabs>li.active>a:focus {

    background-color: #ddd;

}


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

	.negative{
		color:red;	
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


	.footer-total{
		font-size: 0.95em;
		vertical-align: top !important;
	}	

	.smaller-detail-line{
		font-size: 0.8em;
		text-align:center;		
	}	

	.footer-total-negative
	{
		font-size: 1.5em;
		color:red;	
	}


	.modal.modal-wide .modal-dialog {
	  width: 75%;
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
	
	.modalResponsiveTable {
		margin-left: 25px;
		margin-right: 25px;
	}
	
	
	.ajaxRowView .visibleRowEdit, .ajaxRowEdit .visibleRowView { display: none; }
	
	
	.ajax-loading {
	    position: relative;
	}
	.ajax-loading::after {
	    background-image: url("/img/loading.gif");
	    background-position: center top;
	    background-repeat: no-repeat;
	    content: "";
	    display: block;
	    height: 100%;
	    min-height: 100px;
	    position: absolute;
	    top: 0;
	    width: 100%;
	}
	
	.ole {
	    font-family: consolas, courier, monaco, menlo, monospace;
	    background: rgb(255,193,44);
	    padding: 6px 10px;
	    display: inline-block;
	    font-size: 1.2em;
	    border-radius: 4px;
	    border: 0;
	    cursor: pointer;
	    color: #000;
	}	
	
	
	.ole:hover {
		background: dodgerblue;
		color: #fff;
		text-shadow: 1px 1px 1px #000;
		box-shadow: 0 0 0 #555;
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

	.not-as-small-detail-line{
		font-size: 1em;
	}
	tfoot td.smaller-detail-line {
	
		font-weight:bold;
		font-size: 14px;
		text-align:center;
	}
	.border-left {border-left:2px solid #000000;}
	.border-right {border-right:2px solid #000000;}
	.top-padding-30 {
	    padding-top: 30px;
	}
	.clip {
    white-space: nowrap; 
    overflow: hidden; 
    text-overflow: ellipsis; 
	text-align:left;
   }
   
	
	.red{
		font-weight:bold;
		color:red;	
	}

	.blue{
		font-weight:bold;
		color:blue;	
	}
   
</style>
<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.18/sorting/currency.js"></script>

<script type="text/javascript">

	$.fn.dataTable.ext.type.order['currency-pre'] = function ( data ) {
	   
	    var expression = /((\(\$))|(\$\()/g;
	    //Check if its in the proper format
	    if(data.match(expression)){
	        //It matched - strip out parentheses & any characters we dont want and append - at front     
	        data = '-' + data.replace(/[\$\(\),]/g,'');
	    }else{
	      data = data.replace(/[\$\,]/g,'');
	    }
	    return parseInt( data, 10 );
	};
	
	$(document).ready(function() {
		
	
		$('#modalEquipmentVPC').on('show.bs.modal', function(j) {
	
		    //get data-id attribute of the clicked order
		    var CustID = $(j.relatedTarget).data('cust-id');
		    var LCPGP = $(j.relatedTarget).data('lcp-gp');
	 
		    //populate the textbox with the id of the clicked order
		    $(j.currentTarget).find('input[name="txtCustIDToPass"]').val(CustID);
		    $(j.currentTarget).find('input[name="txtLastClosedPeriodGP"]').val(LCPGP);
		    	    
		    var $modal = $(this);
		    //$modal.find('#PleaseWaitPanelModal').show();  
	
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				cache: false,
				data: "action=GetTitleForEquipmentVPCModal&CustID="+encodeURIComponent(CustID)+"&LCPGP="+encodeURIComponent(LCPGP),
				success: function(response)
				 {
	               	 $modal.find('#modalEquipmentVPCTitle').html(response);            	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalEquipmentVPCTitle').html("Failed");
	             }
			});
			
	
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				cache: false,
				data: "action=GetContentForEquipmentVPCModal&CustID="+encodeURIComponent(CustID),
			  	beforeSend: function() {
			     	$('#PleaseWaitPanelModal').show();
			     	$modal.find('#modalCategoryVPCContent').html('');
			  	},
			  	complete: function(){
			     	$('#PleaseWaitPanelModal').hide();
			  	},
				success: function(response)
				 {
	               	 $("#PleaseWaitPanelModal").hide();
	               	 $modal.find('#modalCategoryVPCContent').html(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalCategoryVPCContent').html("Failed");
	             }
			});
		});
	


	
	
		$('#modalEditCustomerNotes').on('show.bs.modal', function(e) {
	
		    //get data-id attribute of the clicked order
		    var CustID = $(e.relatedTarget).data('cust-id');
		    var CategoryID = $(e.relatedTarget).data('category-id');
		    
		    //populate the textbox with the id of the clicked order
		    $(e.currentTarget).find('input[name="txtCustIDToPassToGenerateNotes"]').val(CustID);
		    $(e.currentTarget).find('input[name="txtCustIDToPass"]').val(CustID);
		    $(e.currentTarget).find('input[name="txtCategoryID"]').val(CategoryID);
		    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=GetContentForCustomerNotesModal&CustID="+encodeURIComponent(CustID),
				success: function(response)
				 {
	               	 $modal.find('#modalEditCustomerNotesContent').html(response);
	               	 //alert(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalEditCustomerNotesContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
		});
	

	
	    $("#PleaseWaitPanel").hide();
	    
	});
	
</script>


<%
Response.Write("<div id=""PleaseWaitPanel"" class=""container"">")
Response.Write("<br><br>Creating Leakage Overview Primary Salesman<br><br>Please wait...<br><br>")
Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
Response.Write("</div>")
Response.Flush()

%>

<div class="container-fluid" >
	<div class="row">
		<div class="col-lg-8 col-md-8 col-sm-6 col-xs-12">
			<h3 class="page-header">Salesman <%=FilterSlsmn1 %> - <%=GetSalesmanNameBySlsmnSequence(FilterSlsmn1)%> for Period <%=PeriodBeingEvaluated %>
				&nbsp;&nbsp;(<%=FormatDateTime(GetPeriodBeginDateBySeq(PeriodSeqBeingEvaluated),2)%>&nbsp;-&nbsp;<%=FormatDateTime(GetPeriodEndDateBySeq(PeriodSeqBeingEvaluated),2)%>)
				&nbsp;&nbsp;
			</h3>
		
		</div>
		<div class="col-lg-4 col-md-4 col-sm-6 col-xs-12 top-padding-30">
		<!-- accordion line starts here !-->
<div id="accordion" role="tablist" aria-multiselectable="true">

	<div class="panel panel-default">
		<div class="panel-heading" role="tab" id="headingOne">
			<h5 class="panel-title">
				<a role="button" data-toggle="collapse" data-parent="#accordion" href="#collapseOne" aria-expanded="false" aria-controls="collapseOne">
					Click to view rules 
				</a>
			</h5>
		</div>
		
		<div id="collapseOne" class="panel-collapse collapse" role="tabpanel" aria-labelledby="headingOne">
			<div class="panel-body">

				<div class="table-info">
					<div class="table-responsive">
						<table class="table custom-table">

							<tbody>
								<tr><td>0.&nbsp;&nbsp;Rule Out LCPvs3Pavg Variance < $100 OR<br>
								LCPvs3Pavg Variance < 10% - Rule Out</td></tr>
								<tr><td>1.&nbsp;&nbsp;Rule Out Current (adjusted for days) >= 3Pavg or 12Pavg</td></tr>
								<tr><td>2.&nbsp;&nbsp;Rule Out LCP >= 12pAVG</td></tr>
								<tr><td>3.&nbsp;&nbsp;Rule Out LCP >= SPLY</td></tr>
								<tr><td>4.&nbsp;&nbsp;Rule Out (LCP + PP1 + PP2) / 3 >= 3Pavg </td></tr>
								<tr><td>5.&nbsp;&nbsp;Rule Out (LCP + PP1 + PP2) / 3 >= 12Pavg </td></tr>
								<tr><td>6.&nbsp;&nbsp;Ignore rules 0 thru 5 if 3PROI > 10</td></tr>
							</tbody>	
						</table>
					</div>
				</div>

			</div>
		</div>
	</div>
</div>
<!-- accordion line ends here !-->		
		
		</div>
		
	
	</div>


</div>







</h3>
<% dummy=MUV_WRITE("LOHVAR","Primary")%>

<!--#include file="dashboard_segments_header.asp"-->

<!-- row !-->
<div class="row">

<!-- Nav tabs -->
<ul class="nav nav-tabs filter-tabs" role="tablist">

  <li data-source="segmentTabs/tabup_data.asp?p=<%=FilterSlsmn1%>" data-order="desc" data-col-order="2"><a href="#tab1" id="#tabup" role="tab" data-toggle="tab" >Up&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(1,PeriodSeqBeingEvaluated,FilterSlsmn1)%></span></a></li>
  <li class="active" data-source="segmentTabs/tabdown_data.asp?p=<%=FilterSlsmn1%>" data-order="asc" data-col-order="2"><a href="#tab1" id="tabdown" role="tab" data-toggle="tab">Down&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(2,PeriodSeqBeingEvaluated,FilterSlsmn1)%></span></a></li>
  <li data-source="segmentTabs/tabzero_data.asp?p=<%=FilterSlsmn1%>" data-order="asc" data-col-order="2"><a href="#tab1" id="tabzero" role="tab" data-toggle="tab"><%= PVarShort %> $0 Sales&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(3,PeriodSeqBeingEvaluated,FilterSlsmn1)%></span></a></li>
  <li data-source="segmentTabs/tabruledout_data.asp?p=<%=FilterSlsmn1%>"  data-order="asc" data-col-order="2"><a href="#tab1" id="tabruledout" role="tab" data-toggle="tab">Ruled Out&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(4,PeriodSeqBeingEvaluated,FilterSlsmn1)%></span></a></li>
  <li data-source="segmentTabs/taball_data.asp?p=<%=FilterSlsmn1%>" data-order="asc" data-col-order="2"><a href="#tab1" id="taball" role="tab" data-toggle="tab">All&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(5,PeriodSeqBeingEvaluated,FilterSlsmn1)%></span></a></li>
  <li data-source="segmentTabs/tabtop50_data.asp?p=<%=FilterSlsmn1%>" data-order="desc" data-col-order="10"><a href="#tab1" id="tabtop50" role="tab" data-toggle="tab">Top 50&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(6,PeriodSeqBeingEvaluated,FilterSlsmn1)%></span></a></li>
  <li data-source="segmentTabs/tabbottom50_data.asp?p=<%=FilterSlsmn1%>" data-order="asc" data-col-order="10"><a href="#tab1" id="tabbottom50" role="tab" data-toggle="tab">Bottom 50&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(7,PeriodSeqBeingEvaluated,FilterSlsmn1)%></span></a></li>
  <li data-source="segmentTabs/tabMCS_data.asp?p=<%=FilterSlsmn1%>" data-order="asc" data-col-order="2"><a href="#tab1" id="tabmcs" role="tab" data-toggle="tab">MCS&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(8,PeriodSeqBeingEvaluated,FilterSlsmn1)%></span></a></li>
  <li data-source="segmentTabs/tabhighROI_data.asp?p=<%=FilterSlsmn1%>" data-order="desc" data-col-order="1"><a href="#tab1" id="tabhighroi" role="tab" data-toggle="tab">High ROI&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(5,PeriodSeqBeingEvaluated,FilterSlsmn1)%></span></a></li>
  <li data-source="segmentTabs/tabcategorybreakdown_data.asp?p=<%=FilterSlsmn1%>" data-order="desc" data-col-order="1"><a href="#tab2" id="tabcategorybreakdown" role="tab" data-toggle="tab">Category Breakdown&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(10,PeriodSeqBeingEvaluated,FilterSlsmn1)%></span></a></li>  
  
</ul>
<div class="tab-content">
    <div role="tabpanel" class="tab-pane active" id="tab1">
		<table id="tableData" class="display compact" style="width:100%;">
					<thead>
						<tr>	
							<th colspan="2" class="sorttable numeric smaller-header">&nbsp;</th>
							<th class="td-align1 vpc-variance-header" colspan="4" style="border-right: 2px solid #555 !important;">Variances</th>
							<th class="td-align1 vpc-3pavg-header" colspan="7" style="border-right: 2px solid #555 !important;">Sales</th>
							<th class="td-align1 vpc-lcp-header" colspan="5" style="border-right: 2px solid #555 !important;">MCS</th>
							<th class="td-align1 vpc-current-header" colspan="3" style="border-right: 2px solid #555 !important;">EQUIP ROI</th>
							<th class="td-align1 gen-info-header" colspan="4" style="border-right: 2px solid #555 !important;">General</th>
							
						</tr>

							<% '
							'Setup PP1 & PP2 descriptions
							
							PP1Var = "Period&nbsp;" & GetPeriodBySeq(PeriodSeqBeingEvaluated -1) & "<br>" & GetPeriodYearBySeq(PeriodSeqBeingEvaluated-1) & "&nbsp;$"
							PP1VarShort = "Period&nbsp;" & GetPeriodBySeq(PeriodSeqBeingEvaluated -1)
							PP2Var = "Period&nbsp;" & GetPeriodBySeq(PeriodSeqBeingEvaluated -2) & "<br>" & GetPeriodYearBySeq(PeriodSeqBeingEvaluated-2) & "&nbsp;$"
							PP2VarShort = "Period&nbsp;" & GetPeriodBySeq(PeriodSeqBeingEvaluated -2)
							PVarShort = "Period&nbsp;" & GetPeriodBySeq(PeriodSeqBeingEvaluated)
							%>
						
						<tr>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Acct</th> 
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Client</th> 
							<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><%= PVarShort %> vs<br>3P avg $</th> 
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Day<br>Impact</th>  
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>ADS</th> 
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%= PVarShort %> vs<br>12P avg $</th>
							
							
							<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br><%= PP1VarShort %></th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br><%= PP2VarShort %></th>							
							
							<th class="td-align sorttable_numeric smaller-header not-as-small-detail-line" style="border-left: 2px solid #555 !important; border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><%= PVarShort %> $</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>3P avg $</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>12P avg $</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Current $</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>SPLY $</th> 
							<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>MCS</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%= PVarShort %> <br>vs MCS</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">3P avg vs<br> MCS</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">12P avg vs<br> MCS</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Current vs<br> MCS</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><%= PVarShort %><br>ROI</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">3P avg<br>ROI</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">Equipment<br>Value</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Secondary<br> Slsmn</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Customer<br>Type</th>
							<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Customer<br>Notes</th>
							<th class="td-align smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;">Rules</th>
						</tr>
					</thead>
					
					<tfoot>
						<tr>
							<td>&nbsp;</td>
							<td>Total</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td class="border-left border-right">&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
						</tr>
					</tfoot>
					
					
				</table>

	</div>
    <div role="tabpanel" class="tab-pane" id="tab2">
		<table id="tableData2" class="display compact" style="width:100%;">
			<thead>
				<tr>	
					<th rowspan="1"  class="sorttable numeric smaller-header">&nbsp;</th>
					<th class="td-align1 vpc-3pavg-header" colspan="5" style="border-right: 2px solid #555 !important;">Sales</th>
					<th class="td-align1 vpc-variance-header" colspan="4" style="border-right: 2px solid #555 !important;">Variances</th>
				</tr>
				<tr>	
					<th class="sorttable numeric smaller-header" style="border-top: 2px solid #555 !important;">Category</th>
					<th class="sorttable numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;"><%= PVarShort %><br>Sales $</th>
					<th class="sorttable numeric smaller-header" style="border-top: 2px solid #555 !important;">Current Period<br>(So Far) Sales $</th>
					<th class="sorttable numeric smaller-header" style="border-top: 2px solid #555 !important;">3P Avg<br>Sales $</th>
					<th class="sorttable numeric smaller-header" style="border-top: 2px solid #555 !important;">12P Avg<br>Sales $</th>
					<th class="sorttable numeric smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;"><%= PVarShort %><br>Last Year</th>
					<th class="sorttable numeric smaller-header" style="border-top: 2px solid #555 !important;">Current Period<br>(So Far) Variance</th>
					<th class="sorttable numeric smaller-header" style="border-top: 2px solid #555 !important;">3P Avg<br>Variance</th>
					<th class="sorttable numeric smaller-header" style="border-top: 2px solid #555 !important;">12P Avg<br>Variance</th>
					<th class="sorttable numeric smaller-header" style="border-top: 2px solid #555 !important;"><%= PVarShort %><br>LY Variance</th>
				</tr>
			</thead>
			<tfoot>
				<tr>
					<td>Total</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					
				</tr>
			</tfoot>
		</table>
	
	</div>
  
</div>
  




<%		

	rs.Close	
		
%>


</div>
<!-- eof row !-->



<div class="waitdiv d-none" style="position:fixed;z-index: 999999999; top: 0px; left: 0px; width: 100%; height:80%; background-color:transparent; text-align: center; padding-top: 20%; filter: alpha(opacity=0); opacity:0; "></div>
    <div id="waitdiv" class="waitdiv d-none small" style="padding-bottom: 90px;text-align: center; vertical-align:middle;padding-top:50px;background-color:#ebebeb;width:300px;height:100px;margin: 0 auto; top:40%; left:40%;position:absolute;-webkit-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); -moz-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); z-index:999999999;">
    <img src="<%= BaseURL %>img/loader.gif" alt="" /><br />Request data from server. <br /> Please wait...
</div>

<!-- pencil Modal -->
<div class="modal modal-wide fade" id="modalEditCustomerNotes" tabindex="-1" role="dialog" aria-labelledby="modalEditCustomerNotesLabel">

	<style>
	.modal-header {
	    padding: 15px;
	    border-bottom: 1px solid #e5e5e5;
	    min-height: 35px !important;
	}
	</style>
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	

			<input type="hidden" name="txtCategoryID" id="txtCategoryID">
			<input type="hidden" name="txtCustIDToPassToGenerateNotes" id="txtCustIDToPassToGenerateNotes">
			    
			<div id="modalEditCustomerNotesContent">
				<!-- Content for the modal will be generated and written here -->
				<!-- Content generated by Sub GetContentForCustomerNotesModal() in InsightFuncs_AjaxForBizIntelModals.asp -->
			</div>

				  
			<div class="modal-footer">
				<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
			</div>
	

		</div>
		<!-- eof modal content !-->
</div>
<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->
<%

	FUNCTION getTotalRecordsForTab(tabID,PeriodSeqBeingEvaluated,FilterSlsmn1)
		DIM retData
		retData="0"
		SELECT CASE tabID
			CASE 1
				varSQL="SELECT SUM(d1.LCPSales-d1.ThreePPSales/3) AS totalAmpunt, COUNT(*) AS totalQty FROM (SELECT Distinct CustCatPeriodSales_ReportData.CustNum AS CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats AS LCPSales,CustCatPeriodSales_ReportData.Total3PPSalesAllCats AS ThreePPSales"
				varSQL=varSQL & " FROM CustCatPeriodSales_ReportData, AR_Customer"
				varSQL=varSQL & " WHERE CustCatPeriodSales_ReportData.ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated
				varSQL=varSQL & " AND  CustCatPeriodSales_ReportData.LCPTotSalesAllCats > (CustCatPeriodSales_ReportData.Total3PPSalesAllCats /3)"
				varSQL=varSQL & " AND  CustCatPeriodSales_ReportData.PrimarySalesman =" & FilterSlsmn1
				varSQL=varSQL & " AND CONVERT(varchar,CustCatPeriodSales_ReportData.CustNum)=AR_Customer.CustNum GROUP BY CustCatPeriodSales_ReportData.CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats,CustCatPeriodSales_ReportData.Total3PPSalesAllCats) d1"
			CASE 2
				varSQL="SELECT SUM(d1.LCPSales-d1.ThreePPSales/3) AS totalAmpunt, COUNT(*) AS totalQty FROM (SELECT Distinct CustCatPeriodSales_ReportData.CustNum AS CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats AS LCPSales,CustCatPeriodSales_ReportData.Total3PPSalesAllCats AS ThreePPSales"
				varSQL=varSQL & " FROM CustCatPeriodSales_ReportData, AR_Customer"
				varSQL=varSQL & " WHERE CustCatPeriodSales_ReportData.ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated
				varSQL=varSQL & " AND  CustCatPeriodSales_ReportData.LCPTotSalesAllCats <= (CustCatPeriodSales_ReportData.Total3PPSalesAllCats /3)"
				varSQL=varSQL & " AND  CustCatPeriodSales_ReportData.PrimarySalesman =" & FilterSlsmn1
				varSQL=varSQL & " AND CONVERT(varchar,CustCatPeriodSales_ReportData.CustNum)=AR_Customer.CustNum GROUP BY CustCatPeriodSales_ReportData.CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats,CustCatPeriodSales_ReportData.Total3PPSalesAllCats) d1"
				retData="0"
				Exit function
			CASE 3
				varSQL="SELECT SUM(d1.LCPSales-d1.ThreePPSales/3) AS totalAmpunt, COUNT(*) AS totalQty FROM (SELECT Distinct CustCatPeriodSales_ReportData.CustNum AS CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats AS LCPSales,CustCatPeriodSales_ReportData.Total3PPSalesAllCats AS ThreePPSales"
				varSQL=varSQL & " FROM CustCatPeriodSales_ReportData, AR_Customer"
				varSQL=varSQL & " WHERE CustCatPeriodSales_ReportData.ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated
				varSQL=varSQL & " AND  LCPTotSalesAllCats <= 0 "
				varSQL=varSQL & " AND  CustCatPeriodSales_ReportData.PrimarySalesman =" & FilterSlsmn1
				varSQL=varSQL & " AND CONVERT(varchar,CustCatPeriodSales_ReportData.CustNum)=AR_Customer.CustNum GROUP BY CustCatPeriodSales_ReportData.CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats,CustCatPeriodSales_ReportData.Total3PPSalesAllCats) d1"
			CASE 4
				varSQL="SELECT SUM(d1.LCPSales-d1.ThreePPSales/3) AS totalAmpunt, COUNT(*) AS totalQty FROM (SELECT Distinct CustCatPeriodSales_ReportData.CustNum AS CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats AS LCPSales,CustCatPeriodSales_ReportData.Total3PPSalesAllCats AS ThreePPSales"
				varSQL=varSQL & " FROM CustCatPeriodSales_ReportData, AR_Customer"
				varSQL=varSQL & " WHERE CustCatPeriodSales_ReportData.ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated
				varSQL=varSQL & " AND  LCPTotSalesAllCats <= (Total3PPSalesAllCats /3)"
				varSQL=varSQL & " AND  CustCatPeriodSales_ReportData.PrimarySalesman =" & FilterSlsmn1
				varSQL=varSQL & " AND CONVERT(varchar,CustCatPeriodSales_ReportData.CustNum)=AR_Customer.CustNum GROUP BY CustCatPeriodSales_ReportData.CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats,CustCatPeriodSales_ReportData.Total3PPSalesAllCats) d1"
			CASE 5
				varSQL="SELECT SUM(d1.LCPSales-d1.ThreePPSales/3) AS totalAmpunt, COUNT(*) AS totalQty FROM (SELECT Distinct CustCatPeriodSales_ReportData.CustNum AS CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats AS LCPSales,CustCatPeriodSales_ReportData.Total3PPSalesAllCats AS ThreePPSales"
				varSQL=varSQL & " FROM CustCatPeriodSales_ReportData, AR_Customer"
				varSQL=varSQL & " WHERE CustCatPeriodSales_ReportData.ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated
				varSQL=varSQL & " AND  CustCatPeriodSales_ReportData.PrimarySalesman =" & FilterSlsmn1
				varSQL=varSQL & " AND CONVERT(varchar,CustCatPeriodSales_ReportData.CustNum)=AR_Customer.CustNum GROUP BY CustCatPeriodSales_ReportData.CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats,CustCatPeriodSales_ReportData.Total3PPSalesAllCats) d1"
			CASE 6
				varSQL= "SELECT SUM(d1.LCPSales-d1.ThreePPSales/3) AS totalAmpunt, COUNT(*) AS totalQty FROM ("
				varSQL = varSQL & "SELECT Distinct  TOP 50 "
				varSQL = varSQL & "CustCatPeriodSales_ReportData.CustNum AS CustNum, CustCatPeriodSales_ReportData.LCPTotSalesAllCats AS LCPSales,CustCatPeriodSales_ReportData.Total3PPSalesAllCats AS ThreePPSales, CustCatPeriodSales_ReportData.Total12PPSalesAllCats AS Total12PPSalesAllCats"
				varSQL = varSQL & " FROM CustCatPeriodSales_ReportData,AR_Customer "
				varSQL = varSQL & " WHERE CustCatPeriodSales_ReportData.ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				varSQL = varSQL & " AND  CustCatPeriodSales_ReportData.PrimarySalesman = " & FilterSlsmn1
				varSQL = varSQL & " AND CONVERT(varchar,CustCatPeriodSales_ReportData.CustNum)=AR_Customer.CustNum"
				varSQL = varSQL & " GROUP BY CustCatPeriodSales_ReportData.CustNum, CustCatPeriodSales_ReportData.LCPTotSalesAllCats, CustCatPeriodSales_ReportData.Total3PPSalesAllCats, CustCatPeriodSales_ReportData.Total12PPSalesAllCats ORDER BY 4 DESC) d1"
				
			CASE 7
				varSQL= "SELECT SUM(d1.LCPSales-d1.ThreePPSales/3) AS totalAmpunt, COUNT(*) AS totalQty FROM ("
				varSQL = varSQL & "SELECT Distinct  TOP 50 "
				varSQL = varSQL & "CustCatPeriodSales_ReportData.CustNum AS CustNum, CustCatPeriodSales_ReportData.LCPTotSalesAllCats AS LCPSales,CustCatPeriodSales_ReportData.Total3PPSalesAllCats AS ThreePPSales, CustCatPeriodSales_ReportData.Total12PPSalesAllCats AS Total12PPSalesAllCats"
				varSQL = varSQL & " FROM CustCatPeriodSales_ReportData,AR_Customer "
				varSQL = varSQL & " WHERE CustCatPeriodSales_ReportData.ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				varSQL = varSQL & " AND  CustCatPeriodSales_ReportData.PrimarySalesman = " & FilterSlsmn1
				varSQL = varSQL & " AND CONVERT(varchar,CustCatPeriodSales_ReportData.CustNum)=AR_Customer.CustNum"
				varSQL = varSQL & " GROUP BY CustCatPeriodSales_ReportData.CustNum, CustCatPeriodSales_ReportData.LCPTotSalesAllCats, CustCatPeriodSales_ReportData.Total3PPSalesAllCats, CustCatPeriodSales_ReportData.Total12PPSalesAllCats ORDER BY 4) d1"
			CASE 8
				varSQL="SELECT SUM(d1.LCPSales-d1.ThreePPSales/3) AS totalAmpunt, COUNT(*) AS totalQty FROM (SELECT Distinct CustCatPeriodSales_ReportData.CustNum AS CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats AS LCPSales,CustCatPeriodSales_ReportData.Total3PPSalesAllCats AS ThreePPSales"
				varSQL = varSQL & " FROM CustCatPeriodSales_ReportData,AR_Customer "
				
				varSQL = varSQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
				varSQL=varSQL & " AND CONVERT(varchar,CustCatPeriodSales_ReportData.CustNum)=AR_Customer.CustNum"
				varSQL = varSQL & " AND  CustCatPeriodSales_ReportData.PrimarySalesman = " & FilterSlsmn1 
				varSQL = varSQL & " AND  AR_Customer.MonthlyContractedSalesDollars IS NOT NULL GROUP BY CustCatPeriodSales_ReportData.CustNum,CustCatPeriodSales_ReportData.LCPTotSalesAllCats,CustCatPeriodSales_ReportData.Total3PPSalesAllCats) d1"
			CASE 10
				varSQL ="SELECT 0 AS totalAmpunt,COUNT(Distinct CategoryNameGetTerm) AS totalQty"
				varSQL = varSQL & " FROM CustCatPeriodSales_ReportData "
				varSQL = varSQL & " WHERE PrimarySalesman = " & FilterSlsmn1 & " AND ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated & " AND Category <> 0 "
			
		END SELECT
		Set cnnQty = Server.CreateObject("ADODB.Connection")
		cnnQty.ConnectionTimeout = 120
		cnnQty.open (Session("ClientCnnString"))
	
		Set rsQty = Server.CreateObject("ADODB.Recordset")
		rsQty.CursorLocation = 3
		Set rsQty = cnnQty.Execute(varSQL)
		IF NOT rsQty.EOF Then
			retData=rsQty("totalQty") & ": " 
			If IsNumeric(rsQty("totalAmpunt")) Then
				retData = retData & FormatCurrency(rsQty("totalAmpunt"),0,-2,0)
			Else
				retData = retData & FormatCurrency(0,0,-2,0)
			End If
			retData=rsQty("totalQty")
		Else
			retData = "0: $0"
		END IF
		getTotalRecordsForTab=retData
		rsQty.Close
		cnnQty.Close
		SET rsQty=Nothing
		SET cnnQty=Nothing
	END Function





%>

<!--#include file="../tools/CatAnalByPeriod/CatAnalByPeriod_Modals.asp"-->
<!--#include file="../../inc/footer-main.asp"-->