<%
Server.ScriptTimeout = 900000 'Default value

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))


FilterSlsmn2 = Request.QueryString("p")



%>
<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->
<!--#include file="../../css/fa_animation_styles.css"-->
 

<%
CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Leakage Overview Secondary Salesman"

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
	
	console.log("<%=FilterSlsmn2%>");
	console.log("<%=PeriodSeqBeingEvaluated%>");

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
	var datatableWidgetSecondary;
	
	var ruleColVisible;
	
	
	
	$(window).on("load",function() {
	

		var activeTab=$(".filter-tabs li.active");
		console.log(activeTab);
		
	});
	
	

	$(document).ready(function() {
	    
		$('#tableDataAll').DataTable({
	        scrollY: 500,
	        scrollCollapse: true,
	        paging: false,
	        order: [[ 13, 'asc' ],[ 14, 'asc' ]],
			columnDefs: [
			        { targets: [18, 19], "orderable": false,},
			        { type: 'currency', targets: 13}
			    ]	        
		    }
		);
	
	});	
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
Response.Write("<br><br>Creating Leakage Overview Secondary Salesman<br><br>Please wait...<br><br>")
Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
Response.Write("</div>")
Response.Flush()

%>

<div class="container-fluid" >
	<div class="row">
		<div class="col-lg-8 col-md-8 col-sm-6 col-xs-12">
			<h3 class="page-header">Salesman <%=FilterSlsmn2 %> - <%=GetSalesmanNameBySlsmnSequence(FilterSlsmn2)%> for Period <%=PeriodBeingEvaluated %>
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
<% dummy=MUV_WRITE("LOHVAR","Secondary")%>

<!--#include file="dashboard_segments_header.asp"-->

<!-- row !-->
<div class="row">

<!-- Nav tabs -->
<ul class="nav nav-tabs filter-tabs" role="tablist">

  <!--<li class="active" data-source="segmentTabs/tabdown_datanew.asp?p=<%=FilterSlsmn2%>" data-order="asc" data-col-order="2"><a href="#tab1" id="tabdown" role="tab" data-toggle="tab">Down&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(2,PeriodSeqBeingEvaluated,FilterSlsmn2)%></span></a></li>-->
  <li class="active"><a href="#tabdown" id="tabdown" role="tab" data-toggle="tab" data-order="asc" data-col-order="2">Down&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(2,PeriodSeqBeingEvaluated,FilterSlsmn2)%></span></a></li>
  <li><a href="#taball" id="taball" role="tab" data-toggle="tab" data-order="asc" data-col-order="2">All&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(5,PeriodSeqBeingEvaluated,FilterSlsmn2)%></span></a></li>
  <!--<li data-source="segmentTabs/taball_datanew.asp?p=<%=FilterSlsmn2%>" data-order="asc" data-col-order="2"><a href="#tab1" id="taball" role="tab" data-toggle="tab">All&nbsp;&nbsp;<span class="badge"><%=getTotalRecordsForTab(5,PeriodSeqBeingEvaluated,FilterSlsmn2)%></span></a></li>-->
 
</ul>

<div class="tab-content">

    <div role="tabpanel" class="tab-pane active" id="tabdown">
		<table id="tableDataDown" class="display compact" style="width:100%;">
					<thead>
						<tr>	
							<th colspan="2" class="sorttable numeric smaller-header"></th>
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
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Primary<br> Slsmn</th>
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
		
	
	
    <div role="tabpanel" class="tab-pane" id="taball">
		<table id="tableDataAll" class="display compact" style="width:100%;">
					<thead>
						<tr>	
							<th colspan="2" class="sorttable numeric smaller-header"></th>
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
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Primary<br> Slsmn</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Customer<br>Type</th>
							<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Customer<br>Notes</th>
							<th class="td-align smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;">Rules</th>
						</tr>
					</thead>
					<tbody>
					<%
					
						Segment = FilterSlsmn2
					
						ShowPercentageColumns = False
					
						JSON=""
					
						Select Case MUV_READ("LOHVAR")
						
							Case "Secondary"
						
								SQL = "SELECT * FROM BI_DashboardSegmentTabs WHERE Tab = 'ALL' AND SecondarySalesmanNumber = " & Segment 
								
							Case "Primary"
						
								SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
								SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
								SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
								SQL = SQL & " FROM CustCatPeriodSales_ReportData "
								SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
								SQL = SQL & " AND PrimarySalesman = " & Segment 
								
							Case "CustType"
						
								SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
								SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
								SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
								SQL = SQL & " FROM CustCatPeriodSales_ReportData "
								SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
								SQL = SQL & " AND CustType = " & CustType 
								
						End Select	
						
						Set cnn8 = Server.CreateObject("ADODB.Connection")
						cnn8.ConnectionTimeout = 120
						cnn8.open (Session("ClientCnnString"))
						
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.CursorLocation = 3
						Set rs = cnn8.Execute(SQL)
						
							GrandTotLCPvs3PAvgSales = 0
									
							Do While Not rs.EOF
								
								LCPvs3PAvgPercent = FormatNumber(0,0)
								LCPvs12PAvgPercent = FormatNumber(0,0)
		
								If Not IsNull(rs("MCS")) Then
									MCS = FormatCurrency(rs("MCS"),0)
								Else
									MCS = ""
								End If
								
								If Not IsNull(rs("LCPvMCS")) Then
									LCPvMCS = FormatCurrency(rs("LCPvMCS"),0,-2,0)
								Else
									LCPvMCS = ""
								End If
								
								If Not IsNull(rs("ThreePAvgvMCS")) Then
									ThreePavgvsMCS = FormatCurrency(rs("ThreePAvgvMCS"),0,-2,0)
								Else
									ThreePavgvsMCS = ""
								End If
								
								If Not IsNull(rs("TwelvePAvgvMCS")) Then
									TwelvePAvgvMCS = FormatCurrency(rs("TwelvePAvgvMCS"),0,-2,0)
								Else
									TwelvePAvgvMCS = ""
								End If
								
								If Not IsNull(rs("CPvMCS")) Then
									CurrentvsMCS = FormatCurrency(rs("CPvMCS"),0,-2,0)
								Else
									CurrentvsMCS = ""
								End If
								
								If rs("EqpValue")> 0 Then	
									If IsNumeric(rs("LCPROI")) Then
										LCP_ROI = FormatNumber(rs("LCPROI"),1)
									Else
										LCP_ROI = "No Sales"
									End If
									If IsNumeric(rs("ThreePAvgROI")) Then
										PavgROI = FormatNumber(rs("ThreePAvgROI"),1)
									Else
										PavgROI = ""
									End If
									' Write equipment value regardless of ROI
									TotalEquipmentValue = FormatCurrency(rs("EqpValue"),0)
								Else
									LCP_ROI = ""
									PavgROI = ""
									TotalEquipmentValue ""
								End If
								
								Select Case MUV_READ("LOHVAR")
									Case "Secondary"
									    If Instr(rs("PrimarySalesmanName") ," ") <> 0 Then
											PrimarySalesPerson = Left(rs("PrimarySalesmanName"),Instr(rs("PrimarySalesmanName")," ")+1)
										Else
											PrimarySalesPerson rs("PrimarySalesmanName")
										End If
									Case "Primary"
									    If Instr(rs("SecondarySalesmanName")," ") <> 0 Then
											SecondarySalesPerson = Left(rs("SecondarySalesmanName"),Instr(rs("SecondarySalesmanName")," ")+1)
										Else
											SecondarySalesPerson = ""
										End If
									Case "CustType"
									    If Instr(rs("SecondarySalesmanName")," ") <> 0 Then
											SecondarySalesPerson = Left(rs("SecondarySalesmanName"),Instr(rs("SecondarySalesmanName")," ")+1)
										Else
											SecondarySalesPerson = rs("SecondarySalesmanName")
										End If
								End Select	
								
								CustomerType = rs("CustomerTypeName")
								CustomerNotes = UserHasAnyUnviewedNotes(rs("CustID"))
								rules = "123abc"							
							
							%>
							<tr role="row">
							
								<td class="smaller-detail-line"><a href="tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=<%= rs("CustID") %>&amp;ZDC=0&amp;VB=3Periods&amp;oon=new" target="_blank"><%= rs("CustID") %></a></td> 
								<td class="smaller-detail-line"><a href="tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=<%= rs("CustID") %>&amp;ZDC=0&amp;VB=3Periods&amp;oon=new" target="_blank"><%= rs("CustName") %></a></td> 
								<td class="smaller-detail-line negative sorting_1"><%= FormatCurrency(rs("LCPv3PAvg"),0,-2,0) %></td> 
								<td class="smaller-detail-line"><%= FormatCurrency(rs("DayImpact"),0) %></td>  
								<td class="smaller-detail-line"><%= FormatCurrency(rs("ADS"),0) %></td> 
								<td class="smaller-detail-line"><%= FormatCurrency(rs("LCPv12PAvg"),0) %></td>
								
								
								<td class="smaller-detail-line"><%= FormatCurrency(rs("PP1Sales"),0,-2,0) %></td>
								<td class="smaller-detail-line"><%= FormatCurrency(rs("PP2Sales"),0,-2,0) %></td>							
								
								<td class="smaller-detail-line not-as-small-detail-line border-left border-right"><%= FormatCurrency(rs("LCPSales"),0,-2,0) %> $</td>
								<td class="smaller-detail-line"><%= FormatCurrency(rs("ThreePAvgSales"),0,-2,0) %></td>
								<td class="smaller-detail-line"><%= FormatCurrency(rs("TwelvePAvgSales"),0,-2,0) %></td>
								<td class="smaller-detail-line"><%= FormatCurrency(rs("CPSales"),0,-2,0) %></td>
								<td class="smaller-detail-line"><%= FormatCurrency(rs("SPLYSales"),0,-2,0) %></td> 
								<td class="smaller-detail-line"><%= MCS %></td>
								<td class="smaller-detail-line"><%= LCPvMCS %></td>
								<td class="smaller-detail-line"><%= ThreePavgvsMCS %></td>
								<td class="smaller-detail-line"><%= TwelvePAvgvMCS %></td>
								<td class="smaller-detail-line"><%= CurrentvsMCS %></td>
								<td class="smaller-detail-line"><%= LCP_ROI %></td>
								<td class="smaller-detail-line"><%= PavgROI %></td>
								<td class="smaller-detail-line"><a data-toggle="modal" data-show="true" href="#" data-cust-id="<%= rs("CustID") %>" data-lcp-gp="0" data-target="#modalEquipmentVPC" data-tooltip="true" data-title="View Customer Equipment"><%= TotalEquipmentValue %></a></td>
								<td class="smaller-detail-line"><%= PrimarySalesPerson %></td>
								<td class="smaller-detail-line"><%= CustomerType %></td>
								<td class="smaller-detail-line"><a data-toggle="modal" data-target="#modalEditCustomerNotes" data-category-id="-2" data-cust-id="<%= rs("CustID") %>" class="ole" rel="tooltip" style="cursor:pointer;"><i class="fa fa-file-text-o faa-pulse animated fa-2x" aria-hidden="true"></i></a></td>
								<td class="smaller-detail-line"><%= rules %></td>
								
							</tr>
							<%
							rs.movenext
							Loop
							%>
						
					
					</tbody>
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

	
	
</div>
	
	
	
	
	
  
</div>
  



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

	FUNCTION getTotalRecordsForTab(tabID,PeriodSeqBeingEvaluated,FilterSlsmn2)
		DIM retData
		retData="0"
		SELECT CASE tabID
			CASE 1
				varSQL="SELECT SUM(LCPv3PAvg) AS totalAmpunt, COUNT(*) AS totalQty FROM BI_DashboardSegmentTabs WHERE Tab = 'UP' AND SecondarySalesmanNumber = " & FilterSlsmn2
			CASE 2
				varSQL="SELECT SUM(LCPv3PAvg) AS totalAmpunt, COUNT(*) AS totalQty FROM BI_DashboardSegmentTabs WHERE Tab = 'DOWN' AND SecondarySalesmanNumber = " & FilterSlsmn2
			CASE 3
				varSQL="SELECT SUM(LCPv3PAvg) AS totalAmpunt, COUNT(*) AS totalQty FROM BI_DashboardSegmentTabs WHERE Tab = 'ZEROSALES' AND SecondarySalesmanNumber = " & FilterSlsmn2
			CASE 4
				varSQL="SELECT SUM(LCPv3PAvg) AS totalAmpunt, COUNT(*) AS totalQty FROM BI_DashboardSegmentTabs WHERE Tab = 'RULEDOUT' AND SecondarySalesmanNumber = " & FilterSlsmn2
			CASE 5 
				varSQL="SELECT SUM(LCPv3PAvg) AS totalAmpunt, COUNT(*) AS totalQty FROM BI_DashboardSegmentTabs WHERE Tab = 'ALL' AND SecondarySalesmanNumber = " & FilterSlsmn2
			CASE 6
				varSQL="SELECT COUNT(*) AS totalQty, max(totalAmpunt) as totalAmpunt FROM "
				varSQL = varSQL & " (SELECT TOP (50) totalAmpunt FROM "
				varSQL = varSQL & " (SELECT SUM(TwelvePAvgSales) AS totalAmpunt, CustID FROM "
				varSQL = varSQL & " BI_DashboardSegmentTabs WHERE (TAB = 'ALL') AND (SecondarySalesmanNumber = " & FilterSlsmn2 & ") "
				varSQL = varSQL & " GROUP BY CustID) AS derivedtbl_1 ORDER BY totalAmpunt DESC) AS derivedtbl_2"
			CASE 7 
				varSQL="SELECT COUNT(*) AS totalQty, max(totalAmpunt) as totalAmpunt FROM "
				varSQL = varSQL & " (SELECT TOP (50) totalAmpunt FROM "
				varSQL = varSQL & " (SELECT SUM(TwelvePAvgSales) AS totalAmpunt, CustID FROM "
				varSQL = varSQL & " BI_DashboardSegmentTabs WHERE (TAB = 'ALL') AND (SecondarySalesmanNumber = " & FilterSlsmn2 & ") "
				varSQL = varSQL & " GROUP BY CustID) AS derivedtbl_1 ORDER BY totalAmpunt ASC) AS derivedtbl_2"
			CASE 8
				varSQL="SELECT SUM(LCPv3PAvg) AS totalAmpunt, COUNT(*) AS totalQty FROM BI_DashboardSegmentTabs WHERE Tab = 'MCS' AND SecondarySalesmanNumber = " & FilterSlsmn2
			CASE 9
				varSQL="SELECT SUM(LCPv3PAvg) AS totalAmpunt, COUNT(*) AS totalQty FROM BI_DashboardSegmentTabs WHERE Tab = 'HIGHROI' AND SecondarySalesmanNumber = " & FilterSlsmn2
			CASE 10
				varSQL ="SELECT 0 AS totalAmpunt,COUNT(Distinct CategoryNameGetTerm) AS totalQty"
				varSQL = varSQL & " FROM CustCatPeriodSales_ReportData "
				varSQL = varSQL & " WHERE SecondarySalesman = " & FilterSlsmn2 & " AND ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated & " AND Category <> 0 "
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