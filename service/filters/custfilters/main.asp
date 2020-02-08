<!--#include file="../../../inc/header.asp"-->
<!--#include file="../../../inc/jquery_table_search.asp"-->
<!--#include file="../../../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../../../inc/InSightFuncs_Equipment.asp"--> 
<!--#include file="../../../inc/InSightFuncs_InventoryControl.asp"--> 
<!--#include file="../../../inc/InSightFuncs.asp"--> 
<!--#include file="../../../css/fa_animation_styles.css"-->
<%

Response.Write("<style type='text/css'>")
Response.Write("mark {")
Response.Write("    background-color: yellow;")
Response.Write("    color: black;")
Response.Write("}")
Response.Write("</style>")

ReportDate = Month(Now()) & "/01/" & Year(Now())

Server.ScriptTimeout = 900000 'Default value


Set cnnCustFilters = Server.CreateObject("ADODB.Connection")
cnnCustFilters.open (Session("ClientCnnString"))

CreateAuditLogEntry "Report","Report","Minor",0, MUV_Read("DisplayName") & " ran the report: MCS Report 1"


'Special code for when they are brought here by the automated email
'in this case, it just resets everything to default values and
'runs the page just for the salesperson who logged in
'it does this by writing to the Settings_reports table so the 
'rest of the code can just run normally from that point


If Request.QueryString("qlSls") <> "" Then

	quickloginSalesPerson = Request.QueryString("qlSls")
	
	SQLqlSls = "SELECT * from Settings_Reports where ReportNumber = 2101 AND UserNo = " & Session("userNo")
	
	Set cnnqlSls = Server.CreateObject("ADODB.Connection")
	cnnqlSls.open (Session("ClientCnnString"))
	Set rsqlSls = Server.CreateObject("ADODB.Recordset")
	Set rsqlSls= cnnqlSls.Execute(SQLqlSls)
	
	'Rec does not exist yet, make it quick but empty, update it later
	If rsqlSls.EOF Then
		SQLqlSls = "Insert into Settings_Reports (ReportNumber, UserNo) Values (2101 , " & Session("userNo") & ")"
		rsqlSls.Close
		Set rsqlSls= cnnqlSls.Execute(SQLqlSls)
	End If
	
	'Now update the table with the values
	SQLqlSls = "Update Settings_Reports Set ReportSpecificData1 = '" & quickloginSalesPerson & "', "
	SQLqlSls = SQLqlSls & "ReportSpecificData2 = 'All', " 
	SQLqlSls = SQLqlSls & "ReportSpecificData3 = 'All', "  
	SQLqlSls = SQLqlSls & "ReportSpecificData4 = 'All', "  
	SQLqlSls = SQLqlSls & "ReportSpecificData5 = '100', " 
	SQLqlSls = SQLqlSls & "ReportSpecificData6 = '10'"
	SQLqlSls = SQLqlSls & " WHERE ReportNumber = 2101 AND UserNo = " & Session("userNo")
	Set rsqlSls= cnnqlSls.Execute(SQLqlSls)
	cnnqlSls.Close
	
	Set rsqlSls = Nothing
	Set cnnqlSls = Nothing
	
End If


PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum()


'************************
'Read Settings_Reports
'************************
SQL = "SELECT * from Settings_Reports where ReportNumber = 2101 AND UserNo = " & Session("userNo")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnnCustFilters.Execute(SQL)
UseSettings_Reports = False
If NOT rs.EOF Then
	UseSettings_Reports = True
	FilterSlsmn1 = rs("ReportSpecificData1")
	FilterSlsmn2 = rs("ReportSpecificData2")
	If FilterSlsmn1 <> "All" Then FilterSlsmn1 = CInt(FilterSlsmn1)
	If FilterSlsmn2 <> "All" Then FilterSlsmn2 = CInt(FilterSlsmn2)
End If
'****************************
'End Read Settings_Reports
'****************************

%>

  
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
	#tableSuperSum td {vertical-align:top;}

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
.relative {position:relative;}
.summary-info {
	
  transition: 1s;
}
tr.group {cursor:pointer;}
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

</style>

<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.18/sorting/currency.js"></script>
<script type="text/javascript">
var datatableWidget;

function getfilters(obj) {
	
	if ($(obj).find("span.data-icon img").attr("src")=="/img/details_open.png") {
		
		$(obj).find("span.data-icon img").attr("src", "/img/details_close.png");
				
		var dataID=$(obj).closest('tr').attr("data-name");
		$("tr[data-child-value='"+dataID+"']").css("display","table-row");
	}
	else  {
		$(obj).find("span.data-icon img").attr("src", "/img/details_open.png");
		
		var dataID=$(obj).closest('tr').attr("data-name");
		$("tr[data-child-value='"+dataID+"']").css("display","none");
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
		switch ($("#ViewMode li.active > a").attr("data-grouping")) {
			case "1":
				datatableIni(1);
				break;
			case "0":
				datatableIni(2);
				break;
			case "3":
				datatableIni(3);
				break;
		
		}
		
		
	});
	
	


	$('#modalEditCategoryNotes').on('show.bs.modal', function(e) {

	    //get data-id attribute of the clicked order
	    var CustID = $(e.relatedTarget).data('cust-id');
	    var CategoryID = $(e.relatedTarget).data('category-id');
	    
	    //populate the textbox with the id of the clicked order
	    $(e.currentTarget).find('input[name="txtCustIDToPass"]').val(CustID);
	    $(e.currentTarget).find('input[name="txtCategoryID"]').val(CategoryID);
	    	    	    
	    var $modal = $(this);

    	$.ajax({
			type:"POST",
			url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			data: "action=GetContentForCategoryAnalysisByPeriodNotesModal&CategoryID="+encodeURIComponent(CategoryID)+"&CustID="+encodeURIComponent(CustID),
			success: function(response)
			 {					
               	 $modal.find('#modalEditCategoryNotesContent').html(response);               	 
             },
             failure: function(response)
			 {
			  	$modal.find('#modalEditCategoryNotesContent').html("Failed");
	            //var height = $(window).height() - 600;
	            //$(this).find(".modal-body").css("max-height", height);
             }
		});
    
	});

	function datatableIni(type) {
		console.log(type);
		switch (type) {
			case 1:
				var activeTab=$("#ViewMode li.active");
				datatableWidget=$('#tableSuperSum').on('preXhr.dt', function ( e, settings, data ) {
					$(".waitdiv").removeClass("d-none");
				})
				.on('xhr.dt', function ( e, settings, json, xhr ) {
					$(".waitdiv").addClass("d-none");
				console.log(json);
				$("#ViewMode li.byregional").remove();
				for(j=0;j<json.byRegionData.length;j++) {
					$("#ViewMode").append('<li class="byregional" data-regionid="'+json.byRegionData[j].regionID+'" role="presentation"><a href="#" role="tab" data-toggle="tab" data-grouping="3">'+json.byRegionData[j].region+'&nbsp;&nbsp;<span class="badge">'+json.byRegionData[j].qty+'</span></a></li>');
				}
				$("#ViewMode li").removeClass("active");
				$("#ViewMode li").eq(0).addClass("active");
				
				
				$('#ViewMode a[data-toggle="tab"]').off('shown.bs.tab');
				$('#ViewMode a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
					datatableWidget.clear();
					datatableWidget.destroy();
					switch ($("#ViewMode li.active > a").attr("data-grouping")) {
						case "1":
							datatableIni(1);
							break;
						case "0":
							datatableIni(2);
							break;
						case "3":
							datatableIni(3);
							break;
		
					}
				});
		
		
				
			} )
			
			.DataTable({
	        scrollY: 500,
	        scrollCollapse: true,
	        paging: true,
			ajax: "datatablejson.asp?maxID=1",
			deferRender: true,
			serverSide:true,
			procesing: true,
			lengthMenu: [[10, 25, 50, 75, 100, 500, -1],[10, 25, 50, 75, 100, 500, "All"]],
			pageLength: 100,
			
			createdRow: function ( row, data, index ) {
				if($("#ViewMode li.active > a").attr("data-grouping")=="1") {
				$(row).attr("data-child-value",data.id).css("display","none");
				}
			},	
			columns: [
					{ "data": "id" },
					{ "data": "CustName" },
					{ "data": "CustRegion" },
					{ "data": "FilterID" },
					{ "data": "filterDescription" },
					{ "data": "notes" },
					{ "data": "FrequencyType" },
					{ "data": "FrequencyTime" },
					{ "data": "LastChangeDateTime" },
					{ "data": "NextChangeDateTime" },
					{ "data": "dayTill" },
					{ "data": "Qty" },
					{ "data": "Price" },
					{ "data": "TotalCost" },
					{ "data": "CheckEquipment" },
					{ "data": "ActiveTicketNumber" },
					{ "data": "Action" },

				],
			columnDefs: [
				
				{"orderable": false,"targets": "_all" },
				{"className": "dt-center", "targets": [ 3,4,5,6,7,8,9,10,11,12,13,14,15,16]},
				{"className": "text-left", "targets": [ 0,1,2]},
				{"visible": false, "targets": [ 0,1,2]},
				{"targets": [0,1,2,3,4,6,7,8,9,11,12,13],
				"createdCell": function (td, cellData, rowData, row, col) {
						
						$(td).attr("ondblclick","javascript:doEdit(this);");
					
					}				
				
				},
				{ 			
				"targets": [ 10 ],
				"createdCell": function (td, cellData, rowData, row, col) {
					  
						if ( cellData < 1 ) {
							$(td).css('color', 'red')
						}
					  
					}				
				
				},
				
				{"targets": [ 14 ],	
				"createdCell": function (td, cellData, rowData, row, col) {
					  if($("#ViewMode li.active > a").attr("data-grouping")=="1") {
						$(td).html("");
					  }
					}				
				
				},
				{"targets": [ 15 ],	
				"createdCell": function (td, cellData, rowData, row, col) {
					  if($("#ViewMode li.active > a").attr("data-grouping")=="1" && cellData.length>0) {
						$(td).attr("ticket-data", cellData);
					  }
					}				
				
				},
				{"targets": [ 16 ],	
				"createdCell": function (td, cellData, rowData, row, col) {
					  var nameClient=rowData.CustName;
					  var clientID=rowData.id;
					  var filterID=rowData.FilterID;
					  var filterIntRecID=rowData.FilterIntRecID;
					  if($("#ViewMode li.active > a").attr("data-grouping")=="0" || $("#ViewMode li.active > a").attr("data-grouping")=="3") {
						$(td).html('<div class="dropdown"><button class="btn btn-success dropdown-toggle" type="button" id="dropdownMenu1" data-toggle="dropdown">Action</button><ul class="dropdown-menu dropdown-menu-right relative" aria-labelledby="dropdownMenu1"><li><a data-id="'+clientID+'" href="#" onclick="javascript:doEdit(this);">Edit</a></li><li><a data-id="'+clientID+'" href="#" onclick="javascript:toExclude(this);">Delete from filter Program</a></li><li><a data-filterintrecid="'+filterIntRecID+'" href="#" onclick="javascript:createNewFilterTicket(this);">Create Filter Ticket 1</a></li></ul></div>');
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
		
    }    ,
			
			drawCallback: function ( settings ) {
				
					
					var api = this.api();
					//if(api.column(0).visible()===false) {
						
					//	$("th.sorting").removeClass("sorting");
					//}
					
					var rows = api.rows( {page:'current'} ).nodes();
					var sumData=0;
					var last=null;
					if($("#ViewMode li.active > a").attr("data-grouping")=="1") {
						api.column(0, {page:'current'} ).data().each( function ( group, i ) {
							var rowData= api.row(i).data();
							
							
							if ( last !== group ) {
								var rowData= api.row(i).data();
								
								
								//console.log($(rows).eq(i).context[0].aoData[1]._aFilterData[0]+" "+$(rows).eq(i).context[0].aoData[1]._aFilterData[1]+"</b>, "+$(rows).eq(i).context[0].aoData[1]._aFilterData[2]);
								//console.log($(rows).eq(i).context[0]);
								var nameClient=rowData.CustName;
								var clientID=rowData.id;
								var filterID=rowData.FilterID;
								var filterIntRecID=rowData.FilterIntRecID;
								//console.log($(rows).eq(i).context[0]);
								
									$(rows).eq(i).before(
										'<tr class="group" data-name="'+group+'" class="collapsed" data-id="'+clientID+'"><td class="text-center"><button type="button" class="btn btn-link" onclick="javascript:getfilters(this);"><span class="data-icon"><img src="/img/details_open.png"></span></button></td><td colspan="7" ondblclick="javascript:callEdit(this);">'+clientID+' <b>'+ nameClient +'</b></td><td class="total-qty text-center font-bold"></td><td colspan="2"></td><td align="center">'+rowData.CheckEquipment+'</td><td align="center">&nbsp;</td><td  class="text-right"><div class="dropdown"><button class="btn btn-success dropdown-toggle" type="button" id="dropdownMenu1" data-toggle="dropdown">Action</button><ul class="dropdown-menu dropdown-menu-right relative" aria-labelledby="dropdownMenu1"><li><a data-id="'+clientID+'" href="#" onclick="javascript:doEdit(this);">Edit</a></li><li><a data-id="'+clientID+'" href="#" onclick="javascript:toExclude(this);">Delete from filter Program</a></li><li><a data-custid="'+clientID+'" data-filterintrecid="" href="#" onclick="javascript:createNewFilterTicket(this);">Create Filter Ticket</a></li></ul></div></td></tr>'
									);
								
								if (last!=null) {
									$("tr.group[data-name='"+last+"'] td.total-qty").html(sumData.toString());
									sumData=sumData=parseInt(rowData.Qty);
								}
								else sumData=parseInt(rowData.Qty);
								last = group;
							}
							else sumData+=parseInt(rowData.Qty);
						
						});
						$("tr.group[data-name='"+last+"'] td.total-qty").html(sumData.toString());
					}
					
				
				
				
				
        }
	});
				break;
		
		case 2:
			var activeTab=$("#ViewMode li.active");
			datatableWidget=$('#tableSuperSum').on('preXhr.dt', function ( e, settings, data ) {
			$(".waitdiv").removeClass("d-none");
			})
			.on('xhr.dt', function ( e, settings, json, xhr ) {
				$(".waitdiv").addClass("d-none");
				$("li.byregional").remove();
				for(j=0;j<json.byRegionData.length;j++) {
					$("#ViewMode").append('<li class="byregional" data-regionid="'+json.byRegionData[j].regionID+'" role="presentation "><a href="#" role="tab" data-toggle="tab" data-grouping="3">'+json.byRegionData[j].region+'&nbsp;&nbsp;<span class="badge">'+json.byRegionData[j].qty+'</span></a></li>');
				}
				$("#ViewMode li").removeClass("active");
				$("#ViewMode li").eq(1).addClass("active");
				
				$('#ViewMode a[data-toggle="tab"]').off('shown.bs.tab');
				$('#ViewMode a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
				datatableWidget.clear();
				datatableWidget.destroy();
				switch ($("#ViewMode li.active > a").attr("data-grouping")) {
					case "1":
						datatableIni(1);
						break;
					case "0":
						datatableIni(2);
						break;
					case "3":
						datatableIni(3);
						break;
				
				}
		
		
	});
			} )
			.DataTable({
	        scrollY: 500,
	        scrollCollapse: true,
	        paging: true,
			ajax: "datatablejson.asp?maxID=2",
			deferRender: true,
			serverSide:true,
			procesing: true,
			lengthMenu: [[10, 25, 50, 75, 100, 500, -1],[10, 25, 50, 75, 100, 500, "All"]],
			pageLength: 100,
			order: [[ 9, "asc" ]],
			createdRow: function ( row, data, index ) {
				if($("#ViewMode li.active > a").attr("data-grouping")=="1") {
				$(row).attr("data-child-value",data.id).css("display","none");
				}
				else {
					$(row).attr("data-id", data.id);
					$(row).addClass("group");
				}
			},	
			columns: [
					{ "data": "id" },
					{ "data": "CustName" },
					{ "data": "CustRegion" },
					{ "data": "FilterID" },
					{ "data": "filterDescription" },
					{ "data": "notes" },
					{ "data": "FrequencyType" },
					{ "data": "FrequencyTime" },
					{ "data": "LastChangeDateTime" },
					{ "data": "NextChangeDateTime" },
					{ "data": "dayTill" },
					{ "data": "Qty" },
					{ "data": "Price" },
					{ "data": "TotalCost" },
					{ "data": "CheckEquipment" },
					{ "data": "ActiveTicketNumber" },					
					{ "data": "Action" },

				],
			columnDefs: [
				{"orderable": true,"targets": [0,1,2,3,6,7,8,9,10,11,12,13] },
				{"orderable": false,"targets": [5,14,16] },
				{"className": "dt-center", "targets": [ 2,3,4,5,6,7,8,9,10,11,12,13,14,15,16]},
				{"className": "text-left", "targets": [ 0,1,2]},
				{"visible": false, "targets": [14]},
				{"visible": true, "targets": [0,1,2,3,6,7,8,9,10,11,12,14,15,16]},
				{"targets": [0,1,2,3,4,6,7,8,9,11,12,13],
				"createdCell": function (td, cellData, rowData, row, col) {
						
						$(td).attr("ondblclick","javascript:callEdit(this);");
					
					}				
				
				},
				{ 			
				"targets": [ 10 ],
				"createdCell": function (td, cellData, rowData, row, col) {
					  
						if ( cellData < 1 ) {
							$(td).css('color', 'red')
						}
					  
					}				
				
				},
				
				{"targets": [ 14 ],	
				"createdCell": function (td, cellData, rowData, row, col) {
					  if($("#ViewMode li.active > a").attr("data-grouping")=="1") {
						$(td).html("");
					  }
					}				
				
				},
				{"targets": [ 15 ],	
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData.length>0) $(td).attr("ticket-data", cellData);
					
					}				
				
				},
				{"targets": [ 16 ],	
				"createdCell": function (td, cellData, rowData, row, col) {
					 var nameClient=rowData.CustName;
					 var clientID=rowData.id;
					 var filterID=rowData.FilterID;
					 var filterIntRecID=rowData.FilterIntRecID;
					
					  if($("#ViewMode li.active > a").attr("data-grouping")=="0" || $("#ViewMode li.active > a").attr("data-grouping")=="3") {
						$(td).html('<div class="dropdown"><button class="btn btn-success dropdown-toggle" type="button" id="dropdownMenu1" data-toggle="dropdown">Action</button><ul class="dropdown-menu dropdown-menu-right" aria-labelledby="dropdownMenu1"><li><a data-id="'+clientID+'" href="#" onclick="javascript:doEdit(this);">Edit</a></li><li><a data-id="'+clientID+'" href="#" onclick="javascript:toExclude(this);">Delete from filter Program</a></li><li><a data-custid="" data-filterintrecid="'+filterIntRecID+'" href="#" onclick="javascript:createNewFilterTicket(this);">Create Filter Ticket</a></li></ul></div>');
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
		case 3:
			var activeTab=$("#ViewMode li.active");
			datatableWidget=$('#tableSuperSum').on('preXhr.dt', function ( e, settings, data ) {
			$(".waitdiv").removeClass("d-none");
			})
			.on('xhr.dt', function ( e, settings, json, xhr ) {
				$(".waitdiv").addClass("d-none");
				$("li.byregional").remove();
				for(j=0;j<json.byRegionData.length;j++) {
					$("#ViewMode").append('<li class="byregional" data-regionid="'+json.byRegionData[j].regionID+'" role="presentation "><a href="#" role="tab" data-toggle="tab" data-grouping="3">'+json.byRegionData[j].region+'&nbsp;&nbsp;<span class="badge">'+json.byRegionData[j].qty+'</span></a></li>');
				}
				if ($(activeTab).hasClass("byregional")) {
					
					$("#ViewMode li").removeClass("active");
					$("li.byregional[data-regionid='"+$(activeTab).attr("data-regionid")+"']").addClass("active");
				}
				
				$('#ViewMode a[data-toggle="tab"]').off('shown.bs.tab');
				$('#ViewMode a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
					datatableWidget.clear();
					datatableWidget.destroy();
					switch ($("#ViewMode li.active > a").attr("data-grouping")) {
						case "1":
							datatableIni(1);
							break;
						case "0":
							datatableIni(2);
							break;
						case "3":
							datatableIni(3);
							break;
					
					}
		
		
				});
			} )
			.DataTable({
	        scrollY: 500,
	        scrollCollapse: true,
	        paging: true,
			ajax: "datatablejson.asp?maxID=2&regionID="+$(activeTab).attr("data-regionid"),
			deferRender: true,
			serverSide:true,
			procesing: true,
			lengthMenu: [[10, 25, 50, 75, 100, 500, -1],[10, 25, 50, 75, 100, 500, "All"]],
			pageLength: 100,
			order: [[ 9, "asc" ]],
			createdRow: function ( row, data, index ) {
				if($("#ViewMode li.active > a").attr("data-grouping")=="1") {
				$(row).attr("data-child-value",data.id).css("display","none");
				}
				else {
					$(row).attr("data-id", data.id);
					$(row).addClass("group");
				}
			},	
			columns: [
					{ "data": "id" },
					{ "data": "CustName" },
					{ "data": "CustRegion" },
					{ "data": "FilterID" },
					{ "data": "filterDescription" },
					{ "data": "notes" },
					{ "data": "FrequencyType" },
					{ "data": "FrequencyTime" },
					{ "data": "LastChangeDateTime" },
					{ "data": "NextChangeDateTime" },
					{ "data": "dayTill" },
					{ "data": "Qty" },
					{ "data": "Price" },
					{ "data": "TotalCost" },
					{ "data": "CheckEquipment" },
					{ "data": "ActiveTicketNumber" },					
					{ "data": "Action" },

				],
			columnDefs: [
				{"orderable": true,"targets": [0,1,2,3,6,7,8,9,10,11,12,13] },
				{"orderable": false,"targets": [5,14,16] },
				{"className": "dt-center", "targets": [ 2,3,4,5,6,7,8,9,10,11,12,13,14,15,16]},
				{"className": "text-left", "targets": [ 0,1,2]},
				{"visible": false, "targets": [2,14]},
				{"visible": true, "targets": [0,1,3,6,7,8,9,10,11,12,14,15,16]},
				{"targets": [0,1,2,3,4,6,7,8,9,11,12,13],
				"createdCell": function (td, cellData, rowData, row, col) {
						
						$(td).attr("ondblclick","javascript:callEdit(this);");
					
					}				
				
				},
				{ 			
				"targets": [ 10 ],
				"createdCell": function (td, cellData, rowData, row, col) {
					  
						if ( cellData < 1 ) {
							$(td).css('color', 'red')
						}
					  
					}				
				
				},
				
				{"targets": [ 14 ],	
				"createdCell": function (td, cellData, rowData, row, col) {
					  if($("#ViewMode li.active > a").attr("data-grouping")=="1") {
						$(td).html("");
					  }
					}				
				
				},
				{"targets": [ 15 ],	
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData.length>0) $(td).attr("ticket-data", cellData);
					
					}				
				
				},
				{"targets": [ 16 ],	
				"createdCell": function (td, cellData, rowData, row, col) {
					 var nameClient=rowData.CustName;
					 var clientID=rowData.id;
					 var filterID=rowData.FilterID;
					 var filterIntRecID=rowData.FilterIntRecID;
					
					  if($("#ViewMode li.active > a").attr("data-grouping")=="0" || $("#ViewMode li.active > a").attr("data-grouping")=="3") {
						$(td).html('<div class="dropdown"><button class="btn btn-success dropdown-toggle" type="button" id="dropdownMenu1" data-toggle="dropdown">Action</button><ul class="dropdown-menu dropdown-menu-right" aria-labelledby="dropdownMenu1"><li><a data-id="'+clientID+'" href="#" onclick="javascript:doEdit(this);">Edit</a></li><li><a data-id="'+clientID+'" href="#" onclick="javascript:toExclude(this);">Delete from filter Program</a></li><li><a data-custid="" data-filterintrecid="'+filterIntRecID+'" href="#" onclick="javascript:createNewFilterTicket(this);">Create Filter Ticket</a></li></ul></div>');
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
	$('#modalEquipmentVPC').on('show.bs.modal', function(j) {

	    //get data-id attribute of the clicked order
	    var CustID = $(j.relatedTarget).data('cust-id');
	    var LCPGP = $(j.relatedTarget).data('lcp-gp');
 
	    //populate the textbox with the id of the clicked order
	    $(j.currentTarget).find('input[name="txtCustIDToPass"]').val(CustID);
	    $(j.currentTarget).find('input[name="txtLastClosedPeriodGP"]').val(LCPGP);
	    	    
	    var $modal = $(this);
	    //$modal.find('#PleaseWaitPanelModal').show();  
		console.log("modal ===========================");
    	$.ajax({
			type:"POST",
			url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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
			url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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
			url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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
			url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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


	$('#modalGeneralNotesGroupM').on('show.bs.modal', function(j) {

		var $modal = $(this);
	    //get data-id attribute of the clicked order
	    var CustID = $(j.relatedTarget).data('cust-id');
		var CustName = $(j.relatedTarget).data('cust-name');
	    var MCSvariance = $(j.relatedTarget).data('mcs-variance');
		var SP1 = $(j.relatedTarget).data('mcs-salespersonid1');
		var SP2 = $(j.relatedTarget).data('mcs-salespersonid2');
 		var SPname1 = $(j.relatedTarget).data('mcs-salesperson1');
		var SPname2 = $(j.relatedTarget).data('mcs-salesperson2');
		var MaxMCSCharge = $(j.relatedTarget).data('maxmcscharge');		
		var MCSDollars = $(j.relatedTarget).data('mcsdollars');
		
		var SPhtml = "<option value=\"" + SP1 + "\">" + SPname1 + "</option>\n<option value=\"" + SP2 + "\">" + SPname2 + "</option>\n";
<%
		SQL = "SELECT * from BI_MCSReasons"
		Set rsReason = Server.CreateObject("ADODB.Recordset")
		Set rsReason = cnnCustFilters.Execute(SQL)
		response.write "var ReasonHTML = ""<option value=\""selectreason\"">Select Reason</option>"
		Do While NOT rsReason.EOF
			response.write "<option value=\""" & rsReason("InternalRecordIdentifier") & "\"">" & replace(rsReason("Reason"), """", "\""") & "</option>" 
			rsReason.MoveNext
		Loop
		response.write """;"
		
%>
		var userNo1 = $(j.relatedTarget).data('mcs-userno');
		var MCSMonth = $(j.relatedTarget).data('mcs-month');

		if ($('#GNGMReasons').next().prop("tagName") == 'BR') {
			$('#GNGMReasons').next('br').remove();
			$('#GNGMReasons').next('span').remove();
		}	
		
		$('#GNGMcustname').html(CustName);
		$('#GNGMvariance').html("MCS Variance " + MCSvariance.toString().replace("\-", "\-\$").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","));
	    //populate the textbox with the id of the clicked order
	    $(j.currentTarget).find('input[name="GNGMCustIDToPass"]').val(CustID);
		$(j.currentTarget).find('input[name="GNGMuserno"]').val(userNo1);
		$(j.currentTarget).find('input[name="GNGMMSCMonth"]').val(MCSMonth);
		$(j.currentTarget).find('input[id="GNGMchange_lvf"]').val(MaxMCSCharge.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","));
		$(j.currentTarget).find('input[id="GNGMchange_mcs"]').val(MCSDollars.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","));
		
		if (isNaN(MaxMCSCharge) || MaxMCSCharge == "" ) {
			$(j.currentTarget).find('input[id="GNGMinvoice_amount"]').val(Math.abs(MCSvariance).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","));
			$("#maxallowed").text("");
		} else {
			$(j.currentTarget).find('input[id="GNGMinvoice_amount"]').val(Math.abs(MaxMCSCharge).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ","));
			$("#maxallowed").text("Max Allowed")
		}
		$(j.currentTarget).find('select[id="GNGMsalesperson"]').html(SPhtml);
		$(j.currentTarget).find('select[id="GNGMReasons"]').html(ReasonHTML);
		$(j.currentTarget).find('input:checked[type="radio"]').each(function(){
			$(this).prop('checked', false);  
		});
		$('.hidemsg').hide();
	    //$modal.find('#PleaseWaitPanelModal').show();  



	});

//$("#modalGeneralNotesGroupM").click(function(){
//        $('#modalGeneralNotesGroupM').modal('toggle')
//    }); 	
	
	$("#GNGMSave").click(function(){
		var CustID = $("#GNGMCustIDToPass").val();
		var GNGMuserno = $("#GNGMuserno").val();
		var GNGMMSCMonth = $("#GNGMMSCMonth").val();
		var GNGMinvoice_amount = $("#GNGMinvoice_amount").val();
		var GNGMsalesperson = $("#GNGMsalesperson").val();
		var GNGMMsg = $("#GNGMMsg").val();
		var groupm = $("input[name='groupm']:checked").val();
		var GNGMchange_lvf = $("#GNGMchange_lvf").val();
		var GNGMchange_mcs = $("#GNGMchange_mcs").val();		
		var GNGMReasons = $("#GNGMReasons").val();
		var CustName = $("#GNGMcustname").html();
		var MCSvariance = $("#GNGMvariance").html();

		if (groupm == "no_action_necessary") {
			if(GNGMReasons == "selectreason"){
				$('#GNGMReasons').after('<BR><span class="error"> Please Select Reason</span>');
				return;
			}
		}
		if (groupm == "send_message_to_someone") {
			$("#modalGeneralNotesGroupM").modal('toggle');
			$("#SUParamsFromParent").val("CustID="+encodeURIComponent(CustID)+"&CustName="+encodeURIComponent(CustName)+"&MCSvariance="+encodeURIComponent(MCSvariance)+"&GNGMuserno="+encodeURIComponent(GNGMuserno)+"&GNGMMSCMonth="+encodeURIComponent(GNGMMSCMonth)+"&GNGMinvoice_amount="+encodeURIComponent(GNGMinvoice_amount)+"&GNGMsalesperson="+encodeURIComponent(GNGMsalesperson)+"&GNGMMsg="+encodeURIComponent(GNGMMsg)+"&groupm="+encodeURIComponent(groupm)+"&GNGMchange_lvf="+encodeURIComponent(GNGMchange_lvf)+"&GNGMchange_mcs="+encodeURIComponent(GNGMchange_mcs)+"&GNGMReasons="+encodeURIComponent(GNGMReasons));		
			$("#modalSelectUser").modal('toggle');
			return;
		}
		if (groupm == "notify_selected_sales_person") {
			$("#modalGeneralNotesGroupM").modal('toggle');
			$("#SUParamsFromParent").val("CustID="+encodeURIComponent(CustID)+"&CustName="+encodeURIComponent(CustName)+"&MCSvariance="+encodeURIComponent(MCSvariance)+"&GNGMuserno="+encodeURIComponent(GNGMuserno)+"&GNGMMSCMonth="+encodeURIComponent(GNGMMSCMonth)+"&GNGMinvoice_amount="+encodeURIComponent(GNGMinvoice_amount)+"&GNGMsalesperson="+encodeURIComponent(GNGMsalesperson)+"&GNGMMsg="+encodeURIComponent(GNGMMsg)+"&groupm="+encodeURIComponent(groupm)+"&GNGMchange_lvf="+encodeURIComponent(GNGMchange_lvf)+"&GNGMchange_mcs="+encodeURIComponent(GNGMchange_mcs)+"&GNGMReasons="+encodeURIComponent(GNGMReasons));					
			$("#modalSelectUser").modal('toggle');
			return;
		}

    	$.ajax({
			type:"POST",
			url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			data: "action=SaveGeneralNotesGroupM&CustID="+encodeURIComponent(CustID)+"&GNGMuserno="+encodeURIComponent(GNGMuserno)+"&GNGMMSCMonth="+encodeURIComponent(GNGMMSCMonth)+"&GNGMinvoice_amount="+encodeURIComponent(GNGMinvoice_amount)+"&GNGMsalesperson="+encodeURIComponent(GNGMsalesperson)+"&GNGMMsg="+encodeURIComponent(GNGMMsg)+"&groupm="+encodeURIComponent(groupm)+"&GNGMchange_lvf="+encodeURIComponent(GNGMchange_lvf)+"&GNGMchange_mcs="+encodeURIComponent(GNGMchange_mcs)+"&GNGMReasons="+encodeURIComponent(GNGMReasons),
			success: function(response)
			 {
				if (response.startsWith("Error:")) {				
					swal({
						title: 'Error Saving notes',
						text: response,
						type: 'error'
					})
					return;
				}
				 $("#modalGeneralNotesGroupM").modal('toggle');
				 if (groupm == "no_action_necessary") {
					$('#btn' + CustID).removeClass('btn-success').addClass('btn-default');
					$('#btn' + CustID).addClass('noaction');
					$("#frmMCS_Report1").submit();
					return;
				 } else {
					$('#btn' + CustID).removeClass('btn-success').addClass('btn-default');
				 }
				 if (groupm == 'remove_client') {
					swal({
					  title: 'Are you sure?',
					  text: "You have opted to remove this client from the MCS program. Are you sure?",
					  type: 'warning',
					  showCancelButton: true,
					  confirmButtonColor: '#3085d6',
					  cancelButtonColor: '#d33',
					  confirmButtonText: 'Yes, Remove Client!'
					},
					function(isConfirm){

					  if (isConfirm) {
						$.ajax({
							type:"POST",
							url: "/inc/InSightFuncs_AjaxForBizIntelModals.asp",
							cache: false,
							data: "action=DeleteMCSClientbyCustID&CustID="+encodeURIComponent(CustID),						
							success: function(response) {
								if (response.startsWith("Error:")) {									
									swal({
										title: 'Error Removing Client',
										text: response,
										type: 'error'
									})									
								} else {									
									$("#RemovedCustID").val(CustID);
									//swal({
									//	title: 'Client ' + CustID + ' Removed!',
									//	text: response,
									//	type: 'success'
									//})
									$("#frmMCS_Report1").submit();
								}
							},
							failure: function(response)
							{
									swal({
										title: 'Error Removing Client',
										text: 'Failed',
										type: 'error'
									})
							}															
						});
						
					  } else {
									swal({
										title: 'Client was not removed',
										text: 'You chose not to remove Client',
										type: 'warning'
									})
					  }
					})					
				 }  else {
					$("#frmMCS_Report1").submit();
/*				 	if (groupm == "change_mcs") {						
						$("#frmMCS_Report1").submit();
					} else {
						swal({
							title: 'Changes Saved',
							text: response,
							type: 'success'
						})
					}
*/					
					
				}
				 
             },
             failure: function(response)
			 {
				swal({
					title: 'Error Saving Notes',
					text: 'Failed',
					type: 'error'
				})
             }
		});
		
		        
    });

	$('#modalSelectUser').on('show.bs.modal', function(j) {

		var parentvars = $("#SUParamsFromParent").val();
		pairs = parentvars.split('&');
		for (var i = 0; i < pairs.length; i++) {
			var pair = pairs[i].split('=');
			switch (decodeURIComponent(pair[0])) {
				case "CustID":
					var CustID = decodeURIComponent(pair[1] || '');
				case "CustName":
					var CustName = decodeURIComponent(pair[1] || '');
				case "MCSvariance":
					var MCSvariance = decodeURIComponent(pair[1] || '');
				case "GNGMuserno":
					var GNGMuserno = decodeURIComponent(pair[1] || '');
				case "GNGMMSCMonth":
					var GNGMMSCMonth = decodeURIComponent(pair[1] || '');
				case "GNGMinvoice_amount":
					var GNGMinvoice_amount = decodeURIComponent(pair[1] || '');
				case "GNGMsalesperson":
					var GNGMsalesperson = decodeURIComponent(pair[1] || '');
				case "GNGMMsg":
					var GNGMMsg = decodeURIComponent(pair[1] || '');
				case "groupm":
					var groupm = decodeURIComponent(pair[1] || '');
				case "GNGMchange_lvf":
					var GNGMchange_lvf = decodeURIComponent(pair[1] || '');
				case "GNGMchange_mcs":
					var GNGMchange_mcs = decodeURIComponent(pair[1] || '');	
				case "GNGMReasons":
					var GNGMReasons = decodeURIComponent(pair[1] || '');
			}
		}		

		var $modal = $(this);		

		$('#SUcustname').html(CustName);
		$('#SUvariance').html(MCSvariance);
		$("#SUMsg").val("");
		var selectusers = "";
		var selectedusers = "";
		var selectedusers1 = "";
		var selectedusers2 = "";
		var sep = "";

		if (groupm == "notify_selected_sales_person") {
			$.ajax({
				type:"POST",
				url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				cache: false,
				async: false,
				data: "action=getUsersBySalesperson&SalesPersons="+GNGMsalesperson,
				beforeSend: function() {
					$('#PleaseWaitPanelModal').show();
				},
				complete: function(){
					$('#PleaseWaitPanelModal').hide();
				},
				success: function(response)
				 {	
					response = $.parseJSON(response);
					sep = "";
					$.each(response, function (key, value) {
						selectedusers2 += sep + value.UserNo;
						sep = ",";
					});						
				 },
				 failure: function(response)
				 {
					//selectedusers = "";
				 }
			});		
			if (selectedusers2 == "") {
				swal({
					title: 'Error SalesPerson Not Found',
					text: 'Failed',
					type: 'error'
				})				
			}
		}
		var selectusersHTML;
		var selectedusersHTML;
    	$.ajax({
			type:"POST",
			url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			async: false,
			data: "action=getCCUsers&SelectedUsers="+selectedusers2,
		  	beforeSend: function() {
		     	$('#PleaseWaitPanelModal').show();
		  	},
		  	complete: function(){
		     	$('#PleaseWaitPanelModal').hide();
		  	},
			success: function(response)
			 {
				selectedusers = $.parseJSON(response);				
				sep = "";
				$.each(selectedusers, function (key, value) {
					selectedusers1 += sep + value.UserNo;
					sep = ",";
				});
				
             },
             failure: function(response)
			 {
			  	//selectedusers = "";
             }
		});
		$.ajax({
			type:"POST",
			url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			async: false,
			data: "action=getSelectUsers&SelectedUsers="+selectedusers1,
			beforeSend: function() {
				$('#PleaseWaitPanelModal').show();		     	
			},
			complete: function(){
				$('#PleaseWaitPanelModal').hide();
			},
			success: function(response)
			 {																		
				selectusers = $.parseJSON(response);
					
			 },
			 failure: function(response)
			 {
				selectusers = "";
			 }
		});			
		
		if (selectusers) {
			$.each(selectusers, function (key, value) {
				selectusersHTML += "<option value=\"" + value.UserNo + "\">" + value.FullName + "</option>";
			});			
		}
		
		if (selectedusers) {
			$.each(selectedusers, function (key, value) {
//				selectedusersHTML += "<option value=\"" + value.UserNo + "\" disabled>" + value.FullName + "</option>";
				selectedusersHTML += "<option value=\"" + value.UserNo + "\">" + value.FullName + "</option>";
			});					
		}		

		$("#lstSelectUserIDs").html(selectusersHTML);
		$("#lstSelectedUserIDs").html(selectedusersHTML);			
	});
	
	$("#SUSave").click(function(){
		var parentvars = $("#SUParamsFromParent").val();
		pairs = parentvars.split('&');
		for (var i = 0; i < pairs.length; i++) {
			var pair = pairs[i].split('=');
			switch (decodeURIComponent(pair[0])) {
				case "CustID":
					var CustID = decodeURIComponent(pair[1] || '');
				case "CustName":
					var CustName = decodeURIComponent(pair[1] || '');
				case "MCSvariance":
					var MCSvariance = decodeURIComponent(pair[1] || '');
				case "GNGMuserno":
					var GNGMuserno = decodeURIComponent(pair[1] || '');
				case "GNGMMSCMonth":
					var GNGMMSCMonth = decodeURIComponent(pair[1] || '');
				case "GNGMinvoice_amount":
					var GNGMinvoice_amount = decodeURIComponent(pair[1] || '');
				case "GNGMsalesperson":
					var GNGMsalesperson = decodeURIComponent(pair[1] || '');
				case "GNGMMsg":
					var GNGMMsg = decodeURIComponent(pair[1] || '');
				case "groupm":
					var groupm = decodeURIComponent(pair[1] || '');
				case "GNGMchange_lvf":
					var GNGMchange_lvf = decodeURIComponent(pair[1] || '');
				case "GNGMchange_mcs":
					var GNGMchange_mcs = decodeURIComponent(pair[1] || '');	
				case "GNGMReasons":
					var GNGMReasons = decodeURIComponent(pair[1] || '');
			}
		}

		var lstSelectedUserIDs = "";
		var SUMsg = $("#SUMsg").val();
		var src = $("#lstSelectedUserIDs");
		var sep = "";
		$("#lstSelectedUserIDs option").each(function()
		{
			lstSelectedUserIDs += sep + $(this).val();
			sep = ",";
			
		});
			
    	$.ajax({
			type:"POST",
			url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			data: "action=SaveGeneralNotesGroupM&CustID="+encodeURIComponent(CustID)+"&GNGMuserno="+encodeURIComponent(GNGMuserno)+"&GNGMMSCMonth="+encodeURIComponent(GNGMMSCMonth)+"&GNGMinvoice_amount="+encodeURIComponent(GNGMinvoice_amount)+"&GNGMsalesperson="+encodeURIComponent(GNGMsalesperson)+"&GNGMMsg="+encodeURIComponent(GNGMMsg)+"&groupm="+encodeURIComponent(groupm)+"&GNGMchange_lvf="+encodeURIComponent(GNGMchange_lvf)+"&GNGMchange_mcs="+encodeURIComponent(GNGMchange_mcs)+"&GNGMReasons="+encodeURIComponent(GNGMReasons)+"&lstSelectedUserIDs="+encodeURIComponent(lstSelectedUserIDs)+"&SUMsg="+encodeURIComponent(SUMsg),
			success: function(response)
			 {				
				if (response.startsWith("Error:")) {				
					swal({
						title: 'Error Saving notes',
						text: response,
						type: 'error'
					})
					return;
				}
				 $("#modalSelectUser").modal('toggle');
				 $('#btn' + CustID).removeClass('btn-success').addClass('btn-default ');
				 if (groupm == 'remove_client') {
					swal({
					  title: 'Are you sure?',
					  text: "You have opted to remove this client from the MCS program. Are you sure?",
					  type: 'warning',
					  showCancelButton: true,
					  confirmButtonColor: '#3085d6',
					  cancelButtonColor: '#d33',
					  confirmButtonText: 'Yes, Remove Client!'
					},
					function(isConfirm){

					  if (isConfirm) {
						$.ajax({
							type:"POST",
							url: "/inc/InSightFuncs_AjaxForBizIntelModals.asp",
							cache: false,
							data: "action=DeleteMCSClientbyCustID&CustID="+encodeURIComponent(CustID),						
							success: function(response) {
								if (response.startsWith("Error:")) {									
									swal({
										title: 'Error Removing Client',
										text: response,
										type: 'error'
									})									
								} else {									
									swal({
										title: 'Client ' + CustID + ' Removed!',
										text: response,
										type: 'success'
									})
									
								}
							},
							failure: function(response)
							{
									swal({
										title: 'Error Removing Client',
										text: 'Failed',
										type: 'error'
									})
							}															
						});
						
					  } else {
									swal({
										title: 'Client was not removed',
										text: 'You chose not to remove Client',
										type: 'warning'
									})
					  }
					})					
				 }  else {
					swal({
						title: 'Changes Saved',
						text: response,
						type: 'success'
					})

				}
				 
             },
             failure: function(response)
			 {
				swal({
					title: 'Error Saving Notes',
					text: 'Failed',
					type: 'error'
				})
             }
		});
		        
    });
		
    //$("input[name$='groupm']").click(function() {
    //    if ($(this).val() == "notify_selected_sales_person") {		
	//		$('.hidemsg').show();
	//	} else {
	//		$('.hidemsg').hide();
	//	}
    //});	

	$("#AddMCSClientSave").click(function(){
	
		$("#modalAddMCSClient").modal('toggle');
		var CustID = $("#AddMCSClientCustIDToPass").val();
		var CustName = $("#AddMCSClientCustNameToPass").val();		
		$.ajax({
			type:"GET",
			url: "editFilterDataList.asp",
			cache: false,
			data: "CustomerID="+CustID,
			success: function(response) {
				$("#modalEditCustomerFilter #mode").html("Add");
				$("#modalEditCustomerFilter .modal-body").html(response);
				$("#modalEditCustomerFilter .modal-footer .btn.btn-primary").html("Save & Close Window");
				if($(".filterList tbody tr").length==0) toAddLine();
				$("#modalEditCustomerFilter").modal("show");		
			},
			error: function (response) {
			
			
			},
			complete: function() {
			$("#modalEditCustomerFilter .modal-title #EditCustomerFiltercustid").html($("#EditCustomerFilterCustIDToPass").val());
				$("#modalEditCustomerFilter .modal-title #EditCustomerFiltercustname").html($("#EditCustomerFilterCustName").val());
				$("#modalEditCustomerFilter .filter-program").height($("#modalEditCustomerFilter .modal-body").height()*0.6);
				var divHeight=$("#modalEditCustomerFilter .modal-body").height()*0.4-72;
				console.log(divHeight);
				$("#modalEditCustomerFilter .equipment-list").height(divHeight);
			}
			
		});
		//$("#EditCustomerFiltercustid").text(CustID);
		//$("#EditCustomerFiltercustname").text(CustName);
		//$("#EditCustomerFilterCustIDToPass").val(CustID);
		//$("#modalEditCustomerFilter").modal('toggle');
	})

	$('#modalEditCustomerFilter').on('show.bs.modal', function(j) {

		var $modal = $(this);
	    //get data-id attribute of the clicked order
		var CustID = $("#AddMCSClientCustIDToPass").val();
		var CustName = $("#AddMCSClientCustNameToPass").val();		
		$modal.find(".filter-program").height($modal.find(".modal-body").height()*0.6);
		$modal.find(".equipment-list").height($modal.find(".modal-body").height()*0.4);
		$("#EditCustomerFiltercustid").text(CustID);
		$("#EditCustomerFiltercustname").text(CustName);
		$("#EditCustomerFilterMonthlyContractedSalesDollars").val("");
		$("#EditCustomerFilterMaxMCSCharge").val("");
		
	});

	$('#modalEditCustomerFilter').on('hidden.bs.modal', function(j) {
			if ($('#EditCustomerFilterMonthlyContractedSalesDollars').next().prop("tagName") == 'BR') {
				$('#EditCustomerFilterMonthlyContractedSalesDollars').next('br').remove();
				$('#EditCustomerFilterMonthlyContractedSalesDollars').next('span').remove();
			}
			if ($('#EditCustomerFilterMaxMCSCharge').next().prop("tagName") == 'BR') {
				$('#EditCustomerFilterMaxMCSCharge').next('br').remove();
				$('#EditCustomerFilterMaxMCSCharge').next('span').remove();
			}			
	});	
	
	$("#EditCustomerFilterSave").click(function(){	
		var CustID = $("#EditCustomerFilterCustIDToPass").val();
		var MCSDollars = $("#EditCustomerFilterMonthlyContractedSalesDollars").val();
		var MaxMCSCharge = $("#EditCustomerFilterMaxMCSCharge").val();
		var adderror = 0;
		
		var numberReg =  /^[0-9\.]+$/;
		if(!numberReg.test(MCSDollars)){
            $('#EditCustomerFilterMonthlyContractedSalesDollars').after('<BR><span class="error"> Please Enter Numbers only</span>');
			adderror = 1;
        }
		//if(!numberReg.test(MaxMCSCharge)){
        //    $('#EditCustomerFilterMaxMCSCharge').after('<BR><span class="error">  Please Enter Numbers only</span>');
		//	adderror = 1;
        //}
		if (adderror) {
			return;
		}
    	$.ajax({
			type:"POST",
			url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			cache: false,
			data: "action=AddMCSClientbyCustID&CustID="+encodeURIComponent(CustID)+"&MCSDollars="+encodeURIComponent(MCSDollars)+"&MaxMCSCharge="+encodeURIComponent(MaxMCSCharge),
			success: function(response)
			 {
				if (response.startsWith("Error:")) {				
					swal({
						title: 'Error Adding MCS Client',
						text: response,
						type: 'error'
					})
					return;
				} else {
					$("#AddedCustID").val(CustID);
					$("#frmMCS_Report1").submit();				
/*					swal({
						title: 'Success Client Add',
						text: response,
						type: 'success'
					})
					$.ajax({
						type:"POST",
						url: "../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
						cache: false,
						data: "action=GetTRofNewMCSClientbyCustID&CustID="+encodeURIComponent(CustID),
						success: function(response)
						{
								
								if(response.startsWith("<tr ")) {
									$('#tableSuperSum').find('tbody tr:first').before(response);
								}
						}
					});
*/					
				}
			 },
			failure: function(response)
			 {
				swal({
					title: 'Error Adding MCS Client',
					text: response,
					type: 'error'
				})
             }
		});

		$("#modalEditCustomerFilter").modal('toggle');
    });

		
	$("#GNGMReasons" ).change(function() {				
		$("#GNGMReasons" ).parent().parent().find('input[type="radio"]').prop("checked", true);
	});
	$("#GNGMchange_mcs" ).mousedown(function() {				
		$("#GNGMchange_mcs" ).parent().parent().find('input[type="radio"]').prop("checked", true);
	});	
	$("#GNGMchange_lvf" ).mousedown(function() {				
		$("#GNGMchange_lvf" ).parent().parent().find('input[type="radio"]').prop("checked", true);
	});	
	$("#GNGMinvoice_amount" ).mousedown(function() {				
		$("#GNGMinvoice_amount" ).parent().parent().find('input[type="radio"]').prop("checked", true);
	});	
	$("#GNGMsalesperson" ).change(function() {				
		$("#GNGMsalesperson" ).parent().parent().find('input[type="radio"]').prop("checked", true);
	});	
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
	function summaryinfoShow() {
	
		if ($(".summary-info").hasClass("hidden")) {
			$(".summary-info").removeClass("hidden");
			$("#summaryinfo").html("Hide Summary Info");
		}
		else {
		$(".summary-info").addClass("hidden");
		$("#summaryinfo").html("Show Summary Info");
		}
	
	}

</script>




<h3 class="page-header"><i class="fa fa-filter"></i>&nbsp;Manage Filter Changes
	&nbsp;&nbsp;
	<small><button type="button" class="btn btn-success" id="summaryinfo" onclick="javascript:summaryinfoShow();">Show Summary Info</button></small>
	<small><button type="button" class="btn btn-success" id="btnAddClientToMCS" data-toggle="modal" data-target="#modalAddMCSClient"><font size="2px" >Add Client To Filter Program</font></button></small>
</h3>

 
	
	
<form method="POST" name="frmMCS_Report1" id="frmMCS_Report1" value="frmMCS_Report1" action="main.asp">	
	<input id="frmMCS_Report1submitted" name="frmMCS_Report1submitted" type="hidden" value="1">
	<input id="RemovedCustID" name="RemovedCustID" type="hidden" value="">
	<div class="summary-info hidden">
	<% Call CalcAndDisplaySummary %>
	</div>
</form>	    


<!-- row !-->
<div class="row">
<!-- Nav tabs -->
  <ul class="nav nav-tabs" id="ViewMode" role="tablist">
    <li role="presentation" class="active"><a href="#bycustomer"  role="tab" data-toggle="tab" data-grouping="1">By Customer</a></li>
    <li role="presentation"><a href="#detailed" aria-controls="profile" role="tab" data-toggle="tab" data-grouping="0">Detailed View</a></li>

  </ul>
  

  <div class="container-fluid" style="padding-top:20px;">
		<div class="row">
           <table id="tableSuperSum" class="display compact" style="width:100%;">
              <thead>
                  <tr>	
                		<th class="td-align1 gen-info-header" colspan="3" style="border-right: 2px solid #555 !important;"><%=GetTerm("Customer")%></th>
						<th class="td-align1 vpc-3pavg-header" colspan="3" style="border-right: 2px solid #555 !important;">Filter Information</th>
						<th class="td-align1 vpc-lcp-header" colspan="5" style="border-right: 2px solid #555 !important;">Schedule (Frequency)</th>
						<th class="td-align1 vpc-misc-header" colspan="3"  style="border-right: 2px solid #555 !important;">Price</th>
						<th class="td-align1 vpc-current-header" style="border-right: 2px solid #555 !important;">Equipment</th>
						<th class="td-align1 activities-header" colspan="2" style="border-right: 2px solid #555 !important;">Activities</th>
					</tr>
				
					<tr>
						<th class="td-align smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;"><br>Acct</th>
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;"><br><%=GetTerm("Customer")%></th>
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;"><br>Region</th>

						<th class="td-align smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;"><br>Filter ID</th>
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;"><br>Description</th>
						<th class="td-align smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;"><br>Notes</th>

						
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;">Frequency<br>Type</th>
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;">Frequency<br>Time</th>
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;">Last<br>Changed</th>
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;"> Next Change<br> Date</th>
						<th class="td-align smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;"> Days Until<br> Next Change</th>

						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;"><br>Qty</th>

						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;"><br>Price</th>
						<th class="td-align smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;">Line<br>Total</th>

										
						<th class="td-align smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;">Equipment<br>Value</th>
						
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;">Service<br>Ticket #</th>
						<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;"><br>Action</th>
						<!--<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;">Additional<br>Info</th>-->
						
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




<%		

	rs.Close	
		
%>


</div>
<!-- eof row !-->


<%


Sub CalcAndDisplaySummary

	TotalCustomerInFilterProgram = 0
	TotalNumberOfFilters = 0
	
	Set cnnCustFilters = Server.CreateObject("ADODB.Connection")
	cnnCustFilters.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")

	SQL = "SELECT Count(Distinct CustID) AS CustCount,Sum(Qty) AS FiltCount  FROM FS_CustomerFilters"
	Set rs = cnnCustFilters.Execute(SQL)

	If Not rs.EOF Then
		TotalCustomerInFilterProgram = rs("CustCount")
		TotalNumberOfFilters = rs("FiltCount")
	End If		


	Set cnnFilterChanges = CreateObject("ADODB.Connection")
	cnnFilterChanges.open (Session("ClientCnnString"))
	Set rsFilterChanges = Server.CreateObject("ADODB.Recordset")
	rsFilterChanges.CursorLocation = 3

	'Execute the first one just to get a count for the array
	'SQL = "SELECT COUNT(NextChangeDate) AS MastCount FROM ("
	'SQL = SQL & "SELECT NextChangeDate "
	'SQL = SQL & "FROM "
	'SQL = SQL & "(SELECT "
	'SQL = SQL & "CASE FrequencyType WHEN 'D' THEN dateadd(d,FrequencyTime, LastChangeDateTime) "
	'SQL = SQL & "WHEN 'W' THEN dateadd(d, (FrequencyTime * 7), LastChangeDateTime) "
	'SQL = SQL & "WHEN 'M' THEN dateadd(d, (FrequencyTime * 28), LastChangeDateTime) "
	'SQL = SQL & "END AS NextChangeDate "
	'SQL = SQL & "FROM FS_CustomerFilters) AS derivedtbl_1 "
'	'SQL = SQL & "WHERE (NextChangeDate < DATEADD(m, 12, GETDATE())) AND (NextChangeDate > DATEADD(m, - 3, GETDATE())) "
	'SQL = SQL & "GROUP BY NextChangeDate )  AS derivedtbl_2"

	'Set rsFilterChanges = cnnFilterChanges.Execute(SQL)

	'If NOT rsFilterChanges.EOF Then FilterChangesCount = rsFilterChanges("MastCount") 

	'Redim FilterHeaderArray(FilterChangesCount ,1)
	
	SQL="SELECT MONTH(CASE FrequencyType WHEN 'D' THEN dateadd(d,FrequencyTime, LastChangeDateTime) "
	SQL = SQL & "WHEN 'W' THEN dateadd(d, (FrequencyTime * 7), LastChangeDateTime) "
	SQL = SQL & "WHEN 'M' THEN dateadd(d, (FrequencyTime * 28), LastChangeDateTime) "
	SQL = SQL & "END) AS MonthFromNextChangeDate,"
	SQL = SQL & "YEAR(CASE FrequencyType WHEN 'D' THEN dateadd(d,FrequencyTime, LastChangeDateTime) "
	SQL = SQL & "WHEN 'W' THEN dateadd(d, (FrequencyTime * 7), LastChangeDateTime) "
	SQL = SQL & "WHEN 'M' THEN dateadd(d, (FrequencyTime * 28), LastChangeDateTime) "
	SQL = SQL & "END) YearFromNextChangeDate,"
	SQL = SQL & "COUNT(*) AS Qty"
	SQL = SQL & " FROM FS_CustomerFilters "
	SQL = SQL & " GROUP BY MONTH(CASE FrequencyType WHEN 'D' THEN dateadd(d,FrequencyTime, LastChangeDateTime) "
	SQL = SQL & " WHEN 'W' THEN dateadd(d, (FrequencyTime * 7), LastChangeDateTime) "
	SQL = SQL & " WHEN 'M' THEN dateadd(d, (FrequencyTime * 28), LastChangeDateTime) END),"
	SQL = SQL & " YEAR(CASE FrequencyType WHEN 'D' THEN dateadd(d,FrequencyTime, LastChangeDateTime) "
	SQL = SQL & " WHEN 'W' THEN dateadd(d, (FrequencyTime * 7), LastChangeDateTime)" 
	SQL = SQL & " WHEN 'M' THEN dateadd(d, (FrequencyTime * 28), LastChangeDateTime) END)"
	SQL = SQL & " ORDER BY 2,1"
	
	Set rsFilterChanges = cnnFilterChanges.Execute(SQL)
	FilterChangesCount=0
	
	If not rsFilterChanges.Eof Then
'		
'		ArrCountElement = 0			
'
		Do While Not rsFilterChanges.EOF
			FilterChangesCount=FilterChangesCount+1
'			FilterHeaderArray(ArrCountElement ,0) = rsFilterChanges("MonthFromNextChangeDate") & "/" & rsFilterChanges("YearFromNextChangeDate")
'			FilterHeaderArray(ArrCountElement ,1) = rsFilterChanges("Qty")
'					
'			ArrCountElement = ArrCountElement + 1	
'					
			rsFilterChanges.movenext
		Loop
		Redim FilterHeaderArray(FilterChangesCount ,1)
		rsFilterChanges.MoveFirst
		Do While Not rsFilterChanges.EOF
			
			FilterHeaderArray(ArrCountElement ,0) = rsFilterChanges("MonthFromNextChangeDate") & "/" & rsFilterChanges("YearFromNextChangeDate")
			FilterHeaderArray(ArrCountElement ,1) = rsFilterChanges("Qty")
			ArrCountElement = ArrCountElement + 1
			rsFilterChanges.movenext
		Loop
	End If
			
			rsFilterChanges.close
			Set rsFilterChanges= Nothing
			cnnFilterChanges.Close
			Set cnnFilterChanges = Nothing
				

			' Now actually display all the header info
			' Will use a maximum of x cols & as many
			' rows as are needed		
					
			WayToDisplay = 2
			
			If WayToDisplay = 2 Then MaxColumns = 7 Else MaxColumns = 5
			MaxRows = ubound(FilterHeaderArray) / MaxColumns 
			
			If MaxRows mod MaxColumns  <> 0 Then
				MaxRows = int(MaxRows)
				Remainder = 1
			Else
				Remainder = 0
			End If

%>
<div class='table-responsive table-top'>
	<table class='table table-condensed'>
		<tbody>
		<tr>
		
			<!----- BOX 1 ----->
			<td width="16%">
				<div class="table-striped table-condensed table-hover account-info-table inner-table">
					<table class="table table-striped table-condensed table-hover">
						<tbody>
							<tr>
							
								<td>Total <%= GetTerm("customers")%>&nbsp;&nbsp;&nbsp;<strong><%= TotalCustomerInFilterProgram %></strong></td>
							</tr>
							<tr>
								<td>Total filter changes&nbsp;&nbsp;&nbsp;<strong><%= TotalNumberOfFilters %></strong></td>
							</tr>
						</tbody>
					</table>
				</div>
			</td>
			<!----- END BOX 1 ----->

<% For CurrentCol = 1 to MaxColumns -1 %>

			<!----- BOX 2 ----->
			<td width="<%= 84 / MaxColumns%>%">
				<div class="table-striped table-condensed table-hover account-info-table inner-table">
					<table class="table table-striped table-condensed table-hover">
						<tbody>
							<%
							If CurrentCol = MaxColumns - 1 Then
								EndVal = MaxRows - Remainder
							Else
								EndVal = MaxRows 
							End If
							
							For z = 0 to EndVal %>
								<tr>
									<% 
									If WayToDisplay = 1 Then
										If ArrayMarker <= ubound(FilterHeaderArray) Then %>
											<td width="70%"><small>Filter changes due on </small><strong><%= padDate(MONTH(FilterHeaderArray(ArrayMarker,0)),2) & "/" & padDate(DAY(FilterHeaderArray(ArrayMarker,0)),2) & "/" & padDate(RIGHT(YEAR(FilterHeaderArray(ArrayMarker,0)),2),2)%></strong></td>
											<td width="30%" align="right"><strong><%= FilterHeaderArray(ArrayMarker,1)%></strong></td>
										<% Else %>
												<td width="70%">&nbsp;</td>
											<td width="30%">&nbsp;</td>
										<% End If 
									Else
										If ArrayMarker <= ubound(FilterHeaderArray) Then %>
										<!--
										<td align="center"><strong><%= padDate(MONTH(FilterHeaderArray(ArrayMarker,0)),2) & "/" & padDate(DAY(FilterHeaderArray(ArrayMarker,0)),2) & "/" & padDate(RIGHT(YEAR(FilterHeaderArray(ArrayMarker,0)),2),2)%></strong>
										&nbsp;&nbsp;-&nbsp;&nbsp;<strong><%= FilterHeaderArray(ArrayMarker,1)%></strong></td>
										-->
										<td align="center"><strong><%= FilterHeaderArray(ArrayMarker,0)%></strong>
										&nbsp;&nbsp;-&nbsp;&nbsp;<strong><%= FilterHeaderArray(ArrayMarker,1)%></strong></td>
										<% Else %>
											<td>&nbsp;</td>
										<% End If
									End If %>
								</tr>
								<%
								ArrayMarker = ArrayMarker +1
							Next %>
						</tbody>
					</table>
				</div>
			</td>
			<!----- END BOX 2 ----->

<% Next %>


				
 					 			
			</tr>
		</tbody>
	</table>
</div>

<% End Sub 
%><!-- ************************************************************************** --><!-- MODALS FOR EDITING CATEGORY NOTES, MEMOS AND EQUIPMENT                     --><!-- ************************************************************************** --><!--#include file="./CustomerFilters_Modals.asp"--><!-- ************************************************************************** --><!-- ************************************************************************** --><!--#include file="../../../inc/footer-main.asp"-->