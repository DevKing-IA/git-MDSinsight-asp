<!--#include file="../../../inc/header.asp"-->
<!--#include file="../../../inc/jquery_table_search.asp"-->
<!--#include file="../../../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../../../inc/InSightFuncs_Equipment.asp"--> 
<!--#include file="../../../inc/InsightFuncs_AR_AP.asp"-->
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

If ShowAllCusts = "" Then 

	ShowAllCusts = Request.Form("chkShowAllCusts")
	
	If (ShowAllCusts <> "" AND ShowAllCusts = "on") Then 
		ShowAllCusts = 1 
	Else 
		ShowAllCusts = 0
	End If
End If


If HideNoActionNeeded = "" Then
	HideNoActionNeeded = Request.Form("chkHideNoActionNeeded")
	if (HideNoActionNeeded <> "" AND  HideNoActionNeeded = "on") Then 
		HideNoActionNeeded = 1
	Else
		HideNoActionNeeded = 0
	End IF
	If Request.Form("frmMCS_Report1submitted") = "" Then 
		HideNoActionNeeded = 1
	End If
End IF

If IncludeDeficitCovered = "" Then
	IncludeDeficitCovered = Request.Form("chkIncludeDeficitCovered")
	If (IncludeDeficitCovered <> "" AND  IncludeDeficitCovered = "on") Then 
		IncludeDeficitCovered = 1
	Else
		IncludeDeficitCovered = 0
	End IF
	If Request.Form("frmMCS_Report1submitted") = "" Then 
		IncludeDeficitCovered = 0
	End If
End IF

If ApplyRule = "" Then
	ApplyRule = Request.Form("chkApplyRule")
	If (ApplyRule <> "" AND  ApplyRule = "on") Then 
		ApplyRule = 1
	Else
		ApplyRule = 0
	End IF
	If Request.Form("frmMCS_Report1submitted") = "" Then 
		ApplyRule = 1
	End If
End IF


If ShowZeroSalesCusts = "" Then 

	ShowZeroSalesCusts = Request.Form("chkShowZeroSalesCusts")
	
	If (ShowZeroSalesCusts <> "" AND ShowZeroSalesCusts = "on") Then 
		ShowZeroSalesCusts = 1 
	Else 
		ShowZeroSalesCusts = 0
	End If
End If

If ShowZeroSalesCusts = 1 Then 
	HideNoActionNeeded = 0
	ShowAllCusts = 0
End If

Server.ScriptTimeout = 900000 'Default value



Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

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
Set rs= cnn8.Execute(SQL)
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

<% 
'See if the data file needs to be updated
'I.e. If it has been rebuilt yet today
On error goto 0
NeedToRebuild = True

SQLMCSData = "SELECT Max(RecordCreationDateTime) AS LastBuilt FROM BI_MCSData"

Set cnnMCSData = Server.CreateObject("ADODB.Connection")
cnnMCSData.open (Session("ClientCnnString"))
Set rsMCSData = Server.CreateObject("ADODB.Recordset")
Set rsMCSData= cnnMCSData.Execute(SQLMCSData)

'Rec does not exist yet, make it quick but empty, update it later
If rsMCSData.EOF Then
	NeedToRebuild = True
Else
	If Day(rsMCSData("LastBuilt")) <> Day(Now()) Then
		NeedToRebuild = True
	End IF
End If

cnnMCSData.Close
Set rsMCSData = Nothing
Set cnnMCSData = Nothing

If MUV_READ("MCSFLAG") = "1" Then NeedToRebuild = True

NeedToRebuild = True
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
	
	.table-top .table > tbody > tr > td,
	.table > tbody > tr > th, 
	.table > tfoot > tr > td, 
	.table > tfoot > tr > th, 
	.table > thead > tr > td, 
	.table > thead > tr > th{
		border: 1px solid #ddd !important;
	}	

	.table-top2	{
		
	}	

	.table-top2 .table > tbody > tr > td
	{
		border: 1px solid #ddd !important;
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

	#tableSuperSumClients .hidden {
   		position: absolute !important;
   		top: -9999px !important;
   		left: -9999px !important;
	    /*display:none; */
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
	
	.page-header {
	    padding-bottom: 9px;
	    margin: 40px 0 10px;;
	    border-style:none;
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

	.generate-pdf{
		display: inline-block;
		padding: 10px 15px 10px 15px;
		background: #5bc0de;
		color: #fff;
		cursor: pointer;
		border: 0px;
		border-radius:5px;
		font-size: 14px;
		float: right;
    	margin-bottom: -20px;
	}
	
	.generate-pdf:hover{
		opacity:0.8;
	}
		
</style>

<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />

<style>
	/* these styles change the datatable search box */
	/* they must go AFTER the datatable CSS file */
	.dataTables_filter {
		float:left !important;
		text-align:left !important;
	}
	.dataTables_filter input {
	    border-radius: 7px;
	    border: 1px solid #B1B1B1;
	    padding: 15px; 
	    width: 325px;
	    height: 20px;   		
	}
	/* chain filter style */
	td.details-control {
		background: url(/img/details_open.png) no-repeat center center;
    cursor: pointer;
	}
	tr.shown td.details-control {
		background: url(/img/details_close.png) no-repeat center center;
	}
	tr.fold td{
		padding-left: 0px;
    	padding-right: 0px;
	}
		
</style>

<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.18/sorting/currency.js"></script>
<script type="text/javascript">


$(document).ready(function() {

	$(function(){
		$('.fold-table tr.view').on('click', 'td.details-control', function(){
			$(this).parent().toggleClass("shown").next(".fold").toggleClass("open");
			// $('table.display1').resize();
		});
	});
	
    $("#PleaseWaitPanel").hide();
	
    $("#AddedCustID").val("");
	
    $("[rel='tooltip']").tooltip('destroy');
	$("[rel='tooltip']").tooltip({ placement: 'left' });
	

    //$('[data-tooltip="tooltip"]').tooltip();
    
	$('#tableSuperSum').DataTable({
	        scrollY: 1200,
	        scrollCollapse: true,
	        paging: false,
			//order: [ 11, 'asc' ]
	        order: [[ 13, 'asc' ]],
			columnDefs: [
					{ targets: [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18, 19,20,21,22,23], "orderable": false,},
			        { type: 'currency', targets: 13}
			    ]	        
	    }
	);

	// $('table.display1').DataTable({
	//         scrollY: 1500,
	//         scrollCollapse: true,
	// 		sDom: 'lrtip',
	//         paging: false,
	// 		//order: [ 11, 'asc' ]
	//         order: [[ 1, 'asc' ]],
	// 		columnDefs: [
	// 				{ targets: [0,18, 19], "orderable": false,},
	// 		        { type: 'currency', targets: 13},

	// 				{ width: '65px', 'targets': [0] },
    //                 { width: '75px', 'targets': [1] },
    //                 { width: '63px', 'targets': [2] },
    //                 { width: '75px', 'targets': [3] },
    //                 { width: '94px', 'targets': [4] },
    //                 { width: '64px', 'targets': [5] },
    //                 { width: '57px', 'targets': [6] },
    //                 { width: '55px', 'targets': [7] },
    //                 { width: '55px', 'targets': [8] },
    //                 { width: '57px', 'targets': [9] },
    //                 { width: '84px', 'targets': [10] },
    //                 { width: '61px', 'targets': [11] },
    //                 { width: '82px', 'targets': [12] },
    //                 { width: '57px', 'targets': [13] },
    //                 { width: '83px', 'targets': [14] },
    //                 { width: '83px', 'targets': [15] },
    //                 { width: '96px', 'targets': [16] },
    //                 { width: '81px', 'targets': [17] },
    //                 { width: '77px', 'targets': [18] },
	// 				{ width: '84px', 'targets': [19] },
    //                 { width: '51px', 'targets': [20] },
    //                 { width: '116px', 'targets': [21] },
    //                 { width: '68px', 'targets': [22] },
    //                 { width: '95px', 'targets': [23] }
                                       
                    
	// 		    ]	        
	//     }
	// );
	

	$('#tableSuperSumClients').DataTable({
	        scrollY: 500,
	        scrollCollapse: true,
	        paging: false,
			//order: [ 11, 'asc' ]
	        order: [[ 0, 'asc' ]],
			columnDefs: [
					{ targets: [18, 19], "orderable": false,},
			        { type: 'currency', targets: 13}
			    ]	        
	    }
	);
	
	$("#tableSuperSum_info").hide();

	$('#tableSuperSumPendingCharges').DataTable({
	        scrollY: 500,
	        scrollCollapse: true,
	        paging: false,
	        order: [[ 4, 'asc' ],[ 0, 'asc' ]],
			columnDefs: [
			        { type: 'currency', targets: 4}
			    ]	        
	    }
	);
	
	$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
	
	  var target = $(e.target).attr("href") // activated tab
	  //alert(target);
	  
	  if (target == "#mcs") {
	  	$("#tableSuperSum").resize();
		  $('table.display1').resize();

	  }

	  if (target == "#clients") {
	  	$("#tableSuperSumClients").resize();
	  }

	  if (target == "#pendingcharges") {
	  	$("#tableSuperSumPendingCharges").resize();
	  }
	  
	}); 	
	
	$("#chkShowAllCusts").change(function() {
		$("#frmMCS_Report1").submit();
	});
	
	$("#chkIncludeDeficitCovered").change(function() {
		$("#frmMCS_Report1").submit();
	});
	
	$("#chkApplyRule").change(function() {
		$("#frmMCS_Report1").submit();
	});

	$("#chkShowZeroSalesCusts").change(function() {
		$("#frmMCS_Report1").submit();
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
			url: "../../../inc/InSightFuncs_AjaxForARAP.asp",
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
			url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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
			url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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
		Set rsReason = cnn8.Execute(SQL)
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
			url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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
				url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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
			url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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
			url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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
			url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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
		$("#AddMCSClient2custid").text(CustID);
		$("#AddMCSClient2custname").text(CustName);
		$("#AddMCSClient2CustIDToPass").val(CustID);
		$("#modalAddMCSClient2").modal('toggle');
	})

	$('#modalAddMCSClient2').on('show.bs.modal', function(j) {

		var $modal = $(this);
	    //get data-id attribute of the clicked order
		var CustID = $("#AddMCSClientCustIDToPass").val();
		var CustName = $("#AddMCSClientCustNameToPass").val();		

		$("#AddMCSClient2custid").text(CustID);
		$("#AddMCSClient2custname").text(CustName);
		$("#AddMCSClient2MonthlyContractedSalesDollars").val("");
		$("#AddMCSClient2MaxMCSCharge").val("");
		
	});

	$('#modalAddMCSClient2').on('hidden.bs.modal', function(j) {
			if ($('#AddMCSClient2MonthlyContractedSalesDollars').next().prop("tagName") == 'BR') {
				$('#AddMCSClient2MonthlyContractedSalesDollars').next('br').remove();
				$('#AddMCSClient2MonthlyContractedSalesDollars').next('span').remove();
			}
			if ($('#AddMCSClient2MaxMCSCharge').next().prop("tagName") == 'BR') {
				$('#AddMCSClient2MaxMCSCharge').next('br').remove();
				$('#AddMCSClient2MaxMCSCharge').next('span').remove();
			}			
	});	
	
	$("#AddMCSClient2Save").click(function(){	
		var CustID = $("#AddMCSClient2CustIDToPass").val();
		var MCSDollars = $("#AddMCSClient2MonthlyContractedSalesDollars").val();
		var MaxMCSCharge = $("#AddMCSClient2MaxMCSCharge").val();
		var adderror = 0;
		
		var numberReg =  /^[0-9\.]+$/;
		if(!numberReg.test(MCSDollars)){
            $('#AddMCSClient2MonthlyContractedSalesDollars').after('<BR><span class="error"> Please Enter Numbers only</span>');
			adderror = 1;
        }
		//if(!numberReg.test(MaxMCSCharge)){
        //    $('#AddMCSClient2MaxMCSCharge').after('<BR><span class="error">  Please Enter Numbers only</span>');
		//	adderror = 1;
        //}
		if (adderror) {
			return;
		}
    	$.ajax({
			type:"POST",
			url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
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

		$("#modalAddMCSClient2").modal('toggle');
    });

	$("#chkHideNoActionNeeded").click(function(){		
		$(".noaction").each(function(i, obj) {			
			if (obj.hasAttribute('data-cust-id')) {
				var CustID = obj.getAttribute("data-cust-id");
				if($("#chkHideNoActionNeeded").prop('checked')) {					
					//$("#CUST"+CustID).hide();
				} else {					
					$("#CUST"+CustID).show();
				}
			}
		});	
	});

	$('.noaction').each(function(i, obj) {
		if (obj.hasAttribute('data-cust-id')) {
			var CustID = obj.getAttribute("data-cust-id");
			if($("#chkHideNoActionNeeded").prop('checked')) {					
				//$("#CUST"+CustID).hide();
			} else {					
				$("#CUST"+CustID).show();
			}
		}
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
	
	
	
	
	$("#btnGeneratePDFPendingCharges").click(function(){	
	
		//alert("Button clicked to generate pdf");
	
		$.ajax({
			type:"POST",
			url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
			data: "action=GenerateMCSPendingChargesPDF",
			beforeSend: function() {
				$('#PleaseWaitPanel').show();		     	
			},
			complete: function(){
				$('#PleaseWaitPanel').hide();
			},
			success: function(response)
			 {																		
				//Upon success PDF will open in a new browser tab
				window.open(response,'_blank');
					
			 },
			 failure: function(response)
			 {
				swal({
					title: 'Error Generating MCS Pending Charges PDF',
					text: response,
					type: 'error'
				});

			 }
		});	
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


</script>


<%
If NeedToRebuild = True Then
	Response.Write("<div id=""PleaseWaitPanel"" class=""container"">")
	Response.Write("<br><br>Updating MCS with the latest data <br><br>Please wait...<br><br>")
	Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
	Response.Write("</div>")
	Response.Flush()
Else
	Response.Write("<div id=""PleaseWaitPanel"" class=""container"">")
	Response.Write("<br><br>Running MCS Analysis <br><br>Please wait...<br><br>")
	Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
	Response.Write("</div>")
	Response.Flush()
End If



%>

<h3 class="page-header"><i class="fa fa-graduation-cap"></i><%=MonthName(Month(DateAdd("m",-1,ReportDate))) %>&nbsp;<%= Year(DateAdd("m",-1,ReportDate))%>&nbsp;MCS Analysis
&nbsp;&nbsp;
<small><button type="button" class="btn btn-success" id="btnAddClientToMCS" data-toggle="modal" data-target="#modalAddMCSClient"><font size="2px" >Add Client To MCS</font></button></small>
</h3>

 

	
<form method="POST" name="frmMCS_Report1" id="frmMCS_Report1" value="frmMCS_Report1" action="MCS_Report1.asp">	
	<input id="frmMCS_Report1submitted" name="frmMCS_Report1submitted" type="hidden" value="1">
	<input id="RemovedCustID" name="RemovedCustID" type="hidden" value="">

<%

	 If NeedToRebuild = True Then
		Call RebuildMCSData("")
		NeedToRebuild = False
	End If
%>

	<% Call CalcAndDisplaySummary %>

</form>	    


<ul class="nav nav-tabs">
	<li class="active">
        <a href="#clients" data-toggle="tab">Clients</a>
	</li>
	<li>
        <a href="#mcs" data-toggle="tab">Chains</a>
	</li>
	<li>
		<a href="#pendingcharges" data-toggle="tab">Pending Charges</a>
	</li>
</ul>

<div class="tab-content">
	<div class="tab-pane active" id="clients">
		<!--#include file="MCS_Report1_Tab_MCS_client.asp"-->
	</div>
	<div class="tab-pane" id="mcs">
		<!--#include file="MCS_Report1_Tab_MCS.asp"-->
	</div>
	<div class="tab-pane" id="pendingcharges">
		<!--#include file="MCS_Report1_Tab_Pending_Charges.asp"-->
	</div>
</div>


<!-- row !-->
<div class="row">
<%		
	rs.Close		
%>
</div>
<!-- eof row !-->


<%


Sub CalcAndDisplaySummary

	TotalApplyRuleCount = 0
	TotalSalesAllMCSCustomers = 0
	TotalPendingLVF = 0
	
	SQL = "SELECT * FROM AR_Customer INNER JOIN BI_MCSData ON BI_MCSData.CustID = AR_Customer.CustNum WHERE MonthlyContractedSalesDollars > 0" 
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	Set rs = cnn8.Execute(SQL)

	If Not rs.Eof Then
	
		Do While Not rs.EOF

			ShowThisRecord = True

				
			If ShowThisRecord <> False Then			
			
				PrimarySalesMan =  ""
				SecondarySalesMan =  ""
				SelectedCustomerID = rs("CustNum")
				CustName = rs("Name")

				PrimarySalesMan = rs("Salesman")
				SecondarySalesMan = rs("SecondarySalesman")
				CustMonthlyContractedSalesDollars = rs("MonthlyContractedSalesDollars")
					
				'Decide if this record meets the filter criteria
				If FilterSlsmn1 <> "" And FilterSlsmn1 <> "All" Then
					If CInt(FilterSlsmn1) <> Cint(rs("Salesman")) Then ShowThisRecord = False
				End If
				If FilterSlsmn2 <> "" And FilterSlsmn2 <> "All" Then
					If CInt(FilterSlsmn2) <> Cint(rs("SecondarySalesman")) Then ShowThisRecord = False
				End If
		
			End If
			

			Month3Sales_NoRent = rs("Month3Sales_NoRent") - rs("Month3Cat21Sales") 
				
			If ShowAllCusts <> 1 Then
				If Month3Sales_NoRent >= rs("MonthlyContractedSalesDollars") Then ShowThisRecord = False
			End If

			TotalSalesAllMCSCustomers = TotalSalesAllMCSCustomers + Month3Sales_NoRent
			TotalMCSClients = TotalMCSClients + 1
			TotalMCSCommitment = TotalMCSCommitment + rs("MonthlyContractedSalesDollars")
			
			TotalLVFLastMonth = TotalLVFLastMonth + rs("LVFHolder")
			
			If Month3Sales_NoRent >= rs("MonthlyContractedSalesDollars") Then
				TotalCustomersOver = TotalCustomersOver + 1
				TotalOverDollars = TotalOverDollars + (Month3Sales_NoRent - rs("MonthlyContractedSalesDollars"))
			End If
			
			'If Month3Sales_NoRent < rs("MonthlyContractedSalesDollars") Then 'And Month3Sales_NoRent <> 0 Then
				'TotalCustomersUnder = TotalCustomersUnder + 1
				'TotalUnderDollars = TotalUnderDollars + (rs("MonthlyContractedSalesDollars") - Month3Sales_NoRent)
			'End If

			RuleInEffect = False
			
			' Calc under by the current month recovered the deficit
			

			
			If Month3Sales_NoRent < rs("MonthlyContractedSalesDollars") Then 
			
				Month3LVF = TotalPostedLVFByCustByMonthByYear(rs("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
				If Not IsNumeric(Month3LVF ) Then Month3LVF = 0
				If Month3LVF < 1 Then Month3LVF = 0
				M3Stemp = rs("Month3Sales_NoRent") - (rs("Month3Cat21Sales") + Month3LVF)
				If rs("CurrentHolder") >= rs("MonthlyContractedSalesDollars") +  ABS((M3Stemp - rs("MonthlyContractedSalesDollars"))) Then
					TotalCustomersUnderButRecovered = TotalCustomersUnderButRecovered + 1
				Else
						 If ABS(rs("Month3Sales_NoRent") - rs("MonthlyContractedSalesDollars")) < 100 Then ' Variance
							If rs("Month3Sales_NoRent") <> 0 Then
								VariancePercentHolder = 100 - ((rs("Month3Sales_NoRent")/rs("MonthlyContractedSalesDollars")) * 100)
							Else
								VariancePercentHolder = 100 
							End If
							VariancePercentHolder  = VariancePercentHolder  * -1
							If ApplyRule = 1 Then
							If ABS(VariancePercentHolder) < 10 Then 
								RuleInEffect = True
								TotalApplyRuleCount = TotalApplyRuleCount +1
								' Need to negate the fact that it will added in below, at the end if this if
								'TotalCustomersUnder = TotalCustomersUnder - 1
								'TotalUnderDollars = TotalUnderDollars - (rs("MonthlyContractedSalesDollars") - Month3Sales_NoRent)
							End If
						End If
					End If
				End If
									
				If Month3Sales_NoRent > 0 Then
					TotalCustomersUnder = TotalCustomersUnder + 1
					TotalUnderDollars = TotalUnderDollars + (rs("MonthlyContractedSalesDollars") - Month3Sales_NoRent)
				End If

			End If

			
			
			If Month3Sales_NoRent <= 0 Then
				TotalCustomersZeroSales = TotalCustomersZeroSales + 1
				TotalZeroSalesCommitment = TotalZeroSalesCommitment + rs("MonthlyContractedSalesDollars")
			End If

			
			If ShowThisRecord <> False Then
			
				TotalMonth3Sales = TotalMonth3Sales + Month3Sales_NoRent
				TotalVariance = TotalVariance + Month3Sales_NoRent - rs("MonthlyContractedSalesDollars")
				If Month3Sales_NoRent = 0 Then TotalClientWithZeroSales = TotalClientWithZeroSales + 1
				
				   
		    End If
		    
		    TotalPendingLVF = TotalPendingLVF + rs("PendingLVF")

			rs.movenext
				
		Loop
		
End If

SQLFormat = "SELECT * from Settings_BizIntel"
	
Set rsFormat = Server.CreateObject("ADODB.Recordset")
Set rsFormat = cnn8.Execute(SQLFormat )

MCSUseAlternateHeader = rsFormat("MCSUseAlternateHeader")

rsFormat.Close
Set rsFormat = Nothing

If MCSUseAlternateHeader = 1 Then%><!--#include file="./mcsheader_alternate.asp"--><% Else %><!--#include file="./mcsheader_standard.asp"--><%End If

End Sub 

Sub RebuildMCSData (passedCustID)

	' If passedCustID = "" then it will do the rebuild for all customers, otherwise, just the onde

	dummy = MUV_WRITE("MCSFLAG","0")

	Set cnnMCSData = Server.CreateObject("ADODB.Connection")
	cnnMCSData.open (Session("ClientCnnString"))
	Set rsMCSData = Server.CreateObject("ADODB.Recordset")
	Set rsMCSDataForUpdating = Server.CreateObject("ADODB.Recordset")

	If passedCustID = "" Then
		SQLMCSData = "DELETE FROM BI_MCSData"
	Else
		SQLMCSData = "DELETE FROM BI_MCSData WHERE CustID = '" & passedCustID & "'"
	End If
	Set rsMCSData= cnnMCSData.Execute(SQLMCSData)
	
	If passedCustID = "" Then
		SQLMCSData = "INSERT INTO BI_MCSData (CustID) SELECT CustNum FROM AR_Customer WHERE MonthlyContractedSalesDollars <> 0 AND AcctStatus='A'" 
	Else
		SQLMCSData = "INSERT INTO BI_MCSData (CustID) SELECT CustNum FROM AR_Customer WHERE MonthlyContractedSalesDollars <> 0 AND AR_Customer.CustNum = '"  & passedCustID & "' AND AcctStatus='A'"
	End If
	Set rsMCSData= cnnMCSData.Execute(SQLMCSData)
	
	'Now begin with all the aggregate numbers
	SQLMCSData = "SELECT * FROM BI_MCSData"
	Set rsMCSData= cnnMCSData.Execute(SQLMCSData)
	
	If NOT rsMCSData.EOF Then
		Do While Not rsMCSData.EOF
			
			Month1Sales_NoRent = 0
			Month2Sales_NoRent = 0
			Month3Sales_NoRent = 0
			Month3Cost_NoRent = 0
			LVFHolder = 0
			LVFHolderCurrent = 0
			TotalEquipmentValue = 0
			CurrentHolder = 0
			RentalHolder = 0
			PendingLVF = 0
			
			PendingLVF = cdbl(PendingLVFByCust(rsMCSData("CustID")))
			
			
			Month3Cost_NoRent = TotalCostByCustByMonthByYear_NoRent(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
			If NOT IsNumeric(Month3Cost_NoRent) Then Month3Cost_NoRent = 0
			Month1Sales_NoRent = TotalSalesByCustByMonthByYear_NoRentals(rsMCSData("CustID"),Month(DateAdd("m",-3,ReportDate)),Year(DateAdd("m",-3,ReportDate)))
			Month2Sales_NoRent = TotalSalesByCustByMonthByYear_NoRentals(rsMCSData("CustID"),Month(DateAdd("m",-2,ReportDate)),Year(DateAdd("m",-2,ReportDate)))
			Month3Sales_NoRent = TotalSalesByCustByMonthByYear_NoRentals(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))

			' Remove LVF from Monthly sales
			Month1LVF = TotalPostedLVFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-3,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
			Month2LVF = TotalPostedLVFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-2,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
			Month3LVF = TotalPostedLVFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))

			If Month1LVF > 0 Then Month1Sales_NoRent = Month1Sales_NoRent - Month1LVF 
			If Month2LVF > 0 Then Month2Sales_NoRent = Month2Sales_NoRent - Month2LVF 
			If Month3LVF > 0 Then Month3Sales_NoRent = Month3Sales_NoRent - Month3LVF 
			
			Month1XSF = TotalXSFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-3,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
			Month2XSF = TotalXSFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-2,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
			Month3XSF = TotalXSFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))

			If Month1XSF > 0 Then Month1Sales_NoRent = Month1Sales_NoRent - Month1XSF 
			If Month2XSF > 0 Then Month2Sales_NoRent = Month2Sales_NoRent - Month2XSF 
			If Month3XSF > 0 Then Month3Sales_NoRent = Month3Sales_NoRent - Month3XSF 
				
			CurrentXSF = 0	
			CurrentXSF = TotalXSFByCustByMonthByYear(rsMCSData("CustID"),Month(ReportDate),Year(ReportDate))	
			

			LVFHolder = TotalPostedLVFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
			
			
			LVFHolderCurrent = TotalPostedLVFByCustByMonthByYear(rsMCSData("CustID"),Month(ReportDate),Year(ReportDate))
			TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(rsMCSData("CustID"))
			
			' Must subtract any rentals from Current moth$
			CurrentHolder = GetCurrent_PostedTotal_ByCust(rsMCSData("CustID"),PeriodSeqBeingEvaluated) + GetCurrent_UnPostedTotal_ByCust(rsMCSData("CustID"),PeriodSeqBeingEvaluated)
			CurrentRent = TotalSalesByCustByMonthByYear_RentalsOnly(rsMCSData("CustID"),Month(ReportDate),Year(ReportDate))
			CurrentHolder = CurrentHolder - CurrentRent
			
			
			RentalHolder = TotalSalesByCustByMonthByYear_RentalsOnly(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
			'If Month3XSF <> 0 Then RentalHolder = RentalHolder +  Month3XSF 

			Month1_Cat21Holder = TotalCat21ByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-3,ReportDate)),Year(DateAdd("m",-1,ReportDate)))			
			Month2_Cat21Holder = TotalCat21ByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-2,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
			Month3_Cat21Holder = TotalCat21ByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
			
			
			'Got em all, update the record
			SQLMCSDataForUpdating = "UPDATE BI_MCSData SET "
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & "Month1Sales_NoRent = " & Month1Sales_NoRent
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month2Sales_NoRent = " & Month2Sales_NoRent
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month3Sales_NoRent = " & Month3Sales_NoRent						
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month3Cost_NoRent = " & Month3Cost_NoRent
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", LVFHolder = " & LVFHolder 
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", LVFHolderCurrent = " & LVFHolderCurrent 
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", TotalEquipmentValue = " & TotalEquipmentValue 
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", CurrentHolder = " & CurrentHolder 	
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", RentalHolder = " & RentalHolder 
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month1Cat21Sales = " & Month1_Cat21Holder 
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month2Cat21Sales = " & Month2_Cat21Holder 
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month3Cat21Sales = " & Month3_Cat21Holder 			
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month1XSF = " & Month1XSF 
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month2XSF = " & Month2XSF 
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month3XSF = " & Month3XSF 
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", CurrentXSF = " & CurrentXSF 
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", PendingLVF = " & PendingLVF 
			If GetCustChainIDByCustID(rsMCSData("CustID")) <> "" Then
				SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", ChainID = '" & GetCustChainIDByCustID(rsMCSData("CustID")) & "' "
			End If	
			If GetChainDescByChainNum(GetCustChainIDByCustID(rsMCSData("CustID"))) <> "" Then 
				SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", ChainName = '" & GetChainDescByChainNum(GetCustChainIDByCustID(rsMCSData("CustID"))) & "' "
			End If
			SQLMCSDataForUpdating = SQLMCSDataForUpdating  & " WHERE CustID = " & rsMCSData("CustID")
			
			'Response.Write(SQLMCSDataForUpdating & "<br>")
			
			Set rsMCSDataForUpdating = cnnMCSData.Execute(SQLMCSDataForUpdating)
		
			rsMCSData.MoveNext
		Loop



	
	End If


cnnMCSData.Close
Set rsMCSData = Nothing
Set cnnMCSData = Nothing

End Sub
%>
<!-- ************************************************************************** -->
<!-- MODALS FOR EDITING CATEGORY NOTES, MEMOS AND EQUIPMENT                     -->
<!-- ************************************************************************** -->
<!--#include file="../CatAnalByPeriod/CatAnalByPeriod_Modals.asp"-->
<!-- ************************************************************************** -->
<!-- ************************************************************************** -->
<!--#include file="../../../inc/footer-main.asp"-->