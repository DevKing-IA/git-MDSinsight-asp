<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"--> 
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.asp"-->
<%

Server.ScriptTimeout = 900000 'Default value

StartDate = Request.Form("txtStartDate")
EndDate = Request.Form("txtEndDate")

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


	#PleaseWaitPanelExport{
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
	
	.inventory-header{
		background: #F0AD4E;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}
	
	.cost-header{
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
	
	.actions-header{
		background: #9C02FE;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.export-header{
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

	table.dataTable.compact thead th, table.dataTable.compact thead td {
	    padding: 4px 17px 4px 4px;
	    text-align: center;
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
	    margin: 40px 10px 40px 0px;
	    border-style:none;
	}
	
	.dataTables_wrapper .dataTables_filter input {
	    display: inline-block;
	    width:400px;
	    height: 34px;
	    padding: 6px 12px;
	    font-size: 14px;
	    line-height: 1.42857143;
	    color: #555;
	    background-color: #fff;
	    background-image: none;
	    border: 1px solid #ccc;
	    border-radius: 4px;
	    margin-left: 0.5em;
	    margin-right: 1em;
	    margin-bottom: 1em;
	    -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
	    box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
	    -webkit-transition: border-color ease-in-out .15s,-webkit-box-shadow ease-in-out .15s;
	    -o-transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;
	    transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;
	}	

	.dataTables_filter {margin-top: -45px;}
	
	td.details-control {
	    background: url('../../img/accordion/details_open.png') no-repeat center center;
	    cursor: pointer;
		border-left: 2px solid #555 !important; 
		border-right: 0px !important;
	}
	
	tr.shown td.details-control {
	    background: url('../../img/accordion/details_close.png') no-repeat center center;
	}	

	table.dataTable tbody tr:hover  {
	    background-color:#fff;
	}	
	
	.invoice-title h2, .invoice-title h3, .invoice-title h4 {
	    display: inline-block;
	}
	
	.table > tbody > tr > .no-line {
	    border-top: none;
	}
	
	.table > thead > tr > .no-line {
	    border-bottom: none;
	}
	
	.table > tbody > tr > .thick-line {
	    border-top: 2px solid;
	}	

	.exportcheckbox{
	    display: block;
	    width: 100%;
	    height: 20px;
	    padding: 6px 12px !important;
	    font-size: 14px;
	    line-height: 1.42857143 !important;
	    color: #555;
	    background-color: #fff;
	    border: 1px solid #ccc; 
	    border-radius: 4px;
	}

		
</style>

<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.18/sorting/currency.js"></script>
<script type="text/javascript">


function format (name, value) {

    return '<div>Name: ' + name + '<br />Value: ' + value + '</div>';
    
    
}

function download(filename, text) {
    var element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
    element.setAttribute('download', filename);

    element.style.display = 'none';
    document.body.appendChild(element);

    element.click();

    document.body.removeChild(element);
}


$(document).ready(function() {

    $("#PleaseWaitPanel").hide();
    $("#PleaseWaitPanelExport").hide();
    
    $('#filter').keyup(function () {

	    var rex = new RegExp($(this).val(), 'i');
	    $('.searchable tr').hide();
	    $('.searchable tr').filter(function () {
	        return rex.test($(this).text());
	    }).show();

	});
	
	$('.btn-toggle').click(function() {
	    $(this).find('.btn').toggleClass('active');  
	    
	    if ($(this).find('.btn-primary').size()>0) {
	    	$(this).find('.btn').toggleClass('btn-primary');
	    }
	    if ($(this).find('.btn-danger').size()>0) {
	    	$(this).find('.btn').toggleClass('btn-danger');
	    }
	    if ($(this).find('.btn-success').size()>0) {
	    	$(this).find('.btn').toggleClass('btn-success');
	    }
	    if ($(this).find('.btn-info').size()>0) {
	    	$(this).find('.btn').toggleClass('btn-info');
	    }
	    
	    $(this).find('.btn').toggleClass('btn-default');
	       
	});
	
	$('form').submit(function(){
		//alert($(this["options"]).val());
	    //return false;
	});	
							
	$( "#btnClearDateFilters" ).click(function() {
		$('#txtStartDate').val(moment().subtract(1, 'days').format('MM/DD/YYYY'));
		$('#txtEndDate').val(moment().subtract(1, 'days').format('MM/DD/YYYY'));
		$('#frmSageInvoiceDateRange').submit();

	});
	
	$( "#btnShowExportedInvoices" ).click(function() {
    	$.ajax({
			type:"POST",
			url: "../../inc/InSightFuncs_AjaxForARAP.asp",
			cache: false,
			async: false,
			data: "action=ToggleShowHideExportedSageInvoices&ShowHide=SHOW",
			success: function(response)
			 {		
				$('#frmSageInvoiceDateRange').submit();            	 
             },
             failure: function(response)
			 {
			  	swal("Error showing/hiding exported invoices.")
             }
		});
	});
	
	$( "#btnHideExportedInvoices" ).click(function() {
    	$.ajax({
			type:"POST",
			url: "../../inc/InSightFuncs_AjaxForARAP.asp",
			cache: false,
			async: false,
			data: "action=ToggleShowHideExportedSageInvoices&ShowHide=HIDE",
			success: function(response)
			 {		
				$('#frmSageInvoiceDateRange').submit();            	 
             },
             failure: function(response)
			 {
			  	swal("Error showing/hiding exported invoices.")
             }
		});
	});
		
	$('#chkExportCheckAll').click(function () {    
	    $("input[name^='chkExport']").prop('checked', this.checked);    
	});	
	
	$("#btnExportInvoices").click(function(){
        var invoicesToExport = [];
        $.each($("input[name^='chkExport']:checked"), function(){            
            invoicesToExport.push($(this).val());
        });
        
        invoicesToExportData = invoicesToExport.join(", ");
        
        //alert("Invoices to export are: " + invoicesToExportData);
        
        $("#PleaseWaitPanelExport").show();
        
    	$.ajax({
			type:"POST",
			url: "../../inc/InSightFuncs_AjaxForARAP.asp",
			cache: false,
			data: "action=ExportSelectedSageInvoices&InvoicesToExport="+encodeURIComponent(invoicesToExportData),
			success: function(response)
			 {		
			     var text = response;
			     var d = new Date();
			     var newdate = moment(d);
				 var formattedDate = newdate.format("MM-DD-YYYY-HH:mm:ss");
			     var filename = "VendmaxToSage-" + formattedDate + ".txt"
			    
			     download(filename, text);
			     
			     $("#PleaseWaitPanelExport").hide();
    			
               	 swal("Invoices successfully exported.");            	 
             },
             failure: function(response)
			 {
			 	$("#PleaseWaitPanelExport").hide();
			  	swal("Error exporting selected invoices.")
             }
		});
        
     });	



	$('#modalEditSageInvoice').on('show.bs.modal', function(e) {

	    //get data-id attribute of the clicked order
	    var InvoiceID = $(e.relatedTarget).data('invoice-id');
	    var LineItemNo = $(e.relatedTarget).data('line-item-no');
	    
	    
	    //populate the textbox with the id of the clicked filter
	    $(e.currentTarget).find('input[name="txtInvoiceID"]').val(InvoiceID);
	    $(e.currentTarget).find('input[name="txtLineItemNo"]').val(LineItemNo);

	    var $modal = $(this);

    	$.ajax({
			type:"POST",
			url: "../../inc/InSightFuncs_AjaxForARAP.asp",
			cache: false,
			async: false,
			data: "action=GetContentForEditSageInvoiceModal&InvoiceID="+encodeURIComponent(InvoiceID)+"&LineItemNo="+encodeURIComponent(LineItemNo),
			success: function(response)
			 {					
               	 $modal.find('#modalEditSageInvoiceContent').html(response);              	 
             },
             failure: function(response)
			 {
			  	$modal.find('#modalEditSageInvoiceContent').html("Failed");
             }
		});
		
	});
	
	
	
    
	var table = $('#tableSuperSum').DataTable( {
	//$('#tableSuperSum').DataTable({
        scrollY: 500,
        scrollCollapse: true,
        paging: false,
        order: [[ 3, 'asc' ],[ 4, 'asc' ]],
		columnDefs: [
		        { targets: [0,2,10], "orderable": false,},
		        { type: 'currency', targets: [8]}
		    ]	        
	    }
	);
	
	// Add event listener for opening and closing details
    $('#tableSuperSum tbody').on('click', 'td.details-control', function () {
        var tr = $(this).closest('tr');
        var row = table.row( tr );
 
        if ( row.child.isShown() ) {
            // This row is already open - close it
            row.child.hide();
            tr.removeClass('shown');
        }
        else {
            // Open this row
            //row.child( format(row.data()) ).show();
            
            var InvoiceID = tr.data('child-value');
            
			$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				async: false,
				data: "action=GetContentForSageInvoiceDetailExpansion&InvoiceID="+encodeURIComponent(InvoiceID),
				success: function(response)
				 {					
		            row.child(response).show();
		            tr.addClass('shown');
		         },
		         failure: function(response)
				 {
				  	swal("Row Data Expand Failed");
		         }
			});
            
        }
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

<div id="PleaseWaitPanel" class="containter">
	<br><br>Gathering Invoice Data <br><br>Please wait...<br><br>
	<img src="<%= BaseURL %>img/loading.gif">
</div>

<%
	Response.Flush()
%>

<div id="PleaseWaitPanelExport" class="containter">
	<br><br>Exporting Selected Invoices <br><br>Please wait...<br><br>
	<img src="<%= BaseURL %>img/loading.gif">
</div>


<h3 class="page-header"><i class="fa fa-dollar"></i>&nbsp;Export Invoices To SAGE Accounting&nbsp;<img src="<%= BaseURL %>img/partnericons/sage-logo.png"></h3>


<!-- row !-->
<form method="POST" name="frmSageInvoiceDateRange" id="frmSageInvoiceDateRange" action="main.asp">	
	<div class="row">
	
	    <div class="col-lg-5">
		    <div class="col-box">
			    <!-- date !-->
			    <div class="row date-ranges row-line">
					<div class="col-lg-12">
						<div class="form-group">
							<input type="hidden" id="txtStartDate" name="txtStartDate" value="<%= StartDate %>">
							<input type="hidden" id="txtEndDate" name="txtEndDate" value="<%= EndDate %>">
							<strong>Select Invoice Date Range</strong>: 
							<div class="btn btn-default" id="reportrange">
								<i class="fa fa-calendar"></i> &nbsp;
								<span></span>
								<b class="fa fa-angle-down"></b>
							</div>
							<button type="submit" class="btn btn-primary"><i class="fas fa-check"></i>&nbsp;Apply Date Filter</button>
							<button type="button" class="btn btn-danger" id="btnClearDateFilters"><i class="far fa-backspace"></i>&nbsp;Clear Date Filters</button>
						</div>
				 	</div>
				</div>
			</div>
		</div>
		
	    <div class="col-lg-1">
		    <div class="col-box">
				<button type="button" class="btn btn-success" id="btnExportInvoices"><i class="fas fa-file-export"></i>&nbsp;Export Selected Invoices</button>
			</div>
		</div>
		
	    <div class="col-lg-2" style="margin-left:60px; margin-top:3px;">
		    <div class="col-box">
				<div class="btn-group btn-toggle"> 
					<% If MUV_READ("showExportedSageInvoices") = "HIDE" OR MUV_READ("showExportedSageInvoices") = "" Then %>
				    	<button class="btn btn-sm btn-default" id="btnShowExportedInvoices">SHOW</button>
				    	<button class="btn btn-sm btn-info active" id="btnHideExportedInvoices">HIDE</button>
				    <% Else %>
				    	<button class="btn btn-sm btn-info active" id="btnShowExportedInvoices">SHOW</button>
				    	<button class="btn btn-sm btn-default" id="btnHideExportedInvoices">HIDE</button>				    
					<% End If %>				    
			  	</div>
			  	<strong>Exported Invoices</strong>
  		     </div>
    	</div>
		
	</div>
</form>

<!-- row !-->
<div class="row">

<div class="container-fluid">
    <div class="row">
           <table id="tableSuperSum" class="display compact" style="width:100%;">
              <thead>
                  	<tr>	
	                  	<th class="export-header" colspan="3" style="border-right: 2px solid #555 !important;font-size: 16px;">Export</th>
                		<th class="gen-info-header" colspan="2" style="border-right: 2px solid #555 !important;font-size: 16px;">Invoice</th>
                		<th class="inventory-header" colspan="2" style="border-right: 2px solid #555 !important;font-size: 16px;">Customer</th>
						<th class="cost-header" colspan="3" style="border-right: 2px solid #555 !important;font-size: 16px;">Additional Info</th>
						<th class="actions-header" style="border-right: 2px solid #555 !important;font-size: 16px;">Actions</th>
					</tr>
				
	                <tr>
						<th class="sorttable_numeric" style="border-left:2px solid #555 !important; border-top: 2px solid #555 !important;" id="colExpandCollapse">+</th>
						
						<th class="sorttable_numeric" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="colInvoiceExportedAlready">Exported</th>
						<th class="sorttable_numeric" style="border-top: 2px solid #555 !important;" id="colInvoiceExported"><input type="checkbox" class="exportcheckbox" name="chkExportCheckAll" id="chkExportCheckAll"> <span style="font-size:8pt;">EXP ALL</span></th>
						
						<th class="sorttable_numeric" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="colInvoiceDate">Date</th>
						<th class="sorttable_numeric" style="border-top: 2px solid #555 !important; border-right: 2px solid #555 !important;" id="colInvoiceID">Invoice ID</th>
	
						<th class="sorttable_numeric" style="border-top: 2px solid #555 !important;" id="colCustID">Customer</th>
						<th class="sorttable_numeric" style="border-top: 2px solid #555 !important; border-right: 2px solid #555 !important;" id="colAltCustID">Alt Cust ID</th>
						
						<th class="sorttable_numeric" style="border-top: 2px solid #555 !important;" id="colInvoiceType">Invoice Type</th>	
						<th class="sorttable_numeric" style="border-top: 2px solid #555 !important;" id="colInvoiceTotal">Invoice Total</th>
						<th class="sorttable_numeric" style="border-top: 2px solid #555 !important; border-right: 2px solid #555 !important;" id="colNumLines"># Lines</th>
										
						<th align="center" class="sorttable_numeric" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="colAction">Actions</th>
	                </tr>
                
              </thead>
              
			
   <tbody>
   
   <%
	
	If MUV_READ("showExportedSageInvoices") = "SHOW" Then
		SQL = "SELECT * FROM IN_InvoiceHistHeader WHERE Cast(invoicecreationdate as date) >= '" & StartDate & "' AND cast(invoicecreationdate as date) <= '" & EndDate & "' "
	ElseIf MUV_READ("showExportedSageInvoices") = "HIDE" OR MUV_READ("showExportedSageInvoices") = "" Then
		SQL = "SELECT * FROM IN_InvoiceHistHeader WHERE (cast(invoicecreationdate as date) >= '" & StartDate & "' AND cast(invoicecreationdate as date) <= '" & EndDate & "') "
		SQL = SQL & " AND (InvoiceID NOT IN (SELECT InvoiceID FROM IN_InvoicesExportedSage)) "		
	End If
	
	SQL = SQL & " ORDER BY InvoiceID ASC" 
		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	Set rs = cnn8.Execute(SQL)

	If Not rs.Eof Then
			
		Do While Not rs.EOF
		
			InternalRecordIdentifier = rs("InternalRecordIdentifier")
			InvoiceID = rs("InvoiceID")
			CustID = rs("CustID")
			AlternateCustID = rs("AlternateCustID")
			InvoiceType = rs("InvoiceType")
			InvoiceGrandTotal = rs("InvoiceGrandTotal")
			InvoiceCreationDate = rs("InvoiceCreationDate")
			
		%>
			<tr id="IntRecID<%= InternalRecordIdentifier %>" data-child-value="<%= InvoiceID %>">

				<td class="details-control"></td>

				<%
				LastExportedDate = GetInvoiceExportedToSageLastDate(InvoiceID)
				
				If IsDate(LastExportedDate) Then RndDate = LastExportedDate Else RndDate = ""
				
				If RndDate <> "" Then
					eYear = Year(RndDate)
					If Month(RndDate) < 10 Then eMonth = "0" & Month(RndDate) else eMonth = Month(RndDate)
					If Day(RndDate) < 10 Then eDay = "0" & Day(RndDate) else eDay = Day(RndDate)
					DispayableDate = eMonth & "/" & eDay  & "/" & eYear
					DispayableDate  = cDate(DispayableDate) 
				End If

				If RndDate <> "" Then %>
					<td align="center" style="border-left: 2px solid #555 !important;"><span class="hidden"><%= eYear %><%= eMonth %><%= eDay %></span><%= Left(DispayableDate,Len(DispayableDate)-4) %><%= Right(DispayableDate,2) %></td>				
				<% Else %>
					<td align="center" style="border-left: 2px solid #555 !important;"><span class="hidden"></span>&nbsp;</td>								
				<% End If %>

				
				<td align="center"><input type="checkbox" class="exportcheckbox" name="chkExport<%= InvoiceID %>" id="chkExport<%= InvoiceID %>" value="<%= InvoiceID %>"></td>
				
				<%
				
				Date1 = "01/01/2019"
				Date2 = "05/10/2019"
				iDiff = DateDiff("d", Date1, Date2, vbMonday)
				Randomize
				'RndDate = DateAdd("d", Int((iDiff * Rnd) + 1), Date1)
				
				If IsDate(InvoiceCreationDate) Then RndDate = InvoiceCreationDate Else RndDate = ""
				
				If RndDate <> "" Then
					eYear = Year(RndDate)
					If Month(RndDate) < 10 Then eMonth = "0" & Month(RndDate) else eMonth = Month(RndDate)
					If Day(RndDate) < 10 Then eDay = "0" & Day(RndDate) else eDay = Day(RndDate)
					DispayableDate = eMonth & "/" & eDay  & "/" & eYear
					DispayableDate  = cDate(DispayableDate) 
				End If

				If RndDate <> "" Then %>
					<td align="center" style="border-left: 2px solid #555 !important;"><span class="hidden"><%= eYear %><%= eMonth %><%= eDay %></span><%= Left(DispayableDate,Len(DispayableDate)-4) %><%= Right(DispayableDate,2) %></td>				
				<% Else %>
					<td align="center" style="border-left: 2px solid #555 !important;"><span class="hidden"></span>&nbsp;</td>								
				<% End If %>
				<td align="center" style="border-right: 2px solid #555 !important;"><%= InvoiceID %></td>
				<td align="left" style="padding-left:20px;"><%= GetCustNameByCustNum(CustID) %>(<%= CustID %>)</td>
				<td align="left" style="border-right: 2px solid #555 !important;"><%= AlternateCustID %></td>		
				<td align="center"><%= InvoiceType %></td>					
				<td align="center"><%= InvoiceGrandTotal %></td>			
				<td align="center" style="border-right: 2px solid #555 !important;"><%= GetNumberOfInHistDetailLinesByInvoiceNumber(InvoiceID) %></td>
				
				<td align="center" style="border-right: 2px solid #555 !important;">
					<a data-toggle="modal" data-target="#modalEditSageInvoice" data-line-item-no="<%= InternalRecordIdentifier %>" data-invoice-id="<%= InvoiceID %>" class="btn btn-success" rel="tooltip" data-original-title="Click to edit this filter" style="cursor:pointer;"><i class="fa fa-pencil" aria-hidden="true"></i></a>																																
				</td>

		    </tr>
		<%
		
		rs.movenext
				
		Loop
		
		%>
		</tbody>
		</table>
		</div>
		<%
Else

	Response.Write("Nothing To Report")
End If

%>


            </table>
    </div>
         
          </div>
<!-- eof responsive tables !-->



<!-- eof row !-->

<!-- row !-->
<div class="row">
   <%
'Response.Write("<div class='col-lg-12'><h3>" & "Total Customers Listed:" & TotalCustsReported  & "</h3></div>")
%>

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


<!-- ************************************************************************** -->
<!-- MODALS FOR EDITING SAGE INVOICE IDS                                        -->
<!-- ************************************************************************** -->
<!--#include file="editSageInvoice_Modals.asp"-->
<!-- ************************************************************************** -->
<!-- ************************************************************************** -->
<style type="text/css">
	.datepicker.dropdown-menu {right: auto;}
</style>
<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<!-- Include Bootstrap DaterangePicker For Invoice Date Range Selection -->
<link href="<%= baseURL %>js/bootstrap-daterangepicker/daterangepicker.min.css" rel="stylesheet" type="text/css" />
<script src="<%= baseURL %>js/bootstrap-daterangepicker/daterangepicker.min.js" type="text/javascript"></script>

<script type="text/javascript">
	
    $('#reportrange').daterangepicker({
            opens: 'right',
            startDate: moment(),
            endDate: moment(),
            showWeekNumbers: true,
            timePicker: false,
            linkedCalendars: false,
            autoUpdateInput:false,
            autoApply:true,
            ranges: {
                'Today': [moment(), moment()],
                'Yesterday': [moment().subtract('days', 1), moment().subtract('days', 1)],
                'Last 7 Days': [moment().subtract('days', 6), moment()],
                'Last 30 Days': [moment().subtract('days', 29), moment()],
                'This Month': [moment().startOf('month'), moment().endOf('month')],
                'Last Month': [moment().subtract('month', 1).startOf('month'), moment().subtract('month', 1).endOf('month')]
            },
            buttonClasses: ['btn'],
            applyClass: 'green',
            cancelClass: 'default',
            format: 'MM/DD/YYYY',
            separator: ' to ',
            locale: {
                applyLabel: 'Apply',
                fromLabel: 'From',
                toLabel: 'To',
                customRangeLabel: 'Custom Range',
                daysOfWeek: ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'],
                monthNames: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                firstDay: 1
            }
        },
        function (start, end) {
            $('#reportrange span').html(start.format('MM/DD/YYYY') + ' - ' + end.format('MM/DD/YYYY'));
            $('#txtStartDate').val(start.format('MM/DD/YYYY'));
            $('#txtEndDate').val(end.format('MM/DD/YYYY'));
        }
    );
    
    var StartDate = $("#txtStartDate").val();
    var EndDate = $("#txtEndDate").val();
    
    if (StartDate.length === 0 || EndDate.length === 0) {
	    //Set the initial state of the picker label to yesterday
	    $('#reportrange span').html(moment().subtract(1, 'days').format('MM/DD/YYYY') + ' - ' + moment().subtract(1, 'days').format('MM/DD/YYYY'));
		$('#txtStartDate').val(moment().subtract(1, 'days').format('MM/DD/YYYY'));
		$('#txtEndDate').val(moment().subtract(1, 'days').format('MM/DD/YYYY'));
	}
	else {
	    //Set the initial state of the picker label
	    $('#reportrange span').html(StartDate + ' - ' + EndDate);
		$('#txtStartDate').val(StartDate);
		$('#txtEndDate').val(EndDate);
	}
</script>

<!--#include file="../../inc/footer-main.asp"-->