<!--#include file="../../../inc/header.asp"-->
<!--#include file="../../../inc/InSightFuncs.asp"--> 
<!--#include file="../../../inc/InSightFuncs_Service.asp"-->

<%


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

CreateAuditLogEntry "Filters","Filters","Minor",0, MUV_Read("DisplayName") & " ran Add/Edit Filters"

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
	    margin: 40px 0 10px;;
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
			
</style>

<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.18/sorting/currency.js"></script>
<script type="text/javascript">


$(document).ready(function() {

    $("#PleaseWaitPanel").hide();
    
    $('#filter').keyup(function () {

	    var rex = new RegExp($(this).val(), 'i');
	    $('.searchable tr').hide();
	    $('.searchable tr').filter(function () {
	        return rex.test($(this).text());
	    }).show();

	})

	
    $("#AddedCustID").val("");
	
    $("[rel='tooltip']").tooltip('destroy');
	$("[rel='tooltip']").tooltip({ placement: 'left' });
	
    
	$('#tableSuperSum').DataTable({
        scrollY: 500,
        scrollCollapse: true,
        paging: false,
        order: [[ 0, 'asc' ],[ 1, 'asc' ]],
		columnDefs: [
		        { targets: 8, "orderable": false,},
		        { type: 'currency', targets: [2, 3]}
		    ]	        
	    }
	);


	
	$("#AddFilterSave").click(function(){	
	
		var FilterID = $("#txtFilterID").val();
		var FilterDescription = $("#txtFilterDescription").val();
		var FilterCost = $("#txtFilterCost").val();
		var FilterListPrice = $("#txtFilterListPrice").val();
		var FilterTaxable = $("#selFilterTaxable option:selected").val();
		var FilterInventoried = $("#selFilterInventoried option:selected").val();
		var FilterPickable = $("#selFilterPickable option:selected").val();
		var FilterUPC = $("#txtFilterUPC").val();
		var FilterprodSKU = $("#selFilterprodSKU").val();
		var FilterdisplayOrder = $("#selFilterdisplayOrder").val();

		if (FilterID.length <=0) {
			swal({
				title: 'Error Adding Filter',
				text: 'Please specify a filter ID',
				type: 'error'
			});
			return false;
		}

		if (FilterDescription.length <=0) {
			swal({
				title: 'Error Adding Filter',
				text: 'Please specify a filter description',
				type: 'error'
			});
			return false;
		}
				
		if (FilterTaxable.length <=0) {
			swal({
				title: 'Error Adding Filter',
				text: 'Please specify whether the filter is taxable or not.',
				type: 'error'
			});
			return false;
		}
		
		if (FilterInventoried.length <=0) {
			swal({
				title: 'Error Adding Filter',
				text: 'Please specify whether the filter is inventoried or not.',
				type: 'error'
			});
			return false;
		}

		if (FilterPickable.length <=0) {
			swal({
				title: 'Error Adding Filter',
				text: 'Please specify whether the filter is pickable or not.',
				type: 'error'
			});
			return false;
		}		
		

    	$.ajax({
			type:"POST",
			url: "../../../../inc/InSightFuncs_AjaxForServiceModals.asp",
			cache: false,
			data: "action=CheckForDuplicateFilterIDNewFilter&FilterID="+encodeURIComponent(FilterID),
			
			success: function(response)
			 {
				if (response.startsWith("We are sorry")) {				
					swal({
						title: 'Error Adding New Filter',
						text: response,
						type: 'error'
					})
					return false;
				} 
				
				else {
				
			    	$.ajax({
						type:"POST",
						url: "../../../../inc/InSightFuncs_AjaxForServiceModals.asp",
						cache: false,
						data: "action=CheckForDuplicateFilterUPCCodeNewFilter&FilterUPC="+encodeURIComponent(FilterUPC),
						
						success: function(response)
						 {
							if (response.startsWith("We are sorry")) {				
								swal({
									title: 'Error Adding New Filter',
									text: response,
									type: 'error'
								})
								return false;
							} 
							
							else {
							
						    	$.ajax({
									type:"POST",
									url: "../../../../inc/InSightFuncs_AjaxForServiceModals.asp",
									cache: false,
									data: "action=SaveAddNewFilter&FilterID="+encodeURIComponent(FilterID)+"&FilterDescription="+encodeURIComponent(FilterDescription)+"&FilterCost="+encodeURIComponent(FilterCost)+"&FilterListPrice="+encodeURIComponent(FilterListPrice)+"&FilterTaxable="+encodeURIComponent(FilterTaxable)+"&FilterInventoried="+encodeURIComponent(FilterInventoried)+"&FilterPickable="+encodeURIComponent(FilterPickable)+"&FilterUPC="+encodeURIComponent(FilterUPC)+"&FilterprodSKU="+encodeURIComponent(FilterprodSKU)+"&FilterdisplayOrder="+encodeURIComponent(FilterdisplayOrder),
									
									success: function(response)
									 {
										if (response.startsWith("Error:")) {				
											swal({
												title: 'Error Adding New Filter',
												text: response,
												type: 'error'
											})
											return;
										} 
										else {
											$("#frmAddFilter").submit();				
										}
									 },
									failure: function(response)
									 {
										swal({
											title: 'Error Adding New Filter',
											text: response,
											type: 'error'
										})
						             }
								});
							}
						 },
					});

				}			
			 },
		});

		
    });






	$('#modalEditExistingFilter').on('show.bs.modal', function(e) {

	    //get data-id attribute of the clicked order
	    var IntRecID = $(e.relatedTarget).data('filter-inc-rec-id');
	    
	    //populate the textbox with the id of the clicked filter
	    $(e.currentTarget).find('input[name="txtIntRecID"]').val(IntRecID);

	    var $modal = $(this);

    	$.ajax({
			type:"POST",
			url: "../../../../inc/InSightFuncs_AjaxForServiceModals.asp",
			cache: false,
			async: false,
			data: "action=GetContentForEditFilterModal&IntRecID="+encodeURIComponent(IntRecID),
			success: function(response)
			 {					
               	 $modal.find('#modalEditExistingFilterContent').html(response);              	 
             },
             failure: function(response)
			 {
			  	$modal.find('#modalEditExistingFilterContent').html("Failed");
             }
		});
		
	});
	
	
	
	$('#modalDeleteExistingFilter').on('show.bs.modal', function(e) {

	    //get data-id attribute of the clicked order
	    var IntRecID = $(e.relatedTarget).data('filter-inc-rec-id');
	    
	    //populate the textbox with the id of the clicked filter
	    $(e.currentTarget).find('input[name="txtIntRecID"]').val(IntRecID);

	    var $modal = $(this);

    	$.ajax({
			type:"POST",
			url: "../../../../inc/InSightFuncs_AjaxForServiceModals.asp",
			cache: false,
			async: false,
			data: "action=GetContentForDeleteFilterModal&IntRecID="+encodeURIComponent(IntRecID),
			success: function(response)
			 {					
               	 $modal.find('#modalDeleteExistingFilterContent').html(response);              	 
             },
             failure: function(response)
			 {
			  	$modal.find('#modalDeleteExistingFilterContent').html("Failed");
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
	Response.Write("<div id=""PleaseWaitPanel"" class=""container"">")
	Response.Write("<br><br>Loading Filters <br><br>Please wait...<br><br>")
	Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
	Response.Write("</div>")
	Response.Flush()

%>

<h3 class="page-header"><i class="fa fa-wrench"></i>&nbsp;Add/Edit Filters
&nbsp;&nbsp;
<small><button type="button" class="btn btn-large btn-success" id="btnAddFilter" data-toggle="modal" data-target="#modalAddFilter"><i class="fa fa-plus-circle" aria-hidden="true"></i> Add New Filter</button></small>
</h3>

 
	
	
<form method="POST" name="frmAddFilter" id="frmAddFilter" value="frmAddFilter" action="main.asp">
	
	<input id="frmAddFiltersubmitted" name="frmAddFiltersubmitted" type="hidden" value="1">
	<input id="RemovedCustID" name="RemovedCustID" type="hidden" value="">

</form>	    


<!-- row !-->
<div class="row">


<div class="container-fluid">
    <div class="row">
           <table id="tableSuperSum" class="display compact" style="width:100%;">
              <thead>
                  <tr>	
                		<th class="gen-info-header" colspan="3" style="border-right: 2px solid #555 !important;font-size: 16px;">General</th>
						<th class="cost-header" colspan="3" style="border-right: 2px solid #555 !important;font-size: 16px;">Cost</th>
						<th class="inventory-header" colspan="4" style="border-right: 2px solid #555 !important;font-size: 16px;">Inventory</th>
						<th class="actions-header" style="border-right: 2px solid #555 !important;font-size: 16px;">Actions</th>
				</tr>
				
                <tr>
					<th class="sorttable_numeric" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="colFilterID"><br>Filter ID</th>
					<th class="sorttable_numeric" style="border-top: 2px solid #555 !important;" id="colDescription"><br>Description</th>
					<th class="sorttable_numeric" style="border-top: 2px solid #555 !important; border-right: 2px solid #555 !important;" id="colDescription">Display<br>Order</th>

					<th class="sorttable_numeric" style="border-top: 2px solid #555 !important;" id="colListPrice"><br>List Price</th>
					<th class="sorttable_numeric" style="border-top: 2px solid #555 !important;" id="colDefaultCost"><br>Default Cost</th>
					<th class="sorttable_numeric" style="border-top: 2px solid #555 !important; border-right: 2px solid #555 !important;" id="colTaxable"><br>Taxable</th>
					
					<th class="sorttable_numeric" style="border-top: 2px solid #555 !important;" id="colInventoried"><br>Inventoried</th>	
					<th class="sorttable_numeric" style="border-top: 2px solid #555 !important;" id="colPickable"><br>Pickable</th>
					<th class="sorttable_numeric" style="border-top: 2px solid #555 !important;" id="colUPC"><br>UPC Code</th>
					<th class="sorttable_numeric" style="border-top: 2px solid #555 !important; border-right: 2px solid #555 !important;" id="colProdID"><br>Prod ID</th>									
					
					<th align="center" class="sorttable_numeric" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="colAction"><br>Actions</th>
                </tr>
                
              </thead>
              
			
   <tbody>
   
   <%

	SQL = "SELECT * FROM IC_FILTERS ORDER BY displayOrder ASC" 
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	Set rs = cnn8.Execute(SQL)

	If Not rs.Eof Then
			
		Do While Not rs.EOF
		
			InternalRecordIdentifier = rs("InternalRecordIdentifier")
			FilterID = rs("FilterID")
			Description = rs("Description")
			ListPrice = rs("ListPrice")
			DefaultCost = rs("DefaultCost")
			Taxable = rs("Taxable")
			InventoriedItem = rs("InventoriedItem")
			PickableItem = rs("PickableItem")
			UPCCode = rs("UPCCode")
			prodSKU = rs("prodSKU")	
			displayOrder = rs("displayOrder")		
			
			If InventoriedItem = 1 Then
				InventoriedItem = "Y"
			Else
				InventoriedItem ="N"
			End If

			If PickableItem = 1 Then
				PickableItem = "Y"
			Else
				PickableItem ="N"
			End If

			If Taxable = 1 Then
				Taxable = "Y"
			Else
				Taxable ="N"
			End If
			
			If ListPrice <> "" Then
				ListPrice = FormatCurrency(ListPrice,2)
			Else
				ListPrice = "N/A"
			End If
			
			If DefaultCost <> "" Then
				DefaultCost = FormatCurrency(DefaultCost,2)
			Else
				DefaultCost = "N/A"
			End If
			
			
		%>
			<tr id="IntRecID<%= InternalRecordIdentifier %>">

				<td align="center" style="border-left: 2px solid #555 !important;"><%= FilterID %></td>				
				<td align="left"><%= Description %></td>
				<td align="left" style="border-right: 2px solid #555 !important;"><%= displayOrder %></td>
				<td align="center"><%= ListPrice %></td>
				<td align="center"><%= DefaultCost %></td>
				<td align="center" style="border-right: 2px solid #555 !important;"><%= Taxable %></td>		
				<td align="center"><%= InventoriedItem %></td>					
				<td align="center"><%= PickableItem %></td>			
				<td align="left"><%= UPCCode %></td>
				<td align="left" style="border-right: 2px solid #555 !important;"><%= prodSKU %></td>
				
				<td align="center" style="border-right: 2px solid #555 !important;">
					<a data-toggle="modal" data-target="#modalEditExistingFilter" data-filter-inc-rec-id="<%= InternalRecordIdentifier %>" class="btn btn-success" rel="tooltip" data-original-title="Click to edit this filter" style="cursor:pointer;"><i class="fa fa-pencil" aria-hidden="true"></i></a>																																
					
					<%' Allow delete or display modal
					If NumberCustomerRecsDefinedForFilterID(InternalRecordIdentifier) = 0 Then %>
						<a href="deleteFilterQues.asp?i=<%= InternalRecordIdentifier %>" rel="tooltip" data-original-title="Click to delete this filter" style="cursor:pointer; margin-left:10px;" class="btn btn-danger"><i class="fas fa-trash-alt"></i></a>
					<% Else %>
						<a data-toggle="modal" data-target="#modalDeleteExistingFilter" data-filter-inc-rec-id="<%= InternalRecordIdentifier %>" class="btn btn-danger" rel="tooltip" data-original-title="Click to delete this filter" style="cursor:pointer; margin-left:10px;"><i class="fas fa-trash-alt" aria-hidden="true"></i></a>
					<% End If %>
					
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
'Response.Write("<div class='col-lg-12'><h3>" & "Total filters listed:" & TotalCustsReported  & "</h3></div>")
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
<!-- MODALS FOR ADDING AND EDITING A FILTER                                     -->
<!-- ************************************************************************** -->
<!--#include file="AddEditFilters_Modals.asp"-->
<!-- ************************************************************************** -->
<!-- ************************************************************************** -->


<!--#include file="../../../inc/footer-main.asp"-->