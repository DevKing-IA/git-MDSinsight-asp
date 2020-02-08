<!--#include file="../../../../inc/header.asp"-->
<!--#include file="../../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../../inc/InSightFuncs_BizIntel.asp"-->


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
%>

<script language="JavaScript">
	$(document).ready(function() {

		var table = $('#quotedItemsTable').DataTable({
		        "lengthMenu": [[10, 25, 50, -1], [10, 25, 50, "All"]],
		        "order": [[ 2, "asc" ],[ 0, "asc" ],[ 4, "asc" ]],
		        "stateSave": true,
				"select": true		        
		 });	    

		//table.column(0).visible(false);
		
	    $('#quotedItemsTable tbody').on( 'click', 'tr', function () {
	        if ( $(this).hasClass('selected') ) {
	            $(this).removeClass('selected');
	        }
	        else {
	            table.$('tr.selected').removeClass('selected');
	            $(this).addClass('selected');
	        }
	    } );
	 
	    $('#deleteQuotedItem').click( function () {


				$('tr.selected').each(function(index,item){
				
				    if(parseInt($(item).data('index'))>0){
				    
				        IntRecID = $(item).data('index');
				        
				        alert('about to delete item with rec id '+ IntRecID);
				        
						$.ajax({
							type:"POST",
							url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
							data: "action=DeleteQuotedItemFromCustomer&recid="+encodeURIComponent(IntRecID),
							success: function(msg)
							{
						        table.row('.selected').remove().draw(false);
							}
						}) 
 
				    }
				});			
	    });	
 

	 
	    $('#undoAllChanges').click( function () {

			custID = $('#txtCustID').val();
			alert(custID);
			
			$.ajax({
				type:"POST",
				url: "../../../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				data: "action=UndoQuotedItemChangesForCustomer",
				success: function(msg)
				{
			        location.reload();
				}
			}) 
	
	    });	
   	    	
	    $('#datepicker1').datetimepicker({
	    	format: 'MM/DD/YYYY',
	    	useCurrent: false,
	    	//defaultDate: moment(momentString),
	    	maxDate:moment().add(-1, 'days')
	    });
	    
		$("#datepicker1").on("dp.change", function (e) {
	    	selectedDate = $("#datepicker1").find("input").val();
	        //location.href = 'HistoricalDeliveriesByDriver.asp?date=' + selectedDate;
	    });	  
	} );
   
	
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

</style>
<h1 class="page-header"><i class="fa fa-file-text-o"></i> <%=GetTerm("Customer")%> Quoted Items for Account <%= custID %>, <%= GetCustNameByCustNum(custID) %></h1>
<input type="hidden" name="txtCustID" id="txtCustID" value="<%= custID %>">
<!-- row !-->
<div class="row">


	<div class="col-lg-10">
		<table id="quotedItemsTable" class="table table-striped table-bordered" cellspacing="0" width="100%">
		        <thead>
		            <tr>
		                <th>PRODUCT</th>
		                <th>DESCRIPTION</th>
		                <th>CATEGORY</th>
		                <th>UM</th>
		                <!--<th>SUGG QTY</th>
		                <th>YTD QTY</th>
		                <th>MTD QTY</th>-->
		                <th>LIST FLAG</th>
		                <th>COST</th>
		                <th>DATE QUOTED</th>
		                <th>EXPIRE DATE</th>
		                <th>PRICE</th>
		                <th>GP $</th>
		                <th>GP %</th>
		                <th>NEW PRICE</th>  
		                <th>NEW GP $</th>  
		                <th>NEW GP %</th>
		            </tr>
		        </thead>
		        <tfoot>
		            <tr>
		                <th>PRODUCT</th>
		                <th>DESCRIPTION</th>
		                <th>CATEGORY</th>
		                <th>UM</th>
		                <!--<th>SUGG QTY</th>
		                <th>YTD QTY</th>
		                <th>MTD QTY</th>-->
		                <th>LIST FLAG</th>
		                <th>COST</th>
		                <th>DATE QUOTED</th>
		                <th>EXPIRE DATE</th>
		                <th>PRICE</th>
		                <th>GP $</th>
		                <th>GP %</th>
		                <th>NEW PRICE</th>
		                <th>NEW GP $</th>  
		                <th>NEW GP %</th>		                    
		            </tr>
		        </tfoot>
		        <tbody>
				<%       
				Set rsQuotedItems = Server.CreateObject("ADODB.Recordset")
				rsQuotedItems.CursorLocation = 3 
			
				SQLQuotedItems = "SELECT * FROM zPRC_AccountQuotedItems_" & trim(Session("Userno")) & " WHERE DeleteFlag <> 1"
				'ORDER BY MUST BE DONE USING DATATABLES.JS
				
				'Response.write(SQLQuotedItems)
				
				Set cnnQuotedItems = Server.CreateObject("ADODB.Connection")
				cnnQuotedItems.open (Session("ClientCnnString"))
				Set rsQuotedItems = cnnQuotedItems.Execute(SQLQuotedItems)
				
				If NOT rsQuotedItems.EOF Then
					Do While NOT rsQuotedItems.EOF
						%>
			            <tr id="<%= rsQuotedItems("InternalRecordIdentifier") %>" data-index="<%= rsQuotedItems("InternalRecordIdentifier") %>">
			                <td><%= rsQuotedItems("ProdSKU") %></td>
			                <td><%= rsQuotedItems("Description") %></td>
			                <td><%= GetCategoryByID(rsQuotedItems("Category")) %></td>
			                <td><%= rsQuotedItems("QuoteType") %></td>
	                
			                <!--<td><%= rsQuotedItems("SuggestedQty") %></td>-->
			                <!--<td><%= GetYTDPurchaseQtyByCustByItem(custID, rsQuotedItems("ProdSKU")) %></td>-->
			                <!--<td><%= GetMTDPurchaseQtyByCustByItem(custID, rsQuotedItems("ProdSKU")) %></td>-->
			                
			                <td><%= rsQuotedItems("ListFlag") %></td>
			                <td><%= rsQuotedItems("Cost") %></td>
			                <td><%= rsQuotedItems("DateQuoted") %></td>
			                	
			                <td>
				                <div class="input-group date" id="datepicker1">
				                    <input type="text" class="form-control" name="txtProductPriceExpireDate" id="txtProductPriceExpireDate"  value="<%= rsQuotedItems("ExpireDate") %>">
				                    <span class="input-group-addon">
				                        <span class="glyphicon glyphicon-calendar"></span>
				                    </span>
				                </div>
			                </td>
			                			                
			                <td><%= rsQuotedItems("Price") %></td>
			                
			                <% If rsQuotedItems("Price") <> "" AND NOT IsNull(rsQuotedItems("Price")) AND NOT IsEmpty(rsQuotedItems("Price")) AND rsQuotedItems("Price") > 0 AND rsQuotedItems("Cost") <> ""  Then %>
			                	<td><%= rsQuotedItems("Price") - rsQuotedItems("Cost") %></td>
			                	<td><%= formatNumber(((rsQuotedItems("Price") - rsQuotedItems("Cost"))/rsQuotedItems("Price")) * 100,2) %>%</td>
					        <% Else %>
			                	<td>NA</td>
			                	<td>NA</td>
			                <% End If %>              

			                <td><span class="fa fa-usd dollarSignSpan"></span><input type="text" id="txtNewPrice" name="txtNewPrice" value="" class="form-control last-run-inputs"></td>
			                
			                <td id="newGPDollars">---</td>
			                
			                <td><span class="fa fa-percent dollarSignSpan"></span><input type="text" id="txtNewGPPercent" name="txtNewGPPercent" value="" class="form-control last-run-inputs"></td>
			                
			            </tr>
						<%		
						rsQuotedItems.MoveNext
						Loop
					End If
					
					set rsQuotedItems = Nothing
					%>		
		        </tbody>
		    </table>	
	</div>
	
    <!-- Historical Impact Summary !-->
   	<div class="col-lg-2 reports-box">


		<div class="row" style="margin-bottom:20px;">

		    <button type="button" class="btn btn-primary btn-lg btn-block" style="margin-bottom:10px;" onclick="location.href='addProspectPrequalify.asp';">
		        <i class="fa fa-user"></i>&nbsp;Change Customer
		    </button>
		    
		    <button type="button" class="btn btn-success btn-lg btn-block" style="margin-bottom:10px;" data-toggle="modal" data-target="#myProspectingModalAdd" id="addProspectToGroupSelected">
		        <i class="fa fa-plus"></i>&nbsp;Add New Quoted Item
		    </button>
		    
		    <button type="button" id="deleteQuotedItem" class="btn btn-danger btn-lg btn-block" style="margin-bottom:10px;">
		        <i class="fas fa-trash-alt"></i>&nbsp;Delete Quoted Item
		    </button>
		    		    
		    <button type="button" id="undoAllChanges" class="btn btn-warning btn-lg btn-block" style="margin-bottom:10px;">
		        <i class="fa fa-undo"></i>&nbsp;Undo All Changes
		    </button>		    

		</div>
   	


   		<div class="row">
	    	<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
	    	<p><a href="#" class="title">Historical Impact</a></p>
	        <p align="right"><a href="<%= BaseURL %>############"><button type="button" class="btn btn-primary">Run Simulation</button></a></p>
			
			<!--#include file="historical_impact.asp"-->
	        
	    </div>
	    
    </div> 
    <!-- eof Historical Impact Summary !-->

</div>
<!--#include file="../../../../inc/footer-main.asp"-->