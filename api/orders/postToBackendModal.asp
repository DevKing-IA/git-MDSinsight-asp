<!-- css -->
<style type="text/css">
	.modal-body .table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
  border: 0px !important;
	}
</style>
<!-- eof css -->

<form method="post" action="postToBackendModal_SaveValues.asp" name="frmpostToBackendModal" id="frmpostToBackendModal">
	
	<input type='hidden' id='txtInternalRecordIdentifier' name='txtInternalRecordIdentifier' value='<%=rsOrderHeader("InternalRecordIdentifier")%>'>
	
	<!-- modal starts here -->
	<div class="modal-dialog" role="document">
	    <div class="modal-content">

		    <!-- title / close -->	
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="myModalLabel">Re-send to <%=GetTerm("Backend")%></h4>
			</div>
			<!-- eof title / close -->

			<!-- content -->
			<div class="modal-body">
				<div class="table-responsive">
					<table class="table">
    
						<!-- line -->
						<tr>
						  <td width="20%"><strong>Order ID:</strong></td>
						  <td><%=rsOrderHeader("OrderID")%></td>
						</tr>
						<!-- eof line -->

						<!-- line -->
						<tr>
						  <td width="20%"><strong><%=GetTerm("Customer")%> ID:</strong></td>
						  <td><%=rsOrderHeader("CustID")%></td>
						</tr>
						<!-- eof line -->

						<!-- line -->
						<tr>
						  <td width="20%"><strong>Bill To:</strong></td>
						  <td><%=rsOrderHeader("BillToCompany")%></td>
						</tr>
						<!-- eof line -->

						<!-- line -->
						<tr>
						  <td width="20%"><strong>Ship To:</strong></td>
						  <td><%=rsOrderHeader("ShipToCompany")%></td>
						</tr>
						<!-- eof line -->

						<!-- line -->
						<tr>
						  <td width="20%"><strong>Amount</strong></td>
						  <td><%=FormatCurrency(rsOrderHeader("GrandTotal"),2)%></td>
						</tr>
						<!-- eof line -->

						<!-- line -->
						<tr>
						  <td width="20%"><strong># Lines</strong></td>
						  <td><%=NumberOfAPIOrderLines(rsOrderHeader("OrderID"),GetAPIOrderHighestThread(rsOrderHeader("OrderID")))%></td>
						</tr>
						<!-- eof line -->

					</table>
				</div>
			</div>

			<!-- cancel / send -->
			<div class="modal-footer">
				<button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
		        <button type="submit" class="btn btn-primary">Re-Send</button>
			</div>
			<!-- eof cancel / send -->
		</div>
	</div>
	
</form>
<!-- modal ends here -->