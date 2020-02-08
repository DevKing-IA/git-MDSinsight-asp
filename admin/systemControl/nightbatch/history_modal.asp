<style type="text/css">
	.the-select{
		min-height: 150px;
		max-width: 50%;
	}
</style>

<div class="modal-dialog modal-height">
    <div class="modal-content">
	    <div class="modal-header">
	        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
	        <h4 class="modal-title" id="myModalLabel" align="center">Full Night Batch Log</h4>
		</div>

		<form method="post" action="dispatch_modal_SaveValues.asp" name="frmDispatchModal" id="frmDispatchModal">
	
			<!-- insert content in here !-->
			<div class="modal-body ativa-scroll">
    					<%
   					DataToDisplay = Replace(rs.Fields("NightBatchLogData1"),vbCRLF,"<br>")
   					If Not IsNull(rs.Fields("NightBatchLogData2")) Then
   						If rs.Fields("NightBatchLogData2") <> "" Then
   							DataToDisplay  = DataToDisplay + Replace(rs.Fields("NightBatchLogData1"),vbCRLF,"<br>")
   						End If
   					End If
   					Response.Write(DataToDisplay)
   					%>
 			</div>

			<div class="modal-footer">
				<button type="button" class="btn btn-primary" data-dismiss="modal">CLOSE</button>
			</div>
		</form>
	</div>
</div>