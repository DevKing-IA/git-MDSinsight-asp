<!--#include file="../../inc/InsightFuncs_InventoryControl.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<% InternalRecordIdentifier = Request.QueryString("i") %>
<style type="text/css">
.col-lg-12{
	margin-bottom:20px;
}

.modal-footer{
	margin-top:15px;
}
</style>


<div class="col-lg-12">
	<div class="modal-header">
		<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
		<h4 class="modal-title">Replace Partner Before Deletion</h4>
	</div>
</div>

<form method="post" action="deletePartnerFromModal.asp" name="frmDeletePartnerFromModal" id="frmDeletePartnerFromModal">

	<input type="hidden" name="txtPartnerIntRecIDToReplace" id="txtPartnerIntRecIDToReplace" value="<%=InternalRecordIdentifier %>">

	<div class="col-lg-12">
		There are <strong><%=NumberOfSKUsDefinedForPartner(InternalRecordIdentifier)%> product SKUs</strong> assigned to the partner, <strong><%= GetPartnerNameByIntRecID(InternalRecordIdentifier) %></strong>, you are trying to delete. 
		<br><br>Before this partner can be deleted you must chose a new partner to be assigned to these SKUs from the list below. 
		<br><br>You can also choose to delete this partner AND all associated SKUs (last option below).  
	</div>

	<div class="col-lg-12">
		<div class="form-group">
			<label class="col-sm-3 control-label">Replace <%= GetPartnerNameByIntRecID(InternalRecordIdentifier) %> With:</label>
			<div class="col-sm-9">
			  	<select class="form-control" name='selDeletePartnerFromModal' id='selDeletePartnerFromModal'>
			  			
				      	<% 'Get all stages
				      	  	SQL9 = "SELECT * FROM IC_Partners WHERE InternalRecordIdentifier <> " & InternalRecordIdentifier  & " ORDER BY RecordCreationDateTime Desc"  ' Select all but the one to delete

							Set cnn9 = Server.CreateObject("ADODB.Connection")
							cnn9.open (Session("ClientCnnString"))
							Set rs9 = Server.CreateObject("ADODB.Recordset")
							rs9.CursorLocation = 3 
							Set rs9 = cnn9.Execute(SQL9)
							If not rs9.EOF Then
								Do
									Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("partnerCompanyName") & " - " & rs9("partnerAPIKey") & "</option>")
									rs9.movenext
								Loop until rs9.eof
							End If
							set rs9 = Nothing
							cnn9.close
							set cnn9 = Nothing
						%>
						<option value="DELETE_PARTNER_AND_SKUS">DELETE PARTNER AND ALL ASSOCIATED PRODUCT SKUS</option>
				</select>
			</div>
		</div>
		
	</div>
	
	<div class="col-lg-12">		
		<div class="alert alert-danger">
		  <strong>Warning!</strong> Please note this operation cannot be undone.
		</div>		
	</div>

	<div class="col-lg-12">
		<div class="modal-footer">
			<button type="button" class="btn btn-default" data-dismiss="modal">Cancel Deletion</button>
			<button type="submit" class="btn btn-primary">Confirm Deletion</button>
		</div>
	</div>
</form>