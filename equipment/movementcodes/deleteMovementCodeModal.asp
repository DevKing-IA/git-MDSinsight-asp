<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
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
		<h4 class="modal-title">Replace Movement Code Before Deletion</h4>
	</div>
</div>

<form method="post" action="deleteMovementCodeFromModal.asp" name="frmDeleteMovementCodeFromModal" id="frmDeleteMovementCodeFromModal">

	<input type='hidden' name="txtMovementCodeNoToReplace" id="txtMovementCodeNoToReplace" value='<%=InternalRecordIdentifier %>'>

	<div class="col-lg-12">
		There are <%= NumberEquipmentRecsDefinedForMovementCode(InternalRecordIdentifier)%>&nbsp;<%=GetTerm("Equipment")%>&nbsp;pieces of equipment assigned to the Movement Code you are trying to delete. Before this Movement Code can be deleted you must chose a new Movement Code to be assigned to these&nbsp;<%=GetTerm("Equipment")%>&nbsp;from the list below.  
	</div>

	<div class="col-lg-12">
		<div class="form-group">
			<label class="col-sm-3 control-label">Replace Movement Code with:</label>
			<div class="col-sm-9">
			  	<select class="form-control" name='selDeleteMovementCodeFromModal' id='selDeleteMovementCodeFromModal'>
				      	<% 'Get all Movement Codes
				      	  	SQL9 = "SELECT * FROM EQ_MovementCodes WHERE InternalRecordIdentifier <> " & InternalRecordIdentifier  & " ORDER BY movementDesc"  ' Select all but the one to delete

							Set cnn9 = Server.CreateObject("ADODB.Connection")
							cnn9.open (Session("ClientCnnString"))
							Set rs9 = Server.CreateObject("ADODB.Recordset")
							rs9.CursorLocation = 3 
							Set rs9 = cnn9.Execute(SQL9)
							If not rs9.EOF Then
								Do
									Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("movementCode") & " - " & rs9("movementDesc") & "</option>")
									rs9.movenext
								Loop until rs9.eof
							End If
							set rs9 = Nothing
							cnn9.close
							set cnn9 = Nothing
						%>
				</select>
			</div>
		</div>
	</div>

	<div class="col-lg-12">
		<div class="modal-footer">
			<button type="button" class="btn btn-default" data-dismiss="modal">Cancel Deletion</button>
			<button type="submit" class="btn btn-primary">Replace Movement Code & Delete</button>
		</div>
	</div>
</form>