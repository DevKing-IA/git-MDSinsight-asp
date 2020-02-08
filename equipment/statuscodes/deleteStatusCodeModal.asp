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
		<h4 class="modal-title">Replace Status Code Before Deletion</h4>
	</div>
</div>

<form method="post" action="deleteStatusCodeFromModal.asp" name="frmDeleteStatusCodeFromModal" id="frmDeleteStatusCodeFromModal">

	<input type='hidden' name="txtStatusCodeNoToReplace" id="txtStatusCodeNoToReplace" value='<%=InternalRecordIdentifier %>'>

	<div class="col-lg-12">
		There are <%= NumberEquipmentRecsDefinedForStatusCode(InternalRecordIdentifier)%>&nbsp;<%=GetTerm("Equipment")%>&nbsp;pieces of equipment assigned to the status code you are trying to delete. Before this status code can be deleted you must chose a new status code to be assigned to these&nbsp;<%=GetTerm("Equipment")%>&nbsp;from the list below.  
	</div>

	<div class="col-lg-12">
		<div class="form-group">
			<label class="col-sm-3 control-label">Replace Status Code with:</label>
			<div class="col-sm-9">
			  	<select class="form-control" name='selDeleteStatusCodeFromModal' id='selDeleteStatusCodeFromModal'>
				      	<% 'Get all status codes
				      	  	SQL9 = "SELECT * FROM EQ_StatusCodes WHERE InternalRecordIdentifier <> " & InternalRecordIdentifier  & " ORDER BY statusDesc"  ' Select all but the one to delete

							Set cnn9 = Server.CreateObject("ADODB.Connection")
							cnn9.open (Session("ClientCnnString"))
							Set rs9 = Server.CreateObject("ADODB.Recordset")
							rs9.CursorLocation = 3 
							Set rs9 = cnn9.Execute(SQL9)
							If not rs9.EOF Then
								Do
									Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("statusDesc") & "</option>")
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
			<button type="submit" class="btn btn-primary">Replace Status Code & Delete</button>
		</div>
	</div>
</form>