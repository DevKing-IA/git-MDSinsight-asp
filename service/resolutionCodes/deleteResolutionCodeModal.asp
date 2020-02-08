<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Users.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.asp"-->
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
		<h4 class="modal-title">Replace Resolution Before Deletion</h4>
	</div>
</div>

<form method="post" action="deleteresolutionCodeFromModal.asp" name="frmdeleteresolutionCodeFromModal" id="frmdeleteresolutionCodeFromModal">

	<input type='hidden' name="txtResolutionCodeToReplace" id="txtResolutionCodeToReplace" value='<%=InternalRecordIdentifier %>'>

	<div class="col-lg-12">
		There are <%=NumberOfTicketsByResolutionCode (InternalRecordIdentifier)%>&nbsp;service tickets with the resolution code you are trying to delete. Before this resolution code can be deleted you must chose a new resolution code to be assigned to these service tickets from the list below.  
	</div>

	<div class="col-lg-12">
		<div class="form-group">
			<label class="col-sm-3 control-label">Replace resolution with:</label>
			<div class="col-sm-9">
			  	<select class="form-control" name='seldeleteresolutionCodeFromModal' id='seldeleteresolutionCodeFromModal'>
				      	<% 'Get all resolution codes
				      	  	SQL9 = "SELECT * FROM FS_ResolutionCodes WHERE InternalRecordIdentifier <> " & InternalRecordIdentifier  & " order by ResolutionDescription"  ' Select all but the one to delete

							Set cnn9 = Server.CreateObject("ADODB.Connection")
							cnn9.open (Session("ClientCnnString"))
							Set rs9 = Server.CreateObject("ADODB.Recordset")
							rs9.CursorLocation = 3 
							Set rs9 = cnn9.Execute(SQL9)
							If not rs9.EOF Then
								Do
									Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("ResolutionDescription")& "</option>")
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
			<button type="submit" class="btn btn-primary">Replace Resolution Code &amp; Delete</button>
		</div>
	</div>
</form>