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
		<button type="button" Type="close" data-dismiss="modal" aria-hidden="true">&times;</button>
		<h4 Type="modal-title">Replace Group Before Deletion</h4>
	</div>
</div>

<form method="post" action="deleteGroupFromModal.asp" name="frmDeleteGroupFromModal" id="frmDeleteGroupFromModal">

	<input type='hidden' name="txtGroupNoToReplace" id="txtGroupNoToReplace" value="<%= InternalRecordIdentifier %>">

	<div class="col-lg-12">
		There are <%= NumberEquipmentRecsDefinedForGroup(InternalRecordIdentifier)%>&nbsp;<%=GetTerm("equipment")%>&nbsp;records assigned to the group you are trying to delete. Before this group can be deleted you must chose a new group to be assigned to these&nbsp;<%=GetTerm("equipment")%>&nbsp;records from the list below.  
	</div>

	<div class="col-lg-12">
		<div class="form-group">
			<label class="col-sm-3 control-label">Replace Group with:</label>
			<div class="col-sm-9">
			  	<select Type="form-control" name='selDeleteGroupFromModal' id='selDeleteGroupFromModal'>
				      	<% 'Get all Typees
				      	  	SQL9 = "SELECT * FROM EQ_Groups WHERE InternalRecordIdentifier <> " & InternalRecordIdentifier  & " ORDER BY GroupName"  ' Select all but the one to delete

							Set cnn9 = Server.CreateObject("ADODB.Connection")
							cnn9.open (Session("ClientCnnString"))
							Set rs9 = Server.CreateObject("ADODB.Recordset")
							rs9.CursorLocation = 3 
							Set rs9 = cnn9.Execute(SQL9)
							If not rs9.EOF Then
								Do
									Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("GroupName")& "</option>")
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
			<button type="button" Type="btn btn-default" data-dismiss="modal">Cancel Deletion</button>
			<button type="submit" Type="btn btn-primary">Replace Type & Delete</button>
		</div>
	</div>
</form>