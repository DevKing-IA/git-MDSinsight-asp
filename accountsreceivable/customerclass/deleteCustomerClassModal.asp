<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<% 

InternalRecordIdentifier = Request.QueryString("i")
ClassCode = GetClassCodeByIntRecID(InternalRecordIdentifier)
NumAccountsByClassCode = NumberOfCustomersWithClassCode(ClassCode)

%>
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
		<h4 class="modal-title">Replace Class Code Before Deletion</h4>
	</div>
</div>

<form method="post" action="deleteCustomerClassFromModal.asp" name="frmDeleteClassCodeFromModal" id="frmDeleteClassCodeFromModal">

	<input type="hidden" name="txtClassCodeToBeDeletedIntRecID" id="txtClassCodeToBeDeletedIntRecID" value="<%=InternalRecordIdentifier %>">

	<div class="col-lg-12">
		There are <%= NumAccountsByClassCode %> customers in the <strong><%= ClassCode %></strong> customer class you are trying to delete. Before this class can be deleted you must chose a new customer class to be assigned to these customers from the list below.  
	</div>

	<div class="col-lg-12">
		<div class="form-group">
			<label class="col-sm-3 control-label">Replace customer class with:</label>
			<div class="col-sm-9">
			  	<select class="form-control" name="selDeleteClassCodeFromModal" id="selDeleteClassCodeFromModal">
				      	<% 'Get all stages
				      	  	SQL9 = "SELECT * FROM AR_CustomerClass WHERE InternalRecordIdentifier <> " & InternalRecordIdentifier  & " ORDER BY ClassCode"  ' Select all but the one to delete

							Set cnn9 = Server.CreateObject("ADODB.Connection")
							cnn9.open (Session("ClientCnnString"))
							Set rs9 = Server.CreateObject("ADODB.Recordset")
							rs9.CursorLocation = 3 
							Set rs9 = cnn9.Execute(SQL9)
							If not rs9.EOF Then
								Do
									Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("ClassCode") & " - " & rs9("ClassDescription") & "</option>")
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
			<button type="submit" class="btn btn-primary">Replace Class Code & Delete</button>
		</div>
	</div>
</form>