<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<% 

InternalRecordIdentifier = Request.QueryString("i")
'ClassCode = GetClassCodeByIntRecID(InternalRecordIdentifier)
NumAccountsByTermCode = NumberOfCustomersWithTerm(InternalRecordIdentifier)
CustTypeToBeDeletedDescription = GetCustTermDescByIntRecID(InternalRecordIdentifier)

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
		<h4 class="modal-title">Replace Term Before Deletion</h4>
	</div>
</div>

<form method="post" action="deleteTermFromModal.asp" name="frmDeleteTermFromModal" id="frmDeleteTermFromModal">

	<input type="hidden" name="txtTermToBeDeletedIntRecID" id="txtTermToBeDeletedIntRecID" value="<%=InternalRecordIdentifier %>">

	<div class="col-lg-12">
		There are <%= NumAccountsByTermCode %> customers in the <strong><%= CustTypeToBeDeletedDescription %></strong> term you are trying to delete. Before this type can be deleted you must chose a new term to be assigned to these customers from the list below.  
	</div>

	<div class="col-lg-12">
		<div class="form-group">
			<label class="col-sm-3 control-label">Replace Term with:</label>
			<div class="col-sm-9">
			  	<select class="form-control" name="selDeleteTermFromModal" id="selDeleteTermFromModal">
				      	<% 'Get all stages
				      	  	SQL9 = "SELECT * FROM AR_Terms WHERE InternalRecordIdentifier <> " & InternalRecordIdentifier  & " ORDER BY InternalRecordIdentifier"  ' Select all but the one to delete

							Set cnn9 = Server.CreateObject("ADODB.Connection")
							cnn9.open (Session("ClientCnnString"))
							Set rs9 = Server.CreateObject("ADODB.Recordset")
							rs9.CursorLocation = 3 
							Set rs9 = cnn9.Execute(SQL9)
							If not rs9.EOF Then
								Do
									Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("InternalRecordIdentifier") & " - " & rs9("Description") & "</option>")
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
			<a href="main.asp"><button type="button" class="btn btn-default">Cancel Deletion</button></a>
			<button type="submit" class="btn btn-primary">Replace Term & Delete</button>
		</div>
	</div>
</form>