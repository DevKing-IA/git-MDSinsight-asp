<!--#include file="../../inc/header.asp"-->

<SCRIPT LANGUAGE="JavaScript">

function deletionQuestion(IntRecIdent)
{
swal({
  title: "Delete Stage?",
  text: "Are you sure you wish to delete this stage?",
  type: "warning",
  showCancelButton: true,
  confirmButtonColor: "#DD6B55",
  confirmButtonText: "Yes, delete it.",
  cancelButtonText: "No, cancel.",
  closeOnConfirm: false,
  closeOnCancel: false
},
function(isConfirm){
  if (isConfirm) {
	    window.location = "deleteStages.asp?i=" + IntRecIdent;
  } else {
	    window.location="main.asp";
  }
});
}
</SCRIPT>

<style type="text/css">
.col-lg-12{
	margin-bottom:20px;
}

.modal-footer{
	margin-top:15px;
}
</style>

<%
InternalRecordIdentifier = Request.QueryString("i")
Response.Write("<script language=javascript>deletionQuestion(" & InternalRecordIdentifier & ");</script>")
%>