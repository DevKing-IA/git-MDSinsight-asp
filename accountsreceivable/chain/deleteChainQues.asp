<!--#include file="../../inc/header.asp"-->

<SCRIPT LANGUAGE="JavaScript">

function deletionQuestion(IntRecIdent)
{
swal({
  title: "Delete Customer Chain?",
  text: "Are you sure you wish to delete this chain?",
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
	    window.location = "deleteChain.asp?i=" + IntRecIdent;
  } else {
	    window.location="main.asp";
  }
});
}
</SCRIPT>

<%
InternalRecordIdentifier = Request.QueryString("i")
Response.Write("<script language=javascript>deletionQuestion(" & InternalRecordIdentifier & ");</script>")
%>
