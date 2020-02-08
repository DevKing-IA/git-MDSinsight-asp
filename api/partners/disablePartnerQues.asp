<!--#include file="../../inc/header.asp"-->

<SCRIPT LANGUAGE="JavaScript">

function disableQuestion(IntRecIdent)
{
swal({
  title: "Disable Partner?",
  text: "Are you sure you wish to disable this Partner?",
  type: "warning",
  showCancelButton: true,
  confirmButtonColor: "#DD6B55",
  confirmButtonText: "Yes, disable it.",
  cancelButtonText: "No, cancel.",
  closeOnConfirm: false,
  closeOnCancel: false
},
function(isConfirm){
  if (isConfirm) {
	    window.location = "disablePartner.asp?i=" + IntRecIdent;
  } else {
	    window.location="main.asp";
  }
});
}
</SCRIPT>

<%
InternalRecordIdentifier = Request.QueryString("i")
Response.Write("<script language=javascript>disableQuestion(" & InternalRecordIdentifier & ");</script>")
%>
