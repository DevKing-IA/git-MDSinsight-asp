<!--#include file="../../inc/header.asp"-->

<SCRIPT LANGUAGE="JavaScript">

function enableQuestion(IntRecIdent)
{
swal({
  title: "Enable Partner?",
  text: "Are you sure you wish to enable this Partner?",
  type: "warning",
  showCancelButton: true,
  confirmButtonColor: "#DD6B55",
  confirmButtonText: "Yes, enable it.",
  cancelButtonText: "No, cancel.",
  closeOnConfirm: false,
  closeOnCancel: false
},
function(isConfirm){
  if (isConfirm) {
	    window.location = "enablePartner.asp?i=" + IntRecIdent;
  } else {
	    window.location="main.asp";
  }
});
}
</SCRIPT>

<%
InternalRecordIdentifier = Request.QueryString("i")
Response.Write("<script language=javascript>enableQuestion(" & InternalRecordIdentifier & ");</script>")
%>
