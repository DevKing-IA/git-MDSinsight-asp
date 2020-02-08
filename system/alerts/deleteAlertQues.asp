<!--#include file="../../inc/header.asp"-->


<SCRIPT LANGUAGE="JavaScript">

function deletionQuestion(alertnum,activetab)
{
swal({
  title: "Delete Alert?",
  text: "Are you sure you wish to delete this alert?",
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
	    window.location = "deleteAlert.asp?a=" + alertnum + "&tab=" + activetab;
  } else {
	    window.location="main.asp#" + activetab;
  }
});
}
</SCRIPT>

<%
InternalAlertRecNumber = Request.QueryString("a")
ActiveTab = Request.QueryString("tab")
Response.Write("<script language=javascript>deletionQuestion(" & InternalAlertRecNumber & ",'" & ActiveTab & "');</script>")
%>
