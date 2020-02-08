<!--#include file="../../inc/header.asp"-->

<SCRIPT LANGUAGE="JavaScript">

function deletionQuestion(usernum,activetab)
{
swal({
  title: "Archive user?",
  text: "Are you sure you wish to move this user to the archive tab?",
  type: "warning",
  showCancelButton: true,
  confirmButtonColor: "#DD6B55",
  confirmButtonText: "Yes, archive user.",
  cancelButtonText: "No, cancel.",
  closeOnConfirm: false,
  closeOnCancel: false
},
function(isConfirm){
  if (isConfirm) {
	    window.location = "archiveuser.asp?un=" + usernum + "&tab=" + activetab;
  } else {
	    window.location="main.asp#" + activetab;
  }
});
}
</SCRIPT>

<%
usernum = Request.QueryString("un")
ActiveTab = Request.QueryString("tab")
Response.Write("<script language=javascript>deletionQuestion(" & usernum & ",'" & ActiveTab & "');</script>")
%>
