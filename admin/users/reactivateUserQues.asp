<!--#include file="../../inc/header.asp"-->

<SCRIPT LANGUAGE="JavaScript">

function deletionQuestion(usernum,activetab)
{
swal({
  title: "Reactivate User?",
  text: "Are you sure you wish to reactivate this user?",
  type: "warning",
  showCancelButton: true,
  confirmButtonColor: "#DD6B55",
  confirmButtonText: "Yes, reactivate.",
  cancelButtonText: "No, cancel.",
  closeOnConfirm: false,
  closeOnCancel: false
},
function(isConfirm){
  if (isConfirm) {
	    window.location = "reactivateUser.asp?un=" + usernum + "&tab=" + activetab;
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
