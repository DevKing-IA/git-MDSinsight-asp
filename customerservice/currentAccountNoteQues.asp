<!--#include file="../inc/header.asp"-->

<SCRIPT LANGUAGE="JavaScript">

function deletionQuestion(notenum)
{
swal({
  title: "Move note?",
  text: "Are you sure you wish to move this note to the current tab?",
  type: "warning",
  showCancelButton: true,
  confirmButtonColor: "#DD6B55",
  confirmButtonText: "Yes, I'd like to move it, move it.",
  cancelButtonText: "No, cancel.",
  closeOnConfirm: false,
  closeOnCancel: false
},
function(isConfirm){
  if (isConfirm) {
	    window.location = "currentMoveaccountnote.asp?nt=" + notenum;
  } else {
	    window.location="main.asp#Archived";
  }
});
}
</SCRIPT>

<%
InternalNoteNumber = Request.QueryString("nt")
Response.Write("<script language=javascript>deletionQuestion(" & InternalNoteNumber & ");</script>")
%>
