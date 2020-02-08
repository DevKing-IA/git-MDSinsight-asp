<!--#include file="../../inc/header.asp"-->

<SCRIPT LANGUAGE="JavaScript">

function clearSKUQuestion(IntRecIdent, SKU, UM, CatID, PartIntRecIdent)
{
swal({
  title: "Clear Equivalent SKUs?",
  text: "Are you sure you wish to clear ALL equivalent SKUs values entered for " + SKU + ", with unit " + UM + " from the table?",
  type: "warning",
  showCancelButton: true,
  confirmButtonColor: "#DD6B55",
  confirmButtonText: "Yes, clear SKUs.",
  cancelButtonText: "No, cancel.",
  closeOnConfirm: false,
  closeOnCancel: false
},
function(isConfirm){
  if (isConfirm) {
	    window.location = "deletePartnerSKUandUM.asp?i=" + IntRecIdent + "&p=" + SKU + "&u=" + UM + "&c=" + CatID + "&x=" + PartIntRecIdent;
  } else {
	    window.location="editPartnerSKUCategoryToEdit.asp?i=" + PartIntRecIdent+ "&c=" + CatID;
  }
});
}
</SCRIPT>

<%
SKUInternalRecordIdentifier = Request.QueryString("i")
PartnerInternalRecordIdentifier = Request.QueryString("x")
SKUtoDelete = Request.QueryString("p")
UMToDelete = Request.QueryString("u")
CategoryID = Request.QueryString("c")

Response.Write("<script language=javascript>clearSKUQuestion('" & SKUInternalRecordIdentifier & "','" & SKUtoDelete & "','" & UMToDelete & "','" & CategoryID & "','" & PartnerInternalRecordIdentifier & "');</script>")
%>
