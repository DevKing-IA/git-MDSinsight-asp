<!--#include file="../../../../inc/header.asp"-->
<%
custID = Request.QueryString("custID")
status = Request.QueryString("status")
%>


<SCRIPT LANGUAGE="JavaScript">



function Success()
{
	swal({
	  title: "Sweet!",
	  text: "The updated price data has been sent to your Metroplex system.",
	  type: "success",
	  showCancelButton: false,
	  closeOnConfirm: false
	},
	function(){
	  window.location = "reports.asp";
	});
}


function Nope()
{
	swal({
	  title: "Warning!",
	  text: "An error was encountered while sending the updated price data to your Metroplex system, please try again and contact support if the problem persists.",
	  type: "warning",
	  showCancelButton: false,
	  confirmButtonColor: "#DD6B55",
	  confirmButtonText: "OK!",
	  closeOnConfirm: false
	},
	function(){
	  window.location = "quoteditemstool.asp?custID=" + <%=custID%>;
	});
}

</SCRIPT>

<%
If status = "ok" Then
	Response.Write("<script language=javascript>Success();</script>")
Else
	Response.Write("<script language=javascript>Nope();</script>")
End IF
%>
