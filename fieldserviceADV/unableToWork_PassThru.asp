<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%
SelectedMemoNumber = Request.Form("txtTicketNumber")

If SelectedMemoNumber = "" Then
	SelectedMemoNumber = Request.QueryString("t")
End If


'GeoLoc = Request.Cookies("gps") 

'Response.Write("XX" & GeoLoc)

'Dummy = SetUserGeoLocation(Session("userNo"),GeoLoc)
%>
<form method="post" action="unableToWork.asp" name="frmunableToWork" id="frmunableToWork">
	<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
</form>
<script type="text/javascript">
  document.forms['frmunableToWork'].submit();
</script>