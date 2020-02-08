<!--#include file="../../inc/InSightFuncs.asp"-->
<%
SelectedMemoNumber = Request.Form("txtTicketNumber")

GeoLoc = Request.Cookies("gps") 

'Response.Write("XX" & GeoLoc)

Dummy = SetUserGeoLocation(Session("userNo"),GeoLoc)
%>
<form method="post" action="unableToWork.asp" name="frmunableToWork" id="frmunableToWork">
	<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
</form>
<script type="text/javascript">
  document.forms['frmunableToWork'].submit();
</script>