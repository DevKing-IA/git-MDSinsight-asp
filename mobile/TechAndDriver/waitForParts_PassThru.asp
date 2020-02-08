<!--#include file="../../inc/InSightFuncs.asp"-->
<%
SelectedMemoNumber = Request.Form("txtTicketNumber")

GeoLoc = Request.Cookies("gps") 

'Response.Write("XX" & GeoLoc)

Dummy = SetUserGeoLocation(Session("userNo"),GeoLoc)
%>
<form method="post" action="waitForParts.asp" name="frmwaitForParts" id="frmwaitForParts">
	<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
</form>
<script type="text/javascript">
  document.forms['frmwaitForParts'].submit();
</script>