<!--#include file="../../inc/InSightFuncs.asp"-->
<%
SelectedMemoNumber = Request.Form("txtTicketNumber")

GeoLoc = Request.Cookies("gps") 

'Response.Write("XX" & GeoLoc)

Dummy = SetUserGeoLocation(Session("userNo"),GeoLoc)
%>
<form method="post" action="swap.asp" name="frmswap" id="frmswap">
	<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
</form>
<script type="text/javascript">
  document.forms['frmswap'].submit();
</script>