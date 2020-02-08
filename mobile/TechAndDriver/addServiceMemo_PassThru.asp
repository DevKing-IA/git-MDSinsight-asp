<!--#include file="../../inc/InSightFuncs.asp"-->

<%

GeoLoc = Request.Cookies("gps") 

'Response.Write("XX" & GeoLoc)

Dummy = SetUserGeoLocation(Session("userNo"),GeoLoc)

Response.Redirect("addServiceMemo.asp")

%>