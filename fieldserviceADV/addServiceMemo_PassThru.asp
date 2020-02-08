<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%

'GeoLoc = Request.Cookies("gps") 

'Response.Write("XX" & GeoLoc)

'Dummy = SetUserGeoLocation(Session("userNo"),GeoLoc)

Response.Redirect("addServiceMemo.asp")

%>