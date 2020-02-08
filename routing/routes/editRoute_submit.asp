<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
RouteID = Request.Form("txtRouteID")
RouteID = Replace(RouteID,"'","''")
RouteDescription = Request.Form("txtRouteDescription")
RouteDescription = Replace(RouteDescription,"'","''")
RouteDefaultDriverUserNo = Request.Form("selDefaultDriverUserNo")

ThirdPartyCarrier = Request.Form("chkThirdPartyCarrier")
If ThirdPartyCarrier = "on" then ThirdPartyCarrier = 1 Else ThirdPartyCarrier = 0

ShowOnDBoard = Request.Form("chkShowOnDBoard")
If ShowOnDBoard = "on" then ShowOnDBoard = 1 Else ShowOnDBoard = 0

ShowInPlanner = Request.Form("chkShowInPlanner")
If ShowInPlanner = "on" then ShowInPlanner = 1 Else ShowInPlanner= 0

ShowInWebApp = Request.Form("chkShowInWebApp")
If ShowInWebApp = "on" then ShowInWebApp = 1 Else ShowInWebApp = 0

RouteMonday = Request.Form("chkRouteMonday")
If RouteMonday = "on" then RouteMonday = 1 Else RouteMonday = 0

RouteTuesday = Request.Form("chkRouteTuesday")
If RouteTuesday = "on" then RouteTuesday = 1 Else RouteTuesday = 0

RouteWednesday = Request.Form("chkRouteWednesday")
If RouteWednesday = "on" then RouteWednesday = 1 Else RouteWednesday = 0

RouteThursday = Request.Form("chkRouteThursday")
If RouteThursday = "on" then RouteThursday = 1 Else RouteThursday = 0

RouteFriday = Request.Form("chkRouteFriday")
If RouteFriday = "on" then RouteFriday = 1 Else RouteFriday = 0

RouteSaturday = Request.Form("chkRouteSaturday")
If RouteSaturday = "on" then RouteSaturday = 1 Else RouteSaturday = 0

RouteSunday = Request.Form("chkRouteSunday")
If RouteSunday = "on" then RouteSunday = 1 Else RouteSunday = 0


'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM RT_Routes where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then	
	Orig_RouteID = rs("RouteID")
	Orig_RouteDescription = rs("RouteDescription")
	Orig_DefaultDriverUserNo = rs("DefaultDriverUserNo")
	Orig_ThirdPartyCarrier = rs("ThirdPartyCarrier")
	Orig_ShowOnDBoard = rs("ShowOnDBoard")
	Orig_ShowInWebApp = rs("ShowInWebApp")
	Orig_ShowInPlanner = rs("ShowInPlanner")
	Orig_RouteMonday = rs("Monday")
	Orig_RouteTuesday = rs("Tuesday")
	Orig_RouteWednesday = rs("Wednesday")
	Orig_RouteThursday = rs("Thursday")
	Orig_RouteFriday = rs("Friday")
	Orig_RouteSaturday = rs("Saturday")
	Orig_RouteSunday = rs("Sunday")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

SQL = "UPDATE RT_Routes SET "
SQL = SQL &  "RouteID = '" & RouteID & "', RouteDescription = '" & RouteDescription & "', "
SQL = SQL &  "DefaultDriverUserNo = " & RouteDefaultDriverUserNo & ", ThirdPartyCarrier = " & ThirdPartyCarrier & ", "
SQL = SQL &  "ShowOnDBoard = " & ShowOnDBoard & ", ShowInPlanner = " & ShowInPlanner & ", "
SQL = SQL &  "ShowInWebApp = " & ShowInWebApp & ", Monday = " & RouteMonday & ", Tuesday = " & RouteTuesday & ", Wednesday = " & RouteWednesday & ", "
SQL = SQL &  "Thursday = " & RouteThursday & ", Friday = " & RouteFriday & ", Saturday = " & RouteSaturday & ", Sunday = " & RouteSunday
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""


If Orig_RouteID <> RouteID Then
	Description = GetTerm("Routing") & " Route ID changed from " & Orig_RouteID & " to " & RouteID
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_RouteDescription <> RouteDescription Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") Description changed from " & Orig_RouteDescription & " to " & RouteDescription
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If


Orig_DefaultDriverName = GetUserFirstAndLastNameByUserNo(Orig_DefaultDriverUserNo)
DefaultDriverName = GetUserFirstAndLastNameByUserNo(DefaultDriverUserNo)

If Orig_DefaultDriverUserNo <> DefaultDriverUserNo Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") Default Driver changed from " & Orig_DefaultDriverName & " to " & DefaultDriverName
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_ThirdPartyCarrier = True Then Orig_ThirdPartyCarrier = "True" else Orig_ThirdPartyCarrier = "False"
If Orig_ShowOnDBoard = True Then Orig_ShowOnDBoard = "True" else Orig_ShowOnDBoard = "False"
If Orig_ShowInWebApp = True Then Orig_ShowInWebApp = "True" else Orig_ShowInWebApp = "False"
If Orig_ShowInPlanner = True Then Orig_ShowInPlanner = "True" else Orig_ShowInPlanner = "False"
If Orig_RouteMonday = True Then Orig_RouteMonday = "True" else Orig_RouteMonday = "False"
If Orig_RouteTuesday = True Then Orig_RouteTuesday = "True" else Orig_RouteTuesday = "False"
If Orig_RouteWednesday = True Then Orig_RouteWednesday = "True" else Orig_RouteWednesday = "False"
If Orig_RouteThursday = True Then Orig_RouteThursday = "True" else Orig_RouteThursday = "False"
If Orig_RouteFriday = True Then Orig_RouteFriday = "True" else Orig_RouteFriday = "False"
If Orig_RouteSaturday = True Then Orig_RouteSaturday = "True" else Orig_RouteSaturday = "False"
If Orig_RouteSunday = True Then Orig_RouteSunday = "True" else Orig_RouteSunday = "False"

If ThirdPartyCarrier = 1 then ThirdPartyCarrier = "True" Else ThirdPartyCarrier = "False"
If ShowOnDBoard = 1 then ShowOnDBoard = "True" Else ShowOnDBoard = "False"
If ShowInPlanner = 1 then ShowInPlanner = "True" Else ShowInPlanner = "False"
If ShowInWebApp = 1 then ShowInWebApp = "True" Else ShowInWebApp = "False"
If RouteMonday = 1 then RouteMonday = "True" Else RouteMonday = "False"
If RouteTuesday = 1 then RouteTuesday = "True" Else RouteTuesday = "False"
If RouteWednesday = 1 then RouteWednesday = "True" Else RouteWednesday = "False"
If RouteThursday = 1 then RouteThursday = "True" Else RouteThursday = "False"
If RouteFriday = 1 then RouteFriday = "True" Else RouteFriday = "False"
If RouteSaturday = 1 then RouteSaturday = "True" Else RouteSaturday = "False"
If RouteSunday = 1 then RouteSunday = "True" Else RouteSunday = "False"


If Orig_ThirdPartyCarrier <> ThirdPartyCarrier Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") Third Party Carrier changed from " & Orig_ThirdPartyCarrier & " to " & ThirdPartyCarrier 
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_ShowOnDBoard <> ShowOnDBoard Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") Show on Delivery Board changed from " & Orig_ShowOnDBoard & " to " & ShowOnDBoard 
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_ShowInWebApp <> ShowInWebApp Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") Show In Web App changed from " & Orig_ShowInWebApp & " to " & ShowInWebApp 
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_ShowInPlanner <> ShowInPlanner Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") Show on Delivery Board Planner changed from " & Orig_ShowInPlanner & " to " & ShowInPlanner 
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_RouteMonday <> RouteMonday Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") route runs on Monday changed from " & Orig_RouteMonday & " to " & RouteMonday 
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_RouteTuesday <> RouteTuesday Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") route runs on Tuesday changed from " & Orig_RouteTuesday & " to " & RouteTuesday 
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_RouteWednesday <> RouteWednesday Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") route runs on Wednesday changed from " & Orig_RouteWednesday & " to " & RouteWednesday 
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_RouteThursday <> RouteThursday Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") route runs on Thursday changed from " & Orig_RouteThursday & " to " & RouteThursday 
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_RouteFriday <> RouteFriday Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") route runs on Friday changed from " & Orig_RouteFriday & " to " & RouteFriday 
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_RouteSaturday <> RouteSaturday Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") route runs on Saturday changed from " & Orig_RouteSaturday & " to " & RouteSaturday 
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

If Orig_RouteSunday <> RouteSunday Then
	Description = GetTerm("Routing") & " (route " & RouteID & ") route runs on Sunday changed from " & Orig_RouteSunday & " to " & RouteSunday 
	CreateAuditLogEntry GetTerm("Routing") & " Route Edited",GetTerm("Routing") & " Route Edited","Minor",0,Description
End If

Response.Redirect("main.asp")

%>