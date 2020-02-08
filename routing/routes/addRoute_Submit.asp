<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

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


SQL = "INSERT INTO RT_Routes (RecordSource, RouteID, RouteDescription, DefaultDriverUserNo, ThirdPartyCarrier, "
SQL = SQL &  " ShowOnDBoard, ShowInWebApp, ShowInPlanner, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday) "
SQL = SQL &  " VALUES (" 
SQL = SQL & "'INSIGHT','" & RouteID & "','" & RouteDescription & "'," & RouteDefaultDriverUserNo & "," & ThirdPartyCarrier & ", "
SQL = SQL & ShowOnDBoard & "," & ShowInWebApp & "," & ShowInPlanner & "," & RouteMonday & "," & RouteTuesday & ","
SQL = SQL & RouteWednesday & "," & RouteThursday & "," & RouteFriday & "," & RouteSaturday & "," & RouteSunday & ")"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Route") & " : " & RouteID & " with description, " & RouteDescription & "."

CreateAuditLogEntry GetTerm("Routing") & " Route Added",GetTerm("Routing") & " Route Added","Minor",0,Description

Response.Redirect("main.asp")

%>















