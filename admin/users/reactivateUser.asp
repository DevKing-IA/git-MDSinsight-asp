<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%
UserNo = Request.QueryString("un")
ActiveTab = Request.QueryString("tab")

If UserNo <> "" Then
	
	Set cnnReactivateUser = Server.CreateObject("ADODB.Connection")
	cnnReactivateUser.open (Session("ClientCnnString"))
	Set rsReactivateUser = Server.CreateObject("ADODB.Recordset")
	rsReactivateUser.CursorLocation = 3 

	Description =  "User " & GetUserDisplayNameByUserNo(UserNo) & " was moved from archived and reactivated by " & GetUserDisplayNameByUserNo(Session("UserNo"))
 
	CreateAuditLogEntry "User Reactivated","User Reactivated","Major",0,Description

	SQLReactivateUser = "UPDATE tblUsers Set userArchived = 0, userEnabled = 1 WHERE UserNo = " & UserNo
	
	Set rsReactivateUser = cnnReactivateUser.Execute(SQLReactivateUser)
	
	set rsReactivateUser = Nothing
	cnnReactivateUser.Close
	set cnnReactivateUser = Nothing
	
End If


Response.Redirect("main.asp#" & ActiveTab)

%>