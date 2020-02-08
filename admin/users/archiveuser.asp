<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->

<%
UserNo = Request.QueryString("un")
ActiveTab = Request.QueryString("tab")

If UserNo <> "" Then
	
	Set cnnArchiveUser = Server.CreateObject("ADODB.Connection")
	cnnArchiveUser.open (Session("ClientCnnString"))
	Set rsArchiveUser = Server.CreateObject("ADODB.Recordset")
	rsArchiveUser.CursorLocation = 3 
	
	Description = ""
	Description = Description & "The user " & GetUserDisplayNameByUserNo(UserNo) & " was changed to archived by "  & GetUserDisplayNameByUserNo(Session("UserNo"))
 
	CreateAuditLogEntry "User Archived","User Archived","Major",0,Description

	SQLArchiveUser = "UPDATE tblUsers Set userArchived = 1, userEnabled = 0  WHERE UserNo = " & UserNo
	Set rsArchiveUser = cnnArchiveUser.Execute(SQLArchiveUser)
	
	set rsArchiveUser = Nothing
	cnnArchiveUser.Close
	set cnnArchiveUser = Nothing
	
End If

Response.Redirect("main.asp#" & ActiveTab)
%>