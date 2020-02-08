<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

TeamName = Request.Form("txtTeamName")
TeamName = Replace(TeamName, "'", "''")

TeamUserNos = Request.Form("lstSelectedNewTeamUserIDs")



If TeamName <> "" AND TeamUserNos <> "" Then

	SQL = "INSERT INTO USER_Teams (TeamName,TeamUserNos)"
	SQL = SQL &  " VALUES (" 
	SQL = SQL & "'" & TeamName & "','" & TeamUserNos & "')"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	
	Set rs8 = Server.CreateObject("ADODB.Recordset")
	rs8.CursorLocation = 3 
	Response.Write(SQL)
	Set rs8 = cnn8.Execute(SQL)
	set rs8 = Nothing
	
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the team: " & TeamName
	CreateAuditLogEntry "User Team Added","User Team Added","Minor",0,Description

End If

Response.Redirect("main.asp")

%>















