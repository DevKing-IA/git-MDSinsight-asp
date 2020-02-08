<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM USER_Teams WHERE InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_TeamName = rs("TeamName")
	Orig_TeamUserNos = rs("TeamUserNos")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

TeamName = Request.Form("txtTeamName")
TeamName = Replace(TeamName, "'", "''")

TeamUserNos = Request.Form("lstSelectedNewTeamUserIDs")

SQL = "UPDATE USER_Teams SET "
SQL = SQL &  "TeamName = '" & TeamName & "', "
SQL = SQL &  "TeamUserNos = '" & TeamUserNos & "' "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_TeamName <> TeamName Then
	Description = Description & "The user team name changed from " & Orig_TeamName & " to " & TeamName 
End If

If Orig_TeamUserNos <> TeamUserNos Then
	Description = Description & "The user team, " & TeamName & ", changed members from users " & Orig_TeamUserNos & " to " & TeamUserNos 
End If

CreateAuditLogEntry "User Team Edited","User Team Edited","Minor",0,Description

Response.Redirect("main.asp")

%>















