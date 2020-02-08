<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

Activity = Request.Form("txtActivity")
ApptOrMeet = Request.Form("optApptOrMeet")

Appointment = 0
Meeting = 0
If ApptOrMeet ="Appointment" Then Appointment = 1
If ApptOrMeet ="Meeting" Then Meeting = 1

SQL = "INSERT INTO PR_Activities (Activity,CreateAppointment,CreateMeeting)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & Activity& "'," & Appointment & "," & Meeting & ")"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " activity: " & Activity
CreateAuditLogEntry GetTerm("Prospecting") & " activity added",GetTerm("Prospecting") & " activity added","Minor",0,Description

Response.Redirect("main.asp")

%>















