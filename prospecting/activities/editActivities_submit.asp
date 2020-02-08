<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
ApptOrMeet = Request.Form("optApptOrMeet")
Appointment = 0
Meeting = 0
If ApptOrMeet ="Appointment" Then Appointment = 1
If ApptOrMeet ="Meeting" Then Meeting = 1


'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM PR_Activities where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Activity = rs("Activity")
	Orig_CreateAppointment = rs("CreateAppointment")
	Orig_CreateMeeting = rs("CreateMeeting")
End If
If Orig_CreateAppointment <> 1 Then Orig_CreateAppointment = 0
If Orig_CreateMeeting <> 1 Then Orig_CreateMeeting = 0
set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

Activity = Request.Form("txtActivity")

SQL = "UPDATE PR_Activities SET "
SQL = SQL &  "Activity = '" & Activity & "', CreateAppointment = " & Appointment & ", CreateMeeting = " & Meeting & "  "
SQL = SQL &  " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_Activity  <> Activity  Then
	Description = Description & GetTerm("Prospecting") & " activity changed from " & Orig_Activity & " to " & Activity
End If

If Appointment = 1 Then Appointment = "True" else Appointment = "False"
If Orig_CreateAppointment <> Appointment Then
	Description = Description & "  Create appointment changed from " & Orig_CreateAppointment & " to " & Appointment 
End If

If Meeting = 1 Then Meeting = "True" else Meeting = "False"
If Orig_CreateMeeting <> Meeting Then
	Description = Description & "  Create meeting changed from " & Orig_CreateMeeting & " to " & Meeting 
End If

CreateAuditLogEntry GetTerm("Prospecting") & " activity edited",GetTerm("Prospecting") & " activity edited","Minor",0,Description
Response.Redirect("main.asp")

%>















