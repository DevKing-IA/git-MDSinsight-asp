<!--#include file="../inc/header.asp"-->



<% If Session("MAILOFF") = 1 Then
	Session("MAILOFF") = 0
	StopStart = "started."
Elseif Session("MAILOFF") = 0 Then
	Session("MAILOFF") = 1
	StopStart = "stopped."
End If


'Now that we took care of the immediate need, Session variable, write it to the settings table
SQLtoggle = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_Global Set STOPALLEMAIL = " & Session("MAILOFF")
Set cnntoggle = Server.CreateObject("ADODB.Connection")
cnntoggle.open (Session("ClientCnnString"))
Set rstoggle = Server.CreateObject("ADODB.Recordset")
rstoggle.CursorLocation = 3 
Set rstoggle = cnntoggle.Execute(SQLtoggle)
set rstoggle = Nothing
cnntoggle.close
set cnntoggle = Nothing

CreateAuditLogEntry "Insight Mail Start-Stop", "Insight Mail Start-Stop", "Major", 1, "MDS Insight Outbound email was " & left(StopStart,Len(StopStart)-1)  & " by user: " & MUV_Read("DisplayName")

redirURL = "../main/default.asp"

%>

<script>
swal("MDS Insight email has been <%=StopStart%>");
window.location = "<%=redirURL%>";
</script>	

