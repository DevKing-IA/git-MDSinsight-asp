<!--#include file="settings.asp"-->
<!--#include file="InSightFuncs.asp"-->
<!--#include file="InsightFuncs_Users.asp"-->
<%
action = Request("action")

Select Case action
	Case "CheckIfServerSessionAlive" 
		CheckIfServerSessionAlive()			
End Select

Sub CheckIfServerSessionAlive()
	
	If Session("UserNo") = "" Then ' In case it has timed out already
		Response.Write("SERVERTIMEOUT")
	Else
		Response.Write("ALIVE")
	End If

End Sub
%>

