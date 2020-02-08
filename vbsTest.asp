<%
Response.Write("-Start<br>")
set objshell=server.createobject("WScript.Shell")
objshell.Run "c:\test.bat"
set objshell=nothing
Response.Write("-End<br>")
%>