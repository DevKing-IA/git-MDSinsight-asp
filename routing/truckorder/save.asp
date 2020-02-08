<%
dim fs,tfile
set fs=Server.CreateObject("Scripting.FileSystemObject")
set tfile=fs.CreateTextFile(Server.MapPath(".")&"\"&Session("Userno")&".txt")
tfile.WriteLine(Request.Form("truckorder"))
tfile.close
set tfile=nothing
set fs=nothing
%>