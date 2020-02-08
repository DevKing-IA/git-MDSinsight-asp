<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->

<%


action = Request("action")

Select Case action

	Case "checkIfUserExists"
		checkIfUserExists()
		
End Select


Sub checkIfUserExists()
	
	result = "success"
	
	userEmail = Request.Form("userEmail")
	userPassword= Request.Form("userPassword")
	
	
	SQL = "SELECT * FROM tblUsers WHERE userEmail = '" & userEmail & "' AND userPassword = '" & userPassword & "'"	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		result = "failure"
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	Response.write(result)

End Sub
%>















