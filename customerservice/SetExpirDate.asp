<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InsightFuncs.asp"-->

<%
NoteNum = Request.Form("notnum")
ExpirDate= Request.Form("ExpirDate")

SQL = "SELECT * FROM tblCustomerNotes WHERE InternalNoteNumber = " & NoteNum
Set cnnToggle = Server.CreateObject("ADODB.Connection")
cnnToggle.open (Session("ClientCnnString"))

Set rsToggle = Server.CreateObject("ADODB.Recordset")
rsToggle.CursorLocation = 3 
Set rsToggle = cnnToggle.Execute(SQL)

If Not rsToggle.Eof Then
	SQL = "UPDATE tblCustomerNotes Set ExpirationDate = '" & ExpirDate & "' WHERE InternalNoteNumber= " & NoteNum
	Set rsToggle = cnnToggle.Execute(SQL)
End If

set rsToggle = Nothing
cnnToggle.Close
Set cnnToggle = Nothing
	

%>
