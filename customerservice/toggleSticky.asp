<%
NoteNum = Request.Form("notnum")

SQL = "SELECT * FROM tblCustomerNotes WHERE InternalNoteNumber = " & NoteNum
Set cnnToggle = Server.CreateObject("ADODB.Connection")
cnnToggle.open (Session("ClientCnnString"))

Set rsToggle = Server.CreateObject("ADODB.Recordset")
rsToggle.CursorLocation = 3 
Set rsToggle = cnnToggle.Execute(SQL)

If Not rsToggle.Eof Then
	If rsToggle("Sticky") = 0 Then Sticky = 1 Else Sticky = 0
	SQL = "UPDATE tblCustomerNotes Set Sticky = " & Sticky & " WHERE InternalNoteNumber= " & NoteNum
	Set rsToggle = cnnToggle.Execute(SQL)
End If

set rsToggle = Nothing
cnnToggle.Close
Set cnnToggle = Nothing

%>
