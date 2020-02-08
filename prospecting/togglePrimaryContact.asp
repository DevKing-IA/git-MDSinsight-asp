<%
InternalRecordIdentifier = Request.Form("recnum")

SQL = "SELECT * FROM PR_ProspectContacts WHERE InternalRecordIdentifier = " & InternalRecordIdentifier 
Set cnnToggle = Server.CreateObject("ADODB.Connection")
cnnToggle.open (Session("ClientCnnString"))

Set rsToggle = Server.CreateObject("ADODB.Recordset")
rsToggle.CursorLocation = 3 
Set rsToggle = cnnToggle.Execute(SQL)

If Not rsToggle.Eof Then
	If rsToggle("PrimaryContact") = vbFalse Then PrimaryContact = vbTrue Else PrimaryContact = vbFalse
	 SQL = "UPDATE PR_ProspectContacts Set PrimaryContact = " & PrimaryContact & " WHERE InternalRecordIdentifier= " & InternalRecordIdentifier
	 Set rsToggle = cnnToggle.Execute(SQL)
	'If setting to true, need to set other records to false
	If PrimaryContact = vbTrue Then
		SQL = "UPDATE PR_ProspectContacts Set PrimaryContact = " & vbFalse & " WHERE InternalRecordIdentifier <> " & InternalRecordIdentifier
		Set rsToggle = cnnToggle.Execute(SQL)
	End If	 
End If

set rsToggle = Nothing
cnnToggle.Close
Set cnnToggle = Nothing

%>
