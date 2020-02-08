<%
userno = Request.Form("userno")

Session("MultiUseVar")=""

'Format of multiuse var in this case
'{Y or N}emailaddress~~{Y or N}textnumber

SQL = "SELECT * FROM tblUsers WHERE UserNo = " & userno

Set cnnDispatch = Server.CreateObject("ADODB.Connection")
cnnDispatch.open (Session("ClientCnnString"))

Set rsDispatch = Server.CreateObject("ADODB.Recordset")
rsDispatch.CursorLocation = 3 
Set rsDispatch = cnnDispatch.Execute(SQL)

If Not rsDispatch.Eof Then
	If rsDispatch("userEmail") <> "" Then Session("MultiUseVar") = "Y" & rsDispatch("userEmail") Else Session("MultiUseVar") = "N"
	Session("MultiUseVar") = Session("MultiUseVar") & "~~"
	If rsDispatch("userCellNumber") <> "" Then Session("MultiUseVar") = Session("MultiUseVar") & "Y" & rsDispatch("userCellNumber") Else Session("MultiUseVar") = Session("MultiUseVar") & "N"
End If

set rsDispatch = Nothing
cnnDispatch.Close
Set cnnDispatch = Nothing

%>
